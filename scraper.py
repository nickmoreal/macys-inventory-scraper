import os
import re
import time
import json
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
from curl_cffi import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload


# ============================================================
# CONFIG
# ============================================================

INPUT_FILE = os.getenv("INPUT_FILE", "Macys_URLs.xlsx")
CHUNK_SIZE = int(os.getenv("CHUNK_SIZE", "25"))
DEBUG_MODE = os.getenv("DEBUG_MODE", "true").lower() == "true"

IST = ZoneInfo("Asia/Kolkata")
RUN_TIME = datetime.now(IST)
RUN_ID = RUN_TIME.strftime("%Y%m%d_%H%M%S")
RUN_LABEL = RUN_TIME.strftime("%Y-%m-%d %H:%M:%S IST")

BACKUP_FILE = f"Macys_Inventory_{RUN_ID}.csv"

# Google configuration is supplied through GitHub Actions Secrets.
GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "")
GOOGLE_SHEET_NAME = os.getenv("GOOGLE_SHEET_NAME", "Inventory")
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID", "")

SHEET_HEADERS = [
    "RunTimeIST",
    "SKU",
    "UPC",
    "URL",
    "RegularPrice",
    "DiscountPrice",
    "OfferPrice",
    "Stock",
]

if DEBUG_MODE:
    SHEET_HEADERS += [
        "PriceSource",
        "MatchedOfferSKU",
        "MatchedColor",
        "MatchedSize",
    ]


# ============================================================
# GOOGLE AUTH
# ============================================================

def get_google_services():
    if not GOOGLE_SERVICE_ACCOUNT_JSON:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON secret is missing.")
    if not GOOGLE_SHEET_ID:
        raise RuntimeError("GOOGLE_SHEET_ID secret is missing.")
    if not GOOGLE_DRIVE_FOLDER_ID:
        raise RuntimeError("GOOGLE_DRIVE_FOLDER_ID secret is missing.")

    info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    credentials = service_account.Credentials.from_service_account_info(
        info,
        scopes=scopes,
    )

    sheets = build("sheets", "v4", credentials=credentials)
    drive = build("drive", "v3", credentials=credentials)

    return sheets, drive


def ensure_sheet_header(sheets):
    range_name = f"{GOOGLE_SHEET_NAME}!A1:{chr(64 + len(SHEET_HEADERS))}1"

    result = sheets.spreadsheets().values().get(
        spreadsheetId=GOOGLE_SHEET_ID,
        range=range_name,
    ).execute()

    values = result.get("values", [])

    if not values or values[0] != SHEET_HEADERS:
        sheets.spreadsheets().values().update(
            spreadsheetId=GOOGLE_SHEET_ID,
            range=f"{GOOGLE_SHEET_NAME}!A1",
            valueInputOption="RAW",
            body={"values": [SHEET_HEADERS]},
        ).execute()

        print("✅ Google Sheet header ready.")


def append_to_google_sheet(rows):
    if not rows:
        return

    sheets, _ = get_google_services()

    values = []
    for row in rows:
        values.append([row.get(header, "") for header in SHEET_HEADERS])

    sheets.spreadsheets().values().append(
        spreadsheetId=GOOGLE_SHEET_ID,
        range=f"{GOOGLE_SHEET_NAME}!A1",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": values},
    ).execute()

    print(f"☁️ Google Sheet saved: {len(rows)} rows")


# ============================================================
# GOOGLE DRIVE BACKUP
# ============================================================

def create_drive_backup():
    """
    Creates one CSV file for this run in Drive.
    The same file is updated after every successful chunk.
    """
    _, drive = get_google_services()

    metadata = {
        "name": BACKUP_FILE,
        "parents": [GOOGLE_DRIVE_FOLDER_ID],
        "description": f"Macy's inventory backup - {RUN_LABEL}",
        "mimeType": "text/csv",
    }

    empty_csv = BytesIO(b"")
    media = MediaIoBaseUpload(
        empty_csv,
        mimetype="text/csv",
        resumable=True,
    )

    created = drive.files().create(
        body=metadata,
        media_body=media,
        fields="id,name,webViewLink",
    ).execute()

    print(f"💾 Drive backup created: {created.get('name')}")
    print(f"🔗 Drive file: {created.get('webViewLink', '')}")

    return created["id"]


def update_drive_backup(drive_file_id, csv_path):
    _, drive = get_google_services()

    media = MediaFileUpload(
        csv_path,
        mimetype="text/csv",
        resumable=True,
    )

    drive.files().update(
        fileId=drive_file_id,
        media_body=media,
    ).execute()

    print(f"💾 Drive backup updated: {csv_path}")


# ============================================================
# TEXT / UPC HELPERS
# ============================================================

def normalize_upc(upc):
    upc = str(upc or "").strip()
    upc = re.sub(r"\D", "", upc)

    if len(upc) < 12:
        return upc.zfill(12)
    elif len(upc) > 12:
        return upc[-12:]

    return upc


def digits_only(value):
    return re.sub(r"\D", "", str(value or "").strip())


def norm_text(value):
    return re.sub(r"\s+", " ", str(value or "").strip()).lower()


def extract_product_id(url):
    match = re.search(r"ID=(\d+)", str(url))
    return match.group(1) if match else None


# ============================================================
# PRICE LOGIC
# ============================================================

def pick_prices_from_tiered(tiered):
    regular_price = "0"
    discount_price = "0"

    for t in tiered or []:
        label = norm_text(t.get("label", ""))
        values = t.get("values", [])

        if not values:
            continue

        first = values[0]
        value = str(first.get("value", "0"))
        value_type = norm_text(first.get("type", ""))

        if (
            value_type == "regular"
            or "orig" in label
            or "reg" in label
            or "regular" in label
        ):
            regular_price = value

        if (
            value_type == "discount"
            or "sale" in label
            or "now" in label
            or "discount" in label
        ):
            discount_price = value

    return regular_price, discount_price


def get_fallback_prices(product_data):
    regular_price = "0"
    discount_price = "0"
    offer_price = "0"

    try:
        colors = (
            product_data
            .get("traits", {})
            .get("colors", {})
            .get("colorMap", {})
        )

        for c in colors.values():
            pricing = c.get("pricing", {}).get("price", {})
            tiered = pricing.get("tieredPrice", [])

            r, d = pick_prices_from_tiered(tiered)

            if r != "0" and regular_price == "0":
                regular_price = r

            if d != "0" and discount_price == "0":
                discount_price = d

            final_price = pricing.get("finalPrice", {})

            if isinstance(final_price, dict):
                vals = final_price.get("values", [])

                if vals and offer_price == "0":
                    offer_price = str(vals[0].get("value", "0"))

    except Exception:
        pass

    try:
        price = product_data.get("price", {})

        if regular_price == "0" and "regularPrice" in price:
            regular_price = str(price["regularPrice"])

        if discount_price == "0" and "salePrice" in price:
            discount_price = str(price["salePrice"])

        final_price = price.get("finalPrice", {})

        if isinstance(final_price, dict):
            vals = final_price.get("values", [])

            if vals and offer_price == "0":
                offer_price = str(vals[0].get("value", "0"))

    except Exception:
        pass

    return regular_price, discount_price, offer_price


def get_product_mkt_data(full_json, product_data):
    if isinstance(full_json, dict):
        pmd = full_json.get("productMktData")

        if isinstance(pmd, dict):
            return pmd

        products = full_json.get("product", [])

        if products and isinstance(products, list):
            maybe = products[0].get("productMktData")

            if isinstance(maybe, dict):
                return maybe

    if isinstance(product_data, dict):
        pmd = product_data.get("productMktData")

        if isinstance(pmd, dict):
            return pmd

    return {}


def get_offer_price_from_product_mkt(product_mkt_data, row_upc, row_sku=""):
    offers = product_mkt_data.get("offers", [])

    row_upc_norm = normalize_upc(row_upc)
    row_sku_digits = digits_only(row_sku)

    # 1) Exact SKU digits match
    if row_sku_digits:
        for offer in offers:
            offer_sku = str(offer.get("SKU", "")).strip()
            offer_digits = digits_only(offer_sku)
            offer_price = offer.get("price", None)

            if not offer_price:
                continue

            if offer_digits == row_sku_digits:
                return str(offer_price), offer, "offer_exact_sku"

    # 2) Exact UPC match
    for offer in offers:
        offer_sku = str(offer.get("SKU", "")).strip()
        offer_digits = digits_only(offer_sku)
        offer_upc_norm = normalize_upc(offer_digits)
        offer_price = offer.get("price", None)

        if not offer_price:
            continue

        if offer_upc_norm == row_upc_norm:
            return str(offer_price), offer, "offer_exact_upc"

    # 3) Last 10 digits fallback
    for offer in offers:
        offer_sku = str(offer.get("SKU", "")).strip()
        offer_digits = digits_only(offer_sku)
        offer_price = offer.get("price", None)

        if not offer_price:
            continue

        if (
            offer_digits
            and row_upc_norm
            and offer_digits[-10:] == row_upc_norm[-10:]
        ):
            return str(offer_price), offer, "offer_last10_upc"

    return None, None, "no_offer_match"


def get_offer_stock_from_product_mkt(matched_offer):
    if not matched_offer:
        return None

    availability = str(
        matched_offer.get("availability", "")
    ).lower()

    if "instock" in availability:
        return "2"

    if "outofstock" in availability:
        return "0"

    return None


def resolve_row_price(product_mkt_data, product_data, row_upc, row_sku=""):
    regular_price, discount_price, fallback_offer_price = (
        get_fallback_prices(product_data)
    )

    offer_price, matched_offer, match_type = (
        get_offer_price_from_product_mkt(
            product_mkt_data,
            row_upc,
            row_sku,
        )
    )

    if offer_price is not None:
        return {
            "RegularPrice": regular_price,
            "DiscountPrice": discount_price,
            "OfferPrice": offer_price,
            "MatchedOffer": matched_offer,
            "PriceSource": match_type,
        }

    if fallback_offer_price != "0":
        return {
            "RegularPrice": regular_price,
            "DiscountPrice": discount_price,
            "OfferPrice": fallback_offer_price,
            "MatchedOffer": None,
            "PriceSource": "fallback_finalPrice",
        }

    if discount_price != "0":
        return {
            "RegularPrice": regular_price,
            "DiscountPrice": discount_price,
            "OfferPrice": discount_price,
            "MatchedOffer": None,
            "PriceSource": "fallback_discount",
        }

    return {
        "RegularPrice": regular_price,
        "DiscountPrice": discount_price,
        "OfferPrice": regular_price,
        "MatchedOffer": None,
        "PriceSource": "fallback_regular",
    }


# ============================================================
# LOCAL CSV CHECKPOINT
# ============================================================

def append_local_checkpoint(rows, csv_path):
    if not rows:
        return

    df = pd.DataFrame(rows)

    # Always keep a stable column order.
    columns = SHEET_HEADERS
    df = df.reindex(columns=columns)

    write_header = not os.path.exists(csv_path)

    df.to_csv(
        csv_path,
        mode="a",
        header=write_header,
        index=False,
        encoding="utf-8-sig",
    )


# ============================================================
# SCRAPER
# ============================================================

def scrape_inventory(input_file, on_chunk_saved):
    df = pd.read_excel(input_file)

    required_columns = {"URL", "UPC"}

    missing = required_columns - set(df.columns)

    if missing:
        raise ValueError(
            f"Excel is missing required columns: {sorted(missing)}"
        )

    df = df.sort_values(by="URL")
    grouped = df.groupby("URL")

    unique_urls = list(grouped.groups.keys())
    total_urls = len(unique_urls)

    print(f"🌐 Total URLs: {total_urls}")

    results = []

    session = requests.Session(
        impersonate="chrome110"
    )

    headers = {
        "accept": (
            "application/json,text/html,"
            "application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        ),
        "referer": "https://www.macys.com/",
        "user-agent": "Mozilla/5.0",
    }

    for i, url in enumerate(unique_urls, 1):

        if i % 100 == 0:
            session = requests.Session(
                impersonate="chrome110"
            )
            print("🔄 Session refreshed")

        percent = round((i / total_urls) * 100, 2)
        remaining = total_urls - i

        print(
            f"\n🌐 {i}/{total_urls} "
            f"({percent}%) | Remaining: {remaining}"
        )

        product_id = extract_product_id(url)

        if not product_id:
            print("⚠ Product ID not found")
            continue

        api_url = (
            f"https://www.macys.com/"
            f"xapi/digital/v1/product/{product_id}"
        )

        success = False
        res = None

        for attempt in range(3):
            try:
                res = session.get(
                    api_url,
                    headers=headers,
                    timeout=30,
                )

                if res.status_code == 200:
                    success = True
                    break

                print(
                    f"⚠ API status {res.status_code}, "
                    f"attempt {attempt + 1}/3"
                )

            except Exception as exc:
                print(
                    f"⚠ API request error: {exc}, "
                    f"attempt {attempt + 1}/3"
                )

            time.sleep(1)

        if not success or res is None:
            print("❌ API failed")
            continue

        try:
            full_json = res.json()
            product_data = full_json["product"][0]

        except Exception as exc:
            print(f"❌ JSON parse failed: {exc}")
            continue

        product_mkt_data = get_product_mkt_data(
            full_json,
            product_data,
        )

        subset = grouped.get_group(url)

        upcs = (
            product_data
            .get("relationships", {})
            .get("upcs", {})
        )

        time.sleep(0.35)

        for _, row in subset.iterrows():

            upc = normalize_upc(row["UPC"])

            sku = (
                str(row["SKU"]).strip()
                if "SKU" in row.index
                else ""
            )

            final_stock = "0"

            # Existing stock logic
            for info in upcs.values():

                api_upc = normalize_upc(
                    info
                    .get("identifier", {})
                    .get("upcNumber", "")
                )

                if api_upc and api_upc == upc:

                    final_stock = (
                        "2"
                        if info
                        .get("availability", {})
                        .get("available")
                        else "0"
                    )

                    break

            price_info = resolve_row_price(
                product_mkt_data,
                product_data,
                upc,
                sku,
            )

            matched_offer = price_info["MatchedOffer"]

            # Offer-based stock override
            offer_stock = get_offer_stock_from_product_mkt(
                matched_offer
            )

            if offer_stock is not None:
                final_stock = offer_stock

            result_row = {
                "RunTimeIST": RUN_LABEL,
                "SKU": sku,
                "UPC": upc,
                "URL": url,
                "RegularPrice": price_info["RegularPrice"],
                "DiscountPrice": price_info["DiscountPrice"],
                "OfferPrice": price_info["OfferPrice"],
                "Stock": final_stock,
            }

            if DEBUG_MODE:
                result_row.update({
                    "PriceSource": price_info["PriceSource"],
                    "MatchedOfferSKU": (
                        matched_offer.get("SKU", "")
                        if matched_offer
                        else ""
                    ),
                    "MatchedColor": (
                        matched_offer
                        .get("itemOffered", {})
                        .get("color", "")
                        if matched_offer
                        else ""
                    ),
                    "MatchedSize": (
                        matched_offer
                        .get("itemOffered", {})
                        .get("size", "")
                        if matched_offer
                        else ""
                    ),
                })

            results.append(result_row)

        if len(results) >= CHUNK_SIZE:
            on_chunk_saved(results)
            results.clear()

    return results


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 70)
    print("🚀 Macy's Inventory Automation")
    print(f"🕒 Run: {RUN_LABEL}")
    print(f"📁 Input: {INPUT_FILE}")
    print(f"💾 Backup: {BACKUP_FILE}")
    print("=" * 70)

    sheets, _ = get_google_services()
    ensure_sheet_header(sheets)

    drive_file_id = create_drive_backup()

    def save_chunk(rows):
        print(f"\n📦 Saving chunk: {len(rows)} rows")

        # 1. Local checkpoint FIRST
        append_local_checkpoint(
            rows,
            BACKUP_FILE,
        )

        # 2. Google Sheet
        append_to_google_sheet(rows)

        # 3. Google Drive backup
        update_drive_backup(
            drive_file_id,
            BACKUP_FILE,
        )

        print("✅ Chunk safely stored in Sheet + Drive\n")

    try:
        results = scrape_inventory(
            INPUT_FILE,
            save_chunk,
        )

        if results:
            save_chunk(results)

        # Final Drive sync even if no final pending chunk.
        if os.path.exists(BACKUP_FILE):
            update_drive_backup(
                drive_file_id,
                BACKUP_FILE,
            )

        print("=" * 70)
        print("✅ DONE")
        print(f"📄 Backup: {BACKUP_FILE}")
        print("=" * 70)

    except Exception as exc:
        print("=" * 70)
        print(f"❌ SCRAPER ERROR: {exc}")
        print("=" * 70)

        # The already-saved chunks have already been pushed to
        # Google Sheet and Drive. Preserve the local checkpoint too.
        if os.path.exists(BACKUP_FILE):
            try:
                update_drive_backup(
                    drive_file_id,
                    BACKUP_FILE,
                )
                print("💾 Partial backup uploaded successfully.")
            except Exception as backup_error:
                print(
                    f"⚠ Could not update Drive backup: "
                    f"{backup_error}"
                )

        raise


if __name__ == "__main__":
    main()
