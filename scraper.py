import os
import re
import time
import json
import logging
from datetime import datetime
import pandas as pd
from curl_cffi import requests

# Google APIs
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Logging configuration
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# File names & Configuration
INPUT_FILE = "Macys_URLS.xlsx"
OUTPUT_FILE = "Macys_Inventory_Output.csv"
CHUNK_SIZE = 25
DEBUG_MODE = True

# Google API Scopes
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

PROXY_URL = os.environ.get("PROXY_URL", "").strip() or None


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
    match = re.search(r'ID=(\d+)', str(url), re.IGNORECASE)
    if match:
        return match.group(1)
    match_path = re.search(r'/(\d{5,10})', str(url))
    if match_path:
        return match_path.group(1)
    return None


def save_chunk(data):
    df = pd.DataFrame(data)
    if not os.path.exists(OUTPUT_FILE):
        df.to_csv(OUTPUT_FILE, index=False)
    else:
        df.to_csv(OUTPUT_FILE, mode='a', index=False, header=False)


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

        if value_type == "regular" or any(k in label for k in ["orig", "reg", "regular"]):
            regular_price = value

        if value_type == "discount" or any(k in label for k in ["sale", "now", "discount"]):
            discount_price = value

    return regular_price, discount_price


def get_fallback_prices(product_data):
    regular_price = "0"
    discount_price = "0"
    offer_price = "0"

    try:
        colors = product_data.get("traits", {}).get("colors", {}).get("colorMap", {})
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

    # 2) Exact UPC(12) match with offer SKU digits
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

        if offer_digits and row_upc_norm and offer_digits[-10:] == row_upc_norm[-10:]:
            return str(offer_price), offer, "offer_last10_upc"

    return None, None, "no_offer_match"


def get_offer_stock_from_product_mkt(matched_offer):
    if not matched_offer:
        return None

    availability = str(matched_offer.get("availability", "")).lower()
    if "instock" in availability:
        return "2"
    if "outofstock" in availability:
        return "0"
    return None


def resolve_row_price(product_mkt_data, product_data, row_upc, row_sku=""):
    regular_price, discount_price, fallback_offer_price = get_fallback_prices(product_data)

    offer_price, matched_offer, match_type = get_offer_price_from_product_mkt(
        product_mkt_data, row_upc, row_sku
    )

    if offer_price is not None:
        return {
            "RegularPrice": regular_price,
            "DiscountPrice": discount_price,
            "OfferPrice": offer_price,
            "MatchedOffer": matched_offer,
            "PriceSource": match_type
        }

    if fallback_offer_price != "0":
        return {
            "RegularPrice": regular_price,
            "DiscountPrice": discount_price,
            "OfferPrice": fallback_offer_price,
            "MatchedOffer": None,
            "PriceSource": "fallback_finalPrice"
        }

    if discount_price != "0":
        return {
            "RegularPrice": regular_price,
            "DiscountPrice": discount_price,
            "OfferPrice": discount_price,
            "MatchedOffer": None,
            "PriceSource": "fallback_discount"
        }

    return {
        "RegularPrice": regular_price,
        "DiscountPrice": discount_price,
        "OfferPrice": regular_price,
        "MatchedOffer": None,
        "PriceSource": "fallback_regular"
    }


def build_session():
    """Builds a curl_cffi session configured to bypass Akamai bot detection."""
    proxies = {"http": PROXY_URL, "https": PROXY_URL} if PROXY_URL else None
    session = requests.Session(impersonate="chrome124", proxies=proxies)
    return session


def get_api_headers():
    return {
        "accept": "application/json, text/plain, */*",
        "accept-language": "en-US,en;q=0.9",
        "x-macys-webservice-client-id": "macys-mcom",
        "x-macys-country-code": "US",
        "x-macys-currency-code": "USD",
        "referer": "https://www.macys.com/",
        "origin": "https://www.macys.com",
        "sec-ch-ua": '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"',
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-origin",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    }


def scrape_inventory(input_file):
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file '{input_file}' not found.")

    df = pd.read_excel(input_file)
    df = df.sort_values(by="URL")
    grouped = df.groupby("URL")

    unique_urls = list(grouped.groups.keys())
    total_urls = len(unique_urls)
    logging.info(f"🌐 Total Unique URLs: {total_urls}")

    results = []
    session = build_session()
    headers = get_api_headers()
    current_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for i, url in enumerate(unique_urls, 1):
        if i % 50 == 0:
            session = build_session()
            logging.info("🔄 Refreshed HTTP Session")

        percent = round((i / total_urls) * 100, 2)
        product_id = extract_product_id(url)
        logging.info(f"[{i}/{total_urls} - {percent}%] ID: {product_id} | URL: {url}")

        if not product_id:
            logging.warning(f"⚠ Could not extract Product ID from: {url}")
            continue

        api_url = f"https://www.macys.com/xapi/digital/v1/product/{product_id}"

        success = False
        res = None

        for attempt in range(1, 4):
            try:
                res = session.get(api_url, headers=headers, timeout=25)
                if res.status_code == 200:
                    success = True
                    break
                elif res.status_code == 404:
                    logging.warning(f"⚠ Product ID {product_id} is 404 Not Found")
                    break
                elif res.status_code == 403:
                    logging.warning(f"⚠ Status 403 on attempt {attempt}/3. Refreshing session...")
                    session = build_session()
            except Exception as e:
                logging.warning(f"  Attempt {attempt}/3 failed: {e}")
            time.sleep(1.5)

        if not success or res is None or res.status_code != 200:
            logging.error(f"❌ Failed to fetch product data for {product_id}")
            continue

        try:
            full_json = res.json()
            product_list = full_json.get("product", [])
            if not product_list:
                logging.warning(f"⚠ Empty product list for ID {product_id}")
                continue
            product_data = product_list[0]
        except Exception as e:
            logging.error(f"❌ JSON parse error for {product_id}: {e}")
            continue

        product_mkt_data = get_product_mkt_data(full_json, product_data)
        subset = grouped.get_group(url)
        upcs = product_data.get("relationships", {}).get("upcs", {})

        time.sleep(0.3)

        for _, row in subset.iterrows():
            upc = normalize_upc(row["UPC"])
            sku = str(row["SKU"]).strip() if "SKU" in row.index else ""

            final_stock = "0"

            # Check UPC availability
            for info in upcs.values():
                api_upc = normalize_upc(info.get("identifier", {}).get("upcNumber", ""))
                if api_upc and api_upc == upc:
                    final_stock = "2" if info.get("availability", {}).get("available") else "0"
                    break

            price_info = resolve_row_price(product_mkt_data, product_data, upc, sku)
            matched_offer = price_info["MatchedOffer"]

            # Offer stock override
            offer_stock = get_offer_stock_from_product_mkt(matched_offer)
            if offer_stock is not None:
                final_stock = offer_stock

            result_row = {
                "SKU": sku,
                "UPC": upc,
                "URL": url,
                "RegularPrice": price_info["RegularPrice"],
                "DiscountPrice": price_info["DiscountPrice"],
                "OfferPrice": price_info["OfferPrice"],
                "Stock": final_stock,
                "LastUpdated": current_time_str
            }

            if DEBUG_MODE:
                result_row.update({
                    "PriceSource": price_info["PriceSource"],
                    "MatchedOfferSKU": matched_offer.get("SKU", "") if matched_offer else "",
                    "MatchedColor": matched_offer.get("itemOffered", {}).get("color", "") if matched_offer else "",
                    "MatchedSize": matched_offer.get("itemOffered", {}).get("size", "") if matched_offer else ""
                })

            results.append(result_row)

        if len(results) >= CHUNK_SIZE:
            save_chunk(results)
            results.clear()
            logging.info("💾 Saved chunk to local CSV")

    return results


def get_google_creds():
    sa_json = os.environ.get("GCP_SA_KEY", "").strip()
    if not sa_json:
        logging.warning("⚠ GCP_SA_KEY not found in environment. Skipping Google services.")
        return None
    try:
        creds_dict = json.loads(sa_json)
        return Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    except Exception as e:
        logging.error(f"❌ Failed to parse GCP_SA_KEY: {e}")
        return None


def upload_to_google_sheet(csv_file):
    spreadsheet_id = os.environ.get("GOOGLE_SHEET_ID", "").strip()
    if not spreadsheet_id:
        logging.warning("⚠ GOOGLE_SHEET_ID not provided. Skipping Google Sheet update.")
        return

    creds = get_google_creds()
    if not creds:
        return

    try:
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(spreadsheet_id)
        worksheet = sh.get_worksheet(0)

        df = pd.read_csv(csv_file).fillna("")
        data_matrix = [df.columns.values.tolist()] + df.values.tolist()

        worksheet.clear()
        worksheet.update(values=data_matrix, range_name="A1")
        logging.info("✅ Google Sheet updated successfully!")
    except Exception as e:
        logging.error(f"❌ Google Sheet update error: {e}")


def upload_to_google_drive(csv_file):
    folder_id = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "").strip()
    if not folder_id:
        logging.warning("⚠ GOOGLE_DRIVE_FOLDER_ID not provided. Skipping Drive backup.")
        return

    creds = get_google_creds()
    if not creds:
        return

    try:
        drive_service = build("drive", "v3", credentials=creds)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        drive_file_name = f"Macys_Inventory_{timestamp}.csv"

        file_metadata = {
            "name": drive_file_name,
            "parents": [folder_id]
        }
        media = MediaFileUpload(csv_file, mimetype="text/csv")
        file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id"
        ).execute()

        logging.info(f"✅ Google Drive Backup saved! File ID: {file.get('id')}")
    except Exception as e:
        logging.error(f"❌ Google Drive upload error: {e}")


if __name__ == "__main__":
    if os.path.exists(OUTPUT_FILE):
        os.remove(OUTPUT_FILE)

    logging.info("🚀 Starting Macy's Inventory Scraper...")
    results = scrape_inventory(INPUT_FILE)

    if results:
        save_chunk(results)

    if os.path.exists(OUTPUT_FILE):
        logging.info("📤 Syncing data to Google Cloud...")
        upload_to_google_sheet(OUTPUT_FILE)
        upload_to_google_drive(OUTPUT_FILE)
    else:
        logging.warning("⚠ No output file generated.")

    logging.info("🎉 Scraping and sync process finished!")
