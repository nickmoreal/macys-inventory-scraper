# Mac Inventory Automation

This project runs the Mac inventory/price scraper automatically through GitHub Actions.

## Schedule

- 4:00 PM IST every day
- 10:00 PM IST every day
- Manual run is also available from GitHub Actions > Run workflow.

The workflow uses GitHub's timezone-aware schedule with `Asia/Kolkata`.

## Files

- `scraper.py` - scraper + Google Sheets + Google Drive backup
- `requirements.txt` - Python dependencies

## Input Excel

The Excel file must contain:

- `URL`
- `UPC`

`SKU` is optional.

## Google Sheet

The script appends each successfully completed chunk to the configured worksheet.

Columns:

RunTimeIST, SKU, UPC, URL, RegularPrice, DiscountPrice, OfferPrice, Stock

Debug columns are also included by default:

PriceSource, MatchedOfferSKU, MatchedColor, MatchedSize

## Backup

For each run, one CSV file is created in the configured Google Drive folder.


The same CSV is updated after every 25-row chunk.

Therefore, if the scraper stops after processing part of the input, the successfully saved portion remains in Google Drive and Google Sheets.

## GitHub Secrets

Create these repository secrets:

1. `GOOGLE_SERVICE_ACCOUNT_JSON`
2. `GOOGLE_SHEET_ID`
3. `GOOGLE_SHEET_NAME`
4. `GOOGLE_DRIVE_FOLDER_ID`

Recommended value:

`GOOGLE_SHEET_NAME = Inventory`

## Google setup

Create a Google Cloud project and enable:

- Google Sheets API
- Google Drive API

Create a service account and download its JSON credentials.

Share the target Google Sheet with the service-account email address as Editor.

Also share the Google Drive backup folder with the same service-account email address as Editor.

Do NOT commit the service-account JSON file to GitHub.

## Important

GitHub Actions scheduled workflows run from the default branch. Keep the workflow on the repository's default branch.

The scraper intentionally saves every completed chunk before moving on, reducing data loss if an API request or runner fails later in the run.
