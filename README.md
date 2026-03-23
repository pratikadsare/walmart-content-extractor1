# Walmart Content Extractor

A Streamlit dashboard for approved Pattern users to extract Walmart product page content into Excel.

## Included files

- `app.py`
- `auth.py`
- `requirements.txt`
- `.streamlit/config.toml`
- `walmart_input_template_10_rows.xlsx`
- `README.md`

## Login setup

This version includes a simple login screen.

- Approved users must be listed in `auth.py`
- Email domain must be `@pattern.com`
- Fixed password is defined in `auth.py`
- Welcome message is generated automatically from the email address

Example:
- `pratik.adsare@pattern.com` -> `Welcome Pratik`
- `kunal.adsare@pattern.com` -> `Welcome Kunal`

## What the app currently extracts

Input columns:
- SKU
- Walmart URL

Output columns:
- SKU
- URL
- Title
- Description
- Bullet 1
- Bullet 2
- Bullet 3
- Bullet 4
- Bullet 5
- Bullet 6+ when found on the page

## UI features

- Login screen for approved users
- Clean dashboard layout
- Default 10 input rows
- Supports up to 1000 rows
- Spreadsheet-like input grid for paste/edit
- Compact scrollable table for larger batches
- Optional Excel/CSV upload to prefill the input grid
- Live progress updates during scraping
- Custom output filename box
- Downloadable Excel output
- Run Log sheet for failed rows
- Dynamic bullet columns when more than 5 bullets are found

## Deploy on Streamlit Community Cloud

1. Upload all extracted files to your GitHub repo root.
2. In Streamlit Community Cloud, deploy `app.py`.
3. If you need to approve more users, edit `auth.py` and redeploy.

## Local setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Notes

- This cloud-ready version uses direct page requests instead of a browser runtime.
- Some listings may load differently or trigger anti-bot verification, so failed rows are recorded in the `Run Log` sheet.
- The export always includes Bullet 1 to Bullet 5 and adds more bullet columns when found.
