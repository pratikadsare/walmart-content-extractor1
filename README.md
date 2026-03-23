# Walmart Content Extractor

Streamlit app for internal Walmart PDP extraction with:
- login screen backed by `auth.py`
- light theme via `.streamlit/config.toml`
- SKU + Walmart URL input grid
- checkbox-based attribute selection
- optional expansion for extra bullets and extra images
- product image filtering tuned to keep only PDP/item-related image URLs
- Excel export with unavailable selected fields marked as `Unable to find on PDP` and highlighted in light pink
- Run Log sheet for rows that return warnings or errors

## Files
- `app.py` - main Streamlit app
- `auth.py` - approved user list and fixed password
- `requirements.txt` - Python dependencies
- `.streamlit/config.toml` - light theme configuration
- `walmart_input_template_10_rows.xlsx` - blank upload template

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud
1. Upload all files and folders from this package to your GitHub repo.
2. Make sure `.streamlit/config.toml` is included.
3. Point Streamlit Community Cloud to `app.py`.
4. Redeploy the app.

## Manage logins
Edit `auth.py`:
- add approved `@pattern.com` users to `ALLOWED_USERS`
- update `DEFAULT_PASSWORD` if needed

## Notes
- The browser console warnings about unsupported iframe features come from the hosting environment and do not come from your scraper logic.
- The app uses direct page requests plus HTML parsing.
- Some Walmart PDPs may still block requests or expose fewer fields than others.
- When a selected field is not found, the export writes `Unable to find on PDP`.
