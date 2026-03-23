# Walmart Content Extractor

Streamlit app for internal Walmart PDP extraction with:
- login screen backed by `auth.py`
- SKU + Walmart URL input grid
- checkbox-based attribute selection
- optional expansion for extra bullets and extra images
- Excel export with unavailable selected fields marked as `Unable to find on PDP` and highlighted in light pink
- Run Log sheet for rows that return warnings or errors

## Files
- `app.py` - main Streamlit app
- `auth.py` - approved user list and fixed password
- `requirements.txt` - Python dependencies

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud
1. Upload `app.py`, `auth.py`, `requirements.txt`, and `README.md` to your GitHub repo.
2. Point Streamlit Community Cloud to `app.py`.
3. Redeploy the app.

## Manage logins
Edit `auth.py`:
- add approved `@pattern.com` users to `ALLOWED_USERS`
- update `DEFAULT_PASSWORD` if needed

## Notes
- The app uses direct page requests plus HTML parsing.
- Some Walmart PDPs may still block requests or expose fewer fields than others.
- When a selected field is not found, the export writes `Unable to find on PDP`.
