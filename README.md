# Walmart Content Extractor - Accuracy Version

This package contains a stricter, section-locked version of the Walmart Content Extractor.

## Key updates
- Narrower centered login card
- Slower, accuracy-first scraping
- Stricter section-based extraction for specs, directions, indications, ingredients, and related PDP fields
- Stricter product-image filtering to avoid related-product, icon, logo, and duplicate image URLs
- Duplicate image removal
- Light theme forced via `.streamlit/config.toml`

## Files to upload
Upload all files and folders in this package to your GitHub repo, including the `.streamlit` folder.

## Important note
This version is designed to prefer blanks / `Unable to find on PDP` over guessed or unrelated values.
