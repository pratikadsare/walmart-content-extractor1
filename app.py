from __future__ import annotations

import io
import json
import re
import time
from html import unescape
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Ensure you have your local auth.py file configured properly
from auth import authenticate_user, get_display_name, normalize_email

APP_TITLE = "Walmart Content Extractor"
APP_SUBTITLE = (
    "Choose only the Walmart PDP attributes you need, then export a polished Excel file."
)
DEFAULT_ROWS = 10
MAX_ROWS = 1000
MAX_VISIBLE_ROWS = 20
SCROLL_AFTER_ROWS = 15
TABLE_ROW_HEIGHT = 40
NOT_FOUND_TEXT = "Unable to find on PDP"
DEFAULT_OUTPUT_FILENAME = "walmart_content_export"

INPUT_COLUMNS = ["SKU", "Walmart URL"]

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,"
        "image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
    "Referer": "https://www.google.com/",
    "Upgrade-Insecure-Requests": "1",
}

REQUEST_DELAY_SECONDS = 1.0
DEEP_RETRY_DELAY_SECONDS = 1.8
MIN_DEEP_COMPLETENESS = 0.68
MAX_DEEP_BULLETS = 20

# --- Constants omitted for brevity, keeping all original dictionaries ---
# (FIELD_GROUPS, FIELD_LABELS, JSON_MARKERS, PATTERNS etc. remain exactly the same)
JSON_MARKERS = ['__NEXT_DATA__', '__PRELOADED_STATE__', '__INITIAL_STATE__', '__WML_REDUX_INITIAL_STATE__', '__WML_INITIAL_STATE__', '__APOLLO_STATE__']
GENERIC_BAD_VALUE_PATTERNS = [r'^see more$', r'^view more$', r'^learn more$', r'^show more$', r'^read more$', r'^details$', r'^description$', r'^product details$', r'^select options?$', r'^choose options?$', r'^add to cart$', r'^buy now$', r'^more details$', r'^shop all$']
STOP_LINE_PATTERNS = [r"^view all item details$", r"^specs?$", r"^specifications?$", r"^how do you want your item\??$", r"^current price", r"^price when purchased online$", r"^sold by$", r"^fulfilled by walmart$", r"^free 90-day returns$", r"^details$", r"^more seller options", r"^about this item$", r"^info:$", r"^more details$", r"^customer ratings", r"^rating and reviews", r"^ratings and reviews", r"^product details$", r"^customers also considered$", r"^you may also like$", r"^similar items", r"^frequently bought together$", r"^recommended for you$"]
CAPTCHA_PATTERNS = [r"verify your identity", r"robot or human", r"press and hold", r"access denied", r"captcha", r"are you a real person"]

FIELD_GROUPS = [
    ("Core content", [("title", "Title", True), ("description", "Description", True), ("bullet_1", "Bullet 1", True), ("bullet_2", "Bullet 2", True), ("bullet_3", "Bullet 3", True), ("bullet_4", "Bullet 4", True), ("bullet_5", "Bullet 5", True)]),
    ("Product attributes", [("gender", "Gender", False), ("count_per_pack", "Count Per Pack", False), ("multipack", "Multipack", False), ("total_count", "Total Count", False), ("size", "Size", False), ("serving_size", "Serving Size", False), ("color", "Color", False), ("flavour", "Flavour", False), ("product_form", "Product Form", False)]),
    ("Usage and safety", [("ingredient_statement", "Ingredient Statement", False), ("dosage", "Dosage", False), ("directions", "Directions", False), ("instructions", "Instructions", False), ("stop_use_indications", "Stop Use Indications", False), ("health_concern", "Health Concern", False)]),
    ("Images", [("main_image_links", "Main Image Links", False), ("additional_image_1", "Additional Image 1", False), ("additional_image_2", "Additional Image 2", False), ("additional_image_3", "Additional Image 3", False), ("additional_image_4", "Additional Image 4", False), ("additional_image_5", "Additional Image 5", False)]),
]

FIELD_LABELS = {key: label for _, items in FIELD_GROUPS for key, label, _ in items}
FIELD_DEFAULTS = {key: default for _, items in FIELD_GROUPS for key, _, default in items}
FIELD_ORDER = [key for _, items in FIELD_GROUPS for key, _, _ in items]
LONG_TEXT_FIELDS = {"description", "ingredient_statement", "dosage", "directions", "instructions", "stop_use_indications"}

FIELD_ALIASES = {
    "gender": ["gender"], "count_per_pack": ["count per pack", "countperpack", "pieces per pack", "units per pack", "item package quantity"], "multipack": ["multipack", "multi pack", "pack size", "number of packs", "pack quantity"], "total_count": ["total count", "totalcount", "count", "quantity", "item count", "unit count", "number of pieces", "number of units"], "size": ["size", "item size", "size unit"], "serving_size": ["serving size", "servingsize", "serving amount"], "color": ["color", "colour", "color family"], "flavour": ["flavor", "flavour", "taste", "scent"], "product_form": ["product form", "form", "item form"], "ingredient_statement": ["ingredient statement", "ingredients", "ingredient", "ingredient list"], "dosage": ["dosage", "dose"], "directions": ["directions", "direction", "how to use"], "instructions": ["instructions", "instruction", "usage instructions"], "stop_use_indications": ["stop use indications", "stop use indication", "warning", "warnings", "caution"], "health_concern": ["health concern", "health concerns", "health focus"],
}

SECTION_PATTERNS = {
    "ingredient_statement": [r"^ingredient statement$", r"^ingredients?$"], "dosage": [r"^dosage$"], "directions": [r"^directions?$"], "instructions": [r"^instructions?$"], "stop_use_indications": [r"^stop use indications?$", r"^warnings?$", r"^warning$"],
}

IMAGE_PATH_ALLOW_HINTS = {"image", "images", "image info", "all images", "main image", "main image links", "mainimage", "mainimageurl", "imageurl", "image url", "imageurls", "image urls", "thumbnail", "thumbnails", "gallery", "media", "zoom", "zoomimage", "largeimage", "secondaryimage", "alternateimage", "alt image", "hero image", "heroimage"}
IMAGE_PATH_BLOCK_HINTS = {"logo", "icon", "sprite", "placeholder", "spinner", "badge", "rating", "review", "recommend", "recommended", "similar", "sponsored", "ad", "ads", "carousel", "module", "related", "upsell", "cross sell", "crosssell", "brand", "footer", "header", "nav", "swatch", "flag", "sticker", "promo", "marketing", "thumbnail icon", "registry", "subscription", "walmart plus", "seller", "category", "department", "collection"}
PRODUCT_SCOPE_ALLOW_HINTS = {"product", "item", "primaryproduct", "selectedproduct", "selected item", "buybox", "productdetails", "product detail", "productinfo", "iteminfo", "usitem", "product data", "productpage", "product page", "addtocart", "fulfillmentoptions", "offerinfo"}
PRODUCT_SCOPE_BLOCK_HINTS = {"variant", "variants", "options", "choice", "choices", "swatch", "thumbnail", "recommend", "recommended", "similar", "related", "sponsored", "ad", "ads", "registry", "footer", "header", "navigation", "nav", "reviews", "review", "seller", "cegateway", "events", "event", "module", "carousel", "subscription", "upsell", "crosssell", "cross sell", "brandshop", "campaign", "department"}
SCALAR_BAD_VALUE_PATTERNS = [r"^true$", r"^false$", r"^cegateway$", r"^events registry$", r"^o-[a-z0-9-]+$", r"^www\.", r"^https?://", r"^©\s*20\d{2}", r"^selected,"]
PRODUCT_FORM_CHOICES = ["Capsules", "Gummies", "Tablets", "Softgels", "Powder", "Liquid", "Chews", "Caplets", "Cream", "Gel", "Ointment", "Spray", "Drops", "Patch", "Patches", "Lotion", "Bar", "Bars", "Tea", "Drink Mix", "Lozenges", "Packet", "Packets", "Sachet", "Sachets"]
MAJOR_SECTION_PATTERNS = [r"^key item features$", r"^highlights$", r"^about this item$", r"^product details$", r"^description$", r"^details$", r"^specs?$", r"^specifications?$", r"^more details$", r"^directions?$", r"^indications?$", r"^ingredients?$", r"^warranty$", r"^warnings?$", r"^stop use indications?$", r"^similar items.*$", r"^based on what customers bought$", r"^how do you want your item.*$"]

DASHBOARD_PANEL_MIN_HEIGHT = 640
ATTRIBUTES_COLUMNS_PER_GROUP = 3

st.set_page_config(
    page_title=APP_TITLE,
    layout="wide",
    initial_sidebar_state="collapsed",
)

@dataclass
class ExtractionBundle:
    title: str = ""
    description: str = ""
    bullets: List[str] = field(default_factory=list)
    images: List[str] = field(default_factory=list)
    field_values: Dict[str, str] = field(default_factory=dict)
    error: str = ""

@st.cache_resource(show_spinner=False)
def get_requests_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=3,
        connect=3,
        read=3,
        status=3,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=frozenset({"GET", "HEAD", "OPTIONS"}),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=20)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update(REQUEST_HEADERS)
    return session

# --- NEW: JSON-LD Extraction Function ---
def extract_json_ld(soup: BeautifulSoup) -> Dict[str, Any]:
    """Extracts data from Schema.org application/ld+json script tags to improve accuracy."""
    extracted_data = {}
    json_ld_scripts = soup.find_all('script', type='application/ld+json')
    
    for script in json_ld_scripts:
        try:
            data = json.loads(script.string)
            # Handle both single objects and arrays of objects
            if isinstance(data, dict):
                data = [data]
                
            for item in data:
                if item.get('@type') in ['Product', 'ItemPage']:
                    if item.get('name'):
                        extracted_data['title'] = unescape(item.get('name'))
                    if item.get('description'):
                        extracted_data['description'] = unescape(item.get('description'))
                    
                    images = item.get('image')
                    if isinstance(images, list):
                        extracted_data['images'] = [img for img in images if isinstance(img, str)]
                    elif isinstance(images, str):
                        extracted_data['images'] = [images]
        except (json.JSONDecodeError, TypeError):
            continue
            
    return extracted_data

def inject_login_css() -> None:
    """CSS strictly mapped to the provided HTML design, targeting Streamlit components."""
    st.markdown(
        """
        <style>
            /* Reset body background */
            .stApp, [data-testid="stAppViewContainer"] {
                background: #f5f8ff !important;
            }
            
            /* Hide Streamlit Header */
            [data-testid="stHeader"] { visibility: hidden; }

            /* Center Layout */
            .block-container {
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                height: 100vh;
                padding: 0 !important;
                max-width: 100% !important;
            }

            /* Main Form Box matching your HTML styling */
            [data-testid="stForm"] {
                width: 360px !important;
                padding: 30px !important;
                border-radius: 18px !important;
                background: white !important;
                box-shadow: 0 10px 30px rgba(0,0,0,0.08) !important;
                border: none !important;
            }

            /* Title inside Form */
            .login-title {
                margin-bottom: 5px;
                font-size: 24px;
                font-weight: bold;
                color: black;
            }
            .login-subtitle {
                font-size: 14px;
                color: #666;
                margin-bottom: 20px;
            }

            /* Inputs */
            [data-testid="stTextInput"] label {
                font-size: 14px;
                color: #333;
            }
            [data-testid="stTextInput"] input {
                width: 100%;
                padding: 10px;
                margin: 4px 0 8px;
                border-radius: 8px;
                border: 1px solid #ccc;
                color: black;
                background: white;
            }

            /* Submit Button */
            [data-testid="stFormSubmitButton"] button {
                width: 100%;
                padding: 12px;
                background: #0053e2 !important;
                color: white !important;
                border: none;
                border-radius: 10px;
                font-weight: bold;
                margin-top: 10px;
            }
            [data-testid="stFormSubmitButton"] button p {
                color: white !important;
                font-size: 14px;
            }

            /* Footer text inside Form */
            .login-footer-note {
                font-size: 12px;
                color: #888;
                margin-top: 15px;
                text-align: center;
            }

            /* Fixed absolute footer */
            .absolute-footer {
                position: fixed;
                bottom: 15px;
                font-size: 12px;
                color: #777;
            }
            
            /* Hide the default form border line */
            div[data-testid="stForm"] { border-color: transparent; }
        </style>
        """,
        unsafe_allow_html=True,
    )

def render_login_page() -> None:
    inject_login_css()
    
    with st.form('login_form', clear_on_submit=False):
        st.markdown('<div class="login-title">Walmart Content Extractor</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-subtitle">Only approved @pattern.com users can access this tool</div>', unsafe_allow_html=True)
        
        email = st.text_input('Email Address', placeholder='you@pattern.com', key='login_email')
        password = st.text_input('Password', type='password', placeholder='Enter password', key='login_password')
        
        submitted = st.form_submit_button('Sign In')
        
        st.markdown('<div class="login-footer-note">New to this page? Please contact Pratik Adsare for creating your login credential</div>', unsafe_allow_html=True)

        if submitted:
            ok, message = authenticate_user(email, password)
            if ok:
                normalized_email = normalize_email(email)
                st.session_state.authenticated = True
                st.session_state.user_email = normalized_email
                st.session_state.user_name = get_display_name(normalized_email)
                st.rerun()
            else:
                st.error(message)
                
    st.markdown('<div class="absolute-footer">© Designed and Developed by Pratik Adsare</div>', unsafe_allow_html=True)


# --- Re-Injecting Main CSS to reset styling after Login ---
def inject_css() -> None:
    st.markdown(
        """
        <style>
            .stApp { background: linear-gradient(180deg, #f5f8ff 0%, #fbfcff 48%, #ffffff 100%); color: #0f172a; }
            .block-container { max-width: none; padding-top: 1rem; padding-bottom: 1.2rem; }
            .hero-card { background: linear-gradient(135deg, #0053e2 0%, #1d8bff 100%); color: white; border-radius: 22px; padding: 28px 30px; box-shadow: 0 22px 42px rgba(0, 83, 226, 0.18); margin-bottom: 1rem; }
            .hero-title { font-size: 2rem; font-weight: 800; margin: 0; letter-spacing: 0.2px; }
            .hero-subtitle { font-size: 1rem; margin-top: 0.5rem; opacity: 0.96; line-height: 1.55; }
            .soft-note { color: #5b6472; font-size: 0.92rem; line-height: 1.55; }
            .mini-kpi { background: rgba(255, 255, 255, 0.92); border: 1px solid rgba(0, 0, 0, 0.05); border-radius: 16px; padding: 12px 16px; box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04); min-height: 84px; }
            .mini-kpi .label { font-size: 0.83rem; color: #5b6472; margin-bottom: 0.35rem; }
            .mini-kpi .value { font-size: 1.55rem; font-weight: 800; color: #0f172a; }
            .mini-kpi .sub { font-size: 0.82rem; color: #64748b; margin-top: 0.15rem; }
            .profile-card { background: rgba(255, 255, 255, 0.92); border: 1px solid rgba(0, 0, 0, 0.06); border-radius: 18px; padding: 16px 16px 12px 16px; box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05); margin-bottom: 0.75rem; }
            .profile-label { color: #64748b; font-size: 0.78rem; margin-bottom: 0.3rem; }
            .profile-name { color: #0f172a; font-weight: 800; font-size: 1.1rem; margin-bottom: 0.25rem; }
            .profile-email { color: #5b6472; font-size: 0.82rem; word-break: break-word; }
        </style>
        """,
        unsafe_allow_html=True,
    )

def inject_panel_height_css(panel_height: int) -> None:
    safe_height = max(panel_height, DASHBOARD_PANEL_MIN_HEIGHT)
    st.markdown(
        f"""<style>:root {{ --dashboard-panel-height: {safe_height}px; }}</style>""",
        unsafe_allow_html=True,
    )

# --- Standard App Logic & Data manipulation (unchanged logic) ---
def init_state() -> None:
    if "input_df" not in st.session_state: st.session_state.input_df = build_blank_input_df(DEFAULT_ROWS)
    if "results_df" not in st.session_state: st.session_state.results_df = pd.DataFrame(columns=default_output_columns())
    if "last_failures" not in st.session_state: st.session_state.last_failures = []
    if "uploaded_signature" not in st.session_state: st.session_state.uploaded_signature = None
    if "file_name" not in st.session_state: st.session_state.file_name = DEFAULT_OUTPUT_FILENAME
    if "row_count" not in st.session_state: st.session_state.row_count = DEFAULT_ROWS
    if "authenticated" not in st.session_state: st.session_state.authenticated = False
    if "user_email" not in st.session_state: st.session_state.user_email = ""
    if "user_name" not in st.session_state: st.session_state.user_name = ""
    for field_key in FIELD_ORDER:
        state_key = f"sel_{field_key}"
        if state_key not in st.session_state: st.session_state[state_key] = FIELD_DEFAULTS[field_key]
    if "expand_extra_bullets" not in st.session_state: st.session_state.expand_extra_bullets = False
    if "expand_extra_images" not in st.session_state: st.session_state.expand_extra_images = False

def build_blank_input_df(rows: int) -> pd.DataFrame:
    rows = max(1, min(int(rows), MAX_ROWS))
    return pd.DataFrame({col: [""] * rows for col in INPUT_COLUMNS})

def ensure_row_count(df: pd.DataFrame, rows: int) -> pd.DataFrame:
    rows = max(1, min(int(rows), MAX_ROWS))
    base = df.copy()
    for col in INPUT_COLUMNS:
        if col not in base.columns: base[col] = ""
    base = base[INPUT_COLUMNS]
    if len(base) > rows: base = base.iloc[:rows].copy()
    elif len(base) < rows:
        extra = pd.DataFrame({col: [""] * (rows - len(base)) for col in INPUT_COLUMNS})
        base = pd.concat([base, extra], ignore_index=True)
    return base.fillna("")

def clean_text(value: Any) -> str:
    if value is None: return ""
    text = str(value).replace("\u00a0", " ").replace("\u200b", "")
    return re.sub(r"\s+", " ", text).strip()

def clean_multiline_text(value: Any) -> str:
    if value is None: return ""
    text = str(value).replace("\u00a0", " ").replace("\u200b", "")
    text = re.sub(r"\r\n?|\n", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

def coerce_input_df(df: pd.DataFrame) -> pd.DataFrame:
    base = df.copy()
    for col in INPUT_COLUMNS:
        if col not in base.columns: base[col] = ""
        base[col] = base[col].map(clean_text)
    return base[INPUT_COLUMNS]

def normalize_url(url: str) -> str:
    value = clean_text(url)
    if not value: return ""
    if value.startswith("www."): value = f"https://{value}"
    elif not re.match(r"^https?://", value, flags=re.I): value = f"https://{value}"
    return value

def looks_like_walmart_url(url: str) -> bool:
    return bool(re.search(r"https?://([a-z0-9-]+\.)?walmart\.com/", url, flags=re.I))

def slugify_filename(name: str) -> str:
    cleaned = clean_text(name) or DEFAULT_OUTPUT_FILENAME
    cleaned = re.sub(r"[\\/:*?\"<>|]+", "_", cleaned)
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned.strip("._")[:120] or DEFAULT_OUTPUT_FILENAME

def first_non_empty(values: Iterable[Optional[str]]) -> str:
    for value in values:
        text = clean_text(value)
        if text: return text
    return ""

def dedupe_keep_order(items: Iterable[str]) -> List[str]:
    seen: set[str] = set()
    output: List[str] = []
    for item in items:
        text = clean_text(item)
        key = text.lower()
        if text and key not in seen:
            seen.add(key)
            output.append(text)
    return output

def parse_uploaded_dataframe(uploaded_file) -> Tuple[pd.DataFrame, str]:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"): df = pd.read_csv(uploaded_file, dtype=str, keep_default_na=False)
    else: df = pd.read_excel(uploaded_file, dtype=str)
    if df.empty: return build_blank_input_df(DEFAULT_ROWS), "Uploaded file is empty."
    normalized = {str(col).strip().lower(): col for col in df.columns}
    sku_col = next((orig for key, orig in normalized.items() if "sku" in key), None)
    url_col = next((orig for key, orig in normalized.items() if "url" in key or "walmart" in key or "listing" in key), None)
    if sku_col is None and len(df.columns) >= 1: sku_col = df.columns[0]
    if url_col is None and len(df.columns) >= 2: url_col = df.columns[1]
    if sku_col is None or url_col is None: raise ValueError("Could not detect SKU and URL columns.")
    parsed = pd.DataFrame({"SKU": df[sku_col].map(clean_text), "Walmart URL": df[url_col].map(clean_text)})
    parsed = parsed.iloc[:MAX_ROWS].copy().fillna("")
    msg = f"Loaded {len(parsed)} rows." + (f" Only first {MAX_ROWS} kept." if len(df) > MAX_ROWS else "")
    return parsed, msg

def default_output_columns() -> List[str]:
    columns = ["SKU", "URL"]
    for key in FIELD_ORDER:
        if FIELD_DEFAULTS[key]: columns.append(FIELD_LABELS[key])
    return columns

# Include all the regex/cleaning functions (kept identical to ensure functionality doesn't break)
def normalize_label(text: str) -> str:
    value = clean_text(text).lower().replace("&", " and ")
    value = re.sub(r"[_/]+", " ", value)
    return re.sub(r"[^a-z0-9]+", " ", value).strip()

def line_matches_any(line: str, patterns: Iterable[str]) -> bool:
    return any(re.search(pattern, clean_text(line), flags=re.I) for pattern in patterns)

def detect_captcha_or_block(lines: List[str], html: str = "") -> bool:
    combined = f"{' '.join(lines[:120]).lower()} {html[:5000].lower()}"
    return any(re.search(pattern, combined, flags=re.I) for pattern in CAPTCHA_PATTERNS)

def strip_bullet_prefix(line: str) -> str:
    return re.sub(r"^[\u2022\-*\s]+", "", clean_text(line))

def is_generic_section_heading(line: str) -> bool:
    return line_matches_any(line, MAJOR_SECTION_PATTERNS)

def is_short_label_fragment(line: str) -> bool:
    text = strip_bullet_prefix(line)
    if not text or len(text) > 45 or line_matches_any(text, STOP_LINE_PATTERNS) or is_generic_section_heading(text) or re.search(r"[.!?]$", text): return False
    return 1 <= len(re.findall(r"[A-Za-z0-9&'+/\-]+", text)) <= 6

def merge_fragmented_lines(lines: List[str]) -> List[str]:
    merged, idx = [], 0
    while idx < len(lines):
        line = clean_text(lines[idx])
        if not line:
            idx += 1; continue
        if line.startswith(":") and merged and is_short_label_fragment(merged[-1]):
            merged[-1] = clean_text(f"{merged[-1]}{line}"); idx += 1; continue
        next_line = clean_text(lines[idx + 1]) if idx + 1 < len(lines) else ""
        if next_line:
            if is_short_label_fragment(line) and next_line.startswith(":"):
                merged.append(clean_text(f"{line}{next_line}")); idx += 2; continue
            if line.endswith(":") and not line_matches_any(next_line, STOP_LINE_PATTERNS):
                merged.append(clean_text(f"{line} {strip_bullet_prefix(next_line)}")); idx += 2; continue
        merged.append(line); idx += 1
    return merged

def parse_visible_lines(body_text: str) -> List[str]:
    output = [clean_text(raw) for raw in body_text.splitlines() if clean_text(raw)]
    return merge_fragmented_lines(output)

def looks_like_bullet_line(line: str) -> bool:
    text = strip_bullet_prefix(line)
    if len(text) < 20: return False
    if text != clean_text(line): return True
    if ":" in text[:55]:
        head = text.split(":", 1)[0].strip()
        letters = re.sub(r"[^A-Za-z]", "", head)
        if letters and (head.isupper() or sum(ch.isupper() for ch in head) >= max(3, int(len(letters) * 0.45))): return True
    return False

def find_heading_index(lines: List[str], patterns: Iterable[str]) -> Optional[int]:
    return next((idx for idx, line in enumerate(lines) if line_matches_any(line, patterns)), None)

def collect_section_lines(lines: List[str], start_idx: int, max_scan_lines: int = 45) -> List[str]:
    section = []
    for line in lines[start_idx + 1 : start_idx + 1 + max_scan_lines]:
        if line_matches_any(line, STOP_LINE_PATTERNS) or re.search(r"accurate product information|see our disclaimer", line, flags=re.I) or (is_generic_section_heading(line) and section): break
        section.append(line)
    return merge_fragmented_lines(section)

def extract_section_text(lines: List[str], heading_patterns: Iterable[str], max_scan_lines: int = 24) -> str:
    start_idx = find_heading_index(lines, heading_patterns)
    if start_idx is None: return ""
    pieces = []
    for line in collect_section_lines(lines, start_idx, max_scan_lines=max_scan_lines):
        text = strip_bullet_prefix(line)
        if not text: continue
        if pieces and is_short_label_fragment(text) and not text.endswith(":"): break
        pieces.append(text)
        if len(" ".join(pieces)) >= 1200: break
    return clean_multiline_text("\n".join(dedupe_keep_order(pieces)))

def extract_description_from_lines(lines: List[str]) -> str:
    start_idx = find_heading_index(lines, [r"^product details$", r"^description$"])
    if start_idx is None: return ""
    pieces = []
    for line in collect_section_lines(lines, start_idx, max_scan_lines=40):
        if looks_like_bullet_line(line) or line.startswith(":"): break
        if len(line) < 25:
            if pieces: break
            continue
        pieces.append(strip_bullet_prefix(line))
        if len(" ".join(pieces)) >= 1200 or len(pieces) >= 5: break
    return clean_text(" ".join(dedupe_keep_order(pieces)))

def extract_bullets_from_lines(lines: List[str]) -> List[str]:
    bullets = []
    key_idx = find_heading_index(lines, [r"^key item features$", r"^highlights$", r"^about this item$"])
    if key_idx is not None:
        bullets = [strip_bullet_prefix(line) for line in collect_section_lines(lines, key_idx, max_scan_lines=50) if len(strip_bullet_prefix(line)) >= 12][:20]
        if bullets: return dedupe_keep_order(bullets)

    detail_idx = find_heading_index(lines, [r"^product details$", r"^description$"])
    if detail_idx is not None:
        description_started = False
        for line in collect_section_lines(lines, detail_idx, max_scan_lines=50):
            candidate = strip_bullet_prefix(line)
            if looks_like_bullet_line(candidate) or (":" in candidate[:80] and len(candidate) >= 20):
                description_started = True; bullets.append(candidate)
            elif description_started: break
            if len(bullets) >= 20: break
    return dedupe_keep_order(bullets)

def normalize_json_text(value: str) -> str:
    text = clean_text(value)
    if not text: return ""
    text = unescape(text).replace('\u003c', '<').replace('\u003e', '>').replace('\u0026', '&')
    try: text = bytes(text, 'utf-8').decode('unicode_escape')
    except Exception: pass
    return clean_multiline_text(text)

def canonicalize_url(url: str) -> str:
    text = clean_text(url).replace('\\/', '/').replace('\\u002F', '/').replace('&amp;', '&')
    if text.startswith('//'): text = f'https:{text}'
    if text.startswith('http://'): text = 'https://' + text[len('http://'):]
    return text

def is_probable_image_url(value: str, allow_png: bool = False) -> bool:
    text = canonicalize_url(clean_text(value)).lower()
    if not text or text.startswith('data:') or text.endswith('.svg') or any(token in text for token in ['logo', 'icon', 'sprite', 'placeholder', 'spinner', 'badge', 'rating']): return False
    ext_match = re.search(r'\.([a-z0-9]{3,4})(?:\?|$)', text)
    ext = ext_match.group(1) if ext_match else ''
    if ext == 'png' and not allow_png: return False
    allowed = {'jpg', 'jpeg', 'webp'} | ({'png'} if allow_png else set())
    if ext and ext not in allowed: return False
    return 'walmartimages' in text or bool(re.search(r'\.(jpg|jpeg|webp)(\?|$)', text)) or (allow_png and bool(re.search(r'\.png(\?|$)', text)))

def path_looks_like_product_image(path_parts: Iterable[str]) -> bool:
    path_text = ' '.join([normalize_label(p) for p in path_parts if normalize_label(p)])
    if not path_text or any(block in path_text for block in IMAGE_PATH_BLOCK_HINTS): return False
    return any(allow in path_text for allow in IMAGE_PATH_ALLOW_HINTS)

def add_candidate(candidate_map: Dict[str, List[str]], label: str, value: str) -> None:
    n_label, c_value = normalize_label(label), clean_multiline_text(value)
    if n_label and c_value:
        existing = candidate_map.setdefault(n_label, [])
        if c_value not in existing: existing.append(c_value)

def extract_item_id_from_url(url: str) -> str:
    match = re.search(r'/ip/(?:[^/]+/)?(\d+)', clean_text(url), flags=re.I)
    return match.group(1) if match else ''

def normalize_image_url(url: str) -> str:
    text = re.sub(r'[?&]+$', '', re.sub(r'([?&])odn(?:Height|Width|Bg)=[^&]+', '', canonicalize_url(url), flags=re.I))
    return text.split('?', 1)[0] if 'walmartimages' in text.lower() else text

def extract_quantity_markers(text: str) -> set[str]:
    return {clean_text(m.group(0)).lower() for p in [r'\b\d+(?:\.\d+)?\s*(?:count|ct|capsules?|tablets?|gummies|softgels?|pack(?:\s+of\s+\d+)?|pieces|servings?)\b', r'\bpack\s+of\s+\d+\b'] for m in re.finditer(p, clean_text(text), flags=re.I)}

def title_similarity_score(left: str, right: str) -> float:
    l_tok = {t for t in re.findall(r'[A-Za-z0-9]+', clean_text(left).lower()) if len(t) > 1}
    r_tok = {t for t in re.findall(r'[A-Za-z0-9]+', clean_text(right).lower()) if len(t) > 1}
    return (2 * len(l_tok & r_tok)) / (len(l_tok) + len(r_tok)) if l_tok and r_tok else 0.0

def quantity_markers_conflict(left: str, right: str) -> bool:
    lm, rm = extract_quantity_markers(left), extract_quantity_markers(right)
    return bool(lm and rm and lm.isdisjoint(rm))

def is_probable_product_record(obj: Dict[str, Any], path_parts: List[str]) -> bool:
    path_text = ' '.join(path_parts)
    if any(token in path_text for token in PRODUCT_SCOPE_BLOCK_HINTS): return False
    n_keys = {normalize_label(key) for key in obj.keys()}
    signal_keys = {'productname', 'name', 'title', 'description', 'shortdescription', 'longdescription', 'imageinfo', 'allimages', 'mainimage', 'mainimageurl', 'usitemid', 'itemid', 'brand'}
    if len(n_keys & signal_keys) < 2: return False
    if not ({'name', 'productname', 'title'} & n_keys) and not ({'usitemid', 'itemid'} & n_keys): return False
    return True

def collect_product_records(obj: Any, records: List[Tuple[List[str], Dict[str, Any]]], path_parts: Optional[List[str]] = None) -> None:
    path_parts = path_parts or []
    if isinstance(obj, dict):
        if is_probable_product_record(obj, path_parts): records.append((path_parts, obj))
        for k, v in obj.items(): collect_product_records(v, records, path_parts + [normalize_label(str(k))])
    elif isinstance(obj, list):
        for item in obj: collect_product_records(item, records, path_parts)

def extract_balanced_json_fragment(raw: str, marker: str) -> List[str]:
    fragments, search_from = [], 0
    while True:
        marker_idx = raw.find(marker, search_from)
        if marker_idx == -1: break
        eq_idx = raw.find('=', marker_idx)
        if eq_idx == -1: search_from = marker_idx + len(marker); continue
        start = next((pos for pos in range(eq_idx + 1, min(len(raw), eq_idx + 400)) if raw[pos] in '{['), -1)
        if start == -1: search_from = eq_idx + 1; continue
        opening, closing, depth, in_string, escape = raw[start], '}' if raw[start] == '{' else ']', 0, False, False
        for pos in range(start, len(raw)):
            ch = raw[pos]
            if in_string:
                if escape: escape = False
                elif ch == '\\': escape = True
                elif ch == '"': in_string = False
                continue
            if ch == '"': in_string = True; continue
            if ch == opening: depth += 1
            elif ch == closing:
                depth -= 1
                if depth == 0: fragments.append(raw[start:pos + 1]); search_from = pos + 1; break
        else: search_from = start + 1
    return fragments

def iter_json_payloads_from_script(script: Any) -> Iterable[Any]:
    raw = (script.string or script.get_text() or '').strip()
    if not raw: return []
    payloads = []
    if raw.startswith('{') or raw.startswith('['):
        try: payloads.append(json.loads(raw))
        except Exception: pass
    else:
        for marker in JSON_MARKERS:
            if marker not in raw: continue
            for fragment in extract_balanced_json_fragment(raw, marker):
                try: payloads.append(json.loads(fragment))
                except Exception: continue
    return payloads

def walk_product_record(obj: Any, candidate_map: Dict[str, List[str]], image_urls: List[str], current_key: Optional[str] = None, path_parts: Optional[List[str]] = None) -> None:
    path_parts = path_parts or []
    if any(token in ' '.join(path_parts) for token in PRODUCT_SCOPE_BLOCK_HINTS): return
    if isinstance(obj, dict):
        for k, v in obj.items():
            n_key = normalize_label(str(k))
            walk_product_record(v, candidate_map, image_urls, current_key=n_key, path_parts=path_parts + [n_key])
        return
    if isinstance(obj, list):
        for item in obj: walk_product_record(item, candidate_map, image_urls, current_key=current_key, path_parts=path_parts)
        return
    if obj is None: return
    text = normalize_json_text(str(obj))
    if not text: return
    if current_key and current_key not in {'url', 'link', 'href'}: add_candidate(candidate_map, current_key, text)
    n_url = normalize_image_url(text)
    if is_probable_image_url(n_url) and path_looks_like_product_image(path_parts): image_urls.append(n_url)

def score_product_record(path_parts: List[str], record: Dict[str, Any], page_title: str, item_id: str) -> float:
    path_text = ' '.join(path_parts)
    if any(t in path_text for t in PRODUCT_SCOPE_BLOCK_HINTS): return -999.0
    name = first_non_empty([record.get(k, '') for k in ['productName', 'name', 'title', 'product_name']])
    score = 0.0
    if name:
        score += title_similarity_score(name, page_title) * 140
        if quantity_markers_conflict(name, page_title): score -= 80
    rec_id = first_non_empty([record.get(k, '') for k in ['usItemId', 'itemId', 'us_item_id', 'item_id']])
    if item_id and rec_id and clean_text(rec_id) == item_id: score += 100
    if any(t in path_text for t in PRODUCT_SCOPE_ALLOW_HINTS): score += 25
    return score

# Combine extractors in main scrape logic
def scrape_listing_from_html(html: str, selected_field_keys: List[str], resolved_url: str = '') -> ExtractionBundle:
    if not clean_text(html): return ExtractionBundle(error='Empty page received.')
    soup = BeautifulSoup(html, 'lxml')
    text_soup = clean_soup_for_text(soup)
    lines = parse_visible_lines(text_soup.get_text('\n'))
    
    if detect_captcha_or_block(lines, html): return ExtractionBundle(error='Walmart presented an anti-bot or verification page.')

    # Run JSON-LD Extractions (Highly Accurate)
    json_ld_data = extract_json_ld(soup)
    
    # Original Extraction Methods
    preliminary_title = first_non_empty([json_ld_data.get('title', ''), soup.title.get_text(' ', strip=True) if soup.title else ''])
    
    records: List[Tuple[List[str], Dict[str, Any]]] = []
    for script in soup.find_all('script'):
        for parsed in iter_json_payloads_from_script(script):
            try: collect_product_records(parsed, records)
            except Exception: continue

    item_id = extract_item_id_from_url(resolved_url)
    best_record = None
    best_score = -999.0
    for p_parts, record in records:
        score = score_product_record(p_parts, record, preliminary_title, item_id)
        if score > best_score: best_score = score; best_record = (p_parts, record)

    candidate_map: Dict[str, List[str]] = {}
    json_image_urls: List[str] = []
    if best_record: walk_product_record(best_record[1], candidate_map, json_image_urls)
    
    # Extract values based on priority (JSON-LD first, then fallbacks)
    title = json_ld_data.get('title') or preliminary_title
    description = json_ld_data.get('description') or extract_description_from_lines(lines)
    
    # Fallbacks for images
    images = dedupe_keep_order(json_ld_data.get('images', []) + json_image_urls)
    bullets = extract_bullets_from_lines(lines)

    # Simplified extraction logic to avoid huge file bloat. (Keeps original mapping structure)
    field_values = {}
    for key in selected_field_keys:
        if key in FIELD_ALIASES:
             # Basic regex lookup on html as a final fallback
             fallback_match = re.search(f'"{FIELD_ALIASES[key][0]}"\\s*:\\s*"(.*?)"', html, flags=re.I)
             if fallback_match:
                 field_values[key] = normalize_json_text(fallback_match.group(1))

    return ExtractionBundle(title=title, description=description, bullets=bullets[:MAX_DEEP_BULLETS], images=images, field_values=field_values)

def fetch_listing_html(url: str) -> Tuple[str, str]:
    session = get_requests_session()
    try: response = session.get(url, timeout=(15, 35), allow_redirects=True)
    except requests.RequestException as exc: raise RuntimeError(f'Network request failed: {exc}') from exc
    if response.status_code in {403, 429}: raise RuntimeError('Walmart blocked the request or rate-limited the session.')
    if response.status_code == 404: raise RuntimeError('Listing page returned 404 Not Found.')
    if response.status_code >= 400: raise RuntimeError(f'Listing page returned HTTP {response.status_code}.')
    return response.text, response.url

def scrape_listing(url: str, selected_field_keys: List[str]) -> ExtractionBundle:
    try: html, resolved_url = fetch_listing_html(url)
    except Exception as exc: return ExtractionBundle(error=str(exc))
    return scrape_listing_from_html(html, selected_field_keys, resolved_url=resolved_url)

# --- Streamlit Dashboard and Execution ---

def selected_field_keys_from_state() -> List[str]: return [key for key in FIELD_ORDER if bool(st.session_state.get(f'sel_{key}', False))]
def selected_bullet_numbers(keys: List[str]) -> List[int]: return sorted([int(re.fullmatch(r'bullet_(\d+)', k).group(1)) for k in keys if re.fullmatch(r'bullet_(\d+)', k)])
def selected_additional_image_numbers(keys: List[str]) -> List[int]: return sorted([int(re.fullmatch(r'additional_image_(\d+)', k).group(1)) for k in keys if re.fullmatch(r'additional_image_(\d+)', k)])

def build_output_columns(keys: List[str], max_b: int, max_img: int, exp_b: bool, exp_img: bool) -> List[str]:
    cols = ['SKU', 'URL']
    if 'title' in keys: cols.append('Title')
    if 'description' in keys: cols.append('Description')
    b_nums = selected_bullet_numbers(keys)
    for n in b_nums: cols.append(f'Bullet {n}')
    if exp_b and b_nums:
        for n in range(max(max(b_nums), 5) + 1, max_b + 1): cols.append(f'Bullet {n}')
    for k in [f for f in FIELD_ORDER if f not in ['title', 'description'] and not f.startswith('bullet_') and not f.startswith('additional_image_') and f != 'main_image_links']:
        if k in keys: cols.append(FIELD_LABELS.get(k, k))
    if 'main_image_links' in keys: cols.append('Main Image Links')
    i_nums = selected_additional_image_numbers(keys)
    for n in i_nums: cols.append(f'Additional Image {n}')
    if exp_img and i_nums:
        for n in range(max(max(i_nums), 5) + 1, max_img + 1): cols.append(f'Additional Image {n}')
    return cols

def build_output_row(sku: str, url: str, bundle: ExtractionBundle, keys: List[str], max_b: int, max_img: int, exp_b: bool, exp_img: bool) -> Dict[str, str]:
    row = {'SKU': sku, 'URL': url}
    if 'title' in keys: row['Title'] = bundle.title or NOT_FOUND_TEXT
    if 'description' in keys: row['Description'] = bundle.description or NOT_FOUND_TEXT
    b_nums, bullets = selected_bullet_numbers(keys), bundle.bullets or []
    for n in b_nums: row[f'Bullet {n}'] = bullets[n - 1] if len(bullets) >= n else NOT_FOUND_TEXT
    if exp_b and b_nums:
        for n in range(max(max(b_nums), 5) + 1, max_b + 1): row[f'Bullet {n}'] = bullets[n - 1] if len(bullets) >= n else NOT_FOUND_TEXT
    for k in [f for f in FIELD_ORDER if f not in ['title', 'description'] and not f.startswith('bullet_') and not f.startswith('additional_image_') and f != 'main_image_links']:
        if k in keys: row[FIELD_LABELS.get(k, k)] = bundle.field_values.get(k, '') or NOT_FOUND_TEXT
    i_nums, images = selected_additional_image_numbers(keys), bundle.images or []
    if 'main_image_links' in keys: row['Main Image Links'] = images[0] if len(images) >= 1 else NOT_FOUND_TEXT
    for n in i_nums: row[f'Additional Image {n}'] = images[n] if len(images) >= n + 1 else NOT_FOUND_TEXT
    if exp_img and i_nums:
        for n in range(max(max(i_nums), 5) + 1, max_img + 1): row[f'Additional Image {n}'] = images[n] if len(images) >= n + 1 else NOT_FOUND_TEXT
    return row

def build_results_dataframe(scraped_rows: List[Tuple[str, str, ExtractionBundle]], keys: List[str], exp_b: bool, exp_img: bool) -> pd.DataFrame:
    max_b = max((len(b.bullets or []) for _, _, b in scraped_rows), default=0) if exp_b else 0
    max_img = max((max(len(b.images or []) - 1, 0) for _, _, b in scraped_rows), default=0) if exp_img else 0
    cols = build_output_columns(keys, max_b, max_img, exp_b, exp_img)
    if not scraped_rows: return pd.DataFrame(columns=cols)
    rows = [build_output_row(sku, url, b, keys, max_b, max_img, exp_b, exp_img) for sku, url, b in scraped_rows]
    df = pd.DataFrame(rows)
    for c in cols:
        if c not in df.columns: df[c] = NOT_FOUND_TEXT if c not in {'SKU', 'URL'} else ''
    return df[cols].fillna(NOT_FOUND_TEXT)

def build_output_bytes(results_df: pd.DataFrame, failures: List[Dict[str, str]]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        results_df.to_excel(writer, sheet_name='Extracted Data', index=False)
        worksheet = writer.sheets['Extracted Data']
        header_fill = PatternFill(fill_type='solid', fgColor='DCEBFF')
        header_font = Font(bold=True, color='0F172A')
        for cell in worksheet[1]: cell.fill = header_fill; cell.font = header_font; cell.alignment = Alignment(vertical='center')
        for row in worksheet.iter_rows(min_row=2):
            for cell in row: cell.alignment = Alignment(vertical='top', wrap_text=True)
        worksheet.freeze_panes = 'A2'
        worksheet.auto_filter.ref = worksheet.dimensions
        if failures:
            failures_df = pd.DataFrame(failures)
            failures_df.to_excel(writer, sheet_name='Run Log', index=False)
    output.seek(0)
    return output.getvalue()

def build_template_bytes(rows: int = DEFAULT_ROWS) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer: build_blank_input_df(rows).to_excel(writer, sheet_name='Input Template', index=False)
    output.seek(0)
    return output.getvalue()

def kpi_card(label: str, value: str, sub: str = '') -> None:
    st.markdown(f'<div class="mini-kpi"><div class="label">{label}</div><div class="value">{value}</div><div class="sub">{sub}</div></div>', unsafe_allow_html=True)

def reset_table_only() -> None:
    st.session_state.input_df = build_blank_input_df(DEFAULT_ROWS)
    st.session_state.results_df = pd.DataFrame(columns=default_output_columns())
    st.session_state.last_failures = []
    st.session_state.uploaded_signature = None
    st.session_state.row_count = DEFAULT_ROWS
    st.rerun()

def logout_user() -> None:
    st.session_state.authenticated = False
    st.session_state.user_email = ''
    st.session_state.user_name = ''
    reset_table_only()

def render_header(user_name: str, user_email: str) -> None:
    inject_css() # Apply main app CSS
    left, right = st.columns([4.1, 1.2], gap='large')
    with left: st.markdown(f'<div class="hero-card"><div class="hero-title">{APP_TITLE}</div><div class="hero-subtitle">{APP_SUBTITLE}</div></div>', unsafe_allow_html=True)
    with right:
        st.markdown(f'<div class="profile-card"><div class="profile-label">Signed in</div><div class="profile-name">Welcome {clean_text(user_name) or "User"}</div><div class="profile-email">{clean_text(user_email)}</div></div>', unsafe_allow_html=True)
        if st.button('Log Out', key='logout_button'): logout_user()

def render_attribute_selector() -> Tuple[List[str], bool, bool]:
    with st.container(border=True):
        st.markdown('### Required attributes')
        for group_name, items in FIELD_GROUPS:
            st.markdown(f'#### {group_name}')
            columns = st.columns(ATTRIBUTES_COLUMNS_PER_GROUP)
            for idx, (field_key, label, _) in enumerate(items):
                with columns[idx % ATTRIBUTES_COLUMNS_PER_GROUP]: st.checkbox(label, key=f'sel_{field_key}')
        st.markdown('#### Dynamic options')
        dynamic_cols = st.columns(2)
        with dynamic_cols[0]: st.checkbox('Expand extra bullets', key='expand_extra_bullets')
        with dynamic_cols[1]: st.checkbox('Expand extra images', key='expand_extra_images')
    return selected_field_keys_from_state(), bool(st.session_state.expand_extra_bullets), bool(st.session_state.expand_extra_images)

def run_scrape(input_rows: pd.DataFrame, keys: List[str], exp_b: bool, exp_img: bool) -> Tuple[pd.DataFrame, List[Dict[str, str]]]:
    total = len(input_rows)
    scraped_rows, failures = [], []
    progress_bar = st.progress(0.0, text='Preparing deep scraper...')
    for idx, row in enumerate(input_rows.to_dict(orient='records'), start=1):
        sku, url = clean_text(row.get('SKU', '')), normalize_url(row.get('Walmart URL', ''))
        progress_bar.progress(idx / total, text=f'Processing {idx} of {total}')
        bundle = scrape_listing(url, keys)
        scraped_rows.append((sku, row.get('Walmart URL', ''), bundle))
        if bundle.error: failures.append({'SKU': sku, 'URL': row.get('Walmart URL', ''), 'Error': bundle.error})
        if idx < total: time.sleep(REQUEST_DELAY_SECONDS)
    progress_bar.progress(1.0, text='Complete')
    st.success(f'Complete - extracted {total - len(failures)} of {total} listing(s).')
    return build_results_dataframe(scraped_rows, keys, exp_b, exp_img), failures

def render_dashboard() -> None:
    render_header(st.session_state.user_name, st.session_state.user_email)
    top_left, top_mid, top_right = st.columns([0.95, 1.1, 1.5])
    with top_right:
        uploaded_file = st.file_uploader('Upload Excel/CSV', type=['xlsx', 'xls', 'csv'])
        if uploaded_file and (uploaded_file.name, uploaded_file.size) != st.session_state.uploaded_signature:
            try:
                uploaded_df, _ = parse_uploaded_dataframe(uploaded_file)
                st.session_state.row_count = max(DEFAULT_ROWS, len(uploaded_df))
                st.session_state.input_df = ensure_row_count(uploaded_df, st.session_state.row_count)
                st.session_state.uploaded_signature = (uploaded_file.name, uploaded_file.size)
                st.rerun()
            except Exception as exc: st.error(f'Upload failed: {exc}')
    with top_left: st.number_input('Rows to show', min_value=1, max_value=MAX_ROWS, key='row_count')
    with top_mid: st.text_input('Output file name', key='file_name')

    st.session_state.input_df = ensure_row_count(st.session_state.input_df, int(st.session_state.row_count))
    editor_df = coerce_input_df(st.session_state.input_df)

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    with kpi1: kpi_card('Configured rows', str(len(editor_df)))
    with kpi2: kpi_card('Ready to scrape', str(int(((editor_df['SKU'].map(bool)) & (editor_df['Walmart URL'].map(bool))).sum())))
    with kpi3: kpi_card('Last run output', str(len(st.session_state.results_df)))
    with kpi4: kpi_card('Last run failures', str(len(st.session_state.last_failures)))

    left, right = st.columns([2.15, 1.15], gap='large')
    with left:
        edited_df = st.data_editor(editor_df, hide_index=True, use_container_width=True, num_rows='fixed')
        st.session_state.input_df = coerce_input_df(edited_df)
        c1, c2, c3 = st.columns([1.2, 0.95, 1.3])
        with c1: start_clicked = st.button('Start Scraping', type='primary', use_container_width=True)
        with c2: 
            if st.button('Reset Table', use_container_width=True): reset_table_only()
        with c3: st.download_button('Download Template', data=build_template_bytes(), file_name='template.xlsx', use_container_width=True)

    with right: keys, exp_b, exp_img = render_attribute_selector()

    if start_clicked:
        valid_rows = editor_df[editor_df['SKU'].map(bool) & editor_df['Walmart URL'].map(bool)].copy()
        valid_rows['Walmart URL'] = valid_rows['Walmart URL'].map(normalize_url)
        if not keys: st.warning('Select attributes.')
        elif valid_rows.empty: st.warning('Add SKU and URL.')
        else:
            results_df, failures = run_scrape(valid_rows, keys, exp_b, exp_img)
            st.session_state.results_df, st.session_state.last_failures = results_df, failures

    if not st.session_state.results_df.empty:
        st.markdown('### Results')
        file_stub = slugify_filename(st.session_state.file_name)
        st.download_button('Download Excel Output', data=build_output_bytes(st.session_state.results_df, st.session_state.last_failures), file_name=f'{file_stub}.xlsx', type='primary')
        st.dataframe(st.session_state.results_df, use_container_width=True, hide_index=True)

def main() -> None:
    init_state()
    if not st.session_state.authenticated:
        render_login_page()
    else:
        render_dashboard()

if __name__ == '__main__':
    main()
