from __future__ import annotations

import io
import json
import re
from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

from auth import authenticate_user, get_display_name, normalize_email
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

APP_TITLE = "Walmart Content Extractor"
APP_SUBTITLE = (
    "Securely extract Walmart product page content into a polished Excel file."
)
DEFAULT_ROWS = 10
MAX_ROWS = 1000
MAX_VISIBLE_ROWS = 20
SCROLL_AFTER_ROWS = 15
TABLE_ROW_HEIGHT = 40
MIN_BULLET_COLUMNS = 5
MAX_EXTRACTED_BULLETS = 20

INPUT_COLUMNS = ["SKU", "Walmart URL"]
BASE_OUTPUT_COLUMNS = ["SKU", "URL", "Title", "Description"]

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
    "Referer": "https://www.walmart.com/",
    "Upgrade-Insecure-Requests": "1",
}

STOP_LINE_PATTERNS = [
    r"^view all item details$",
    r"^specs?$",
    r"^specifications?$",
    r"^how do you want your item\??$",
    r"^current price",
    r"^price when purchased online$",
    r"^sold by$",
    r"^fulfilled by walmart$",
    r"^free 90-day returns$",
    r"^details$",
    r"^more seller options",
    r"^about this item$",
    r"^info:$",
    r"^more details$",
    r"^indications$",
    r"^directions$",
    r"^ingredients$",
    r"^warnings?$",
    r"^customer ratings",
    r"^rating and reviews",
    r"^ratings and reviews",
    r"^product details$",
    r"^customers also considered$",
    r"^you may also like$",
    r"^similar items",
    r"^frequently bought together$",
    r"^recommended for you$",
]

CAPTCHA_PATTERNS = [
    r"verify your identity",
    r"robot or human",
    r"press and hold",
    r"access denied",
    r"captcha",
    r"are you a real person",
]


@dataclass
class ScrapeResult:
    title: str = ""
    description: str = ""
    bullets: List[str] | None = None
    error: str = ""

    def as_row(self, sku: str, url: str) -> Dict[str, str]:
        row = {
            "SKU": sku,
            "URL": url,
            "Title": self.title,
            "Description": self.description,
        }
        for idx, bullet in enumerate(self.bullets or [], start=1):
            row[f"Bullet {idx}"] = clean_text(bullet)
        return row


st.set_page_config(
    page_title=APP_TITLE,
    layout="wide",
    initial_sidebar_state="collapsed",
)


@st.cache_resource(show_spinner=False)
def get_requests_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=2,
        connect=2,
        read=2,
        status=2,
        backoff_factor=0.8,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=frozenset({"GET", "HEAD", "OPTIONS"}),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=20)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update(REQUEST_HEADERS)
    return session


def inject_css() -> None:
    st.markdown(
        """
        <style>
            .stApp {
                background: linear-gradient(180deg, #f5f8ff 0%, #fbfcff 48%, #ffffff 100%);
            }
            .hero-card {
                background: linear-gradient(135deg, #0053e2 0%, #1d8bff 100%);
                color: white;
                border-radius: 22px;
                padding: 28px 30px;
                box-shadow: 0 22px 42px rgba(0, 83, 226, 0.18);
                margin-bottom: 1rem;
            }
            .hero-title {
                font-size: 2rem;
                font-weight: 800;
                margin: 0;
                letter-spacing: 0.2px;
            }
            .hero-subtitle {
                font-size: 1rem;
                margin-top: 0.5rem;
                opacity: 0.96;
                line-height: 1.55;
            }
            .card {
                background: rgba(255, 255, 255, 0.92);
                border: 1px solid rgba(0, 0, 0, 0.06);
                border-radius: 18px;
                padding: 18px 18px 10px 18px;
                box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
            }
            .card h3 {
                margin-top: 0.15rem;
                margin-bottom: 0.2rem;
                font-size: 1.02rem;
            }
            .soft-note {
                color: #5b6472;
                font-size: 0.92rem;
                line-height: 1.5;
            }
            .mini-kpi {
                background: rgba(255, 255, 255, 0.92);
                border: 1px solid rgba(0, 0, 0, 0.05);
                border-radius: 16px;
                padding: 12px 16px;
                box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04);
                min-height: 84px;
            }
            .mini-kpi .label {
                font-size: 0.83rem;
                color: #5b6472;
                margin-bottom: 0.35rem;
            }
            .mini-kpi .value {
                font-size: 1.55rem;
                font-weight: 800;
                color: #0f172a;
            }
            .mini-kpi .sub {
                font-size: 0.82rem;
                color: #64748b;
                margin-top: 0.15rem;
            }
            .table-note {
                font-size: 0.88rem;
                color: #576172;
                margin-top: 0.35rem;
            }
            div[data-testid="stDownloadButton"] button,
            div[data-testid="stButton"] button {
                width: 100%;
                border-radius: 12px;
                font-weight: 700;
                padding-top: 0.65rem;
                padding-bottom: 0.65rem;
            }
            div[data-testid="stDataEditor"] {
                border-radius: 18px;
                overflow: hidden;
            }
            .status-box {
                background: #0f172a;
                color: white;
                border-radius: 16px;
                padding: 14px 16px;
            }
            .status-line {
                font-size: 0.94rem;
                line-height: 1.55;
            }
            .login-top-space {
                height: 7vh;
            }
            .login-card {
                background: rgba(255, 255, 255, 0.96);
                border: 1px solid rgba(0, 0, 0, 0.06);
                border-radius: 22px;
                padding: 28px 26px 18px 26px;
                box-shadow: 0 18px 38px rgba(15, 23, 42, 0.08);
            }
            .login-title {
                font-size: 1.7rem;
                font-weight: 800;
                color: #0f172a;
                margin-bottom: 0.35rem;
            }
            .login-subtitle {
                font-size: 0.96rem;
                color: #5b6472;
                line-height: 1.5;
                margin-bottom: 1rem;
            }
            .login-note {
                font-size: 0.8rem;
                color: #6b7280;
                line-height: 1.5;
                margin-top: 1rem;
                text-align: center;
            }
            .login-footer {
                position: fixed;
                bottom: 14px;
                left: 0;
                right: 0;
                text-align: center;
                font-size: 0.78rem;
                color: #7a8597;
            }
            .profile-card {
                background: rgba(255, 255, 255, 0.92);
                border: 1px solid rgba(0, 0, 0, 0.06);
                border-radius: 18px;
                padding: 16px 16px 12px 16px;
                box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
                margin-bottom: 0.75rem;
            }
            .profile-label {
                color: #64748b;
                font-size: 0.78rem;
                margin-bottom: 0.3rem;
            }
            .profile-name {
                color: #0f172a;
                font-weight: 800;
                font-size: 1.1rem;
                margin-bottom: 0.25rem;
            }
            .profile-email {
                color: #5b6472;
                font-size: 0.82rem;
                word-break: break-word;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def init_state() -> None:
    if "input_df" not in st.session_state:
        st.session_state.input_df = build_blank_input_df(DEFAULT_ROWS)
    if "results_df" not in st.session_state:
        st.session_state.results_df = pd.DataFrame(columns=ordered_output_columns())
    if "run_log" not in st.session_state:
        st.session_state.run_log = []
    if "last_failures" not in st.session_state:
        st.session_state.last_failures = []
    if "uploaded_signature" not in st.session_state:
        st.session_state.uploaded_signature = None
    if "file_name" not in st.session_state:
        st.session_state.file_name = "walmart_content_export"
    if "row_count" not in st.session_state:
        st.session_state.row_count = DEFAULT_ROWS
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "user_email" not in st.session_state:
        st.session_state.user_email = ""
    if "user_name" not in st.session_state:
        st.session_state.user_name = ""


def build_blank_input_df(rows: int) -> pd.DataFrame:
    rows = max(1, min(int(rows), MAX_ROWS))
    return pd.DataFrame({col: [""] * rows for col in INPUT_COLUMNS})


def ensure_row_count(df: pd.DataFrame, rows: int) -> pd.DataFrame:
    rows = max(1, min(int(rows), MAX_ROWS))
    base = df.copy()
    for col in INPUT_COLUMNS:
        if col not in base.columns:
            base[col] = ""
    base = base[INPUT_COLUMNS]
    if len(base) > rows:
        base = base.iloc[:rows].copy()
    elif len(base) < rows:
        extra = pd.DataFrame({col: [""] * (rows - len(base)) for col in INPUT_COLUMNS})
        base = pd.concat([base, extra], ignore_index=True)
    return base.fillna("")


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\u00a0", " ")
    text = text.replace("\u200b", "")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def coerce_input_df(df: pd.DataFrame) -> pd.DataFrame:
    base = df.copy()
    for col in INPUT_COLUMNS:
        if col not in base.columns:
            base[col] = ""
        base[col] = base[col].map(clean_text)
    return base[INPUT_COLUMNS]


def normalize_url(url: str) -> str:
    value = clean_text(url)
    if not value:
        return ""
    if value.startswith("www."):
        value = f"https://{value}"
    elif not re.match(r"^https?://", value, flags=re.I):
        value = f"https://{value}"
    return value


def looks_like_walmart_url(url: str) -> bool:
    return bool(re.search(r"https?://([a-z0-9-]+\.)?walmart\.com/", url, flags=re.I))


def slugify_filename(name: str) -> str:
    cleaned = clean_text(name) or "walmart_content_export"
    cleaned = re.sub(r"[\\/:*?\"<>|]+", "_", cleaned)
    cleaned = re.sub(r"\s+", "_", cleaned)
    cleaned = cleaned.strip("._")
    return cleaned[:120] or "walmart_content_export"


def first_non_empty(values: Iterable[Optional[str]]) -> str:
    for value in values:
        text = clean_text(value)
        if text:
            return text
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


def bullet_columns(count: int = MIN_BULLET_COLUMNS) -> List[str]:
    total = max(MIN_BULLET_COLUMNS, int(count or 0))
    return [f"Bullet {idx}" for idx in range(1, total + 1)]


def ordered_output_columns(bullet_count: int = MIN_BULLET_COLUMNS) -> List[str]:
    return BASE_OUTPUT_COLUMNS + bullet_columns(bullet_count)


def build_results_dataframe(rows_output: List[Dict[str, str]]) -> pd.DataFrame:
    max_bullets = 0
    for row in rows_output:
        for key in row.keys():
            match = re.fullmatch(r"Bullet (\d+)", str(key))
            if match:
                max_bullets = max(max_bullets, int(match.group(1)))
    columns = ordered_output_columns(max_bullets)
    if not rows_output:
        return pd.DataFrame(columns=columns)
    results_df = pd.DataFrame(rows_output)
    for col in columns:
        if col not in results_df.columns:
            results_df[col] = ""
    return results_df[columns].fillna("")


def parse_uploaded_dataframe(uploaded_file) -> Tuple[pd.DataFrame, str]:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, dtype=str, keep_default_na=False)
    else:
        df = pd.read_excel(uploaded_file, dtype=str)

    if df.empty:
        return build_blank_input_df(DEFAULT_ROWS), "Uploaded file is empty."

    normalized = {str(col).strip().lower(): col for col in df.columns}
    sku_col = None
    url_col = None
    for key, original in normalized.items():
        if sku_col is None and "sku" in key:
            sku_col = original
        if url_col is None and ("url" in key or "walmart" in key or "listing" in key):
            url_col = original

    if sku_col is None and len(df.columns) >= 1:
        sku_col = df.columns[0]
    if url_col is None and len(df.columns) >= 2:
        url_col = df.columns[1]

    if sku_col is None or url_col is None:
        raise ValueError("Could not detect SKU and URL columns in the uploaded file.")

    parsed = pd.DataFrame(
        {
            "SKU": df[sku_col].map(clean_text),
            "Walmart URL": df[url_col].map(clean_text),
        }
    )
    parsed = parsed.iloc[:MAX_ROWS].copy()
    parsed = parsed.fillna("")

    message = f"Loaded {len(parsed)} rows from the uploaded file."
    if len(df) > MAX_ROWS:
        message += f" Only the first {MAX_ROWS} rows were kept."
    return parsed, message


def detect_captcha_or_block(lines: List[str], html: str = "") -> bool:
    joined = " ".join(lines[:120]).lower()
    combined = f"{joined} {html[:5000].lower()}"
    return any(re.search(pattern, combined, flags=re.I) for pattern in CAPTCHA_PATTERNS)


def line_matches_any(line: str, patterns: Iterable[str]) -> bool:
    text = clean_text(line)
    return any(re.search(pattern, text, flags=re.I) for pattern in patterns)


def strip_bullet_prefix(line: str) -> str:
    text = clean_text(line)
    text = re.sub(r"^[\u2022\-*\s]+", "", text)
    return text


def looks_like_bullet_line(line: str) -> bool:
    text = strip_bullet_prefix(line)
    if len(text) < 25:
        return False
    if text != clean_text(line):
        return True
    if ":" in text[:45]:
        head = text.split(":", 1)[0].strip()
        letters = re.sub(r"[^A-Za-z]", "", head)
        if letters and (
            head.isupper()
            or sum(ch.isupper() for ch in head) >= max(3, int(len(letters) * 0.45))
        ):
            return True
    return False


def is_generic_section_heading(line: str) -> bool:
    return line_matches_any(
        line,
        [
            r"^key item features$",
            r"^highlights$",
            r"^about this item$",
            r"^product details$",
            r"^description$",
            r"^details$",
            r"^specs?$",
            r"^specifications?$",
            r"^more details$",
        ],
    )


def is_short_label_fragment(line: str) -> bool:
    text = strip_bullet_prefix(line)
    if not text or len(text) > 45:
        return False
    if line_matches_any(text, STOP_LINE_PATTERNS) or is_generic_section_heading(text):
        return False
    if re.search(r"[.!?]$", text):
        return False
    words = re.findall(r"[A-Za-z0-9&'+/\-]+", text)
    return 1 <= len(words) <= 6


def merge_fragmented_bullet_lines(lines: List[str]) -> List[str]:
    merged: List[str] = []
    idx = 0
    while idx < len(lines):
        line = clean_text(lines[idx])
        if not line:
            idx += 1
            continue

        if line.startswith(":") and merged and is_short_label_fragment(merged[-1]):
            merged[-1] = clean_text(f"{merged[-1]}{line}")
            idx += 1
            continue

        next_line = clean_text(lines[idx + 1]) if idx + 1 < len(lines) else ""
        if next_line:
            if is_short_label_fragment(line) and next_line.startswith(":"):
                merged.append(clean_text(f"{line}{next_line}"))
                idx += 2
                continue
            if line.endswith(":") and not line_matches_any(next_line, STOP_LINE_PATTERNS):
                merged.append(clean_text(f"{line} {strip_bullet_prefix(next_line)}"))
                idx += 2
                continue

        merged.append(line)
        idx += 1
    return merged


def collect_section_lines(lines: List[str], start_idx: int, max_scan_lines: int = 50) -> List[str]:
    section: List[str] = []
    for line in lines[start_idx + 1 : start_idx + 1 + max_scan_lines]:
        if line_matches_any(line, STOP_LINE_PATTERNS):
            break
        if re.search(r"accurate product information|see our disclaimer", line, flags=re.I):
            break
        section.append(line)
    return merge_fragmented_bullet_lines(section)


def parse_visible_lines(body_text: str) -> List[str]:
    output: List[str] = []
    for raw in body_text.splitlines():
        line = clean_text(raw)
        if not line:
            continue
        output.append(line)
    return output


def find_heading_index(lines: List[str], patterns: Iterable[str]) -> Optional[int]:
    for idx, line in enumerate(lines):
        if line_matches_any(line, patterns):
            return idx
    return None


def extract_description_from_lines(lines: List[str]) -> str:
    start_idx = find_heading_index(lines, [r"^product details$", r"^description$"])
    if start_idx is None:
        return ""

    pieces: List[str] = []
    for line in collect_section_lines(lines, start_idx, max_scan_lines=40):
        if looks_like_bullet_line(line) or line.startswith(":"):
            break
        if len(line) < 25:
            if pieces:
                break
            continue
        pieces.append(strip_bullet_prefix(line))
        if len(" ".join(pieces)) >= 1200 or len(pieces) >= 4:
            break
    return clean_text(" ".join(dedupe_keep_order(pieces)))


def extract_bullets_from_lines(lines: List[str]) -> List[str]:
    bullets: List[str] = []
    key_idx = find_heading_index(lines, [r"^key item features$", r"^highlights$", r"^about this item$"])
    if key_idx is not None:
        for line in collect_section_lines(lines, key_idx, max_scan_lines=40):
            candidate = strip_bullet_prefix(line)
            if len(candidate) < 12:
                continue
            bullets.append(candidate)
            if len(bullets) >= MAX_EXTRACTED_BULLETS:
                break

    if bullets:
        return dedupe_keep_order(bullets)[:MAX_EXTRACTED_BULLETS]

    detail_idx = find_heading_index(lines, [r"^product details$", r"^description$"])
    if detail_idx is not None:
        description_started = False
        for line in collect_section_lines(lines, detail_idx, max_scan_lines=50):
            candidate = strip_bullet_prefix(line)
            if looks_like_bullet_line(candidate) or (":" in candidate[:80] and len(candidate) >= 20):
                description_started = True
                bullets.append(candidate)
            elif description_started:
                break
            if len(bullets) >= MAX_EXTRACTED_BULLETS:
                break

    return dedupe_keep_order(bullets)[:MAX_EXTRACTED_BULLETS]


def normalize_json_text(value: str) -> str:
    text = clean_text(value)
    if not text:
        return ""
    text = text.replace('\\u003c', '<').replace('\\u003e', '>').replace('\\u0026', '&')
    try:
        text = bytes(text, 'utf-8').decode('unicode_escape')
    except Exception:
        pass
    return clean_text(text)


def extract_product_json_from_soup(soup: BeautifulSoup) -> Dict[str, str]:
    data: Dict[str, str] = {}
    scripts = soup.find_all("script", attrs={"type": "application/ld+json"})

    def visit(obj: Any) -> List[Dict[str, Any]]:
        items: List[Dict[str, Any]] = []
        if isinstance(obj, dict):
            items.append(obj)
            for value in obj.values():
                if isinstance(value, dict):
                    items.extend(visit(value))
                elif isinstance(value, list):
                    for child in value:
                        items.extend(visit(child))
        elif isinstance(obj, list):
            for child in obj:
                items.extend(visit(child))
        return items

    for script in scripts:
        raw = script.string or script.get_text() or ""
        raw = raw.strip()
        if not raw:
            continue
        try:
            parsed = json.loads(raw)
        except Exception:
            continue
        for obj in visit(parsed):
            types = obj.get("@type", [])
            if isinstance(types, str):
                types = [types]
            if any(str(item).lower() == "product" for item in types):
                name = clean_text(obj.get("name", ""))
                description = clean_text(obj.get("description", ""))
                if name and not data.get("title"):
                    data["title"] = name
                if description and not data.get("description"):
                    data["description"] = description
            if data.get("title") and data.get("description"):
                return data
    return data


def extract_embedded_text_candidates(html: str) -> Dict[str, str]:
    candidates: Dict[str, str] = {}
    patterns = {
        "description": [
            r'"shortDescription"\s*:\s*"(.*?)"',
            r'"description"\s*:\s*"(.*?)"',
            r'"productDescription"\s*:\s*"(.*?)"',
        ],
        "title": [
            r'"productName"\s*:\s*"(.*?)"',
            r'"name"\s*:\s*"(.*?)"',
        ],
    }
    for key, regexes in patterns.items():
        for pattern in regexes:
            match = re.search(pattern, html, flags=re.I | re.S)
            if match:
                value = normalize_json_text(match.group(1))
                if value:
                    candidates[key] = value
                    break
    return candidates


def clean_soup_for_text(soup: BeautifulSoup) -> BeautifulSoup:
    cloned = BeautifulSoup(str(soup), "lxml")
    for tag in cloned(["script", "style", "noscript", "svg", "template"]):
        tag.decompose()
    return cloned


def extract_title_from_soup(soup: BeautifulSoup, json_data: Dict[str, str], html_candidates: Dict[str, str]) -> str:
    h1 = soup.find("h1")
    og = soup.find("meta", attrs={"property": "og:title"})
    meta_title = soup.find("meta", attrs={"name": "title"})
    page_title = soup.find("title")

    title = first_non_empty(
        [
            h1.get_text(" ", strip=True) if h1 else "",
            og.get("content", "") if og else "",
            meta_title.get("content", "") if meta_title else "",
            page_title.get_text(" ", strip=True) if page_title else "",
            json_data.get("title", ""),
            html_candidates.get("title", ""),
        ]
    )
    title = re.sub(r"\s*-\s*Walmart\.com\s*$", "", title, flags=re.I)
    return clean_text(title)


def extract_meta_description(soup: BeautifulSoup) -> str:
    tag = soup.find("meta", attrs={"name": "description"})
    if not tag:
        tag = soup.find("meta", attrs={"property": "og:description"})
    if not tag:
        return ""
    return clean_text(tag.get("content", ""))


def scrape_listing_from_html(html: str, resolved_url: str = "") -> ScrapeResult:
    if not clean_text(html):
        return ScrapeResult(error="Empty page received.")

    soup = BeautifulSoup(html, "lxml")
    json_data = extract_product_json_from_soup(soup)
    html_candidates = extract_embedded_text_candidates(html)
    text_soup = clean_soup_for_text(soup)
    body_text = text_soup.get_text("\n")
    lines = parse_visible_lines(body_text)

    if detect_captcha_or_block(lines, html):
        return ScrapeResult(error="Walmart presented an anti-bot or verification page.")

    title = extract_title_from_soup(soup, json_data, html_candidates)
    description = first_non_empty(
        [
            extract_description_from_lines(lines),
            json_data.get("description", ""),
            html_candidates.get("description", ""),
            extract_meta_description(soup),
        ]
    )
    description = re.sub(r"\s+", " ", description).strip()

    bullets = extract_bullets_from_lines(lines)

    if not title and not description and not bullets:
        if resolved_url and not re.search(r"/ip/", resolved_url):
            return ScrapeResult(error="The URL did not resolve to a Walmart product page.")
        return ScrapeResult(error="Could not extract title, description, or bullets from the page.")

    return ScrapeResult(title=title, description=description, bullets=bullets)


def fetch_listing_html(url: str) -> Tuple[str, str]:
    session = get_requests_session()
    try:
        response = session.get(url, timeout=(15, 35), allow_redirects=True)
    except requests.RequestException as exc:
        raise RuntimeError(f"Network request failed: {exc}") from exc

    if response.status_code in {403, 429}:
        raise RuntimeError("Walmart blocked the request or rate-limited the session.")
    if response.status_code == 404:
        raise RuntimeError("Listing page returned 404 Not Found.")
    if response.status_code >= 400:
        raise RuntimeError(f"Listing page returned HTTP {response.status_code}.")

    content_type = response.headers.get("content-type", "")
    if "html" not in content_type.lower():
        raise RuntimeError(f"Unexpected response type: {content_type or 'unknown'}")

    return response.text, response.url


def scrape_listing(url: str) -> ScrapeResult:
    try:
        html, resolved_url = fetch_listing_html(url)
    except Exception as exc:
        return ScrapeResult(error=str(exc))
    return scrape_listing_from_html(html, resolved_url=resolved_url)


def build_output_bytes(results_df: pd.DataFrame, failures: List[Dict[str, str]]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        results_df.to_excel(writer, sheet_name="Extracted Data", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Extracted Data"]
        header_fill = PatternFill(fill_type="solid", fgColor="DCEBFF")
        header_font = Font(bold=True, color="0F172A")
        wrap_alignment = Alignment(vertical="top", wrap_text=True)

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(vertical="center")

        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = wrap_alignment

        width_map = {
            "SKU": 18,
            "URL": 42,
            "Title": 42,
            "Description": 70,
        }
        for col_idx, col_name in enumerate(results_df.columns, start=1):
            width = width_map.get(col_name, 48 if str(col_name).startswith("Bullet ") else min(max(len(str(col_name)) + 2, 18), 60))
            worksheet.column_dimensions[get_column_letter(col_idx)].width = width

        url_col_idx = results_df.columns.get_loc("URL") + 1 if "URL" in results_df.columns else None
        if url_col_idx is not None:
            for row_idx in range(2, worksheet.max_row + 1):
                url_cell = worksheet.cell(row=row_idx, column=url_col_idx)
                if clean_text(url_cell.value):
                    url_cell.hyperlink = clean_text(url_cell.value)
                    url_cell.style = "Hyperlink"
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        if failures:
            failures_df = pd.DataFrame(failures)
            failures_df.to_excel(writer, sheet_name="Run Log", index=False)
            log_sheet = writer.sheets["Run Log"]
            for cell in log_sheet[1]:
                cell.fill = header_fill
                cell.font = header_font
            for col_idx, col_name in enumerate(failures_df.columns, start=1):
                width = min(max(len(col_name) + 2, 18), 60)
                log_sheet.column_dimensions[get_column_letter(col_idx)].width = width
            for row in log_sheet.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = wrap_alignment

        workbook.properties.creator = APP_TITLE

    output.seek(0)
    return output.getvalue()


def build_template_bytes(rows: int = DEFAULT_ROWS) -> bytes:
    template_df = build_blank_input_df(rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        template_df.to_excel(writer, sheet_name="Input Template", index=False)
        worksheet = writer.sheets["Input Template"]
        worksheet.freeze_panes = "A2"
        worksheet.column_dimensions["A"].width = 22
        worksheet.column_dimensions["B"].width = 46
        for cell in worksheet[1]:
            cell.fill = PatternFill(fill_type="solid", fgColor="E7F1FF")
            cell.font = Font(bold=True)
    output.seek(0)
    return output.getvalue()


def kpi_card(label: str, value: str, sub: str = "") -> None:
    st.markdown(
        f"""
        <div class="mini-kpi">
            <div class="label">{label}</div>
            <div class="value">{value}</div>
            <div class="sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def logout_user() -> None:
    st.session_state.authenticated = False
    st.session_state.user_email = ""
    st.session_state.user_name = ""
    st.session_state.input_df = build_blank_input_df(DEFAULT_ROWS)
    st.session_state.results_df = pd.DataFrame(columns=ordered_output_columns())
    st.session_state.last_failures = []
    st.session_state.run_log = []
    st.session_state.uploaded_signature = None
    st.session_state.file_name = "walmart_content_export"
    st.session_state.row_count = DEFAULT_ROWS
    st.rerun()


def render_login_page() -> None:
    st.markdown('<div class="login-top-space"></div>', unsafe_allow_html=True)
    left, center, right = st.columns([1.15, 0.95, 1.15])
    with center:
        st.markdown(
            f"""
            <div class="login-card">
                <div class="login-title">{APP_TITLE}</div>
                <div class="login-subtitle">Only approved @pattern.com users can access this tool</div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("login_form", clear_on_submit=False):
            email = st.text_input("Email Address", placeholder="you@pattern.com")
            password = st.text_input("Password", type="password", placeholder="Enter password")
            submitted = st.form_submit_button("Sign In", use_container_width=True, type="primary")
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
        st.markdown(
            """
                <div class="login-note">New to this page? Please contact Pratik Adsare for creating your login credential</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown(
        '<div class="login-footer">© Designed and Developed by Pratik Adsare</div>',
        unsafe_allow_html=True,
    )


def render_header(user_name: str, user_email: str) -> None:
    left, right = st.columns([4.1, 1.2], gap="large")
    with left:
        st.markdown(
            f"""
            <div class="hero-card">
                <div class="hero-title">{APP_TITLE}</div>
                <div class="hero-subtitle">{APP_SUBTITLE}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with right:
        st.markdown(
            f"""
            <div class="profile-card">
                <div class="profile-label">Signed in</div>
                <div class="profile-name">Welcome {clean_text(user_name) or 'User'}</div>
                <div class="profile-email">{clean_text(user_email)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Log Out", key="logout_button"):
            logout_user()


def run_scrape(input_rows: pd.DataFrame) -> Tuple[pd.DataFrame, List[Dict[str, str]]]:
    total = len(input_rows)
    progress_box = st.container()
    rows_output: List[Dict[str, str]] = []
    failures: List[Dict[str, str]] = []
    status_placeholder = progress_box.empty()
    progress_bar = progress_box.progress(0.0, text="Preparing scraper...")
    live_log = progress_box.empty()
    live_lines: List[str] = []

    for idx, row in enumerate(input_rows.to_dict(orient="records"), start=1):
        sku = clean_text(row.get("SKU", ""))
        display_url = clean_text(row.get("Walmart URL", ""))
        url = normalize_url(display_url)

        status_placeholder.markdown(
            f"""
            <div class="status-box">
                <div class="status-line"><strong>Processing {idx} of {total}</strong></div>
                <div class="status-line">SKU: <code>{sku or '—'}</code></div>
                <div class="status-line">URL: {display_url or '—'}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        progress_bar.progress(idx / total, text=f"Processing {idx} of {total}")

        result = scrape_listing(url)
        rows_output.append(result.as_row(sku=sku, url=display_url))
        if result.error:
            failures.append({"SKU": sku, "URL": display_url, "Error": result.error})
            live_lines.append(f"{idx}. {sku or '[no SKU]'} - failed: {result.error}")
        else:
            live_lines.append(f"{idx}. {sku or '[no SKU]'} - extracted successfully")
        live_log.code("\n".join(live_lines[-10:]), language="text")

    progress_bar.progress(1.0, text="Complete")
    success_count = total - len(failures)
    status_placeholder.success(f"Complete - extracted {success_count} of {total} listing(s).")
    results_df = build_results_dataframe(rows_output)
    return results_df, failures


def render_dashboard() -> None:
    render_header(st.session_state.user_name, st.session_state.user_email)

    controls_left, controls_mid, controls_right = st.columns([0.95, 1.1, 1.5])
    with controls_right:
        uploaded_file = st.file_uploader(
            "Optional: upload Excel/CSV with SKU and URL columns",
            type=["xlsx", "xls", "csv"],
            help="If you upload a file, the editor below will be filled automatically.",
        )

    if uploaded_file is not None:
        signature = (uploaded_file.name, uploaded_file.size)
        if signature != st.session_state.uploaded_signature:
            try:
                uploaded_df, upload_message = parse_uploaded_dataframe(uploaded_file)
                desired_rows = max(DEFAULT_ROWS, len(uploaded_df))
                st.session_state.row_count = desired_rows
                st.session_state.input_df = ensure_row_count(uploaded_df, desired_rows)
                st.session_state.uploaded_signature = signature
                st.success(upload_message)
            except Exception as exc:
                st.error(f"Could not read the uploaded file: {exc}")

    with controls_left:
        st.number_input(
            "Rows to show",
            min_value=1,
            max_value=MAX_ROWS,
            step=1,
            key="row_count",
            help="Default is 10. You can scale this up to 1000 rows.",
        )
    with controls_mid:
        file_name = st.text_input(
            "Output file name",
            value=st.session_state.file_name,
            help="This name will be used for the downloaded Excel file.",
        )
        st.session_state.file_name = file_name

    st.session_state.input_df = ensure_row_count(st.session_state.input_df, int(st.session_state.row_count))
    editor_df = coerce_input_df(st.session_state.input_df)

    ready_rows = int(((editor_df["SKU"].map(bool)) & (editor_df["Walmart URL"].map(bool))).sum())
    completed_rows = len(st.session_state.results_df)
    failed_rows = len(st.session_state.last_failures)

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    with kpi1:
        kpi_card("Configured rows", str(len(editor_df)), f"Default {DEFAULT_ROWS} · max {MAX_ROWS}")
    with kpi2:
        kpi_card("Ready to scrape", str(ready_rows), "Rows with both SKU and Walmart URL")
    with kpi3:
        kpi_card("Last run output", str(completed_rows), "Rows exported in the last scrape")
    with kpi4:
        kpi_card("Last run failures", str(failed_rows), "See Run Log sheet if any rows fail")

    left, right = st.columns([2.2, 1.1], gap="large")

    with left:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("### Input grid")
        st.caption(
            "Paste directly from Excel or type manually. The table stays compact and becomes scrollable for larger batches."
        )
        if len(editor_df) <= SCROLL_AFTER_ROWS:
            visible_rows = len(editor_df)
        elif len(editor_df) <= MAX_VISIBLE_ROWS:
            visible_rows = SCROLL_AFTER_ROWS
        else:
            visible_rows = MAX_VISIBLE_ROWS
        editor_height = 46 + max(1, visible_rows) * TABLE_ROW_HEIGHT

        edited_df = st.data_editor(
            editor_df,
            hide_index=True,
            height=editor_height,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "SKU": st.column_config.TextColumn(
                    "SKU",
                    width="medium",
                    help="Your internal SKU or item identifier.",
                    required=False,
                ),
                "Walmart URL": st.column_config.TextColumn(
                    "Walmart URL",
                    width="large",
                    help="Paste a Walmart product page URL.",
                    required=False,
                ),
            },
            key="input_editor",
        )
        st.session_state.input_df = coerce_input_df(edited_df)
        if len(editor_df) > SCROLL_AFTER_ROWS:
            st.markdown(
                (
                    f'<div class="table-note">Showing a compact scroll area for {len(editor_df)} rows. '
                    "Scroll inside the grid to review everything.</div>"
                ),
                unsafe_allow_html=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("### Actions")
        start_clicked = st.button("Start Scraping", type="primary")
        if st.button("Reset Table"):
            st.session_state.input_df = build_blank_input_df(DEFAULT_ROWS)
            st.session_state.results_df = pd.DataFrame(columns=ordered_output_columns())
            st.session_state.last_failures = []
            st.session_state.run_log = []
            st.session_state.uploaded_signature = None
            st.session_state.row_count = DEFAULT_ROWS
            st.rerun()

        template_name = f"walmart_input_template_{DEFAULT_ROWS}_rows.xlsx"
        st.download_button(
            label="Download Blank Input Template",
            data=build_template_bytes(DEFAULT_ROWS),
            file_name=template_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown("### Notes")
        st.markdown(
            """
            <div class="soft-note">
                • Paste up to 1000 SKU + URL rows.<br>
                • Output columns: SKU, URL, Title, Description, and Bullet 1 onward.<br>
                • The export always includes Bullet 1 to Bullet 5, and adds Bullet 6, 7, 8, and beyond when they are found.<br>
                • Failed rows are recorded on a separate <strong>Run Log</strong> sheet in the output file.<br>
                • This cloud-ready version uses direct page requests instead of a browser runtime, which makes deployment simpler on Streamlit Cloud.
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)

    if start_clicked:
        current_input = coerce_input_df(st.session_state.input_df)
        current_input["Walmart URL"] = current_input["Walmart URL"].map(normalize_url)
        ready_mask = current_input["SKU"].map(bool) & current_input["Walmart URL"].map(bool)
        valid_rows = current_input.loc[ready_mask].copy()

        invalid_url_mask = ~valid_rows["Walmart URL"].map(looks_like_walmart_url)
        if invalid_url_mask.any():
            bad_urls = valid_rows.loc[invalid_url_mask, "Walmart URL"].tolist()[:3]
            st.error(
                "Only Walmart product URLs are supported. Please fix these example URL(s): "
                + ", ".join(bad_urls)
            )
        elif valid_rows.empty:
            st.warning("Add at least one row with both SKU and Walmart URL before starting.")
        else:
            incomplete_count = len(
                current_input[(current_input["SKU"].map(bool) ^ current_input["Walmart URL"].map(bool))]
            )
            if incomplete_count:
                st.info(f"Ignoring {incomplete_count} incomplete row(s) that do not have both SKU and URL.")
            results_df, failures = run_scrape(valid_rows)
            st.session_state.results_df = results_df
            st.session_state.last_failures = failures

    if not st.session_state.results_df.empty:
        st.markdown("### Results")
        result_cols = st.columns([1.2, 1.2, 1.1])
        with result_cols[0]:
            st.success(f"Rows ready for export: {len(st.session_state.results_df)}")
        with result_cols[1]:
            st.info(f"Failures recorded: {len(st.session_state.last_failures)}")
        with result_cols[2]:
            file_stub = slugify_filename(st.session_state.file_name)
            st.download_button(
                label="Download Excel Output",
                data=build_output_bytes(st.session_state.results_df, st.session_state.last_failures),
                file_name=f"{file_stub}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

        preview_tabs = st.tabs(["Output Preview", "Failures"])
        with preview_tabs[0]:
            st.dataframe(
                st.session_state.results_df,
                use_container_width=True,
                hide_index=True,
                height=min(520, 46 + min(len(st.session_state.results_df), 10) * 40),
            )
        with preview_tabs[1]:
            if st.session_state.last_failures:
                st.dataframe(
                    pd.DataFrame(st.session_state.last_failures),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.success("No failures in the last run.")


def main() -> None:
    inject_css()
    init_state()
    if not st.session_state.authenticated:
        render_login_page()
        return
    render_dashboard()


if __name__ == "__main__":
    main()
