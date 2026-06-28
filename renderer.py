"""
renderer.py — NextGen VIP Onboarding email renderer.

All placeholders are matched using the EXACT bytes found in the source HTML files,
which include non-ASCII characters (curly apostrophe U+2019, zero-width space U+200B).
"""
import base64
import re
import logging
from pathlib import Path

log = logging.getLogger(__name__)

# Absolute path to the template root
TEMPLATE_BASE = (
    Path(__file__).parent
    / "templates"
    / "NextGen VIP Onboarding_HTML"
    / "NextGen VIP Onboarding_HTML"
)

SUPPORTED_LANGS = ("ENG", "BM")

# ---------------------------------------------------------------------------
# Field definitions (single source of truth — drives UI form + API response)
# ---------------------------------------------------------------------------

FIELDS = [
    # ── Customer ──────────────────────────────────────────────────────────────
    {
        "group": "customer",
        "name": "customer_name",
        "label": "Customer Full Name",
        "label_bm": "Nama Penuh Pelanggan",
        "type": "text",
        "placeholder": "e.g. Ahmad bin Abdullah",
        "required": True,
    },
    {
        "group": "customer",
        "name": "customer_phone",
        "label": "Customer Phone Number",
        "label_bm": "Nombor Telefon Pelanggan",
        "type": "text",
        "placeholder": "e.g. 012-3456789",
        "required": False,
    },
    {
        "group": "customer",
        "name": "home_office",
        "label": "Location Type",
        "label_bm": "Jenis Lokasi",
        "type": "select",
        "options": [
            {"value": "home", "label": "Home / Rumah"},
            {"value": "office", "label": "Office / Pejabat"},
        ],
        "required": False,
    },
    # ── Service Details ────────────────────────────────────────────────────────
    {
        "group": "service",
        "name": "unifi_login",
        "label": "Unifi Login ID",
        "label_bm": "ID Log Masuk Unifi",
        "type": "text",
        "placeholder": "e.g. ahmad123@unifi",
        "required": False,
    },
    {
        "group": "service",
        "name": "package",
        "label": "Package",
        "label_bm": "Pakej",
        "type": "text",
        "placeholder": "e.g. Unifi 800Mbps",
        "required": False,
    },
    {
        "group": "service",
        "name": "address",
        "label": "Service Address",
        "label_bm": "Alamat Perkhidmatan",
        "type": "text",
        "placeholder": "e.g. No. 1, Jalan ABC, 50000 KL",
        "required": False,
    },
    {
        "group": "service",
        "name": "connection",
        "label": "Connection Quality",
        "label_bm": "Kualiti Sambungan",
        "type": "text",
        "placeholder": "e.g. Stable / Good",
        "required": False,
    },
    {
        "group": "service",
        "name": "router",
        "label": "Router Model",
        "label_bm": "Model Router",
        "type": "text",
        "placeholder": "e.g. Huawei HG8145V5",
        "required": False,
    },
    {
        "group": "service",
        "name": "internetusage",
        "label": "Internet Usage",
        "label_bm": "Penggunaan Internet",
        "type": "text",
        "placeholder": "e.g. High / 95% utilised",
        "required": False,
    },
    # ── NEXTGen Crew (Sender) ──────────────────────────────────────────────────
    {
        "group": "crew",
        "name": "crew_name",
        "label": "NEXTGen Crew Full Name",
        "label_bm": "Nama Penuh NEXTGen Crew",
        "type": "text",
        "placeholder": "e.g. Siti binti Ahmad",
        "required": True,
    },
    {
        "group": "crew",
        "name": "crew_phone",
        "label": "Crew Phone / Contact",
        "label_bm": "Telefon / Hubungi Crew",
        "type": "text",
        "placeholder": "e.g. 012-3456789",
        "required": False,
    },
]


def get_fields() -> list:
    return FIELDS


# ---------------------------------------------------------------------------
# Substitution map
# Defined here as constants so the exact bytes are explicit and auditable.
# Key findings from audit.py:
#   - ENG customer name placeholder uses U+2019 RIGHT SINGLE QUOTATION MARK
#   - Both xxxxx@unifi occurrences are followed by U+200B ZERO WIDTH SPACE
# ---------------------------------------------------------------------------

# U+2019 RIGHT SINGLE QUOTATION MARK  (e2 80 99 in UTF-8)
_RSQM = "\u2019"
# U+200B ZERO WIDTH SPACE  (e2 80 8b in UTF-8)
_ZWS = "\u200b"

# ENG placeholders (exact bytes from the source file)
_ENG_CUSTOMER_NAME  = "{{Customer" + _RSQM + "s Full Name}}"   # {{Customer's Full Name}}
_ENG_HOME_OFFICE    = "{{home/office}}"

# BM placeholders (plain ASCII-equivalent UTF-8)
_BM_CUSTOMER_NAME   = "{{Nama Penuh Pelanggan}}"
_BM_HOME_OFFICE     = "{{rumah/pejabat}}"

# Shared placeholders
_DOLLAR_FIELDS      = ["package", "address", "connection", "router", "internetusage"]
_CREW_NAME_ENG      = "[NEXTGen Full Name]"
_CREW_NAME_BM_ALT   = "[NEXTGen Crew Full Name]"
_CREW_PHONE_TAG     = "[Hand phone Contact Information]"
_UNIFI_LOGIN        = "xxxxx@unifi" + _ZWS   # template appends a zero-width space
_UNIFI_LOGIN_PLAIN  = "xxxxx@unifi"           # fallback without ZWS
_CUSTOMER_PHONE_PAT = ">XXXXXX<"              # surrounded by HTML tags in template


def _load_html(lang: str) -> str:
    lang = lang.upper()
    if lang not in SUPPORTED_LANGS:
        raise ValueError(f"Unsupported language: '{lang}'. Choose from {SUPPORTED_LANGS}")
    path = TEMPLATE_BASE / lang / "index.html"
    if not path.exists():
        raise FileNotFoundError(f"Template not found: {path}")
    return path.read_text(encoding="utf-8")


# Module-level image cache: (lang_upper, img_name) -> full data-URI src attribute string
_IMAGE_CACHE: dict = {}


def _embed_images(html: str, lang: str) -> str:
    """Replace relative `images/xxx` src paths with base64 data URIs (cached per server run)."""
    lang = lang.upper()
    img_dir = TEMPLATE_BASE / lang / "images"
    _mime_map = {
        "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
        "gif": "image/gif", "svg": "image/svg+xml", "webp": "image/webp",
    }

    def _replace(match):
        src = match.group(1)
        if not src.startswith("images/"):
            return match.group(0)
        img_name = src[len("images/"):]
        cache_key = (lang, img_name)
        if cache_key not in _IMAGE_CACHE:
            img_path = img_dir / img_name
            if not img_path.exists():
                log.warning("Image not found: %s", img_path)
                _IMAGE_CACHE[cache_key] = match.group(0)
            else:
                ext  = img_path.suffix.lower().lstrip(".")
                mime = _mime_map.get(ext, "image/png")
                b64  = base64.b64encode(img_path.read_bytes()).decode()
                _IMAGE_CACHE[cache_key] = f'src="data:{mime};base64,{b64}"'
        return _IMAGE_CACHE[cache_key]

    return re.sub(r'src="([^"]*)"', _replace, html)


def _substitute(html: str, values: dict, lang: str) -> str:
    """
    Replace every dynamic placeholder in the HTML template.
    All substitutions are idempotent and logged when the source value is empty.
    """
    lang = lang.upper()
    v = {k: (str(v).strip() if v else "") for k, v in values.items()}

    def _rep(original: str, replacement: str, label: str) -> str:
        if not replacement:
            log.warning("Empty value for '%s' — placeholder '%s' left as-is", label, original)
        return html.replace(original, replacement)

    # 1. $(variable) fields — e.g. $(package)
    for key in _DOLLAR_FIELDS:
        html = html.replace(f"$({key})", v.get(key, ""))

    # 2. Customer name  (NOTE: ENG uses curly apostrophe U+2019)
    customer_name = v.get("customer_name", "")
    if lang == "ENG":
        html = html.replace(_ENG_CUSTOMER_NAME, customer_name)
    else:
        html = html.replace(_BM_CUSTOMER_NAME, customer_name)

    # 3. Home / office
    loc = v.get("home_office", "home")
    if lang == "ENG":
        html = html.replace(_ENG_HOME_OFFICE, "home" if loc == "home" else "office")
    else:
        html = html.replace(_BM_HOME_OFFICE,  "rumah" if loc == "home" else "pejabat")

    # 4. NEXTGen crew name (appears in body paragraph AND sign-off)
    crew_name = v.get("crew_name", "")
    html = html.replace(_CREW_NAME_ENG,    crew_name)
    html = html.replace(_CREW_NAME_BM_ALT, crew_name)

    # 5. NEXTGen crew phone / contact
    html = html.replace(_CREW_PHONE_TAG, v.get("crew_phone", ""))

    # 6. Unifi login ID  (template has a zero-width space appended)
    unifi_login = v.get("unifi_login", "")
    html = html.replace(_UNIFI_LOGIN,       unifi_login)
    html = html.replace(_UNIFI_LOGIN_PLAIN, unifi_login)   # safety fallback

    # 7. Customer phone number  (appears as <strong>XXXXXX</strong>)
    customer_phone = v.get("customer_phone", "")
    html = html.replace(_CUSTOMER_PHONE_PAT, f">{customer_phone}<")

    # 8. Audit: warn about any remaining unresolved placeholders
    remaining_dbl = re.findall(r"\{\{[^}]+\}\}", html)
    remaining_dol = re.findall(r"\$\([^)]+\)", html)
    if remaining_dbl:
        log.warning("[%s] Unresolved {{...}} placeholders: %s", lang, remaining_dbl)
    if remaining_dol:
        log.warning("[%s] Unresolved $(...) placeholders: %s", lang, remaining_dol)

    return html


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def render(values: dict, lang: str = "ENG") -> str:
    """
    Render the NextGen VIP Onboarding email as a self-contained HTML string.

    Args:
        values:  Dict of field values keyed by field name (see FIELDS).
        lang:    "ENG" or "BM".

    Returns:
        Complete HTML string with all placeholders substituted and all images
        embedded as base64 data URIs.
    """
    lang = lang.upper()
    html = _load_html(lang)
    html = _substitute(html, values, lang)
    html = _embed_images(html, lang)
    return html
