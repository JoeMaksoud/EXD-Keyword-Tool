import streamlit as st
import google.genai as genai
import requests
import csv
import json
import re
import io
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="EXD Keyword Research Tool", page_icon="⚡", layout="centered")

# ── Password protection ───────────────────────────────────────────
def check_password():
    try:
        app_password = st.secrets.get("APP_PASSWORD", "") or os.environ.get("APP_PASSWORD", "")
    except:
        app_password = os.environ.get("APP_PASSWORD", "")

    if not app_password:
        return True  # No password set — allow access

    if st.session_state.get("authenticated"):
        return True

    st.markdown("""
    <style>
    .login-box { max-width: 380px; margin: 6rem auto; text-align: center; }
    .login-box h2 { font-size: 22px; font-weight: 700; color: #fff; margin-bottom: 0.5rem; }
    .login-box p { font-size: 14px; color: #888; margin-bottom: 2rem; }
    </style>
    <div class="login-box">
        <div style="font-size:36px;margin-bottom:1rem">⚡</div>
        <h2>EXD Keyword Research Tool</h2>
        <p>Enter your password to continue</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        pwd = st.text_input("Password", type="password", label_visibility="collapsed",
                            placeholder="Enter password...")
        if st.button("Login", use_container_width=True):
            if pwd == app_password:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password. Please try again.")
    return False

if not check_password():
    st.stop()
# ─────────────────────────────────────────────────────────────────

st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
  html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
  .block-container { max-width: 800px; padding: 2rem 2rem 4rem; }
  .section-hdr { display:flex; align-items:center; gap:12px; margin:2.5rem 0 1rem; }
  .section-num {
    min-width:28px; height:28px; border-radius:50%;
    background:#2a2a2a; color:#fff; font-size:13px; font-weight:600;
    display:flex; align-items:center; justify-content:center;
  }
  .section-ttl { font-size:15px; font-weight:600; color:#fff; }
  .section-sub { font-size:13px; color:#666; font-weight:400; }
  div[data-testid="stButton"] > button {
    background:#1e1e1e !important; border:1.5px solid #333 !important;
    border-radius:10px !important; color:#fff !important;
    font-weight:500 !important; transition:border-color 0.15s !important;
  }
  div[data-testid="stButton"] > button:hover {
    border-color:#f97316 !important; background:#1e1e1e !important; color:#fff !important;
  }
  .gen-btn div[data-testid="stButton"] > button {
    background:#f97316 !important; border:none !important;
    border-radius:10px !important; font-size:15px !important;
    font-weight:600 !important; color:#fff !important; padding:0.75rem !important;
  }
  .gen-btn div[data-testid="stButton"] > button:hover { background:#ea6c0a !important; }
  .add-btn div[data-testid="stButton"] > button {
    background:#f97316 !important; border:none !important;
    border-radius:8px !important; color:#fff !important; font-weight:600 !important;
  }
  .add-btn div[data-testid="stButton"] > button:hover { background:#ea6c0a !important; }
  div[data-testid="stDownloadButton"] > button {
    background:#f97316 !important; color:#fff !important;
    border:none !important; border-radius:10px !important; font-weight:600 !important;
  }
  div[data-testid="stDownloadButton"] > button:hover { background:#ea6c0a !important; }
  .tag-pill {
    display:inline-flex; align-items:center; gap:6px;
    background:#1e1e1e; border:1px solid #444; border-radius:20px;
    padding:5px 14px; font-size:13px; color:#ddd; margin:3px;
    cursor:pointer;
  }
  hr { border-color:#2a2a2a !important; }
  #MainMenu, footer { visibility:hidden; }
</style>
""", unsafe_allow_html=True)

DFS_LOCATIONS = {
    "Saudi Arabia": "Saudi Arabia",
    "UAE":          "United Arab Emirates",
    "Egypt":        "Egypt",
    "Kuwait":       "Kuwait",
    "Qatar":        "Qatar",
    "Bahrain":      "Bahrain",
    "Oman":         "Oman",
    "Jordan":       "Jordan",
    "Lebanon":      "Lebanon",
    "United Kingdom": "United Kingdom",
    "United States":  "United States",
    "France":       "France",
    "Germany":      "Germany",
    "India":        "India",
    "China":        "China",
    "Japan":        "Japan",
}
MARKETS = {
    "Saudi Arabia":"sa","UAE":"ae","Egypt":"eg","Kuwait":"kw","Qatar":"qa",
    "Bahrain":"bh","Oman":"om","Jordan":"jo","Lebanon":"lb",
    "United Kingdom":"uk","United States":"us","France":"fr",
    "Germany":"de","India":"in","China":"cn","Japan":"jp"
}
LANGUAGES = [
    {"code":"en","name":"English", "native":"English",  "flag":"🇬🇧"},
    {"code":"ar","name":"Arabic",  "native":"العربية",   "flag":"🇸🇦"},
    {"code":"fr","name":"French",  "native":"Français",  "flag":"🇫🇷"},
    {"code":"de","name":"German",  "native":"Deutsch",   "flag":"🇩🇪"},
    {"code":"hi","name":"Hindi",   "native":"हिन्दी",    "flag":"🇮🇳"},
    {"code":"zh","name":"Chinese", "native":"简体中文",  "flag":"🇨🇳"},
    {"code":"ja","name":"Japanese","native":"日本語",    "flag":"🇯🇵"},
]
QTYPES = [
    {"code":"direct",   "title":"Direct",   "desc":"Short, high-intent. 1–3 words."},
    {"code":"longtail", "title":"Long-tail", "desc":"Specific phrases. 4–8 words."},
    {"code":"mix",      "title":"Mix",       "desc":"Balanced blend of both."},
]
QTYPE_INSTRUCTIONS = {
    "direct":   "Generate SHORT HIGH-INTENT keywords only. Each must be 1-3 words maximum.",
    "longtail": "Generate LONG-TAIL keywords only. Each must be 4-8 words. Focus on specific questions and use-cases.",
    "mix":      "Generate a MIX: roughly half 1-3 words (direct) and half 4-8 words (long-tail).",
}

if "nav_tags"       not in st.session_state: st.session_state.nav_tags = []
if "selected_langs" not in st.session_state: st.session_state.selected_langs = ["en"]
if "selected_qt"    not in st.session_state: st.session_state.selected_qt = "mix"

try:
    gemini_key   = st.secrets.get("GEMINI_KEY", "")   or os.environ.get("GEMINI_KEY", "")
    dfs_login    = st.secrets.get("DFS_LOGIN", "")    or os.environ.get("DFS_LOGIN", "")
    dfs_password = st.secrets.get("DFS_PASSWORD", "") or os.environ.get("DFS_PASSWORD", "")
except:
    gemini_key   = os.environ.get("GEMINI_KEY", "")
    dfs_login    = os.environ.get("DFS_LOGIN", "")
    dfs_password = os.environ.get("DFS_PASSWORD", "")

with st.sidebar:
    st.markdown("## ⚡ EXD Keywords")
    st.markdown("Generate pitch keyword lists in seconds.")
    st.divider()
    st.markdown("**How to use**")
    st.caption("1. Fill in account setup")
    st.caption("2. Select languages")
    st.caption("3. Choose keyword type")
    st.caption("4. Add tags")
    st.caption("5. Click Generate")
    st.caption("6. Download Excel or CSV")
    st.divider()
    st.markdown("**API Keys**")
    if gemini_key:
        st.success("Gemini ✓")
    else:
        gemini_key = st.text_input("Gemini API key", type="password")
    st.caption("DataForSEO credentials (optional — for search volumes)")
    if dfs_login and dfs_password:
        st.success("DataForSEO ✓")
        # Allow override in case secrets are wrong
        override = st.checkbox("Override credentials")
        if override:
            dfs_login    = st.text_input("DataForSEO login (email)", value=dfs_login, key="dfs_login_input")
            dfs_password = st.text_input("DataForSEO password", type="password", key="dfs_pass_input")
    else:
        dfs_login    = st.text_input("DataForSEO login (email)", key="dfs_login_input")
        dfs_password = st.text_input("DataForSEO password", type="password", key="dfs_pass_input")

st.markdown("# ⚡ EXD Keyword Research Tool")
st.markdown("Generate, validate, and export pitch keyword lists in seconds.")
st.divider()

# SECTION 1
st.markdown('<div class="section-hdr"><div class="section-num">1</div><div class="section-ttl">Account Setup</div></div>', unsafe_allow_html=True)
c1, c2 = st.columns(2)
with c1:
    client_name  = st.text_input("Client name", placeholder="e.g. Al Shifa Honey")
with c2:
    market_label = st.selectbox("Target market", list(MARKETS.keys()), index=1)
website = st.text_input("Client website", placeholder="https://www.alshifa.com  (optional — validates keyword relevance)")
c3, c4 = st.columns(2)
with c3:
    branded_count = st.number_input("Branded keywords", min_value=0, max_value=1000, value=10, step=1)
with c4:
    generic_count = st.number_input("Generic keywords", min_value=0, max_value=1000, value=30, step=1)

# SECTION 2
st.markdown('<div class="section-hdr"><div class="section-num">2</div><div class="section-ttl">Languages</div></div>', unsafe_allow_html=True)
st.caption("Click to select — multiple languages supported.")
lang_cols = st.columns(len(LANGUAGES))
for i, lang in enumerate(LANGUAGES):
    with lang_cols[i]:
        is_active = lang["code"] in st.session_state.selected_langs
        border = "#f97316" if is_active else "#333"
        tick = "✓" if is_active else ""
        st.markdown(f"""<div style="background:#1e1e1e;border:1.5px solid {border};
            border-radius:10px 10px 0 0;padding:10px 4px 6px;text-align:center;">
            <div style="font-size:20px">{lang['flag']}</div>
            <div style="font-size:11px;font-weight:600;color:#fff;margin-top:3px">{lang['name']}</div>
            <div style="font-size:10px;color:#666">{lang['native']}</div>
            <div style="font-size:11px;color:#f97316;font-weight:700;min-height:14px;margin-top:3px">{tick}</div>
        </div>""", unsafe_allow_html=True)
        btn_bg = "#f97316" if is_active else "#2a2a2a"
        st.markdown(f"""<style>div[data-testid="column"]:nth-of-type({i+1}) div[data-testid="stButton"] > button {{
            background:{btn_bg} !important; border:1.5px solid {border} !important;
            border-top:none !important; border-radius:0 0 10px 10px !important;
            color:#fff !important; font-size:12px !important; font-weight:400 !important;
            padding:4px !important; width:100% !important;
        }}</style>""", unsafe_allow_html=True)
        if st.button("Select", key=f"lang_{lang['code']}"):
            if lang["code"] in st.session_state.selected_langs:
                if len(st.session_state.selected_langs) > 1:
                    st.session_state.selected_langs.remove(lang["code"])
            else:
                st.session_state.selected_langs.append(lang["code"])
            st.rerun()
selected_lang_names = [l["name"] for l in LANGUAGES if l["code"] in st.session_state.selected_langs]
if len(selected_lang_names) > 1:
    st.caption(f"✓ {len(selected_lang_names)} languages: {', '.join(selected_lang_names)}")

# SECTION 3
st.markdown('<div class="section-hdr"><div class="section-num">3</div><div class="section-ttl">Keyword Type</div></div>', unsafe_allow_html=True)
qt_cols = st.columns(3)
for i, qt in enumerate(QTYPES):
    with qt_cols[i]:
        is_active = st.session_state.selected_qt == qt["code"]
        border = "#f97316" if is_active else "#333"
        tick = " ✓" if is_active else ""
        st.markdown(f"""<div style="background:#1e1e1e;border:1.5px solid {border};
            border-radius:10px 10px 0 0;padding:14px 12px 8px;min-height:75px;">
            <div style="font-size:13px;font-weight:600;color:#fff">{qt['title']}<span style="color:#f97316">{tick}</span></div>
            <div style="font-size:11px;color:#888;margin-top:5px;line-height:1.4">{qt['desc']}</div>
        </div>""", unsafe_allow_html=True)
        btn_bg = "#f97316" if is_active else "#2a2a2a"
        st.markdown(f"""<style>div[data-testid="column"]:nth-of-type({i+1}) div[data-testid="stButton"] > button {{
            background:{btn_bg} !important; border:1.5px solid {border} !important;
            border-top:none !important; border-radius:0 0 10px 10px !important;
            color:#fff !important; font-size:12px !important; font-weight:400 !important;
            padding:4px !important; width:100% !important;
        }}</style>""", unsafe_allow_html=True)
        if st.button("Select", key=f"qt_{qt['code']}"):
            st.session_state.selected_qt = qt["code"]
            st.rerun()

# SECTION 4 — Tags
st.markdown('<div class="section-hdr"><div class="section-num">4</div><div class="section-ttl">Tags</div></div>', unsafe_allow_html=True)

st.markdown("""<style>
div[data-testid="stForm"] { border: none !important; padding: 0 !important; }
div[data-testid="stForm"] div[data-testid="stButton"] > button {
    background:#f97316 !important; border:none !important;
    border-radius:8px !important; color:#fff !important;
    font-weight:600 !important; width:100% !important;
}
div[data-testid="stForm"] div[data-testid="stButton"] > button:hover { background:#ea6c0a !important; }
</style>""", unsafe_allow_html=True)

with st.form(key="tag_form", clear_on_submit=True):
    tc, bc = st.columns([5, 1])
    with tc:
        new_tag = st.text_input("tag", placeholder="e.g. Honey, Royal Jelly...", label_visibility="collapsed")
    with bc:
        add_clicked = st.form_submit_button("+ Add")
    if add_clicked and new_tag.strip():
        tag_val = new_tag.strip()
        if tag_val not in st.session_state.nav_tags:
            st.session_state.nav_tags.append(tag_val)
        st.rerun()

if st.session_state.nav_tags:
    pills = "".join(
        f'<span style="display:inline-flex;align-items:center;gap:6px;background:#1e1e1e;'
        f'border:1px solid #444;border-radius:20px;padding:5px 14px;font-size:13px;'
        f'color:#ddd;margin:3px 3px;">{tag}</span>'
        for tag in st.session_state.nav_tags
    )
    st.markdown(f'<div style="display:flex;flex-wrap:wrap;margin-top:6px;">{pills}</div>', unsafe_allow_html=True)
    st.markdown("<p style='font-size:12px;color:#666;margin:10px 0 4px;'>Click to remove:</p>", unsafe_allow_html=True)
    st.markdown("""<style>
    div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] > button {
        background:transparent !important; border:1px solid #444 !important;
        border-radius:20px !important; color:#888 !important;
        font-size:12px !important; padding:3px 10px !important;
        height:auto !important; width:auto !important;
    }
    div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] > button:hover {
        border-color:#f97316 !important; color:#f97316 !important;
    }
    </style>""", unsafe_allow_html=True)
    rm_cols = st.columns(len(st.session_state.nav_tags))
    for idx, tag in enumerate(st.session_state.nav_tags):
        with rm_cols[idx]:
            if st.button(f"✕ {tag}", key=f"rm_{idx}"):
                st.session_state.nav_tags.pop(idx)
                st.rerun()
else:
    st.caption("No tags yet — type above and press Enter or click + Add")

# SECTION 5
st.markdown('<div class="section-hdr"><div class="section-num">5</div><div class="section-ttl">Seed Keywords <span class="section-sub">— optional</span></div></div>', unsafe_allow_html=True)
seeds_input = st.text_area("seeds", height=100, label_visibility="collapsed",
    placeholder="One per line — specific keywords that must be included")

st.divider()

if "generating" not in st.session_state:
    st.session_state.generating = False

st.markdown("""<style>
.gen-btn div[data-testid="stButton"] > button {
    background:#f97316 !important; border:none !important;
    border-radius:10px !important; font-size:15px !important;
    font-weight:600 !important; color:#fff !important; padding:0.7rem !important;
}
.gen-btn div[data-testid="stButton"] > button:hover { background:#ea6c0a !important; }
</style>""", unsafe_allow_html=True)

st.markdown('<div class="gen-btn">', unsafe_allow_html=True)
generate = st.button("⚡  Generate Keywords", key="generate_btn", use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

if generate:
    st.session_state.generating = True
    st.session_state.stop_requested = False
    st.session_state.results = None

    if not gemini_key:
        st.session_state.generating = False
        st.error("Please enter your Gemini API key in the sidebar.")
        st.stop()
    if not client_name:
        st.session_state.generating = False
        st.error("Please enter a client name.")
        st.stop()
    if not st.session_state.nav_tags:
        st.session_state.generating = False
        st.error("Please add at least one tag.")
        st.stop()

    nav_categories    = st.session_state.nav_tags
    seed_keywords     = [x.strip() for x in seeds_input.strip().split("\n") if x.strip()]
    market_code       = MARKETS[market_label]
    tags              = ", ".join(nav_categories)
    seed_note         = f" Must include these seed keywords: {', '.join(seed_keywords)}." if seed_keywords else ""
    web_note          = f" Client website: {website} — only generate keywords relevant to products on this site." if website else ""
    qtype_code        = st.session_state.selected_qt
    selected_lang_obj = [l for l in LANGUAGES if l["code"] in st.session_state.selected_langs]

    import base64
    genai_client = genai.Client(api_key=gemini_key)

    # Test DataForSEO credentials before starting
    dfs_ok = False
    if dfs_login and dfs_password:
        try:
            test_creds = base64.b64encode(f"{dfs_login}:{dfs_password}".encode()).decode()
            test_res = requests.post(
                "https://api.dataforseo.com/v3/keywords_data/google_ads/search_volume/live",
                headers={"Authorization": f"Basic {test_creds}", "Content-Type": "application/json"},
                json=[{"keywords": ["test"], "location_name": DFS_LOCATIONS.get(market_label, market_label), "language_name": "English"}],
                timeout=10
            )
            test_data = test_res.json()
            if test_data.get("status_code") == 20000:
                dfs_ok = True
            else:
                st.warning(f"⚠️ DataForSEO connected but returned error: {test_data.get('status_message', 'Unknown error')}. Keywords will generate without volumes.")
        except Exception as e:
            st.warning(f"⚠️ Could not connect to DataForSEO: {str(e)}. Keywords will generate without volumes.")

    def run_for_language(lang_name):
        prompt = f"""You are an expert SEO strategist. Client: "{client_name}", market: "{market_label}".{web_note}
Nav categories to use as tags: {tags}.
Generate exactly {branded_count} branded and {generic_count} generic keywords in {lang_name}.
{QTYPE_INSTRUCTIONS[qtype_code]}
Rules:
- Branded = includes the brand name or a clear brand variation
- Generic = category/product search with no brand name
- Assign each keyword exactly ONE tag from: {tags}
- All keywords must be in {lang_name}{seed_note}
- Add "validation": "confirmed" if clearly on-brand, "inferred" if plausible but less certain
Return ONLY a raw JSON array, no markdown, no explanation:
[{{"keyword":"...","category":"Branded","tag":"...","validation":"confirmed"}}]"""
        response = genai_client.models.generate_content(
            model="gemini-2.5-flash-lite", contents=prompt)
        match = re.search(r'\[[\s\S]*\]', response.text)
        if not match:
            raise ValueError(f"Could not parse response for {lang_name}")
        kws = json.loads(match.group())
        for kw in kws:
            kw["language"] = lang_name
            kw["volume"]   = None
            kw["combined"] = f"{kw['keyword']}, {kw['category']}, {kw['tag']}"
        return kws

    all_keywords = []
    for li, lang in enumerate(selected_lang_obj):
        with st.spinner(f"⏳ Generating {lang['name']} keywords ({li+1}/{len(selected_lang_obj)})..."):
            try:
                all_keywords.extend(run_for_language(lang["name"]))
            except Exception as e:
                st.session_state.generating = False
                st.error(f"Error for {lang['name']}: {str(e)}")
                st.stop()

    if dfs_ok:
        dfs_creds   = base64.b64encode(f"{dfs_login}:{dfs_password}".encode()).decode()
        dfs_headers = {"Authorization": f"Basic {dfs_creds}", "Content-Type": "application/json"}
        dfs_location = DFS_LOCATIONS.get(market_label, market_label)

        with st.spinner("Fetching search volumes from DataForSEO..."):
            try:
                all_kw_texts = [kw["keyword"] for kw in all_keywords]
                payload = [{"keywords": all_kw_texts, "location_name": dfs_location, "language_name": "English"}]
                res  = requests.post(
                    "https://api.dataforseo.com/v3/keywords_data/google_ads/search_volume/live",
                    headers=dfs_headers, json=payload, timeout=60)
                data = res.json()
                task  = data.get("tasks", [{}])[0]
                items = task.get("result", []) or []
                vol_map = {}
                for item in items:
                    kw  = item.get("keyword")
                    vol = item.get("search_volume")
                    if kw and vol is not None:
                        vol_map[kw] = vol
            except Exception as e:
                st.warning(f"⚠️ Volume fetch failed: {str(e)}")
                vol_map = {}

            # Assign volumes — collect zero-volume keywords for substitution
            zero_kws = []
            for i, kw in enumerate(all_keywords):
                vol = vol_map.get(kw["keyword"])
                if vol and vol > 0:
                    all_keywords[i]["volume"] = vol
                else:
                    all_keywords[i]["volume"] = None
                    zero_kws.append(i)

            # Handle zero-volume keywords — get substitutes and batch fetch them
            if zero_kws:
                def get_substitutes_batch(keywords_with_meta):
                    sub_prompt = (
                        f"These keywords have zero search volume in {market_label}. "
                        f"For each, suggest 1 better alternative that means the same thing and is more commonly searched. "
                        f"Keywords: {json.dumps([k['keyword'] for k in keywords_with_meta])}. "
                        f"Return ONLY a JSON array of strings, one substitute per keyword in the same order: "
                        f'["sub1","sub2","sub3"]'
                    )
                    try:
                        r = genai_client.models.generate_content(model="gemini-2.5-flash-lite", contents=sub_prompt)
                        m = re.search(r'\[[\s\S]*?\]', r.text)
                        return json.loads(m.group()) if m else []
                    except:
                        return []

                zero_meta   = [all_keywords[i] for i in zero_kws]
                substitutes = get_substitutes_batch(zero_meta)

                if substitutes:
                    # Batch fetch volumes for all substitutes at once
                    try:
                        sub_payload = [{"keywords": substitutes, "location_name": dfs_location, "language_name": "English"}]
                        sub_res  = requests.post(
                            "https://api.dataforseo.com/v3/keywords_data/google_ads/search_volume/live",
                            headers=dfs_headers, json=sub_payload, timeout=60)
                        sub_data  = sub_res.json()
                        sub_items = sub_data.get("tasks", [{}])[0].get("result", [{}])[0].get("items", [])
                        sub_vol_map = {item["keyword"]: item.get("search_volume") for item in sub_items if item.get("keyword")}
                    except:
                        sub_vol_map = {}

                    zero_replaced = 0
                    for idx, orig_i in enumerate(zero_kws):
                        if idx < len(substitutes):
                            sub = substitutes[idx]
                            sub_vol = sub_vol_map.get(sub)
                            if sub_vol and sub_vol > 0:
                                all_keywords[orig_i]["original_keyword"] = all_keywords[orig_i]["keyword"]
                                all_keywords[orig_i]["keyword"]  = sub
                                all_keywords[orig_i]["volume"]   = sub_vol
                                all_keywords[orig_i]["combined"] = f"{sub}, {all_keywords[orig_i]['category']}, {all_keywords[orig_i]['tag']}"
                                zero_replaced += 1

                    if zero_replaced:
                        st.info(f"ℹ️ {zero_replaced} keyword(s) replaced due to zero search volume.")

    st.session_state.generating = False
    st.success(f"✅ {len(all_keywords)} keywords generated!")
    st.divider()

    branded_n   = sum(1 for k in all_keywords if k["category"] == "Branded")
    generic_n   = sum(1 for k in all_keywords if k["category"] == "Generic")
    confirmed_n = sum(1 for k in all_keywords if k.get("validation") == "confirmed")
    total_vol   = sum(k.get("volume") or 0 for k in all_keywords)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total", len(all_keywords))
    m2.metric("Branded", branded_n)
    m3.metric("Generic", generic_n)
    m4.metric("Total volume" if dfs_ok else "Confirmed",
              f"{total_vol:,}" if dfs_ok else confirmed_n)
    st.divider()

    import pandas as pd
    df = pd.DataFrame([{
        "Keyword":               k["keyword"],
        "Original (if replaced)":k.get("original_keyword",""),
        "Category":              k["category"],
        "Tag":                   k["tag"],
        "Language":              k["language"],
        "Search Volume":         k.get("volume") or "",
        "Validation":            k.get("validation") or "",
        "Combined Entry":        k["combined"]
    } for k in all_keywords])

    st.dataframe(df, use_container_width=True, height=450, column_config={
        "Keyword":               st.column_config.TextColumn(width="large"),
        "Original (if replaced)":st.column_config.TextColumn(width="medium"),
        "Category":              st.column_config.TextColumn(width="small"),
        "Tag":                   st.column_config.TextColumn(width="medium"),
        "Language":              st.column_config.TextColumn(width="small"),
        "Search Volume":         st.column_config.NumberColumn(width="small"),
        "Validation":            st.column_config.TextColumn(width="small"),
        "Combined Entry":        st.column_config.TextColumn(width="large"),
    })
    st.divider()

    timestamp = datetime.now().strftime("%Y%m%d-%H%M")
    slug      = client_name.lower().replace(" ","-")

    csv_buf = io.StringIO()
    w = csv.DictWriter(csv_buf, fieldnames=["keyword","original_keyword","category","tag","language","volume","validation","combined"])
    w.writeheader()
    for kw in all_keywords:
        w.writerow({k: kw.get(k,"") for k in ["keyword","original_keyword","category","tag","language","volume","validation","combined"]})

    wb = Workbook(); ws = wb.active; ws.title = "Keywords"
    DARK,WHITE="1A1A1A","FFFFFF"; LIGHT_GREY="F5F5F5"
    BLUE_BG="E3F2FD"; BLUE_TXT="0D47A1"; GREEN_BG="E8F5E9"; GREEN_TXT="1B5E20"
    AMBER_BG="FFF8E1"; AMBER_TXT="E65100"
    thin=Side(style="thin",color="E0E0E0"); border=Border(left=thin,right=thin,top=thin,bottom=thin)

    ws.merge_cells("A1:H1"); t=ws["A1"]
    t.value=f"Keyword Research — {client_name} | {market_label} | {', '.join(selected_lang_names)}"
    t.font=Font(name="Calibri",bold=True,size=13,color=WHITE)
    t.fill=PatternFill("solid",fgColor=DARK)
    t.alignment=Alignment(horizontal="left",vertical="center",indent=1)
    ws.row_dimensions[1].height=30

    ws.merge_cells("A2:H2"); meta=ws["A2"]
    meta.value=f"Generated {datetime.now().strftime('%d %b %Y')}  |  {qtype_code}  |  {len(all_keywords)} keywords  |  {branded_n} branded / {generic_n} generic"
    meta.font=Font(name="Calibri",size=10,color="888888")
    meta.fill=PatternFill("solid",fgColor="F0F0F0")
    meta.alignment=Alignment(horizontal="left",vertical="center",indent=1)
    ws.row_dimensions[2].height=18

    for col,h in enumerate(["Keyword","Original (if replaced)","Category","Tag","Language","Search Volume","Validation","Combined Entry"],1):
        c=ws.cell(row=3,column=col,value=h)
        c.font=Font(name="Calibri",bold=True,size=10,color=WHITE)
        c.fill=PatternFill("solid",fgColor=DARK)
        c.alignment=Alignment(horizontal="center",vertical="center")
        c.border=border
    ws.row_dimensions[3].height=22

    for r,kw in enumerate(all_keywords,4):
        rf=PatternFill("solid",fgColor=WHITE if r%2==0 else LIGHT_GREY)
        vals=[kw.get("keyword",""),kw.get("original_keyword",""),kw.get("category",""),
              kw.get("tag",""),kw.get("language",""),kw.get("volume",""),kw.get("validation",""),kw.get("combined","")]
        for col,val in enumerate(vals,1):
            c=ws.cell(row=r,column=col,value=val)
            c.font=Font(name="Calibri",size=10)
            c.alignment=Alignment(vertical="center",wrap_text=(col==8))
            c.border=border; c.fill=rf
        cat=ws.cell(row=r,column=3); is_brand=kw.get("category")=="Branded"
        cat.fill=PatternFill("solid",fgColor=BLUE_BG if is_brand else GREEN_BG)
        cat.font=Font(name="Calibri",size=10,bold=True,color=BLUE_TXT if is_brand else GREEN_TXT)
        cat.alignment=Alignment(horizontal="center",vertical="center")
        vc=ws.cell(row=r,column=7)
        if kw.get("validation")=="confirmed":
            vc.fill=PatternFill("solid",fgColor=GREEN_BG); vc.font=Font(name="Calibri",size=10,color=GREEN_TXT)
        elif kw.get("validation")=="inferred":
            vc.fill=PatternFill("solid",fgColor=AMBER_BG); vc.font=Font(name="Calibri",size=10,color=AMBER_TXT)
        vc.alignment=Alignment(horizontal="center",vertical="center")
        volc=ws.cell(row=r,column=6)
        volc.alignment=Alignment(horizontal="right",vertical="center")
        if isinstance(kw.get("volume"),int): volc.number_format="#,##0"
        ws.row_dimensions[r].height=18

    for i,w in enumerate([32,24,12,18,12,14,12,48],1):
        ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes="A4"

    xlsx_buf=io.BytesIO(); wb.save(xlsx_buf); xlsx_buf.seek(0)

    dl1,dl2=st.columns(2)
    with dl1:
        st.download_button("⬇️ Download Excel",data=xlsx_buf,
            file_name=f"keywords-{slug}-{market_code}-{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with dl2:
        st.download_button("⬇️ Download CSV",data=csv_buf.getvalue().encode("utf-8-sig"),
            file_name=f"keywords-{slug}-{market_code}-{timestamp}.csv",
            mime="text/csv",use_container_width=True)
