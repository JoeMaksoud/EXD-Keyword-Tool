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
    gemini_key  = st.secrets.get("GEMINI_KEY", "") or os.environ.get("GEMINI_KEY", "")
    semrush_key = st.secrets.get("SEMRUSH_KEY", "") or os.environ.get("SEMRUSH_KEY", "")
except:
    gemini_key  = os.environ.get("GEMINI_KEY", "")
    semrush_key = os.environ.get("SEMRUSH_KEY", "")

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
    if semrush_key:
        st.success("SEMrush ✓")
    else:
        semrush_key = st.text_input("SEMrush API key (optional)", type="password")

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
if "stop_requested" not in st.session_state:
    st.session_state.stop_requested = False

st.markdown("""<style>
.gen-btn div[data-testid="stButton"] > button {
    background:#f97316 !important; border:none !important;
    border-radius:10px !important; font-size:15px !important;
    font-weight:600 !important; color:#fff !important; padding:0.7rem !important;
}
.gen-btn div[data-testid="stButton"] > button:hover { background:#ea6c0a !important; }
.stop-btn div[data-testid="stButton"] > button {
    background:#2a2a2a !important; border:1.5px solid #555 !important;
    border-radius:10px !important; font-size:15px !important;
    font-weight:600 !important; color:#888 !important; padding:0.7rem !important;
}
.stop-btn div[data-testid="stButton"] > button:hover {
    background:#3a1a1a !important; border-color:#f97316 !important; color:#f97316 !important;
}
</style>""", unsafe_allow_html=True)

gb_col, sb_col = st.columns([3, 1])
with gb_col:
    st.markdown('<div class="gen-btn">', unsafe_allow_html=True)
    generate = st.button("⚡  Generate Keywords", key="generate_btn", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
with sb_col:
    st.markdown('<div class="stop-btn">', unsafe_allow_html=True)
    stop_clicked = st.button("⏹ Stop", key="stop_btn", use_container_width=True,
                             disabled=not st.session_state.generating)
    st.markdown('</div>', unsafe_allow_html=True)
    if stop_clicked:
        st.session_state.stop_requested = True

if generate:
    st.session_state.generating = True
    st.session_state.stop_requested = False

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

    client = genai.Client(api_key=gemini_key)

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
        response = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
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
        if st.session_state.stop_requested:
            st.warning("⏹ Generation stopped by user.")
            st.session_state.generating = False
            st.stop()
        with st.spinner(f"Generating keywords in {lang['name']} ({li+1}/{len(selected_lang_obj)})..."):
            try:
                all_keywords.extend(run_for_language(lang["name"]))
            except Exception as e:
                st.session_state.generating = False
                st.error(f"Error for {lang['name']}: {str(e)}")
                st.stop()

    if semrush_key:
        progress = st.progress(0, text="Fetching search volumes from SEMrush...")

        def fetch_volume(keyword):
            try:
                res = requests.get("https://api.semrush.com/", params={
                    "type":"phrase_this","key":semrush_key,
                    "export_columns":"Ph,Nq","database":market_code,
                    "phrase":keyword,"export_escape":"1"
                }, timeout=10)
                lines = res.text.strip().split("\n")
                if len(lines) >= 2:
                    vol = lines[1].split(";")[1].replace('"','').strip()
                    return int(vol) if vol.isdigit() else None
            except:
                return None

        def get_substitutes(original, tag, category, lang_name):
            try:
                sub_prompt = (f'Keyword "{original}" has zero search volume in {market_label}. '
                              f'Suggest 3 alternatives with same meaning, more commonly searched. '
                              f'Category: {category}, Tag: {tag}, Language: {lang_name}. '
                              f'Return ONLY a JSON array: ["alt1","alt2","alt3"]')
                r = client.models.generate_content(model="gemini-2.5-flash", contents=sub_prompt)
                m = re.search(r'\[[\s\S]*?\]', r.text)
                if m: return json.loads(m.group())
            except:
                pass
            return []

        zero_replaced = 0
        for i, kw in enumerate(all_keywords):
            vol = fetch_volume(kw["keyword"])
            if not vol:
                for sub in get_substitutes(kw["keyword"], kw["tag"], kw["category"], kw["language"]):
                    sub_vol = fetch_volume(sub)
                    if sub_vol and sub_vol > 0:
                        all_keywords[i]["original_keyword"] = kw["keyword"]
                        all_keywords[i]["keyword"]  = sub
                        all_keywords[i]["volume"]   = sub_vol
                        all_keywords[i]["combined"] = f"{sub}, {kw['category']}, {kw['tag']}"
                        zero_replaced += 1
                        break
                else:
                    all_keywords[i]["volume"] = vol
            else:
                all_keywords[i]["volume"] = vol
            if st.session_state.stop_requested:
                progress.empty()
                st.warning("⏹ Stopped during volume fetch. Partial results shown below.")
                st.session_state.generating = False
                break
            progress.progress((i+1)/len(all_keywords),
                text=f"Fetching volumes {i+1}/{len(all_keywords)}..." +
                     (f" ({zero_replaced} replaced)" if zero_replaced else ""))
        progress.empty()
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
    m4.metric("Total volume" if semrush_key else "Confirmed",
              f"{total_vol:,}" if semrush_key else confirmed_n)
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
