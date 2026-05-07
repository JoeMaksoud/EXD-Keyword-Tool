import streamlit as st
import google.generativeai as genai
import requests
import csv
import json
import re
import io
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

  /* Section headers */
  .section-header {
    display: flex; align-items: center; gap: 12px;
    margin: 2.5rem 0 1rem;
  }
  .section-num {
    width: 28px; height: 28px; border-radius: 50%;
    background: #333; color: #fff;
    font-size: 13px; font-weight: 600;
    display: flex; align-items: center; justify-content: center;
    flex-shrink: 0;
  }
  .section-title { font-size: 15px; font-weight: 600; color: #fff; }
  .section-sub { font-size: 13px; color: #888; font-weight: 400; margin-left: 4px; }

  /* Field cards (language + query type) — match input field dark bg */
  .card-field {
    background: #1e1e1e;
    border: 1.5px solid #333;
    border-radius: 10px;
    padding: 12px 10px;
    text-align: center;
    cursor: pointer;
    position: relative;
    transition: border-color 0.15s;
    user-select: none;
  }
  .card-field:hover { border-color: #555; }
  .card-field.active { border-color: #f97316; }
  .card-field .tick {
    position: absolute; top: 6px; right: 6px;
    width: 18px; height: 18px; border-radius: 50%;
    background: #fff; display: flex; align-items: center; justify-content: center;
    font-size: 10px; color: #f97316; font-weight: 700;
    display: none;
  }
  .card-field.active .tick { display: flex; }

  /* Tag pills */
  .tag-pill {
    display: inline-flex; align-items: center; gap: 6px;
    background: #1e1e1e; border: 1px solid #333; border-radius: 20px;
    padding: 5px 12px; font-size: 13px; color: #ddd;
    margin: 4px;
  }
  .tag-x {
    cursor: pointer; color: #888; font-size: 13px; line-height: 1;
    background: none; border: none; padding: 0; font-family: inherit;
    transition: color 0.1s;
  }
  .tag-x:hover { color: #f97316; }
  .tags-wrap { display: flex; flex-wrap: wrap; margin-top: 8px; }
  .tag-empty { font-size: 13px; color: #555; margin-top: 8px; font-style: italic; }

  /* Generate button */
  div[data-testid="stButton"] > button {
    background: #f97316 !important; color: #fff !important;
    border: none !important; border-radius: 10px !important;
    padding: 0.75rem 1.5rem !important;
    font-size: 15px !important; font-weight: 600 !important;
    width: 100% !important; letter-spacing: 0.01em !important;
  }
  div[data-testid="stButton"] > button:hover { background: #ea6c0a !important; }

  /* Download buttons */
  div[data-testid="stDownloadButton"] > button {
    background: #f97316 !important; color: #fff !important;
    border: none !important; border-radius: 10px !important;
    font-weight: 600 !important; width: 100% !important;
  }
  div[data-testid="stDownloadButton"] > button:hover { background: #ea6c0a !important; }

  hr { border-color: #2a2a2a !important; }
  section[data-testid="stSidebar"] { background: #111; }

  /* Hide streamlit branding */
  #MainMenu { visibility: hidden; }
  footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────
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
    "direct":   "Generate SHORT HIGH-INTENT keywords only. Each must be 1-3 words maximum. Focus on product names, brand terms, and core category searches.",
    "longtail": "Generate LONG-TAIL keywords only. Each must be 4-8 words. Focus on specific questions, comparisons, and use-case searches.",
    "mix":      "Generate a MIX: roughly half 1-3 words (direct) and half 4-8 words (long-tail). Distribute naturally.",
}

# ── Session state ─────────────────────────────────────────────────
if "nav_tags"       not in st.session_state: st.session_state.nav_tags = []
if "selected_langs" not in st.session_state: st.session_state.selected_langs = ["en"]
if "selected_qt"    not in st.session_state: st.session_state.selected_qt = "mix"

# ── API Keys ──────────────────────────────────────────────────────
gemini_key  = st.secrets.get("GEMINI_KEY", "")
semrush_key = st.secrets.get("SEMRUSH_KEY", "")

# ── Sidebar ───────────────────────────────────────────────────────
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

# ── Header ────────────────────────────────────────────────────────
st.markdown("# ⚡ EXD Keyword Research Tool")
st.markdown("Generate, validate, and export pitch keyword lists in seconds.")
st.divider()

# ══════════════════════════════════════════════════════════════════
# SECTION 1 — Account Setup
# ══════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header"><div class="section-num">1</div><div class="section-title">Account Setup</div></div>', unsafe_allow_html=True)

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
    generic_count = st.number_input("Generic keywords",  min_value=0, max_value=1000, value=30, step=1)

# ══════════════════════════════════════════════════════════════════
# SECTION 2 — Languages (multi-select, no buttons under cards)
# ══════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header"><div class="section-num">2</div><div class="section-title">Languages</div></div>', unsafe_allow_html=True)

lang_cols = st.columns(len(LANGUAGES))
for i, lang in enumerate(LANGUAGES):
    with lang_cols[i]:
        is_active = lang["code"] in st.session_state.selected_langs
        border_c  = "#f97316" if is_active else "#333"
        tick_html = f'<div style="position:absolute;top:6px;right:6px;width:18px;height:18px;border-radius:50%;background:#fff;display:flex;align-items:center;justify-content:center;font-size:10px;color:#f97316;font-weight:700;">✓</div>' if is_active else ""
        st.markdown(f"""
        <div style="background:#1e1e1e;border:1.5px solid {border_c};border-radius:10px;
            padding:12px 4px;text-align:center;position:relative;cursor:pointer;">
            {tick_html}
            <div style="font-size:20px">{lang['flag']}</div>
            <div style="font-size:11px;font-weight:600;color:#fff;margin-top:4px">{lang['name']}</div>
            <div style="font-size:10px;color:#888">{lang['native']}</div>
        </div>""", unsafe_allow_html=True)
        # Invisible button over the card to capture click
        if st.button(" ", key=f"lang_{lang['code']}", use_container_width=True,
                     help=f"Select {lang['name']}"):
            if lang["code"] in st.session_state.selected_langs:
                if len(st.session_state.selected_langs) > 1:
                    st.session_state.selected_langs.remove(lang["code"])
            else:
                st.session_state.selected_langs.append(lang["code"])
            st.rerun()

# Show selected languages summary
selected_lang_names = [l["name"] for l in LANGUAGES if l["code"] in st.session_state.selected_langs]
if len(selected_lang_names) > 1:
    st.caption(f"✓ {len(selected_lang_names)} languages selected — tool will generate keywords for each: {', '.join(selected_lang_names)}")

# ══════════════════════════════════════════════════════════════════
# SECTION 3 — Keyword Type (click card, orange border, no button)
# ══════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header"><div class="section-num">3</div><div class="section-title">Keyword Type</div></div>', unsafe_allow_html=True)

qt_cols = st.columns(3)
for i, qt in enumerate(QTYPES):
    with qt_cols[i]:
        is_active = st.session_state.selected_qt == qt["code"]
        border_c  = "#f97316" if is_active else "#333"
        tick_html = f'<div style="position:absolute;top:6px;right:6px;width:18px;height:18px;border-radius:50%;background:#fff;display:flex;align-items:center;justify-content:center;font-size:10px;color:#f97316;font-weight:700;">✓</div>' if is_active else ""
        st.markdown(f"""
        <div style="background:#1e1e1e;border:1.5px solid {border_c};border-radius:10px;
            padding:14px;min-height:90px;position:relative;cursor:pointer;">
            {tick_html}
            <div style="font-size:14px;font-weight:600;color:#fff">{qt['title']}</div>
            <div style="font-size:11px;color:#888;margin-top:6px">{qt['desc']}</div>
        </div>""", unsafe_allow_html=True)
        if st.button(" ", key=f"qt_{qt['code']}", use_container_width=True,
                     help=qt["title"]):
            st.session_state.selected_qt = qt["code"]
            st.rerun()

# ══════════════════════════════════════════════════════════════════
# SECTION 4 — Tags (single field, pill shows with X to remove)
# ══════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header"><div class="section-num">4</div><div class="section-title">Tags</div></div>', unsafe_allow_html=True)

tag_col, btn_col = st.columns([5, 1])
with tag_col:
    new_tag = st.text_input("Add tag", placeholder="e.g. Honey, Royal Jelly, Supplements...",
                             label_visibility="collapsed", key="tag_input")
with btn_col:
    add_clicked = st.button("+ Add", key="add_tag_btn", use_container_width=True)

if (add_clicked or new_tag) and new_tag.strip():
    tag_val = new_tag.strip()
    if tag_val not in st.session_state.nav_tags:
        st.session_state.nav_tags.append(tag_val)
        st.rerun()

# Render tags inline — each pill has an X button right next to it
if st.session_state.nav_tags:
    # Build one HTML row of pills, each with a unique form key for the X
    for idx, tag in enumerate(st.session_state.nav_tags):
        col_tag, col_x, col_spacer = st.columns([3, 0.5, 8])
        with col_tag:
            st.markdown(f"""
            <div style="background:#1e1e1e;border:1px solid #333;border-radius:20px;
                padding:6px 14px;font-size:13px;color:#ddd;display:inline-block;
                white-space:nowrap;margin-top:4px;">
                {tag}
            </div>""", unsafe_allow_html=True)
        with col_x:
            if st.button("✕", key=f"rm_{idx}", help=f"Remove '{tag}'"):
                st.session_state.nav_tags.pop(idx)
                st.rerun()
else:
    st.markdown('<p style="font-size:13px;color:#555;font-style:italic;margin-top:8px;">No tags added yet — type above and press Enter or click + Add</p>', unsafe_allow_html=True)

nav_input = "\n".join(st.session_state.nav_tags)

# ══════════════════════════════════════════════════════════════════
# SECTION 5 — Seed Keywords
# ══════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header"><div class="section-num">5</div><div class="section-title">Seed Keywords <span class="section-sub">— optional</span></div></div>', unsafe_allow_html=True)
seeds_input = st.text_area("Seed keywords", height=100, label_visibility="collapsed",
    placeholder="One per line — specific keywords that must be included in the output")

st.divider()
generate = st.button("⚡ Generate Keywords", use_container_width=True)

# ── Generation logic ──────────────────────────────────────────────
if generate:
    if not gemini_key:
        st.error("Please enter your Gemini API key in the sidebar.")
        st.stop()
    if not client_name:
        st.error("Please enter a client name.")
        st.stop()
    if not st.session_state.nav_tags:
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

    genai.configure(api_key=gemini_key)
    model = genai.GenerativeModel("gemini-2.5-flash")

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
        response = model.generate_content(prompt)
        match    = re.search(r'\[[\s\S]*\]', response.text)
        if not match:
            raise ValueError(f"Could not parse Gemini response for {lang_name}")
        kws = json.loads(match.group())
        for kw in kws:
            kw["language"] = lang_name
            kw["volume"]   = None
            kw["combined"] = f"{kw['keyword']}, {kw['category']}, {kw['tag']}"
        return kws

    # Generate for each selected language
    all_keywords = []
    total_langs  = len(selected_lang_obj)
    for li, lang in enumerate(selected_lang_obj):
        with st.spinner(f"Generating keywords in {lang['name']} ({li+1}/{total_langs})..."):
            try:
                kws = run_for_language(lang["name"])
                all_keywords.extend(kws)
            except Exception as e:
                st.error(f"Error generating for {lang['name']}: {str(e)}")
                st.stop()

    # Fetch SEMrush volumes
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
            except: return None

        def get_substitutes(original, tag, category, lang_name):
            try:
                r = model.generate_content(
                    f'The keyword "{original}" has zero search volume in {market_label}. '
                    f'Suggest 3 alternative keywords with same meaning that are more commonly searched. '
                    f'Category: {category}, Tag: {tag}, Language: {lang_name}. '
                    f'Return ONLY a JSON array: ["alt1","alt2","alt3"]'
                )
                m = re.search(r'\[[\s\S]*?\]', r.text)
                if m: return json.loads(m.group())
            except: pass
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
            progress.progress((i+1)/len(all_keywords),
                text=f"Fetching volumes {i+1}/{len(all_keywords)}..." +
                     (f" ({zero_replaced} replaced)" if zero_replaced else ""))
        progress.empty()
        if zero_replaced:
            st.info(f"ℹ️ {zero_replaced} keyword(s) with zero search volume were automatically replaced.")

    # ── Results ──────────────────────────────────────────────────
    st.success(f"✅ {len(all_keywords)} keywords generated across {total_langs} language(s)!")
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

    # CSV
    csv_buf = io.StringIO()
    w = csv.DictWriter(csv_buf, fieldnames=["keyword","original_keyword","category","tag","language","volume","validation","combined"])
    w.writeheader()
    for kw in all_keywords:
        w.writerow({k: kw.get(k,"") for k in ["keyword","original_keyword","category","tag","language","volume","validation","combined"]})

    # Excel
    wb = Workbook(); ws = wb.active; ws.title = "Keywords"
    DARK,WHITE = "1A1A1A","FFFFFF"; LIGHT_GREY="F5F5F5"
    BLUE_BG="E3F2FD"; BLUE_TXT="0D47A1"; GREEN_BG="E8F5E9"; GREEN_TXT="1B5E20"
    AMBER_BG="FFF8E1"; AMBER_TXT="E65100"
    thin=Side(style="thin",color="E0E0E0"); border=Border(left=thin,right=thin,top=thin,bottom=thin)

    ws.merge_cells("A1:H1"); t=ws["A1"]
    t.value=f"Keyword Research — {client_name} | {market_label} | {', '.join(selected_lang_names)}"
    t.font=Font(name="Calibri",bold=True,size=13,color=WHITE)
    t.fill=PatternFill("solid",fgColor=DARK)
    t.alignment=Alignment(horizontal="left",vertical="center",indent=1)
    ws.row_dimensions[1].height=30

    ws.merge_cells("A2:H2"); m=ws["A2"]
    m.value=f"Generated {datetime.now().strftime('%d %b %Y')}  |  {qtype_code}  |  {len(all_keywords)} keywords  |  {branded_n} branded / {generic_n} generic  |  {', '.join(selected_lang_names)}"
    m.font=Font(name="Calibri",size=10,color="888888")
    m.fill=PatternFill("solid",fgColor="F0F0F0")
    m.alignment=Alignment(horizontal="left",vertical="center",indent=1)
    ws.row_dimensions[2].height=18

    headers = ["Keyword","Original (if replaced)","Category","Tag","Language","Search Volume","Validation","Combined Entry"]
    for col,h in enumerate(headers,1):
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
