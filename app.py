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
  .block-container { max-width: 780px; padding: 2rem 2rem 4rem; }
  .step-header { display:flex; align-items:center; gap:12px; margin:2rem 0 1rem; }
  .step-num {
    width:28px; height:28px; border-radius:50%;
    background:#111; color:#fff; font-size:13px; font-weight:600;
    display:flex; align-items:center; justify-content:center; flex-shrink:0;
  }
  .step-title { font-size:16px; font-weight:600; color:#111; }
  .step-sub { font-size:13px; color:#888; font-weight:400; margin-left:2px; }
  .tag-container { display:flex; flex-wrap:wrap; gap:8px; margin-top:10px; }
  .tag-pill {
    display:inline-flex; align-items:center; gap:6px;
    background:#f0f0f0; border:1px solid #ddd; border-radius:20px;
    padding:5px 14px; font-size:13px; font-weight:500; color:#222;
  }
  .tag-empty { font-size:13px; color:#aaa; margin-top:8px; font-style:italic; }
  div[data-testid="stButton"] > button {
    background:#111 !important; color:#fff !important;
    border:none !important; border-radius:10px !important;
    padding:0.7rem 1.5rem !important; font-size:15px !important;
    font-weight:600 !important; width:100% !important;
  }
  div[data-testid="stButton"] > button:hover { background:#333 !important; }
  div[data-testid="stDownloadButton"] > button {
    background:#111 !important; color:#fff !important;
    border:none !important; border-radius:10px !important;
    font-weight:600 !important; width:100% !important;
  }
  .remove-label { font-size:12px; color:#888; margin-top:10px; margin-bottom:4px; }
  /* Make remove tag buttons look like small pills */
  div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] > button {
    background:#fff !important; color:#555 !important;
    border:1px solid #ddd !important; border-radius:20px !important;
    padding:3px 10px !important; font-size:12px !important;
    font-weight:500 !important;
  }
  div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] > button:hover {
    background:#ffe8e8 !important; color:#cc0000 !important;
    border-color:#ffaaaa !important;
  }
  hr { border-color:#f0f0f0 !important; }
  section[data-testid="stSidebar"] { background:#fafafa; }
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
    {"code":"en","name":"English", "native":"English","flag":"🇬🇧"},
    {"code":"ar","name":"Arabic",  "native":"العربية", "flag":"🇸🇦"},
    {"code":"fr","name":"French",  "native":"Français","flag":"🇫🇷"},
    {"code":"de","name":"German",  "native":"Deutsch", "flag":"🇩🇪"},
    {"code":"hi","name":"Hindi",   "native":"हिन्दी",   "flag":"🇮🇳"},
    {"code":"zh","name":"Chinese", "native":"简体中文", "flag":"🇨🇳"},
    {"code":"ja","name":"Japanese","native":"日本語",   "flag":"🇯🇵"},
]
QTYPES = [
    {"code":"direct",   "title":"Direct",   "desc":"Short, high-intent. 1–3 words. e.g. \"pure honey\""},
    {"code":"longtail", "title":"Long-tail", "desc":"Specific phrases. 4–8 words. e.g. \"best raw honey for immunity\""},
    {"code":"mix",      "title":"Mix",       "desc":"Balanced blend of short and long-tail queries"},
]
QTYPE_INSTRUCTIONS = {
    "direct":   "Generate SHORT HIGH-INTENT keywords only. Each must be 1-3 words maximum. Focus on product names, brand terms, and core category searches.",
    "longtail": "Generate LONG-TAIL keywords only. Each must be 4-8 words. Focus on specific questions, comparisons, and use-case searches.",
    "mix":      "Generate a MIX: roughly half 1-3 words (direct) and half 4-8 words (long-tail). Distribute naturally across branded and generic.",
}

# ── Session state ─────────────────────────────────────────────────
if "nav_tags"      not in st.session_state: st.session_state.nav_tags = []
if "selected_lang" not in st.session_state: st.session_state.selected_lang = "en"
if "selected_qt"   not in st.session_state: st.session_state.selected_qt = "mix"

# ── API Keys ──────────────────────────────────────────────────────
gemini_key  = st.secrets.get("GEMINI_KEY", "")
semrush_key = st.secrets.get("SEMRUSH_KEY", "")

# ── Sidebar ───────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚡ EXD Keywords")
    st.markdown("Generate pitch keyword lists in seconds.")
    st.divider()
    st.markdown("**How to use**")
    st.caption("1. Fill in pitch details")
    st.caption("2. Select language & query type")
    st.caption("3. Add keyword categories")
    st.caption("4. Click Generate")
    st.caption("5. Download Excel or CSV")
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
# STEP 1 — Client details
# ══════════════════════════════════════════════════════════════════
st.markdown('<div class="step-header"><div class="step-num">1</div><div class="step-title">Client details</div></div>', unsafe_allow_html=True)

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
# STEP 2 — Language
# ══════════════════════════════════════════════════════════════════
st.markdown('<div class="step-header"><div class="step-num">2</div><div class="step-title">Language</div></div>', unsafe_allow_html=True)

lang_cols = st.columns(len(LANGUAGES))
for i, lang in enumerate(LANGUAGES):
    with lang_cols[i]:
        is_active  = st.session_state.selected_lang == lang["code"]
        bg         = "#111" if is_active else "#fff"
        border_c   = "#111" if is_active else "#e0e0e0"
        name_c     = "#fff" if is_active else "#111"
        native_c   = "#ccc" if is_active else "#888"
        st.markdown(f"""<div style="background:{bg};border:1.5px solid {border_c};border-radius:10px;
            padding:10px 4px;text-align:center;">
            <div style="font-size:20px">{lang['flag']}</div>
            <div style="font-size:11px;font-weight:600;color:{name_c};margin-top:3px">{lang['name']}</div>
            <div style="font-size:10px;color:{native_c}">{lang['native']}</div>
        </div>""", unsafe_allow_html=True)
        if st.button("✓" if is_active else "Select", key=f"lang_{lang['code']}", use_container_width=True):
            st.session_state.selected_lang = lang["code"]
            st.rerun()

selected_lang_name = next(l["name"] for l in LANGUAGES if l["code"] == st.session_state.selected_lang)

# ══════════════════════════════════════════════════════════════════
# STEP 3 — Query type
# ══════════════════════════════════════════════════════════════════
st.markdown('<div class="step-header"><div class="step-num">3</div><div class="step-title">Query type</div></div>', unsafe_allow_html=True)

qt_cols = st.columns(3)
for i, qt in enumerate(QTYPES):
    with qt_cols[i]:
        is_active = st.session_state.selected_qt == qt["code"]
        bg        = "#111" if is_active else "#fff"
        border_c  = "#111" if is_active else "#e0e0e0"
        title_c   = "#fff" if is_active else "#111"
        desc_c    = "#bbb" if is_active else "#888"
        st.markdown(f"""<div style="background:{bg};border:1.5px solid {border_c};
            border-radius:10px;padding:14px;min-height:85px;">
            <div style="font-size:14px;font-weight:600;color:{title_c}">{qt['title']}</div>
            <div style="font-size:11px;color:{desc_c};margin-top:5px">{qt['desc']}</div>
        </div>""", unsafe_allow_html=True)
        if st.button("✓" if is_active else "Select", key=f"qt_{qt['code']}", use_container_width=True):
            st.session_state.selected_qt = qt["code"]
            st.rerun()

# ══════════════════════════════════════════════════════════════════
# STEP 4 — Keyword categories
# ══════════════════════════════════════════════════════════════════
st.markdown('<div class="step-header"><div class="step-num">4</div><div class="step-title">Keyword categories <span class="step-sub">— from the client\'s website navigation</span></div></div>', unsafe_allow_html=True)

tag_col, btn_col = st.columns([5, 1])
with tag_col:
    new_tag = st.text_input("Add category", placeholder="e.g. Honey, Royal Jelly, Supplements...",
                             label_visibility="collapsed", key="tag_input")
with btn_col:
    add_clicked = st.button("+ Add", key="add_tag_btn", use_container_width=True)

if (add_clicked or new_tag) and new_tag.strip():
    tag_val = new_tag.strip()
    if tag_val not in st.session_state.nav_tags:
        st.session_state.nav_tags.append(tag_val)
        st.rerun()

# Show tags as HTML pills (visual only)
if st.session_state.nav_tags:
    pills = "".join(f'<span class="tag-pill">{t}</span>' for t in st.session_state.nav_tags)
    st.markdown(f'<div class="tag-container">{pills}</div>', unsafe_allow_html=True)

    # Remove buttons in uniform row below
    st.markdown('<div class="remove-label">Click to remove:</div>', unsafe_allow_html=True)
    n = len(st.session_state.nav_tags)
    max_cols = min(n, 5)
    rm_cols  = st.columns(max_cols)
    for idx, tag in enumerate(st.session_state.nav_tags):
        with rm_cols[idx % max_cols]:
            if st.button(f"✕ {tag}", key=f"rm_{idx}", use_container_width=True):
                st.session_state.nav_tags.pop(idx)
                st.rerun()
else:
    st.markdown('<div class="tag-empty">No categories added yet — type above and press Enter or click + Add</div>', unsafe_allow_html=True)

nav_input = "\n".join(st.session_state.nav_tags)

# ══════════════════════════════════════════════════════════════════
# STEP 5 — Seed keywords
# ══════════════════════════════════════════════════════════════════
st.markdown('<div class="step-header"><div class="step-num">5</div><div class="step-title">Seed keywords <span class="step-sub">— optional</span></div></div>', unsafe_allow_html=True)
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
        st.error("Please add at least one keyword category.")
        st.stop()

    nav_categories = st.session_state.nav_tags
    seed_keywords  = [x.strip() for x in seeds_input.strip().split("\n") if x.strip()]
    market_code    = MARKETS[market_label]
    tags           = ", ".join(nav_categories)
    seed_note      = f" Must include these seed keywords: {', '.join(seed_keywords)}." if seed_keywords else ""
    web_note       = f" Client website: {website} — only generate keywords relevant to products on this site." if website else ""
    qtype_code     = st.session_state.selected_qt

    prompt = f"""You are an expert SEO strategist. Client: "{client_name}", market: "{market_label}".{web_note}
Nav categories to use as tags: {tags}.
Generate exactly {branded_count} branded and {generic_count} generic keywords in {selected_lang_name}.
{QTYPE_INSTRUCTIONS[qtype_code]}
Rules:
- Branded = includes the brand name or a clear brand variation
- Generic = category/product search with no brand name
- Assign each keyword exactly ONE tag from: {tags}
- All keywords must be in {selected_lang_name}{seed_note}
- Add "validation": "confirmed" if clearly on-brand, "inferred" if plausible but less certain
Return ONLY a raw JSON array, no markdown, no explanation:
[{{"keyword":"...","category":"Branded","tag":"...","validation":"confirmed"}}]"""

    with st.spinner(f"Generating {branded_count + generic_count} keywords with Gemini..."):
        try:
            genai.configure(api_key=gemini_key)
            model    = genai.GenerativeModel("gemini-2.5-flash")
            response = model.generate_content(prompt)
            match    = re.search(r'\[[\s\S]*\]', response.text)
            if not match:
                st.error("Could not parse Gemini response. Please try again.")
                st.stop()
            keywords = json.loads(match.group())
            for kw in keywords:
                kw["volume"]   = None
                kw["language"] = selected_lang_name
                kw["combined"] = f"{kw['keyword']}, {kw['category']}, {kw['tag']}"
        except Exception as e:
            st.error(f"Gemini error: {str(e)}")
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
            except: return None

        def get_substitutes(original, tag, category):
            try:
                r = genai.GenerativeModel("gemini-2.5-flash").generate_content(
                    f'The keyword "{original}" has zero search volume in {market_label}. '
                    f'Suggest 3 alternative keywords with same meaning that are more commonly searched. '
                    f'Category: {category}, Tag: {tag}, Language: {selected_lang_name}. '
                    f'Return ONLY a JSON array: ["alt1","alt2","alt3"]'
                )
                m = re.search(r'\[[\s\S]*?\]', r.text)
                if m: return json.loads(m.group())
            except: pass
            return []

        zero_replaced = 0
        for i, kw in enumerate(keywords):
            vol = fetch_volume(kw["keyword"])
            if not vol:
                for sub in get_substitutes(kw["keyword"], kw["tag"], kw["category"]):
                    sub_vol = fetch_volume(sub)
                    if sub_vol and sub_vol > 0:
                        keywords[i]["original_keyword"] = kw["keyword"]
                        keywords[i]["keyword"]  = sub
                        keywords[i]["volume"]   = sub_vol
                        keywords[i]["combined"] = f"{sub}, {kw['category']}, {kw['tag']}"
                        zero_replaced += 1
                        break
                else:
                    keywords[i]["volume"] = vol
            else:
                keywords[i]["volume"] = vol
            progress.progress((i+1)/len(keywords),
                text=f"Fetching volumes {i+1}/{len(keywords)}..." +
                     (f" ({zero_replaced} replaced)" if zero_replaced else ""))

        progress.empty()
        if zero_replaced:
            st.info(f"ℹ️ {zero_replaced} keyword(s) with zero search volume were automatically replaced.")

    st.success(f"✅ {len(keywords)} keywords generated!")
    st.divider()

    branded_n   = sum(1 for k in keywords if k["category"] == "Branded")
    generic_n   = sum(1 for k in keywords if k["category"] == "Generic")
    confirmed_n = sum(1 for k in keywords if k.get("validation") == "confirmed")
    total_vol   = sum(k.get("volume") or 0 for k in keywords)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total", len(keywords))
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
    } for k in keywords])

    st.dataframe(df, use_container_width=True, height=450, column_config={
        "Keyword":               st.column_config.TextColumn(width="large"),
        "Original (if replaced)":st.column_config.TextColumn(width="medium"),
        "Category":              st.column_config.TextColumn(width="small"),
        "Tag":                   st.column_config.TextColumn(width="medium"),
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
    for kw in keywords:
        w.writerow({k: kw.get(k,"") for k in ["keyword","original_keyword","category","tag","language","volume","validation","combined"]})

    wb = Workbook(); ws = wb.active; ws.title = "Keywords"
    DARK,WHITE = "1A1A1A","FFFFFF"; LIGHT_GREY="F5F5F5"
    BLUE_BG="E3F2FD"; BLUE_TXT="0D47A1"; GREEN_BG="E8F5E9"; GREEN_TXT="1B5E20"
    AMBER_BG="FFF8E1"; AMBER_TXT="E65100"
    thin=Side(style="thin",color="E0E0E0"); border=Border(left=thin,right=thin,top=thin,bottom=thin)

    ws.merge_cells("A1:G1"); t=ws["A1"]
    t.value=f"Keyword Research — {client_name} | {market_label} | {selected_lang_name}"
    t.font=Font(name="Calibri",bold=True,size=13,color=WHITE)
    t.fill=PatternFill("solid",fgColor=DARK)
    t.alignment=Alignment(horizontal="left",vertical="center",indent=1)
    ws.row_dimensions[1].height=30

    ws.merge_cells("A2:G2"); m=ws["A2"]
    m.value=f"Generated {datetime.now().strftime('%d %b %Y')}  |  {qtype_code}  |  {len(keywords)} keywords  |  {branded_n} branded / {generic_n} generic"
    m.font=Font(name="Calibri",size=10,color="888888")
    m.fill=PatternFill("solid",fgColor="F0F0F0")
    m.alignment=Alignment(horizontal="left",vertical="center",indent=1)
    ws.row_dimensions[2].height=18

    for col,h in enumerate(["Keyword","Category","Tag","Language","Search Volume","Validation","Combined Entry"],1):
        c=ws.cell(row=3,column=col,value=h)
        c.font=Font(name="Calibri",bold=True,size=10,color=WHITE)
        c.fill=PatternFill("solid",fgColor=DARK)
        c.alignment=Alignment(horizontal="center",vertical="center")
        c.border=border
    ws.row_dimensions[3].height=22

    for r,kw in enumerate(keywords,4):
        rf=PatternFill("solid",fgColor=WHITE if r%2==0 else LIGHT_GREY)
        for col,val in enumerate([kw.get("keyword",""),kw.get("category",""),kw.get("tag",""),
                                   selected_lang_name,kw.get("volume",""),kw.get("validation",""),kw.get("combined","")],1):
            c=ws.cell(row=r,column=col,value=val)
            c.font=Font(name="Calibri",size=10)
            c.alignment=Alignment(vertical="center",wrap_text=(col==7))
            c.border=border; c.fill=rf
        cat=ws.cell(row=r,column=2); is_brand=kw.get("category")=="Branded"
        cat.fill=PatternFill("solid",fgColor=BLUE_BG if is_brand else GREEN_BG)
        cat.font=Font(name="Calibri",size=10,bold=True,color=BLUE_TXT if is_brand else GREEN_TXT)
        cat.alignment=Alignment(horizontal="center",vertical="center")
        vc=ws.cell(row=r,column=6)
        if kw.get("validation")=="confirmed":
            vc.fill=PatternFill("solid",fgColor=GREEN_BG); vc.font=Font(name="Calibri",size=10,color=GREEN_TXT)
        elif kw.get("validation")=="inferred":
            vc.fill=PatternFill("solid",fgColor=AMBER_BG); vc.font=Font(name="Calibri",size=10,color=AMBER_TXT)
        vc.alignment=Alignment(horizontal="center",vertical="center")
        volc=ws.cell(row=r,column=5)
        volc.alignment=Alignment(horizontal="right",vertical="center")
        if isinstance(kw.get("volume"),int): volc.number_format="#,##0"
        ws.row_dimensions[r].height=18

    for i,w in enumerate([38,12,20,14,16,14,50],1):
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
