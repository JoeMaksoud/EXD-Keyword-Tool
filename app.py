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

# ── Page config ───────────────────────────────────────────────────
st.set_page_config(
    page_title="EXD Keyword Research Tool",
    page_icon="⚡",
    layout="wide"
)

# ── Custom CSS ────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { padding-top: 1rem; }
    .block-container { padding-top: 1rem; max-width: 900px; }
    h1 { font-size: 1.6rem !important; font-weight: 700 !important; }
    h3 { font-size: 1rem !important; font-weight: 600 !important; color: #444 !important; }
    .stButton > button {
        width: 100%;
        background: #1a1a1a;
        color: white;
        border: none;
        padding: 0.6rem 1rem;
        font-size: 15px;
        font-weight: 600;
        border-radius: 6px;
    }
    .stButton > button:hover { background: #333; border: none; color: white; }
    .stDownloadButton > button {
        background: #1a1a1a;
        color: white;
        border: none;
        font-weight: 600;
    }
    .stDownloadButton > button:hover { background: #333; border: none; color: white; }
    div[data-testid="stMetricValue"] { font-size: 2rem !important; }
    .stat-box { background: #f8f8f8; border-radius: 8px; padding: 1rem; text-align: center; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────
MARKETS = {
    "Saudi Arabia": "sa", "UAE": "ae", "Egypt": "eg", "Kuwait": "kw",
    "Qatar": "qa", "Bahrain": "bh", "Oman": "om", "Jordan": "jo",
    "Lebanon": "lb", "United Kingdom": "uk", "United States": "us",
    "France": "fr", "Germany": "de", "India": "in", "China": "cn", "Japan": "jp"
}

LANGUAGES = ["English", "Arabic", "French", "German", "Hindi", "Chinese (Simplified)", "Japanese"]

QTYPE_INSTRUCTIONS = {
    "Direct (1-3 words)":    "Generate SHORT HIGH-INTENT keywords only. Each must be 1-3 words maximum. Focus on product names, brand terms, and core category searches.",
    "Long-tail (4-8 words)": "Generate LONG-TAIL keywords only. Each must be 4-8 words. Focus on specific questions, comparisons, and use-case searches reflecting real user behavior.",
    "Mix (balanced)":        "Generate a MIX: roughly half 1-3 words (direct/high-intent) and half 4-8 words (long-tail/specific). Distribute naturally across branded and generic."
}

# ── Header ────────────────────────────────────────────────────────
st.markdown("# ⚡ EXD Keyword Research Tool")
st.markdown("Generate, validate, and export pitch keyword lists in seconds.")
st.divider()

# ── Sidebar — API Keys ────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔑 API Keys")
    gemini_key  = st.text_input("Gemini API key", type="password",
        help="Get a free key at aistudio.google.com")
    semrush_key = st.text_input("SEMrush API key (optional)", type="password",
        help="Leave empty to skip search volumes")
    st.caption("Keys are not stored — entered fresh each session.")
    st.divider()
    st.markdown("### ℹ️ How to use")
    st.caption("1. Enter your API keys above")
    st.caption("2. Fill in the pitch details")
    st.caption("3. Click Generate")
    st.caption("4. Download your Excel or CSV")

# ── Main form ─────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.markdown("### Client details")
    client_name = st.text_input("Client name", placeholder="e.g. Al Shifa Honey")
    market_label = st.selectbox("Target market", list(MARKETS.keys()), index=1)
    website = st.text_input("Client website (optional)", placeholder="https://www.alshifa.com")
    language = st.selectbox("Language", LANGUAGES)

with col2:
    st.markdown("### Keyword settings")
    query_type = st.radio("Query type", list(QTYPE_INSTRUCTIONS.keys()), index=2)
    c1, c2 = st.columns(2)
    with c1:
        branded_count = st.selectbox("Branded keywords", [5, 7, 10], index=2)
    with c2:
        generic_count = st.selectbox("Generic keywords", [20, 30, 40, 50], index=1)

st.divider()

col3, col4 = st.columns(2)
with col3:
    st.markdown("### Nav categories (tags)")
    nav_input = st.text_area(
        "One per line — paste from client's website menu",
        placeholder="Honey\nRoyal Jelly\nSupplements\nBrand",
        height=130,
        label_visibility="collapsed"
    )

with col4:
    st.markdown("### Seed keywords (optional)")
    seeds_input = st.text_area(
        "One per line — keywords that must be included",
        placeholder="Al Shifa honey\npure honey\nraw organic honey",
        height=130,
        label_visibility="collapsed"
    )

st.divider()

# ── Generate button ───────────────────────────────────────────────
generate = st.button("⚡ Generate Keywords", use_container_width=True)

# ── Generation logic ──────────────────────────────────────────────
if generate:
    # Validate inputs
    if not gemini_key:
        st.error("Please enter your Gemini API key in the sidebar.")
        st.stop()
    if not client_name:
        st.error("Please enter a client name.")
        st.stop()
    if not nav_input.strip():
        st.error("Please enter at least one nav category.")
        st.stop()

    nav_categories = [x.strip() for x in nav_input.strip().split("\n") if x.strip()]
    seed_keywords  = [x.strip() for x in seeds_input.strip().split("\n") if x.strip()]
    market_code    = MARKETS[market_label]
    tags           = ", ".join(nav_categories)
    seed_note      = f" Must include these seed keywords: {', '.join(seed_keywords)}." if seed_keywords else ""
    web_note       = f" Client website: {website} — only generate keywords relevant to products on this site." if website else ""

    prompt = f"""You are an expert SEO strategist. Client: "{client_name}", market: "{market_label}".{web_note}
Nav categories to use as tags: {tags}.

Generate exactly {branded_count} branded and {generic_count} generic keywords in {language}.
{QTYPE_INSTRUCTIONS[query_type]}

Rules:
- Branded = includes the brand name or a clear brand variation
- Generic = category/product search with no brand name
- Assign each keyword exactly ONE tag from: {tags}
- All keywords must be in {language}{seed_note}
- Add "validation": "confirmed" if clearly on-brand, "inferred" if plausible but less certain

Return ONLY a raw JSON array, no markdown, no explanation:
[{{"keyword":"...","category":"Branded","tag":"...","validation":"confirmed"}}]"""

    # Step 1: Generate
    with st.spinner(f"Generating {branded_count + generic_count} keywords with Gemini..."):
        try:
            genai.configure(api_key=gemini_key)
            model    = genai.GenerativeModel("gemini-2.5-flash")
            response = model.generate_content(prompt)
            raw      = response.text
            match    = re.search(r'\[[\s\S]*\]', raw)
            if not match:
                st.error("Could not parse Gemini response. Please try again.")
                st.stop()
            keywords = json.loads(match.group())
            for kw in keywords:
                kw["volume"]   = None
                kw["language"] = language
                kw["combined"] = f"{kw['keyword']}, {kw['category']}, {kw['tag']}"
        except Exception as e:
            st.error(f"Gemini error: {str(e)}")
            st.stop()

    # Step 2: Volumes
    if semrush_key:
        progress = st.progress(0, text="Fetching search volumes from SEMrush...")
        for i, kw in enumerate(keywords):
            try:
                res   = requests.get("https://api.semrush.com/", params={
                    "type": "phrase_this", "key": semrush_key,
                    "export_columns": "Ph,Nq", "database": market_code,
                    "phrase": kw["keyword"], "export_escape": "1"
                }, timeout=10)
                lines = res.text.strip().split("\n")
                if len(lines) >= 2:
                    vol = lines[1].split(";")[1].replace('"','').strip()
                    keywords[i]["volume"] = int(vol) if vol.isdigit() else None
            except:
                keywords[i]["volume"] = None
            progress.progress((i + 1) / len(keywords),
                text=f"Fetching volumes {i+1}/{len(keywords)}...")
        progress.empty()

    # ── Results ───────────────────────────────────────────────────
    st.success(f"✅ {len(keywords)} keywords generated!")
    st.divider()

    branded_n   = sum(1 for k in keywords if k["category"] == "Branded")
    generic_n   = sum(1 for k in keywords if k["category"] == "Generic")
    confirmed_n = sum(1 for k in keywords if k.get("validation") == "confirmed")
    total_vol   = sum(k.get("volume") or 0 for k in keywords)

    # Stats
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total keywords", len(keywords))
    m2.metric("Branded", branded_n)
    m3.metric("Generic", generic_n)
    m4.metric("Total volume" if semrush_key else "Confirmed", f"{total_vol:,}" if semrush_key else confirmed_n)

    st.divider()

    # Table
    import pandas as pd
    df = pd.DataFrame([{
        "Keyword":       k["keyword"],
        "Category":      k["category"],
        "Tag":           k["tag"],
        "Language":      k["language"],
        "Search Volume": k.get("volume") or "",
        "Validation":    k.get("validation") or "",
        "Combined Entry":k["combined"]
    } for k in keywords])

    st.dataframe(df, use_container_width=True, height=450,
        column_config={
            "Keyword":       st.column_config.TextColumn(width="large"),
            "Category":      st.column_config.TextColumn(width="small"),
            "Tag":           st.column_config.TextColumn(width="medium"),
            "Search Volume": st.column_config.NumberColumn(width="small"),
            "Validation":    st.column_config.TextColumn(width="small"),
            "Combined Entry":st.column_config.TextColumn(width="large"),
        }
    )

    st.divider()

    # ── Export ────────────────────────────────────────────────────
    timestamp = datetime.now().strftime("%Y%m%d-%H%M")
    slug      = client_name.lower().replace(" ", "-")

    # CSV
    csv_buf = io.StringIO()
    writer  = csv.DictWriter(csv_buf,
        fieldnames=["keyword","category","tag","language","volume","validation","combined"])
    writer.writeheader()
    for kw in keywords:
        writer.writerow({k: kw.get(k, "") for k in
            ["keyword","category","tag","language","volume","validation","combined"]})

    # Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Keywords"

    DARK, WHITE  = "1A1A1A", "FFFFFF"
    LIGHT_GREY   = "F5F5F5"
    BLUE_BG      = "E3F2FD";  BLUE_TXT  = "0D47A1"
    GREEN_BG     = "E8F5E9";  GREEN_TXT = "1B5E20"
    AMBER_BG     = "FFF8E1";  AMBER_TXT = "E65100"
    thin         = Side(style="thin", color="E0E0E0")
    border       = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:G1")
    t = ws["A1"]
    t.value = f"Keyword Research — {client_name} | {market_label} | {language}"
    t.font  = Font(name="Calibri", bold=True, size=13, color=WHITE)
    t.fill  = PatternFill("solid", fgColor=DARK)
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:G2")
    m = ws["A2"]
    m.value = f"Generated {datetime.now().strftime('%d %b %Y')}  |  {query_type}  |  {len(keywords)} keywords  |  {branded_n} branded / {generic_n} generic"
    m.font  = Font(name="Calibri", size=10, color="888888")
    m.fill  = PatternFill("solid", fgColor="F0F0F0")
    m.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18

    for col, h in enumerate(["Keyword","Category","Tag","Language","Search Volume","Validation","Combined Entry"], 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font      = Font(name="Calibri", bold=True, size=10, color=WHITE)
        c.fill      = PatternFill("solid", fgColor=DARK)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = border
    ws.row_dimensions[3].height = 22

    for r, kw in enumerate(keywords, 4):
        rf = PatternFill("solid", fgColor=WHITE if r % 2 == 0 else LIGHT_GREY)
        for col, val in enumerate([
            kw.get("keyword",""), kw.get("category",""), kw.get("tag",""),
            language, kw.get("volume",""), kw.get("validation",""), kw.get("combined","")
        ], 1):
            c = ws.cell(row=r, column=col, value=val)
            c.font = Font(name="Calibri", size=10)
            c.alignment = Alignment(vertical="center", wrap_text=(col==7))
            c.border = border
            c.fill   = rf
        cat = ws.cell(row=r, column=2)
        is_brand  = kw.get("category") == "Branded"
        cat.fill  = PatternFill("solid", fgColor=BLUE_BG if is_brand else GREEN_BG)
        cat.font  = Font(name="Calibri", size=10, bold=True, color=BLUE_TXT if is_brand else GREEN_TXT)
        cat.alignment = Alignment(horizontal="center", vertical="center")
        vc = ws.cell(row=r, column=6)
        if kw.get("validation") == "confirmed":
            vc.fill = PatternFill("solid", fgColor=GREEN_BG)
            vc.font = Font(name="Calibri", size=10, color=GREEN_TXT)
        elif kw.get("validation") == "inferred":
            vc.fill = PatternFill("solid", fgColor=AMBER_BG)
            vc.font = Font(name="Calibri", size=10, color=AMBER_TXT)
        vc.alignment = Alignment(horizontal="center", vertical="center")
        volc = ws.cell(row=r, column=5)
        volc.alignment = Alignment(horizontal="right", vertical="center")
        if isinstance(kw.get("volume"), int):
            volc.number_format = "#,##0"
        ws.row_dimensions[r].height = 18

    for i, w in enumerate([38,12,20,14,16,14,50], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"

    xlsx_buf = io.BytesIO()
    wb.save(xlsx_buf)
    xlsx_buf.seek(0)

    # Download buttons
    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            label="⬇️ Download Excel",
            data=xlsx_buf,
            file_name=f"keywords-{slug}-{market_code}-{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with dl2:
        st.download_button(
            label="⬇️ Download CSV",
            data=csv_buf.getvalue().encode("utf-8-sig"),
            file_name=f"keywords-{slug}-{market_code}-{timestamp}.csv",
            mime="text/csv",
            use_container_width=True
        )
