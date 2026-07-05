"""
✨ DataPolish — Professional CSV / Excel Cleaner
--------------------------------------------------
A single-file Streamlit app that lets a non-technical person upload a
CSV or Excel file, clean it with one click, see a Before vs After
comparison, and download both the cleaned file and a plain-English
cleaning report.
"""

import io
from datetime import datetime

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------
# PAGE CONFIG
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="DataPolish · Clean your data in one click",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------------------------------------------------
# CUSTOM STYLING (unique, professional look)
# ----------------------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&family=Inter:wght@400;500;600&display=swap');

    html, body, [class*="css"]  {
        font-family: 'Inter', sans-serif;
    }

    /* -----------------------------------------------------------------
       LOCK THE WHOLE APP TO ONE LIGHT THEME.
       Streamlit auto-switches backgrounds/text based on the visitor's
       OS/browser dark-mode setting, which caused mismatches (dark
       background + dark text, or light background + light text).
       Instead of chasing every mismatch, we force every major
       container to a light background AND force dark text — so the
       two always match, no matter the visitor's system setting.
       ----------------------------------------------------------------- */
    .stApp,
    [data-testid="stAppViewContainer"],
    [data-testid="stHeader"],
    [data-testid="stSidebar"],
    [data-testid="stSidebarContent"],
    [data-testid="stFileUploaderDropzone"],
    [data-testid="stFileUploaderDropzoneInstructions"],
    [data-testid="stExpander"],
    [data-testid="stWidgetLabel"],
    div[data-baseweb="select"],
    div[data-baseweb="select"] *,
    ul[role="listbox"],
    li[role="option"] {
        background-color: #ffffff !important;
        color: #1e293b !important;
    }

    [data-testid="stSidebar"] {
        border-right: 1px solid #e5e7eb;
    }

    [data-testid="stSidebarContent"] {
        background-color: #f8fafc !important;
    }

    [data-testid="stFileUploaderDropzone"] {
        border: 1px dashed #c7d2fe !important;
        border-radius: 12px;
    }

    [data-testid="stFileUploaderDropzone"] button {
        background-color: #f1f5f9 !important;
        color: #1e293b !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
    }

    /* Catch-all: any text not explicitly re-colored below stays dark */
    .stApp * {
        color: #1e293b !important;
    }

    .main {
        background: linear-gradient(180deg, #f7f9fc 0%, #eef1f8 100%);
    }

    /* Hero banner */
    .hero {
        background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 50%, #db2777 100%);
        padding: 2.2rem 2.5rem;
        border-radius: 18px;
        color: white;
        margin-bottom: 1.6rem;
        box-shadow: 0 10px 30px rgba(79, 70, 229, 0.25);
    }
    .hero h1 {
        font-family: 'Poppins', sans-serif;
        font-size: 2.1rem;
        font-weight: 700;
        margin-bottom: 0.3rem;
    }
    .hero p {
        font-size: 1.02rem;
        opacity: 0.92;
        margin: 0;
    }

    /* Step chips */
    .step-chip {
        display: inline-block;
        background: rgba(255,255,255,0.18);
        border: 1px solid rgba(255,255,255,0.35);
        border-radius: 999px;
        padding: 0.35rem 0.9rem;
        margin-right: 0.5rem;
        font-size: 0.85rem;
        font-weight: 500;
    }

    /* Section card */
    .card {
        background: white;
        border-radius: 16px;
        padding: 1.4rem 1.6rem;
        box-shadow: 0 4px 18px rgba(30, 41, 59, 0.06);
        border: 1px solid #eef0f5;
        margin-bottom: 1.2rem;
    }

    .section-title {
        font-family: 'Poppins', sans-serif;
        font-weight: 600;
        font-size: 1.15rem;
        color: #1e293b;
        margin-bottom: 0.6rem;
    }

    /* Metric cards */
    div[data-testid="stMetric"] {
        background: white;
        border-radius: 14px;
        padding: 0.8rem 0.6rem;
        border: 1px solid #eef0f5;
        box-shadow: 0 2px 10px rgba(30,41,59,0.04);
    }
    div[data-testid="stMetric"] * {
        color: #1e293b !important;
    }
    div[data-testid="stMetricLabel"] {
        color: #64748b !important;
    }

    /* Buttons */
    div.stButton > button, div.stDownloadButton > button {
        border-radius: 10px;
        font-weight: 600;
        padding: 0.65rem 1.2rem;
        border: none;
        background: linear-gradient(135deg, #4f46e5, #7c3aed);
        color: white;
        transition: all 0.15s ease-in-out;
    }
    div.stButton > button:hover, div.stDownloadButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 8px 20px rgba(124, 58, 237, 0.35);
    }

    .badge-before {
        background: #fee2e2; color: #b91c1c;
        padding: 2px 10px; border-radius: 8px; font-size: 0.8rem; font-weight: 600;
    }
    .badge-after {
        background: #dcfce7; color: #15803d;
        padding: 2px 10px; border-radius: 8px; font-size: 0.8rem; font-weight: 600;
    }

    [data-baseweb="menu"], [data-baseweb="menu"] * {
        background-color: #ffffff !important;
        color: #1e293b !important;
    }

    footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------
# HERO HEADER
# ----------------------------------------------------------------------
st.markdown(
    """
    <div class="hero">
        <h1>✨ DataPolish</h1>
        <p>Upload a messy CSV or Excel file → clean it in one click → download the polished result + a report.</p>
        <div style="margin-top:0.9rem;">
            <span class="step-chip">1️⃣ Upload</span>
            <span class="step-chip">2️⃣ Choose fixes</span>
            <span class="step-chip">3️⃣ Clean</span>
            <span class="step-chip">4️⃣ Download</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------
# SIDEBAR — CLEANING OPTIONS
# ----------------------------------------------------------------------
st.sidebar.markdown("## ⚙️ Cleaning Options")

opt_dedup = st.sidebar.checkbox("🗑️ Remove duplicate rows", value=True)
opt_missing = st.sidebar.checkbox("🧩 Remove rows with missing values", value=True)
opt_dates = st.sidebar.checkbox("📅 Fix date formats", value=True)
opt_text_case = st.sidebar.checkbox("🔤 Fix text case", value=True)

text_case_choice = "Title Case"
if opt_text_case:
    text_case_choice = st.sidebar.radio(
        "Convert text columns to:",
        ["UPPERCASE", "lowercase", "Title Case"],
        index=2,
    )

date_output_format = "%Y-%m-%d"
if opt_dates:
    date_output_format = st.sidebar.selectbox(
        "Standard date format:",
        ["%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d %b %Y"],
        format_func=lambda f: {
            "%Y-%m-%d": "YYYY-MM-DD (2026-07-05)",
            "%d-%m-%Y": "DD-MM-YYYY (05-07-2026)",
            "%m/%d/%Y": "MM/DD/YYYY (07/05/2026)",
            "%d %b %Y": "DD Mon YYYY (05 Jul 2026)",
        }[f],
    )

st.sidebar.markdown("---")
st.sidebar.caption("Made with ❤️ using Streamlit")

# ----------------------------------------------------------------------
# FILE UPLOAD
# ----------------------------------------------------------------------
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">📂 Step 1 — Upload your file</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Supports CSV and Excel (.xlsx) files",
    type=["csv", "xlsx"],
)
st.markdown("</div>", unsafe_allow_html=True)


def load_file(file):
    """Read CSV or Excel into a DataFrame, returns (df, file_type)."""
    if file.name.lower().endswith(".xlsx"):
        return pd.read_excel(file), "xlsx"
    return pd.read_csv(file), "csv"


def try_fix_dates(df, out_format):
    """Detect date-like object columns and standardize their format."""
    fixed_cols = []
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == object:
            sample = df[col].dropna().astype(str)
            if len(sample) == 0:
                continue
            parsed = pd.to_datetime(sample, errors="coerce", dayfirst=False)
            success_ratio = parsed.notna().sum() / len(sample)
            if success_ratio >= 0.7:
                full_parsed = pd.to_datetime(df[col], errors="coerce", dayfirst=False)
                df[col] = full_parsed.dt.strftime(out_format)
                fixed_cols.append(col)
    return df, fixed_cols


def fix_text_case(df, mode):
    """Apply case fixing to all text (object, non-date) columns."""
    df = df.copy()
    fixed_cols = []
    for col in df.select_dtypes(include="object").columns:
        try:
            if mode == "UPPERCASE":
                df[col] = df[col].astype(str).str.upper()
            elif mode == "lowercase":
                df[col] = df[col].astype(str).str.lower()
            else:
                df[col] = df[col].astype(str).str.title()
            fixed_cols.append(col)
        except Exception:
            pass
    return df, fixed_cols


def to_download_bytes(df, file_type):
    if file_type == "xlsx":
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Cleaned")
        return buf.getvalue()
    return df.to_csv(index=False).encode("utf-8")


# ----------------------------------------------------------------------
# MAIN LOGIC
# ----------------------------------------------------------------------
if uploaded_file is not None:
    try:
        original_df, file_type = load_file(uploaded_file)
    except Exception as e:
        st.error(f"⚠️ Could not read this file: {e}")
        st.stop()

    st.success(f"✅ File loaded — **{uploaded_file.name}** ({original_df.shape[0]} rows, {original_df.shape[1]} columns)")

    # ---------- BEFORE SUMMARY ----------
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">🔎 Step 2 — Before Cleaning</div>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Rows", f"{original_df.shape[0]:,}")
    c2.metric("Total Columns", original_df.shape[1])
    c3.metric("🔴 Duplicate Rows", int(original_df.duplicated().sum()))
    c4.metric("🔴 Missing Values", int(original_df.isnull().sum().sum()))

    with st.expander("Preview original data"):
        st.dataframe(original_df.head(15), use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # ---------- CLEAN BUTTON ----------
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">🧹 Step 3 — Clean your data</div>', unsafe_allow_html=True)
    clean_clicked = st.button("✨ Clean My File Now", use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    if clean_clicked:
        report_lines = []
        report_lines.append("DataPolish Cleaning Report")
        report_lines.append("=" * 40)
        report_lines.append(f"File name:        {uploaded_file.name}")
        report_lines.append(f"Cleaned on:       {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report_lines.append(f"Original rows:    {original_df.shape[0]}")
        report_lines.append(f"Original columns: {original_df.shape[1]}")
        report_lines.append("-" * 40)

        cleaned_df = original_df.copy()

        # Duplicates
        if opt_dedup:
            before = len(cleaned_df)
            cleaned_df = cleaned_df.drop_duplicates()
            removed = before - len(cleaned_df)
            report_lines.append(f"✔ Removed duplicate rows: {removed}")

        # Missing values
        if opt_missing:
            before = len(cleaned_df)
            cleaned_df = cleaned_df.dropna()
            removed = before - len(cleaned_df)
            report_lines.append(f"✔ Removed rows with missing values: {removed}")

        # Dates
        if opt_dates:
            cleaned_df, date_cols = try_fix_dates(cleaned_df, date_output_format)
            if date_cols:
                report_lines.append(f"✔ Standardized date format ({date_output_format}) in columns: {', '.join(date_cols)}")
            else:
                report_lines.append("✔ Date fix checked — no clear date columns found")

        # Text case
        if opt_text_case:
            cleaned_df, text_cols = fix_text_case(cleaned_df, text_case_choice)
            if text_cols:
                report_lines.append(f"✔ Converted text to {text_case_choice} in columns: {', '.join(text_cols)}")
            else:
                report_lines.append("✔ Text case fix checked — no text columns found")

        report_lines.append("-" * 40)
        report_lines.append(f"Final rows:       {cleaned_df.shape[0]}")
        report_lines.append(f"Final columns:    {cleaned_df.shape[1]}")
        report_lines.append(f"🔴 Remaining duplicates:      {int(cleaned_df.duplicated().sum())}")
        report_lines.append(f"🔴 Remaining missing values:  {int(cleaned_df.isnull().sum().sum())}")
        report_lines.append("=" * 40)
        report_lines.append("Generated automatically by DataPolish ✨")

        report_text = "\n".join(report_lines)

        # ---------- AFTER SUMMARY ----------
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">✅ Step 4 — Before vs After Comparison</div>', unsafe_allow_html=True)

        colL, colR = st.columns(2)
        with colL:
            st.markdown('<span class="badge-before">BEFORE</span>', unsafe_allow_html=True)
            st.metric("Rows", f"{original_df.shape[0]:,}")
            st.metric("Duplicate Rows", int(original_df.duplicated().sum()))
            st.metric("Missing Values", int(original_df.isnull().sum().sum()))
        with colR:
            st.markdown('<span class="badge-after">AFTER</span>', unsafe_allow_html=True)
            st.metric("Rows", f"{cleaned_df.shape[0]:,}", delta=int(cleaned_df.shape[0] - original_df.shape[0]))
            st.metric("Duplicate Rows", int(cleaned_df.duplicated().sum()))
            st.metric("Missing Values", int(cleaned_df.isnull().sum().sum()))

        tab1, tab2 = st.tabs(["📄 Original (first 10 rows)", "✨ Cleaned (first 10 rows)"])
        with tab1:
            st.dataframe(original_df.head(10), use_container_width=True)
        with tab2:
            st.dataframe(cleaned_df.head(10), use_container_width=True)

        st.markdown("</div>", unsafe_allow_html=True)

        # ---------- DOWNLOADS ----------
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">⬇️ Step 5 — Download your results</div>', unsafe_allow_html=True)

        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                label=f"⬇️ Download Cleaned File (.{file_type})",
                data=to_download_bytes(cleaned_df, file_type),
                file_name=f"cleaned_{uploaded_file.name.rsplit('.', 1)[0]}.{file_type}",
                mime="text/csv" if file_type == "csv" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with d2:
            st.download_button(
                label="📋 Download Cleaning Report (.txt)",
                data=report_text.encode("utf-8"),
                file_name=f"cleaning_report_{uploaded_file.name.rsplit('.', 1)[0]}.txt",
                mime="text/plain",
                use_container_width=True,
            )

        st.markdown("</div>", unsafe_allow_html=True)

        st.balloons()
        st.success("🎉 Your file is polished and ready! Use the buttons above to download.")

else:
    st.markdown(
        """
        <div class="card" style="text-align:center; padding:3rem;">
            <h3 style="color:#475569;">👋 Welcome! Upload a CSV or Excel file above to get started.</h3>
            <p style="color:#94a3b8;">Your data never leaves this session — nothing is stored anywhere.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )