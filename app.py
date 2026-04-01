import pandas as pd
import streamlit as st
from gl_engine import (
    analyze_gl,
    infer_mapping,
    export_samples_to_excel,
    export_assurance_to_excel,
    ALL_FIELDS,
)

import streamlit as st

# --- LOGIN STATE ---
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def login():
    st.title("🔐 Audit Tool Login")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if username == "auditor" and password == "auditorbdo@123":
            st.session_state.authenticated = True
            st.success("Login successful")
            st.rerun()
        else:
            st.error("Invalid credentials")

# --- BLOCK APP IF NOT LOGGED IN ---
if not st.session_state.authenticated:
    login()
    st.stop()

st.set_page_config(page_title="GL Insight AI", layout="wide")

if "df" not in st.session_state:
    st.session_state.df = None
if "filename" not in st.session_state:
    st.session_state.filename = None
if "result" not in st.session_state:
    st.session_state.result = None
if "uploaded_file_token" not in st.session_state:
    st.session_state.uploaded_file_token = None

st.title("GL Insight AI")
st.caption("Improved GL analysis with exact source-row sampling, dynamic risk observations, assurance recommendations and HR analytics.")

with st.sidebar:
    st.header("Upload GL")
    uploaded = st.file_uploader("Excel ya CSV upload karo", type=["xlsx", "xls", "csv"])
    if uploaded is not None:
        upload_token = f"{uploaded.name}|{uploaded.size}"
        if st.session_state.uploaded_file_token != upload_token:
            if uploaded.name.lower().endswith(".csv"):
                st.session_state.df = pd.read_csv(uploaded)
            else:
                st.session_state.df = pd.read_excel(uploaded)
            st.session_state.filename = uploaded.name
            st.session_state.result = None
            st.session_state.uploaded_file_token = upload_token
            st.success(f"Loaded: {uploaded.name}")
        else:
            st.caption(f"Loaded: {st.session_state.filename}")


df = st.session_state.df
if df is None:
    st.info("Pehle GL upload karo. Upload ke baad hi analysis chalega.")
    st.stop()

st.subheader(f"Uploaded file: {st.session_state.filename}")
st.write(f"Rows: {len(df):,} | Columns: {len(df.columns)}")

auto_mapping, auto_conf = infer_mapping(df)

with st.expander("Custom Mapping", expanded=True):
    st.write("Agar auto-mapping ghalat ho to yahan columns manually set karo.")
    cols = ["__none__"] + list(df.columns)
    mapping_override = {}
    grid = st.columns(3)
    for i, field in enumerate(ALL_FIELDS):
        default_val = auto_mapping.get(field)
        default_index = cols.index(default_val) if default_val in cols else 0
        with grid[i % 3]:
            mapping_override[field] = st.selectbox(
                f"{field.title()} column",
                cols,
                index=default_index,
                key=f"map_{field}",
                help=f"Auto confidence: {auto_conf.get(field, 0)}"
            )

if st.button("Run Analysis", type="primary", use_container_width=True):
    clean_override = {k: (None if v == "__none__" else v) for k, v in mapping_override.items()}
    try:
        st.session_state.result = analyze_gl(df, mapping_override=clean_override)
    except Exception as e:
        st.error(f"Analysis error: {e}")

result = st.session_state.result
if result is None:
    st.warning("Mapping review karke Run Analysis dabao.")
    st.stop()

st.success("Analysis completed.")
if result.get("warnings"):
    for w in result["warnings"]:
        st.warning(w)

summary = result["summary"]
c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Journals", f"{summary['total_journals']:,}")
c2.metric("High Risk", f"{summary['high_risk_count']:,}")
c3.metric("Medium Risk", f"{summary['medium_risk_count']:,}")
c4.metric("Suggested Samples", f"{summary['suggested_samples']:,}")

tabs = st.tabs([
    "Dashboard",
    "Flagged Entries",
    "Party Analysis",
    "Sampling",
    "Assurance by Head",
    "HR Analytics",
    "Mapping",
    "Source Preview",
])

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = tabs

with tab1:
    left, right = st.columns(2)
    with left:
        st.subheader("Monthly entries")
        me = pd.DataFrame(result["monthly_entries"])
        if not me.empty:
            st.bar_chart(me.set_index("month"))
    with right:
        st.subheader("Monthly debit / credit movement")
        mv = pd.DataFrame(result["monthly_party_movement"])
        if not mv.empty:
            st.line_chart(mv.set_index("month"))
    a, b = st.columns([1.8, 1])
    with a:
        st.subheader("Top observations")
        for idx, item in enumerate(result["top_observations"], start=1):
            st.markdown(f"**{idx}.** {item}")
    with b:
        st.subheader("Risk distribution")
        st.dataframe(pd.DataFrame(result["risk_distribution"]), use_container_width=True, hide_index=True)

with tab2:
    st.subheader("Flagged journal explorer")
    flagged_df = pd.DataFrame(result["flagged_entries"])
    if not flagged_df.empty:
        c1, c2 = st.columns([1, 2])
        with c1:
            risk_choice = st.selectbox("Risk filter", ["All","High","Medium","Low"])
        with c2:
            search = st.text_input("Search by source row / journal / account / party / user")
        filtered = flagged_df.copy()
        if risk_choice != "All":
            filtered = filtered[filtered["risk_label"] == risk_choice]
        if search:
            mask = filtered.astype(str).apply(lambda s: s.str.contains(search, case=False, na=False)).any(axis=1)
            filtered = filtered[mask]
        st.dataframe(filtered, use_container_width=True, hide_index=True)

with tab3:
    st.subheader("Party-wise summary")
    st.dataframe(pd.DataFrame(result["party_summary"]), use_container_width=True, hide_index=True)

with tab4:
    st.subheader("Automatic sampling")
    s1, s2, s3 = st.columns(3)
    s1.metric("100% High Risk", summary["high_risk_samples"])
    s2.metric("Targeted Medium Risk", summary["medium_risk_samples"])
    s3.metric("Random/Residual", summary["low_risk_samples"])

    sample_summary_df = pd.DataFrame(result["sample_summary"])
    if not sample_summary_df.empty:
        st.markdown("### Coverage by account head")
        st.dataframe(sample_summary_df, use_container_width=True, hide_index=True)

    sample_df = pd.DataFrame(result["sample_records"])
    if not sample_df.empty:
        st.markdown("### Suggested sample entries (with exact source row reference)")
        st.dataframe(sample_df, use_container_width=True, hide_index=True)

        st.markdown("### Exact copied lines from uploaded GL")
        source_extract_df = pd.DataFrame(result["sample_source_extract"])
        st.dataframe(source_extract_df, use_container_width=True, hide_index=True)

        excel_bytes = export_samples_to_excel(
            result["sample_records"],
            result["sample_summary"],
            result["sample_source_extract"],
        )
        st.download_button(
            "Download Sample Excel (with Original GL Extract)",
            data=excel_bytes,
            file_name="gl_ai_sample_extract.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.markdown("### Sampling logic")
    st.markdown("- Har account head ke liye at least 15% absolute balance coverage target liya gaya hai.")
    st.markdown("- High risk entries pehle include hoti hain.")
    st.markdown("- Phir medium risk aur zarurat par residual entries add hoti hain.")
    st.markdown("- Ab har sample ke saath source row number aur original GL extract bhi diya jata hai, taa ke management se tie-back asaan ho.")

with tab5:
    st.subheader("Assurance by head")
    st.caption("Yeh section batata hai ke kis head me analytical assurance kaise li ja sakti hai aur kis procedure se le sakte hain.")
    assurance_df = pd.DataFrame(result["assurance_summary"])
    monthly_df = pd.DataFrame(result["assurance_monthly"])
    if not assurance_df.empty:
        st.dataframe(assurance_df, use_container_width=True, hide_index=True)
        with st.expander("Monthly analytical procedure output"):
            st.dataframe(monthly_df, use_container_width=True, hide_index=True)
        assurance_bytes = export_assurance_to_excel(
            result["assurance_summary"],
            result["assurance_monthly"],
            result["hr_summary"],
            result["hr_monthly"],
        )
        st.download_button(
            "Download Assurance & Analytical Procedures Excel",
            data=assurance_bytes,
            file_name="gl_assurance_analytics.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

with tab6:
    st.subheader("HR audit analytical procedures")
    st.caption("HR-related heads ko detect karke batata hai ke analytical procedure kyun lag sakta hai aur practically kaise apply hoga.")
    hr_df = pd.DataFrame(result["hr_summary"])
    hr_monthly_df = pd.DataFrame(result["hr_monthly"])
    if hr_df.empty:
        st.info("Uploaded GL me clear HR-related heads detect nahi hue. Agar HR heads different naming se hain to account mapping / GL narration review karo.")
    else:
        st.dataframe(hr_df, use_container_width=True, hide_index=True)
        with st.expander("HR monthly analytics applied on full selected account population"):
            st.dataframe(hr_monthly_df, use_container_width=True, hide_index=True)
        assurance_bytes = export_assurance_to_excel(
            result["assurance_summary"],
            result["assurance_monthly"],
            result["hr_summary"],
            result["hr_monthly"],
        )
        st.download_button(
            "Download HR Analytical Procedures Excel",
            data=assurance_bytes,
            file_name="hr_analytical_procedures.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

with tab7:
    st.subheader("Auto + custom mapping result")
    mapping_df = pd.DataFrame({
        "Field": list(result["mapping"].keys()),
        "Mapped Column": [result["mapping"][k] for k in result["mapping"].keys()],
        "Auto Confidence": [result["mapping_confidence"].get(k, 0) for k in result["mapping"].keys()],
    })
    st.dataframe(mapping_df, use_container_width=True, hide_index=True)

with tab8:
    st.subheader("Uploaded file preview")
    st.write("Detected source columns:")
    st.write(result["source_columns"])
    st.dataframe(df.head(25), use_container_width=True, hide_index=True)
