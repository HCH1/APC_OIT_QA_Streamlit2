import io
import tempfile
from pathlib import Path
from datetime import datetime

import streamlit as st
from modules.apc1_converter import convert_uploaded_file
from modules.qa_compare import write_result

st.set_page_config(page_title="Engineering Data Tools", page_icon="🧩", layout="wide", initial_sidebar_state="collapsed")
st.title("Engineering Data Tools")
st.caption("APC1 → OIT Converter | OIT vs webDM QA Comparison")

tab1, tab2 = st.tabs(["Tab 1 · APC1 → OIT", "Tab 2 · QA PSV Compare"])

def temp_file(suffix, data):
    f = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    f.write(data); f.flush(); f.close()
    return Path(f.name)

with tab1:
    st.header("APC1 → OIT Converter")
    st.write("Upload APC1 Excel workbook and convert it using the v9 mapping and embedded OIT template.")
    apc = st.file_uploader("Upload APC1 Excel", type=["xlsx", "xlsm"], key="apc_upload")
    if apc is not None:
        st.info(f"Selected: {apc.name} · {apc.size:,} bytes")
        if st.button("Convert APC1 to OIT", type="primary", key="apc_convert"):
            tmp_out = None
            try:
                suffix = Path(apc.name).suffix or ".xlsx"
                m = __import__("re").search(r"PT\d+", apc.name, flags=__import__("re").I)
                pt = m.group(0).upper() if m else "PT_UNKNOWN"
                out_name = f"{datetime.now():%m%d}_OIT_result_{pt}.xlsx"
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
                    tmp_out = Path(f.name)
                count = convert_uploaded_file(apc.getvalue(), apc.name, tmp_out)
                st.session_state["apc_result"] = (tmp_out.read_bytes(), out_name, count)
                st.success(f"Conversion completed. Rows converted: {count}")
            except Exception as e:
                st.error(f"APC1 → OIT conversion failed: {e}")
            finally:
                if tmp_out and tmp_out.exists():
                    tmp_out.unlink(missing_ok=True)
    if "apc_result" in st.session_state:
        data, name, count = st.session_state["apc_result"]
        st.metric("Rows converted", count)
        st.download_button("Download OIT Result", data, file_name=name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="apc_download")

with tab2:
    st.header("QA PSV Comparison")
    st.write("Upload OIT1 Excel and webDM2 raw PSV, PSV-export Excel, or normal Excel Device Listing.")
    c1, c2 = st.columns(2)
    with c1:
        oit = st.file_uploader("Upload OIT1 Excel", type=["xlsx", "xlsm"], key="qa_oit")
    with c2:
        webdm = st.file_uploader("Upload webDM2 (.psv / .xlsx / .xlsm)", type=["psv", "xlsx", "xlsm"], key="qa_webdm")
    if oit is not None and webdm is not None:
        st.caption(f"OIT1: {oit.name} · webDM2: {webdm.name}")
        if st.button("Run QA Comparison", type="primary", key="qa_run"):
            paths = []
            try:
                oit_path = temp_file(Path(oit.name).suffix or ".xlsx", oit.getvalue()); paths.append(oit_path)
                webdm_path = temp_file(Path(webdm.name).suffix or ".psv", webdm.getvalue()); paths.append(webdm_path)
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
                    out_path = Path(f.name)
                try:
                    stats = write_result(oit_path, webdm_path, out_path)
                    out_name = f"{datetime.now():%m%d}_QApsv_result_{__import__('re').search(r'PT\d+', oit.name, flags=__import__('re').I).group(0).upper() if __import__('re').search(r'PT\d+', oit.name, flags=__import__('re').I) else 'PTxxxx'}.xlsx"
                    st.session_state["qa_result"] = (out_path.read_bytes(), out_name, stats)
                finally:
                    out_path.unlink(missing_ok=True)
            except Exception as e:
                st.error(f"QA PSV comparison failed: {e}")
            finally:
                for p in paths:
                    p.unlink(missing_ok=True)
    if "qa_result" in st.session_state:
        data, name, stats = st.session_state["qa_result"]
        st.success("QA comparison completed.")
        cols = st.columns(4)
        cols[0].metric("Matched IDs", stats["matched"])
        cols[1].metric("Different fields", stats["differences"])
        cols[2].metric("OIT-only IDs", stats["oit_only"])
        cols[3].metric("webDM-only IDs", stats["webdm_only"])
        st.metric("Devices with differences", stats["devices_with_differences"])
        st.download_button("Download QA Report", data, file_name=name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="qa_download")

st.divider()
st.caption("Engineering Data Tools · Browser upload/download edition. Use only company-approved hosting for company data.")
