from pathlib import Path
import json
import tempfile
import streamlit as st
from bix.services.extract import extract_project
from bix.services.validation import validate_ir
from bix.services.lineage import build_lineage
from bix.services.compare import compare_ir
from bix.services.translation import translate_expression
from bix.generation.powerbi import generate_powerbi_project

st.set_page_config(page_title="BI-X", page_icon="🔷", layout="wide")
st.title("🔷 BI-X")
st.caption("Bidirectional BI metadata reverse-engineering, translation and generation")

@st.cache_data(show_spinner=False)
def run_extract(data: bytes, filename: str, include_raw: bool):
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / filename
        p.write_bytes(data)
        return extract_project(p, include_raw=include_raw)

tabs = st.tabs(["Extract", "AI Translator", "Generate", "Compare", "Lineage", "Validation"])

with tabs[0]:
    st.subheader("Reverse engineer a BI project")
    uploaded = st.file_uploader("Power BI PBIP/ZIP or Tableau TWB/TWBX", type=["zip","pbip","twb","twbx"])
    include_raw = st.checkbox("Keep raw visual/report definitions", value=False)
    if uploaded and st.button("Extract metadata", type="primary"):
        with st.spinner("Reading project..."):
            result = run_extract(uploaded.getvalue(), uploaded.name, include_raw)
        st.session_state["bix_ir"] = result
        st.success(f"Extracted {result['platform']} project: {result['identity']['name']}")
    ir = st.session_state.get("bix_ir")
    if ir:
        a,b,c,d = st.columns(4)
        a.metric("Platform", ir["platform"])
        a.metric("Tables", len(ir.get("semantic_models",[{}])[0].get("tables",[])))
        a.metric("Pages", len(ir.get("reports",[{}])[0].get("pages",[])))
        d.metric("Visuals", len(ir.get("reports",[{}])[0].get("visuals",[])))
        st.download_button("Download BI-IR JSON", json.dumps(ir, indent=2, ensure_ascii=False), "bix-ir.json", "application/json")
        st.json(ir, expanded=False)

with tabs[1]:
    st.subheader("Calculation translator")
    ir = st.session_state.get("bix_ir")
    if not ir:
        st.info("Extract a project first.")
    else:
        expr = st.text_area("Expression", height=180)
        src = st.selectbox("Source", ["DAX","TABLEAU_CALC"])
        dst = st.selectbox("Target", ["TABLEAU_CALC","DAX"])
        if st.button("Translate"):
            st.session_state["translation"] = translate_expression(expr, src, dst, ir)
        if st.session_state.get("translation"):
            st.json(st.session_state["translation"])

with tabs[2]:
    st.subheader("Generate Power BI PBIP")
    ir = st.session_state.get("bix_ir")
    if not ir:
        st.info("Extract a project first.")
    else:
        name = st.text_input("Project name", ir["identity"]["name"])
        if st.button("Generate PBIP", type="primary"):
            with tempfile.TemporaryDirectory() as td:
                out = generate_powerbi_project(ir, Path(td), name)
                zip_bytes = Path(out["zip"]).read_bytes()
            st.success("PBIP generation completed.")
            st.download_button("Download generated PBIP ZIP", zip_bytes, f"{name}.zip", "application/zip")

with tabs[3]:
    st.subheader("Semantic comparison")
    left = st.text_area("Source BI-IR JSON", value=json.dumps(st.session_state.get("bix_ir", {}), indent=2), height=250)
    right = st.text_area("Target BI-IR JSON", height=250)
    if st.button("Compare"):
        try:
            st.json(compare_ir(json.loads(left), json.loads(right)))
        except Exception as e:
            st.error(str(e))

with tabs[4]:
    st.subheader("Lineage")
    ir = st.session_state.get("bix_ir")
    if ir:
        graph = build_lineage(ir)
        st.dataframe(graph["edges"], use_container_width=True)
        st.download_button("Download DOT", graph["dot"], "lineage.dot", "text/vnd.graphviz")
    else:
        st.info("Extract a project first.")

with tabs[5]:
    st.subheader("Validation & reconstructability")
    ir = st.session_state.get("bix_ir")
    if ir:
        result = validate_ir(ir)
        st.metric("Issues", len(result["issues"]))
        st.dataframe(result["issues"], use_container_width=True)
    else:
        st.info("Extract a project first.")
