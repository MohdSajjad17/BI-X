from pathlib import Path
import sys
import json

# Streamlit Cloud runs app.py from the repository root. Make the src-layout package importable
# even when the repository has not been installed with `pip install -e .` yet.
ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if SRC.is_dir() and str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

import tempfile
import streamlit as st
from bix.services.extract import extract_project
from bix.services.powerbi_export import write_xlsx
from bix.services.validation import validate_ir
from bix.services.lineage import build_lineage
from bix.services.compare import compare_ir
from bix.services.translation import translate_expression
from bix.generation.powerbi import generate_powerbi_project

st.set_page_config(page_title="BI-X", page_icon="🔷", layout="wide")
st.title("🔷 BI-X")
st.caption("Bidirectional BI metadata reverse-engineering, translation and generation")


def run_extract(uploaded_file, filename: str, include_raw: bool):
    # Stream the uploaded object to disk instead of creating another full-size
    # bytes copy in memory. This matters for TWB/TWBX files hundreds of MBs+.
    with tempfile.TemporaryDirectory(prefix="bix-upload-") as td:
        p = Path(td) / Path(filename).name
        with p.open("wb") as target:
            while True:
                chunk = uploaded_file.read(8 * 1024 * 1024)
                if not chunk:
                    break
                target.write(chunk)
        return extract_project(p, include_raw=include_raw)


tabs = st.tabs(["Extract", "AI Translator", "Generate", "Compare", "Lineage", "Validation"])

with tabs[0]:
    st.subheader("Reverse engineer a BI project")
    uploaded = st.file_uploader(
        "Power BI PBIP/ZIP or Tableau TWB/TWBX",
        type=["zip", "pbip", "twb", "twbx"],
    )
    include_raw = st.checkbox("Keep raw visual/report definitions", value=False)

    if uploaded and st.button("Extract metadata", type="primary"):
        with st.spinner("Reading project..."):
            result = run_extract(uploaded, uploaded.name, include_raw)
        st.session_state["bix_ir"] = result
        st.success(f"Extracted {result['platform']} project: {result['identity']['name']}")

    ir = st.session_state.get("bix_ir")
    if ir:
        a, b, c, d = st.columns(4)
        a.metric("Platform", ir["platform"])
        a.metric("Tables", len(ir.get("semantic_models", [{}])[0].get("tables", [])))
        a.metric("Pages", len(ir.get("reports", [{}])[0].get("pages", [])))
        d.metric("Visuals", len(ir.get("reports", [{}])[0].get("visuals", [])))

        st.download_button(
            "Download BI-IR JSON",
            json.dumps(ir, indent=2, ensure_ascii=False),
            "bix-ir.json",
            "application/json",
        )

        # Power BI output compatible with the earlier reverse-engineering workbook:
        # Summary, Tables, Columns, Measures, Relationships, Power Query,
        # Parameters, Pages, Visuals, Visual Fields, Filters, Unresolved Refs, DAX Queries.
        if ir.get("platform") == "powerbi" and ir.get("extensions", {}).get("reference_metadata"):
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                xlsx_path = Path(tmp.name)
            try:
                write_xlsx(ir["extensions"]["reference_metadata"] | {
                    "input": ir["identity"]["source"],
                }, xlsx_path)
                xlsx_bytes = xlsx_path.read_bytes()
            finally:
                xlsx_path.unlink(missing_ok=True)

            st.download_button(
                "Download Power BI Metadata Excel",
                xlsx_bytes,
                f"{ir['identity']['name']}_PowerBI_Metadata.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.json(ir, expanded=False)

with tabs[1]:
    st.subheader("Calculation translator")
    ir = st.session_state.get("bix_ir")
    if not ir:
        st.info("Extract a project first.")
    else:
        expr = st.text_area("Expression", height=180)
        src = st.selectbox("Source", ["DAX", "TABLEAU_CALC"])
        dst = st.selectbox("Target", ["TABLEAU_CALC", "DAX"])
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
    left = st.text_area(
        "Source BI-IR JSON",
        value=json.dumps(st.session_state.get("bix_ir", {}), indent=2),
        height=250,
    )
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
