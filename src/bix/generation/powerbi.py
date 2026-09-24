from pathlib import Path
import json
import shutil
import zipfile

def generate_powerbi_project(ir: dict, out_dir: Path, name: str):
    project = out_dir / name
    if project.exists():
        shutil.rmtree(project)
    report_root = project / f"{name}.Report"
    model_root = project / f"{name}.SemanticModel"
    (report_root / "definition" / "pages").mkdir(parents=True)
    (model_root / "definition").mkdir(parents=True)

    pbip = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0.0",
        "artifacts": [{"report": {"path": f"{name}.Report"}}],
    }
    (project / f"{name}.pbip").write_text(json.dumps(pbip, indent=2), encoding="utf-8")

    model = ir.get("semantic_models", [{"name": name, "tables": []}])[0]
    lines = ["model model", "culture: en-US"]
    for table in model.get("tables", []):
        lines += ["", f"table '{table.get('name')}'"]
        for col in table.get("columns", []):
            lines.append(f"\tcolumn '{col.get('name')}'")
            if col.get("data_type"):
                lines.append(f"\t\tdataType: {col['data_type']}")
        for measure in table.get("measures", []):
            lines.append(f"\tmeasure '{measure.get('name')}' = {measure.get('expression') or '0'}")
    (model_root / "definition.pbism").write_text(json.dumps({"version": "4.0"}, indent=2), encoding="utf-8")
    (model_root / "definition" / "model.tmdl").write_text("\n".join(lines) + "\n", encoding="utf-8")

    report_def = report_root / "definition"
    (report_def / "report.json").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/1.0.0/schema.json"
    }, indent=2), encoding="utf-8")
    (report_def / "pages" / "pages.json").write_text(json.dumps({"pageOrder": []}, indent=2), encoding="utf-8")

    zip_path = out_dir / f"{name}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for file in project.rglob("*"):
            if file.is_file():
                z.write(file, file.relative_to(project))
    return {"folder": str(project), "zip": str(zip_path)}
