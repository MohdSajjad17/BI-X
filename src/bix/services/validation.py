def validate_ir(ir: dict):
    issues = []
    models = ir.get("semantic_models", [])
    reports = ir.get("reports", [])
    for model in models:
        names = {t.get("name") for t in model.get("tables", [])}
        for rel in model.get("relationships", []):
            ft = rel.get("fromTable") or rel.get("from", {}).get("table")
            tt = rel.get("toTable") or rel.get("to", {}).get("table")
            if ft and ft not in names:
                issues.append({"severity": "ERROR", "check": "relationship", "message": f"Unknown source table: {ft}"})
            if tt and tt not in names:
                issues.append({"severity": "ERROR", "check": "relationship", "message": f"Unknown target table: {tt}"})
        for table in model.get("tables", []):
            if not table.get("name"):
                issues.append({"severity": "ERROR", "check": "table", "message": "Table without name"})
            if not table.get("partitions"):
                issues.append({"severity": "WARN", "check": "reconstructability", "message": f"No partition/source captured for {table.get('name')}"})
    for report in reports:
        for page in report.get("pages", []):
            if not page.get("name"):
                issues.append({"severity": "WARN", "check": "page", "message": f"Page {page.get('id')} has no display name"})
    return {
        "status": "PASS" if not any(x["severity"] == "ERROR" for x in issues) else "FAIL",
        "issues": issues,
        "counts": {
            "errors": sum(x["severity"] == "ERROR" for x in issues),
            "warnings": sum(x["severity"] == "WARN" for x in issues),
        },
    }
