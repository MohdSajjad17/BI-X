def compare_ir(left: dict, right: dict):
    def index(ir):
        objects = set()
        for model in ir.get("semantic_models", []):
            for table in model.get("tables", []):
                objects.add(("table", table.get("name")))
                for column in table.get("columns", []):
                    objects.add(("column", table.get("name"), column.get("name")))
                for measure in table.get("measures", []):
                    objects.add(("measure", table.get("name"), measure.get("name")))
            for rel in model.get("relationships", []):
                objects.add(("relationship", str(rel)))
        for report in ir.get("reports", []):
            for page in report.get("pages", []):
                objects.add(("page", page.get("id"), page.get("name")))
            for visual in report.get("visuals", []):
                objects.add(("visual", visual.get("pageId"), visual.get("id"), visual.get("type")))
        return objects
    a, b = index(left), index(right)
    return {
        "added": sorted(b - a, key=str),
        "removed": sorted(a - b, key=str),
        "unchanged": len(a & b),
        "source_objects": len(a),
        "target_objects": len(b),
    }
