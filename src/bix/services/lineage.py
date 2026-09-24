def build_lineage(ir):
    edges = []
    for model in ir.get("semantic_models", []):
        for rel in model.get("relationships", []):
            source = f"{rel.get('fromTable')}.{rel.get('fromColumn','')}"
            target = f"{rel.get('toTable')}.{rel.get('toColumn','')}"
            edges.append({"source": source, "target": target, "type": "relationship"})
        for table in model.get("tables", []):
            for part in table.get("partitions", []):
                if part.get("expression"):
                    edges.append({"source": part["expression"], "target": table["name"], "type": "query"})
    lines = ["digraph BI_X {", "  rankdir=LR;"]
    for edge in edges:
        s = edge["source"].replace('"', "'")
        t = edge["target"].replace('"', "'")
        lines.append(f'  "{s}" -> "{t}" [label="{edge["type"]}"];')
    lines.append("}")
    return {"edges": edges, "dot": "\n".join(lines)}
