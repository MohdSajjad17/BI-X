from __future__ import annotations
import json
import re
import zipfile
from pathlib import Path
from typing import Any

def _safe_extract(zf: zipfile.ZipFile, target: Path) -> None:
    root = target.resolve()
    for info in zf.infolist():
        dest = (target / info.filename).resolve()
        if dest != root and root not in dest.parents:
            raise ValueError(f"Unsafe archive member: {info.filename}")
        if info.is_dir():
            dest.mkdir(parents=True, exist_ok=True)
        else:
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(zf.read(info))

def _json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception:
        return None

def _find(root: Path, names: set[str]):
    for p in root.rglob("*"):
        if p.is_file() and p.name.lower() in names:
            yield p

def _tmdl_table(text: str, fallback: str) -> dict[str, Any]:
    m = re.search(r"(?m)^table\s+(.+?)\s*$", text)
    name = m.group(1).strip().strip('"') if m else fallback
    table = {"name": name, "columns": [], "measures": [], "partitions": [], "hierarchies": [], "metadata": {}}
    current = None
    for line in text.splitlines():
        s=line.strip()
        mc=re.match(r'column\s+(.+?)(?:\s*=\s*(.*))?$',s)
        mm=re.match(r'measure\s+(.+?)(?:\s*=\s*(.*))?$',s)
        if mc:
            table["columns"].append({"name":mc.group(1).strip().strip('"'),"data_type":None,"expression":mc.group(2),"metadata":{}})
        elif mm:
            table["measures"].append({"name":mm.group(1).strip().strip('"'),"expression":mm.group(2),"metadata":{}})
    return table

def _extract_project_root(root: Path, include_raw: bool) -> dict[str, Any]:
    identity={"name": root.stem, "source": str(root)}
    sm={"name": root.stem, "tables": [], "relationships": [], "roles": [], "perspectives": [], "parameters": [], "expressions": [], "data_sources": [], "annotations": {}, "extensions": {}}
    reports=[]
    for p in _find(root, {"model.bim"}):
        data=_json(p)
        model=(data or {}).get("model", data or {})
        sm["name"]=p.parent.name or root.stem
        for t in model.get("tables",[]) or []:
            sm["tables"].append({
                "name":t.get("name",""),
                "columns":[{"name":c.get("name",""),"data_type":c.get("dataType"),"expression":c.get("expression"),"metadata":c} for c in t.get("columns",[]) or []],
                "measures":[{"name":m.get("name",""),"expression":m.get("expression"),"metadata":m} for m in t.get("measures",[]) or []],
                "partitions":t.get("partitions",[]) or [], "hierarchies":t.get("hierarchies",[]) or [], "metadata":t
            })
        sm["relationships"]=model.get("relationships",[]) or []
    for p in _find(root, {"model.tmdl"}):
        sm["name"]=p.parent.name or root.stem
    for p in root.rglob("*.tmdl"):
        if p.name.lower() != "model.tmdl":
            table=_tmdl_table(p.read_text(encoding="utf-8-sig",errors="replace"),p.stem)
            if table["columns"] or table["measures"]:
                sm["tables"].append(table)
    # TMDL model relationships/roles are retained as raw extension data until a full parser is applied.
    for p in _find(root, {"report.json"}):
        data=_json(p)
        if isinstance(data,dict):
            reports.append({"name":p.parent.name or root.stem,"pages":data.get("sections",data.get("pages",[])) or [],"visuals":data.get("visualContainers",[]) or [],"filters":data.get("filters",[]) or [],"bookmarks":data.get("bookmarks",[]) or [],"themes":data.get("themes",[]) or [],"resources":[],"extensions":{"format":"legacy","raw":data if include_raw else None}})
    for p in root.rglob("*.json"):
        n=p.name.lower()
        if n in {"report.json","model.bim","package.json"}:
            continue
        data=_json(p)
        if "sections" in data if isinstance(data,dict) else False:
            reports.append({"name":p.parent.name or root.stem,"pages":data.get("sections",[]),"visuals":data.get("visualContainers",[]),"filters":data.get("filters",[]),"bookmarks":data.get("bookmarks",[]),"themes":data.get("themes",[]),"resources":[],"extensions":{"raw":data if include_raw else None}})
    return {"identity":identity,"platform":"powerbi","format":"pbip","version":None,"metadata":{"parser":"BI-X Power BI adapter"},"semantic_models":[sm],"reports":reports or [{"name":root.stem,"pages":[],"visuals":[],"filters":[],"bookmarks":[],"themes":[],"resources":[],"extensions":{}}],"data_sources":[],"resources":[],"security":{},"lineage":{},"validation":{},"extensions":{}}

def extract(source: Path, include_raw=False):
    source=Path(source)
    if source.is_dir():
        return _extract_project_root(source,include_raw)
    if source.suffix.lower() in {".zip",".pbip"}:
        import tempfile
        with tempfile.TemporaryDirectory() as td:
            target=Path(td)
            with zipfile.ZipFile(source) as zf:
                _safe_extract(zf,target)
            return _extract_project_root(target,include_raw)
    raise ValueError(f"Unsupported Power BI input: {source}")

def parse_tmdl_file(path: Path) -> dict[str,Any]:
    return _tmdl_table(Path(path).read_text(encoding="utf-8-sig",errors="replace"),Path(path).stem)
