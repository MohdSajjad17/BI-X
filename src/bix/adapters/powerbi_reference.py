#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pbip_reverse_engineer.py
========================
Reverse-engineer a Power BI Project (PBIP) into complete, rebuild-ready metadata.

INPUT  (any of these):
  * a .zip that contains the project (nested zips such as  X.Report.zip / X.SemanticModel.zip
    inside the outer zip are unpacked automatically)
  * a plain folder that contains  X.pbip + X.Report + X.SemanticModel
  * the .pbip file itself (its sibling folders are used)

OUTPUT (in the output folder):
  metadata.json          complete, structured metadata (semantic model + report + lineage + validation)
  metadata.xlsx          the same information as filterable sheets (needs `pip install openpyxl`)
  metadata.md            human-readable data dictionary / documentation
  rebuilt_project/       a clean copy of the project (definition files only, cached data removed)
  rebuilt_project.zip    the same folder zipped -> unzip and open the .pbip in Power BI Desktop

Supported formats
  Semantic model : TMDL (definition/*.tmdl)  and  TMSL (model.bim)
  Report         : PBIR (definition/pages/**)  and  legacy PBIR-Legacy (report.json)  [best effort]

USAGE
  python pbip_reverse_engineer.py "Supervisor__2__8466.zip"
  python pbip_reverse_engineer.py "C:/path/to/project_folder" -o my_output
  python pbip_reverse_engineer.py project.zip --no-raw --no-rebuild

Only the Python standard library is required (openpyxl is optional, for the .xlsx).
NOTE: metadata never contains the imported DATA. After rebuilding you must refresh the model
      (with your own credentials) so Power BI can reload data from the sources listed in the output.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import sys
import tempfile
import textwrap
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

try:  # optional
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    Workbook = None

TOOL_VERSION = "1.0"
CACHE_EXT = (".abf",)          # cached data files - never needed to rebuild, often hundreds of MB


def log(msg: str) -> None:
    print(msg, flush=True)


# =====================================================================================
# 1. INPUT HANDLING
# =====================================================================================
def _safe_extract(zf: zipfile.ZipFile, dest: Path, skipped: list) -> None:
    """Extract a zip, skipping cache files and refusing path-traversal ('zip slip')."""
    dest = dest.resolve()
    for info in zf.infolist():
        if info.is_dir():
            continue
        name = info.filename.replace("\\", "/")
        if name.lower().endswith(CACHE_EXT):
            skipped.append({"file": name, "bytes": info.file_size})
            continue
        target = (dest / name).resolve()
        if target != dest and dest not in target.parents:
            raise ValueError(f"Unsafe path inside zip: {name}")
        target.parent.mkdir(parents=True, exist_ok=True)
        with zf.open(info) as src, open(target, "wb") as out:
            shutil.copyfileobj(src, out)


def materialise(inp: str, work: Path, skipped: list) -> Path:
    """Return a folder that contains the project files."""
    p = Path(inp)
    if not p.exists():
        raise SystemExit(f"Input not found: {inp}")
    if p.is_dir():
        return p
    if p.suffix.lower() == ".pbip":
        return p.parent
    if not zipfile.is_zipfile(p):
        raise SystemExit("Input must be a .zip, a .pbip file, or a project folder.")
    root = work / "src"
    root.mkdir(parents=True, exist_ok=True)
    log(f"[1/6] Unpacking {p.name} (cache files are skipped) ...")
    with zipfile.ZipFile(p) as z:
        _safe_extract(z, root, skipped)
    for _ in range(4):                                   # nested zips: X.Report.zip, X.SemanticModel.zip ...
        inner = [z for z in root.rglob("*.zip") if zipfile.is_zipfile(z)]
        if not inner:
            break
        for z in inner:
            with zipfile.ZipFile(z) as zf:
                _safe_extract(zf, z.parent, skipped)
            z.unlink()
    return root


def load_json(p: Path):
    with open(p, encoding="utf-8-sig") as f:
        return json.load(f)


def read_text(p: Path) -> str:
    with open(p, encoding="utf-8-sig", errors="replace") as f:
        return f.read()


def locate_project(root: Path) -> dict:
    """Find the .pbip, the report folder and the semantic-model folder."""
    pbips = sorted(root.rglob("*.pbip"))
    pbip = pbips[0] if pbips else None
    report_dir = model_dir = None

    if pbip:
        try:
            for art in load_json(pbip).get("artifacts", []):
                rp = art.get("report", {}).get("path")
                if rp and (pbip.parent / rp).is_dir():
                    report_dir = (pbip.parent / rp).resolve()
        except Exception:
            pass
    if report_dir is None:
        cands = [d.parent for d in root.rglob("definition.pbir")]
        cands += [d.parent for d in root.rglob("report.json") if d.parent.name.endswith(".Report")]
        if cands:
            report_dir = sorted(set(cands))[0].resolve()

    dataset_ref = None
    if report_dir and (report_dir / "definition.pbir").exists():
        try:
            dataset_ref = load_json(report_dir / "definition.pbir").get("datasetReference")
            bp = (dataset_ref or {}).get("byPath", {}).get("path")
            if bp and (report_dir / bp).resolve().is_dir():
                model_dir = (report_dir / bp).resolve()
        except Exception:
            pass
    if model_dir is None:
        cands = [d.parent for d in root.rglob("definition.pbism")] + [d.parent for d in root.rglob("model.bim")]
        cands += [d.parent.parent for d in root.rglob("model.tmdl")]        # bare TMDL folder (no .pbism)
        if cands:
            model_dir = sorted(set(cands))[0].resolve()

    if not report_dir and not model_dir:
        raise SystemExit("No Power BI project found (no *.Report / *.SemanticModel folders).")
    name = (pbip.stem if pbip else (report_dir or model_dir).name.rsplit(".", 1)[0])
    return {"pbip": pbip, "report_dir": report_dir, "model_dir": model_dir, "dataset_ref": dataset_ref, "name": name}


# =====================================================================================
# 2. TMDL PARSER  (generic indentation-based tree)
# =====================================================================================
FLAGS = {"isHidden", "isKey", "isNameInferred", "isDefaultLabel", "isDefaultImage", "isNullable",
         "isUnique", "isAvailableInMdx", "showAsVariationsOnly", "isPrivate", "isDataTypeInferred"}


class Node:
    def __init__(self, kw, name, depth, has_expr=False, inline=""):
        self.kw, self.name, self.depth = kw, name, depth
        self.has_expr, self.inline, self.cont = has_expr, inline, []
        self.expr = None
        self.props, self.children, self.doc = {}, [], []

    def finish(self):
        if self.has_expr:
            cont = textwrap.dedent("\n".join(self.cont)).strip("\n") if self.cont else ""
            first = self.inline.strip()
            self.expr = (first + ("\n" + cont if cont else "")) if first else cont
            self.expr = self.expr.strip()
        for c in self.children:
            c.finish()

    def kids(self, kw):
        return [c for c in self.children if c.kw == kw]


def _parse_name(rest: str):
    rest = rest.strip()
    if not rest:
        return None, ""
    if rest.startswith("="):
        return None, rest
    if rest.startswith("'"):
        i, out = 1, ""
        while i < len(rest):
            if rest[i] == "'":
                if i + 1 < len(rest) and rest[i + 1] == "'":
                    out += "'"
                    i += 2
                    continue
                break
            out += rest[i]
            i += 1
        return out, rest[i + 1:].strip()
    m = re.match(r"([^\s=]+)\s*(.*)$", rest)
    return m.group(1), m.group(2)


def _depth(raw: str) -> int:
    d = i = 0
    while i < len(raw) and raw[i] in "\t ":
        if raw[i] == "\t":
            d, i = d + 1, i + 1
        else:
            j = i
            while j < len(raw) and raw[j] == " ":
                j += 1
            d, i = d + (j - i) // 4, j
    return d


def parse_tmdl(text: str) -> Node:
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    root = Node("root", None, -1)
    stack, owner, blanks, fence, doc = [root], None, [], False, []
    for raw in lines:
        if not raw.strip():
            if owner is not None:
                blanks.append("")
            continue
        depth, s = _depth(raw), raw.strip()
        if fence:                                            # inside ``` ... ```
            if s == "```":
                fence, owner, blanks = False, None, []
            else:
                owner.cont.extend(blanks)
                blanks = []
                owner.cont.append(raw)
            continue
        if owner is not None and depth >= owner.depth + 2:   # continuation of a multi-line expression
            owner.cont.extend(blanks)
            blanks = []
            owner.cont.append(raw)
            continue
        owner, blanks = None, []
        if s.startswith("///"):
            doc.append(s[3:].strip())
            continue
        while stack[-1].depth >= depth:
            stack.pop()
        parent = stack[-1]
        m = re.match(r"^([A-Za-z]\w*):\s*(.*)$", s)
        if m:                                                # property   key: value
            parent.props[m.group(1)] = m.group(2)
            continue
        m = re.match(r"^([A-Za-z]\w*)(?:\s+(.*))?$", s)
        if not m:
            continue
        kw, rest = m.group(1), (m.group(2) or "")
        if kw in FLAGS and not rest:                         # boolean flag
            parent.props[kw] = True
            continue
        if kw == "ref":                                      # ref table 'X'
            kind, _, nm = rest.partition(" ")
            node = Node("ref", _parse_name(nm)[0], depth)
            node.props["kind"] = kind
            parent.children.append(node)
            continue
        name, after = _parse_name(rest)
        has_expr = after.startswith("=")
        node = Node(kw, name, depth, has_expr, after[1:].strip() if has_expr else "")
        if doc:
            node.doc, doc = doc, []
        parent.children.append(node)
        stack.append(node)
        if has_expr:
            if node.inline.startswith("```"):
                fence = True
                tail = node.inline[3:]
                node.inline = ""
                if tail.strip():
                    node.cont.append(tail)
            owner = node
    root.finish()
    return root


def node_dict(n: Node) -> dict:
    """Lossless dump of any TMDL node (used for object types not modelled explicitly)."""
    return {"type": n.kw, "name": n.name, "expression": n.expr, "description": " ".join(n.doc) or None,
            "properties": n.props, "children": [node_dict(c) for c in n.children]}


def unq(v):
    return v.strip().strip("'") if isinstance(v, str) else v


def truthy(v) -> bool:
    return v is True or str(v).strip().lower() == "true"


def annotations(n: Node) -> dict:
    return {a.name: a.expr for a in n.kids("annotation")}


def extras(n: Node, known: set, skip_kids=()) -> dict:
    ex = {k: v for k, v in n.props.items() if k not in known}
    for ch in n.children:
        if ch.kw in ("annotation",) + tuple(skip_kids):
            continue
        ex.setdefault(ch.kw, []).append({"name": ch.name, "value": ch.expr})
    return ex


def split_ref(ref: str):
    """'Table Name'.column  ->  ('Table Name', 'column')"""
    ref = ref.strip()
    if ref.startswith("'"):
        i = 1
        while True:
            j = ref.index("'", i)
            if j + 1 < len(ref) and ref[j + 1] == "'":
                i = j + 2
                continue
            break
        return ref[1:j].replace("''", "'"), unq(ref[j + 2:])
    t, c = ref.split(".", 1)
    return t, unq(c)


# =====================================================================================
# 3. SEMANTIC MODEL LOADERS  (TMDL  and  model.bim)  ->  one common structure
# =====================================================================================
def is_calc_partition(p):
    return (p.get("type") or "").lower() == "calculated"


def _tmdl_table(t: Node) -> dict:
    tp = t.props
    cols = []
    for c in t.kids("column"):
        p = c.props
        cols.append({
            "name": c.name, "type": p.get("type") or ("Calculated" if c.expr else "Data"),
            "dataType": p.get("dataType"), "formatString": unq(p.get("formatString")) if p.get("formatString") else None,
            "hidden": truthy(p.get("isHidden")), "summarizeBy": p.get("summarizeBy"),
            "sourceColumn": p.get("sourceColumn"), "sortByColumn": p.get("sortByColumn"),
            "displayFolder": p.get("displayFolder"), "dataCategory": p.get("dataCategory"),
            "isKey": truthy(p.get("isKey")), "description": " ".join(c.doc) or p.get("description"),
            "expression": c.expr or None, "lineageTag": p.get("lineageTag"),
            "annotations": annotations(c),
            "extra": extras(c, {"dataType", "formatString", "isHidden", "summarizeBy", "sourceColumn", "sortByColumn",
                                "displayFolder", "dataCategory", "isKey", "description", "lineageTag", "type"}),
        })
    meas = []
    for m in t.kids("measure"):
        p = m.props
        meas.append({
            "name": m.name, "expression": m.expr, "formatString": unq(p.get("formatString")) if p.get("formatString") else None,
            "displayFolder": p.get("displayFolder"), "hidden": truthy(p.get("isHidden")),
            "description": " ".join(m.doc) or p.get("description"), "dataType": p.get("dataType"),
            "lineageTag": p.get("lineageTag"), "annotations": annotations(m),
            "extra": extras(m, {"formatString", "displayFolder", "isHidden", "description", "dataType", "lineageTag"}),
        })
    parts = []
    for pt in t.kids("partition"):
        src = pt.kids("source")
        parts.append({
            "name": pt.name, "type": pt.expr, "mode": pt.props.get("mode"), "queryGroup": unq(pt.props.get("queryGroup")),
            "source": src[0].expr if src else None, "annotations": annotations(pt),
            "extra": extras(pt, {"mode", "queryGroup"}, skip_kids=("source",)),
        })
    hier = []
    for h in t.kids("hierarchy"):
        hier.append({"name": h.name, "hidden": truthy(h.props.get("isHidden")), "displayFolder": h.props.get("displayFolder"),
                     "levels": [{"name": lv.name, "column": lv.props.get("column"), "ordinal": lv.props.get("ordinal")}
                                for lv in h.kids("level")]})
    cg = None
    for g in t.kids("calculationGroup"):
        cg = {"precedence": g.props.get("precedence"), "items": [{
            "name": it.name, "expression": it.expr, "ordinal": it.props.get("ordinal"),
            "properties": {c.kw: c.expr for c in it.children if c.kw != "annotation"}} for it in g.kids("calculationItem")]}
    return {
        "name": t.name, "description": " ".join(t.doc) or tp.get("description"), "hidden": truthy(tp.get("isHidden")),
        "dataCategory": tp.get("dataCategory"), "lineageTag": tp.get("lineageTag"),
        "columns": cols, "measures": meas, "partitions": parts, "hierarchies": hier, "calculationGroup": cg,
        "annotations": annotations(t),
        "extra": extras(t, {"isHidden", "description", "dataCategory", "lineageTag"},
                        skip_kids=("column", "measure", "partition", "hierarchy", "calculationGroup")),
    }


def load_tmdl(mdir: Path, full_culture: bool) -> dict:
    ddir = mdir / "definition"
    tops = defaultdict(list)
    files = sorted(ddir.rglob("*.tmdl"))
    for f in files:
        for n in parse_tmdl(read_text(f)).children:
            tops[n.kw].append(n)

    model_node = (tops.get("model") or [Node("model", "Model", 0)])[0]
    db = (tops.get("database") or [Node("database", None, 0)])[0]
    order = [r.name for r in tops.get("ref", []) if r.props.get("kind") == "table"]
    by_name = {t.name: t for t in tops.get("table", [])}
    ordered = [by_name[n] for n in order if n in by_name] + [t for n, t in by_name.items() if n not in order]

    model_ann = annotations(model_node)
    for a in tops.get("annotation", []):
        model_ann[a.name] = a.expr
    qorder = []
    try:
        qorder = json.loads(model_ann.get("PBI_QueryOrder", "[]"))
    except Exception:
        pass

    exprs = []
    for e in tops.get("expression", []):
        body = e.expr or ""
        exprs.append({"name": e.name, "kind": "Parameter" if "IsParameterQuery=true" in body.replace(" ", "") else "Query / connection",
                      "queryGroup": unq(e.props.get("queryGroup")), "resultType": annotations(e).get("PBI_ResultType"),
                      "definition": body, "lineageTag": e.props.get("lineageTag"), "annotations": annotations(e)})

    rels = []
    for r in tops.get("relationship", []):
        ft, fc = split_ref(r.props.get("fromColumn", "."))
        tt, tc = split_ref(r.props.get("toColumn", "."))
        rels.append({"id": r.name, "fromTable": ft, "fromColumn": fc, "toTable": tt, "toColumn": tc,
                     "fromCardinality": r.props.get("fromCardinality", "many"), "toCardinality": r.props.get("toCardinality", "one"),
                     "crossFilteringBehavior": r.props.get("crossFilteringBehavior", "oneDirection"),
                     "isActive": str(r.props.get("isActive", "true")).lower() != "false",
                     "securityFilteringBehavior": r.props.get("securityFilteringBehavior"),
                     "joinOnDateBehavior": r.props.get("joinOnDateBehavior"),
                     "relyOnReferentialIntegrity": r.props.get("relyOnReferentialIntegrity")})

    roles = [{"name": r.name, "modelPermission": r.props.get("modelPermission"), "description": r.props.get("description"),
              "tablePermissions": [{"table": tp.name, "filterExpression": tp.expr} for tp in r.kids("tablePermission")],
              "members": [m.name for m in r.kids("member")]} for r in tops.get("role", [])]
    persp = []
    for p in tops.get("perspective", []):
        persp.append({"name": p.name, "tables": [{
            "table": pt.name, "columns": [c.name for c in pt.kids("perspectiveColumn")],
            "measures": [c.name for c in pt.kids("perspectiveMeasure")],
            "hierarchies": [c.name for c in pt.kids("perspectiveHierarchy")]} for pt in p.kids("perspectiveTable")]})
    cultures = []
    for c in tops.get("cultureInfo", []):
        ling = next((x.expr for x in c.kids("linguisticMetadata")), None)
        d = {"name": c.name, "hasLinguisticMetadata": bool(ling), "linguisticMetadataChars": len(ling or "")}
        if full_culture:
            d["linguisticMetadata"] = ling
        cultures.append(d)
    handled = {"model", "database", "table", "ref", "annotation", "expression", "relationship", "role", "perspective",
               "cultureInfo", "queryGroup"}
    other = [node_dict(n) for k, v in tops.items() if k not in handled for n in v]

    return {
        "format": "TMDL",
        "info": {"name": model_node.name, "culture": model_node.props.get("culture"),
                 "defaultPowerBIDataSourceVersion": model_node.props.get("defaultPowerBIDataSourceVersion"),
                 "sourceQueryCulture": model_node.props.get("sourceQueryCulture"),
                 "compatibilityLevel": db.props.get("compatibilityLevel"),
                 "dataAccessOptions": [c.kw for n in [model_node] for d in n.kids("dataAccessOptions") for c in d.children],
                 "tmdlFiles": len(files)},
        "annotations": {k: v for k, v in model_ann.items() if k != "PBI_QueryOrder"},
        "queryOrder": qorder,
        "queryGroups": [{"name": g.name, "order": annotations(g).get("PBI_QueryGroupOrder")} for g in tops.get("queryGroup", [])],
        "tables": [_tmdl_table(t) for t in ordered], "relationships": rels, "expressions": exprs,
        "roles": roles, "perspectives": persp, "cultures": cultures, "otherObjects": other,
    }


def _txt(v):
    return "\n".join(v) if isinstance(v, list) else v


def _ann(lst):
    return {a.get("name"): a.get("value") for a in (lst or [])}


def load_bim(path: Path) -> dict:
    j = load_json(path)
    m = j.get("model", j)
    tables = []
    for t in m.get("tables", []):
        cols = [{"name": c.get("name"), "type": "Calculated" if c.get("type") == "calculated" else "Data",
                 "dataType": c.get("dataType"), "formatString": c.get("formatString"), "hidden": bool(c.get("isHidden")),
                 "summarizeBy": c.get("summarizeBy"), "sourceColumn": c.get("sourceColumn"), "sortByColumn": c.get("sortByColumn"),
                 "displayFolder": c.get("displayFolder"), "dataCategory": c.get("dataCategory"), "isKey": bool(c.get("isKey")),
                 "description": c.get("description"), "expression": _txt(c.get("expression")), "lineageTag": c.get("lineageTag"),
                 "annotations": _ann(c.get("annotations")), "extra": {}} for c in t.get("columns", [])]
        meas = [{"name": x.get("name"), "expression": _txt(x.get("expression")), "formatString": x.get("formatString"),
                 "displayFolder": x.get("displayFolder"), "hidden": bool(x.get("isHidden")), "description": x.get("description"),
                 "dataType": x.get("dataType"), "lineageTag": x.get("lineageTag"), "annotations": _ann(x.get("annotations")),
                 "extra": {}} for x in t.get("measures", [])]
        parts = []
        for p in t.get("partitions", []):
            s = p.get("source", {})
            parts.append({"name": p.get("name"), "type": s.get("type"), "mode": p.get("mode"), "queryGroup": p.get("queryGroup"),
                          "source": _txt(s.get("expression") or s.get("query")), "annotations": _ann(p.get("annotations")), "extra": {}})
        hier = [{"name": h.get("name"), "hidden": bool(h.get("isHidden")), "displayFolder": h.get("displayFolder"),
                 "levels": [{"name": lv.get("name"), "column": lv.get("column"), "ordinal": lv.get("ordinal")}
                            for lv in h.get("levels", [])]} for h in t.get("hierarchies", [])]
        cg = None
        if t.get("calculationGroup"):
            g = t["calculationGroup"]
            cg = {"precedence": g.get("precedence"), "items": [{"name": i.get("name"), "expression": _txt(i.get("expression")),
                  "ordinal": i.get("ordinal"), "properties": {}} for i in g.get("calculationItems", [])]}
        tables.append({"name": t.get("name"), "description": t.get("description"), "hidden": bool(t.get("isHidden")),
                       "dataCategory": t.get("dataCategory"), "lineageTag": t.get("lineageTag"), "columns": cols, "measures": meas,
                       "partitions": parts, "hierarchies": hier, "calculationGroup": cg,
                       "annotations": _ann(t.get("annotations")), "extra": {}})
    rels = [{"id": r.get("name"), "fromTable": r.get("fromTable"), "fromColumn": r.get("fromColumn"), "toTable": r.get("toTable"),
             "toColumn": r.get("toColumn"), "fromCardinality": r.get("fromCardinality", "many"),
             "toCardinality": r.get("toCardinality", "one"), "crossFilteringBehavior": r.get("crossFilteringBehavior", "oneDirection"),
             "isActive": r.get("isActive", True), "securityFilteringBehavior": r.get("securityFilteringBehavior"),
             "joinOnDateBehavior": r.get("joinOnDateBehavior"), "relyOnReferentialIntegrity": r.get("relyOnReferentialIntegrity")}
            for r in m.get("relationships", [])]
    exprs = []
    for e in m.get("expressions", []):
        body = _txt(e.get("expression")) or ""
        exprs.append({"name": e.get("name"), "kind": "Parameter" if "IsParameterQuery=true" in body.replace(" ", "") else "Query / connection",
                      "queryGroup": e.get("queryGroup"), "resultType": None, "definition": body,
                      "lineageTag": e.get("lineageTag"), "annotations": _ann(e.get("annotations"))})
    roles = [{"name": r.get("name"), "modelPermission": r.get("modelPermission"), "description": r.get("description"),
              "tablePermissions": [{"table": tp.get("name"), "filterExpression": _txt(tp.get("filterExpression"))}
                                   for tp in r.get("tablePermissions", [])],
              "members": [x.get("memberName") for x in r.get("members", [])]} for r in m.get("roles", [])]
    persp = [{"name": p.get("name"), "tables": [{"table": t.get("name"), "columns": [c.get("name") for c in t.get("perspectiveColumns", [])],
              "measures": [c.get("name") for c in t.get("perspectiveMeasures", [])],
              "hierarchies": [c.get("name") for c in t.get("perspectiveHierarchies", [])]} for t in p.get("perspectiveTables", [])]}
             for p in m.get("perspectives", [])]
    ann = _ann(m.get("annotations"))
    try:
        qorder = json.loads(ann.get("PBI_QueryOrder", "[]"))
    except Exception:
        qorder = []
    return {
        "format": "TMSL (model.bim)",
        "info": {"name": j.get("name"), "culture": m.get("culture"),
                 "defaultPowerBIDataSourceVersion": m.get("defaultPowerBIDataSourceVersion"),
                 "sourceQueryCulture": m.get("sourceQueryCulture"), "compatibilityLevel": j.get("compatibilityLevel"),
                 "dataAccessOptions": list((m.get("dataAccessOptions") or {}).keys())},
        "annotations": {k: v for k, v in ann.items() if k != "PBI_QueryOrder"}, "queryOrder": qorder,
        "queryGroups": [{"name": g.get("folder"), "order": _ann(g.get("annotations")).get("PBI_QueryGroupOrder")} for g in m.get("queryGroups", [])],
        "tables": tables, "relationships": rels, "expressions": exprs, "roles": roles, "perspectives": persp,
        "cultures": [{"name": c.get("name")} for c in m.get("cultures", [])],
        "otherObjects": [{"type": "dataSources", "value": m["dataSources"]}] if m.get("dataSources") else [],
    }


def load_model(mdir: Path, full_culture: bool) -> dict:
    if (mdir / "definition" / "model.tmdl").exists() or list((mdir / "definition").rglob("*.tmdl") if (mdir / "definition").exists() else []):
        model = load_tmdl(mdir, full_culture)
    elif (mdir / "model.bim").exists():
        model = load_bim(mdir / "model.bim")
    else:
        raise SystemExit(f"Semantic model folder has neither definition/*.tmdl nor model.bim: {mdir}")
    try:
        model["info"]["pbismVersion"] = load_json(mdir / "definition.pbism").get("version")
    except Exception:
        pass
    try:
        plat = load_json(mdir / ".platform")
        model["info"]["displayName"] = plat["metadata"]["displayName"]
        model["info"]["logicalId"] = plat["config"]["logicalId"]
    except Exception:
        pass
    dq = []
    for f in sorted((mdir / "DAXQueries").glob("*.dax")) if (mdir / "DAXQueries").exists() else []:
        dq.append({"name": f.name, "content": read_text(f)})
    model["daxQueries"] = dq
    return model


# =====================================================================================
# 4. POWER QUERY (M) ANALYSIS  -> lineage between queries and to external sources
# =====================================================================================
CONNECTORS = ("DatabricksMultiCloud|Databricks|Sql|SharePoint|Sharepoint|Excel|Csv|Web|OData|Odbc|OleDb|AzureStorage|"
              "AzureDataExplorer|Folder|File|Snowflake|PostgreSQL|MySQL|Oracle|GoogleBigQuery|Salesforce|Dataverse|"
              "CommonDataService|PowerPlatform|Fabric|Lakehouse|Json|Xml|Pdf|Access|Teradata|SapHana|AnalysisServices|"
              "ActiveDirectory|Exchange|Cube|Amazon|Value")


def analyse_m(code, self_name, query_names):
    if not code:
        return [], "", []
    deps = []
    for q in query_names:
        if q == self_name:
            continue
        if re.search(r'#"' + re.escape(q) + r'"', code) or re.search(r'(?<![\w"#.])' + re.escape(q) + r'(?![\w"])', code):
            deps.append(q)
    navs = re.findall(r'\{\[(?:Name|Item|Schema)\s*=\s*"([^"]+)"(?:\s*,\s*(?:Kind|Schema)\s*=\s*"([^"]+)")?', code)
    nav_txt = "; ".join(a + (f" ({b})" if b else "") for a, b in navs)
    fns = sorted(set(re.findall(r"\b((?:%s)\.[A-Za-z]+)\(" % CONNECTORS, code)))
    fns = [f for f in fns if not f.startswith("Value.")]
    return deps, nav_txt, fns


def build_power_query(model: dict) -> tuple:
    names = [t["name"] for t in model["tables"]] + [e["name"] for e in model["expressions"]]
    rows = []
    for t in model["tables"]:
        for p in t["partitions"]:
            deps, nav, fns = analyse_m(p["source"], t["name"], names)
            rows.append({"query": t["name"], "kind": "Table partition", "language": "DAX" if is_calc_partition(p) else
                         ("M" if (p.get("type") or "").lower() == "m" else p.get("type")),
                         "partitionType": p.get("type"), "mode": p.get("mode"), "queryGroup": p.get("queryGroup"),
                         "dependsOn": deps, "sourceObjects": nav, "sourceFunctions": fns, "code": p["source"]})
    for e in model["expressions"]:
        deps, nav, fns = analyse_m(e["definition"], e["name"], names)
        rows.append({"query": e["name"], "kind": e["kind"], "language": "M", "partitionType": None, "mode": None,
                     "queryGroup": e["queryGroup"], "dependsOn": deps, "sourceObjects": nav, "sourceFunctions": fns,
                     "code": e["definition"]})
    params = []
    for e in model["expressions"]:
        if e["kind"] != "Parameter":
            continue
        b = e["definition"]
        val, _, meta = b.partition(" meta ")
        g = lambda k: (re.search(k + r'\s*=\s*("[^"]*"|#date\([^)]*\)|\w+)', meta) or [None, None])[1]
        lst = re.search(r'List\s*=\s*\{([^}]*)\}', meta)
        clean = lambda x: x.replace('"', "") if isinstance(x, str) else x
        params.append({"name": e["name"], "currentValue": clean(val.strip()), "type": clean(g("Type")),
                       "defaultValue": clean(g("DefaultValue")), "allowedValues": clean(lst.group(1)) if lst else None,
                       "required": g("IsParameterQueryRequired"), "queryGroup": e["queryGroup"]})
    return rows, params


# =====================================================================================
# 5. REPORT (PBIR and legacy report.json)
# =====================================================================================
def _entity(sr, aliases):
    if not isinstance(sr, dict):
        return None
    if "Entity" in sr:
        return sr["Entity"]
    if aliases and sr.get("Source") in aliases:
        return aliases[sr["Source"]]
    return None


def field_info(fd, aliases=None):
    """(kind, table, field, aggregation_code) from a PBI query-expression, else None."""
    try:
        if not isinstance(fd, dict):
            return None
        agg = None
        if "Aggregation" in fd:
            agg = fd["Aggregation"].get("Function")
            fd = fd["Aggregation"].get("Expression", {})
        if "HierarchyLevel" in fd:
            h = fd["HierarchyLevel"]
            hx = h.get("Expression", {}).get("Hierarchy", {})
            return ("Hierarchy", _entity(hx.get("Expression", {}).get("SourceRef"), aliases), f'{hx.get("Hierarchy")}.{h.get("Level")}', agg)
        for kind in ("Column", "Measure"):
            if kind in fd and isinstance(fd[kind], dict):
                return (kind, _entity(fd[kind].get("Expression", {}).get("SourceRef"), aliases), fd[kind].get("Property"), agg)
    except Exception:
        return None
    return None


def walk_fields(obj, out: set, path="", depth=0, aliases=None):
    """Collect field references embedded in formatting / conditional-format expressions."""
    if isinstance(obj, dict):
        if any(k in obj for k in ("Column", "Measure", "Aggregation", "HierarchyLevel")):
            fi = field_info(obj, aliases)
            if fi and fi[1]:
                out.add((path, fi))
                return
        for k, v in obj.items():
            walk_fields(v, out, f"{path}.{k}" if (depth < 2 and path) else (k if not path else path), depth + 1, aliases)
    elif isinstance(obj, list):
        for v in obj:
            walk_fields(v, out, path, depth + 1, aliases)


def literals(o, out: list):
    if isinstance(o, dict):
        if "Literal" in o and isinstance(o["Literal"], dict):
            v = o["Literal"].get("Value")
            out.append(str(v).strip("'") if isinstance(v, str) else str(v))
        for v in o.values():
            literals(v, out)
    elif isinstance(o, list):
        for v in o:
            literals(v, out)


class ModelIndex:
    def __init__(self, model):
        self.idx = defaultdict(lambda: {"cols": set(), "meas": set()})
        for t in (model or {}).get("tables", []):
            for c in t["columns"]:
                self.idx[t["name"].lower()]["cols"].add((c["name"] or "").lower())
            for m in t["measures"]:
                self.idx[t["name"].lower()]["meas"].add((m["name"] or "").lower())
            for h in t["hierarchies"]:
                self.idx[t["name"].lower()]["cols"].add((h["name"] or "").lower())
        self.has_model = bool(model)

    def status(self, table, field):
        if not self.has_model:
            return "Model not available"
        if table is None:
            return "Unresolved alias"
        t = self.idx.get(table.lower()) if table.lower() in self.idx else None
        if t is None:
            return "Table not in model"
        base = (field or "").split(".")[0].lower()
        return "OK" if (field or "").lower() in t["cols"] | t["meas"] or base in t["cols"] else "Field not in model"


def _lit(v):
    return v.strip("'") if isinstance(v, str) else v


def filter_rows(level, owner, filters, mi: ModelIndex, aliases=None):
    rows = []
    for f in filters or []:
        fi = field_info(f.get("field") or f.get("expression"), aliases)
        vals = []
        literals(f.get("filter", {}), vals)
        rows.append({"level": level, "owner": owner, "filterId": f.get("name"),
                     "table": fi[1] if fi else None, "field": fi[2] if fi else None, "fieldKind": fi[0] if fi else None,
                     "filterType": f.get("type"), "values": ", ".join(dict.fromkeys(vals))[:2000] or None,
                     "hiddenInViewMode": f.get("isHiddenInViewMode"), "lockedInViewMode": f.get("isLockedInViewMode"),
                     "modelCheck": mi.status(fi[1], fi[2]) if fi else None})
    return rows


def visual_title(vis):
    for key in ("visualContainerObjects", "vcObjects"):
        try:
            return _lit(vis[key]["title"][0]["properties"]["text"]["expr"]["Literal"]["Value"])
        except Exception:
            pass
    return None


def visual_text(vis):
    try:
        runs = [r["value"] for p in vis["objects"]["general"][0]["properties"]["paragraphs"] for r in p["textRuns"]]
        return " ".join(x.strip() for x in runs if x.strip()) or None
    except Exception:
        return None


def load_pbir(rdir: Path, mi: ModelIndex, include_raw: bool) -> dict:
    ddir = rdir / "definition"
    rep = load_json(ddir / "report.json")
    order = load_json(ddir / "pages" / "pages.json") if (ddir / "pages" / "pages.json").exists() else {"pageOrder": []}
    page_dirs = sorted(p.name for p in (ddir / "pages").iterdir() if p.is_dir()) if (ddir / "pages").exists() else []
    page_ids = order.get("pageOrder") or page_dirs
    page_ids += [p for p in page_dirs if p not in page_ids]

    pages, visuals, vfields, filters = [], [], [], []
    filters += filter_rows("Report", "(all pages)", rep.get("filterConfig", {}).get("filters"), mi)
    for idx, pid in enumerate(page_ids, 1):
        pf = ddir / "pages" / pid / "page.json"
        if not pf.exists():
            continue
        pj = load_json(pf)
        vjs = {}
        for vf in sorted((ddir / "pages" / pid / "visuals").glob("*/visual.json")):
            v = load_json(vf)
            vjs[v.get("name", vf.parent.name)] = v
        groups = {k: v["visualGroup"].get("displayName") for k, v in vjs.items() if "visualGroup" in v}
        pb = pj.get("pageBinding", {})
        page = {"order": idx, "id": pid, "name": pj.get("displayName"), "activeOnOpen": pid == order.get("activePageName"),
                "visibility": pj.get("visibility", "Visible"),
                "pageType": "Standard" if pb.get("type") in (None, "Default") else pb.get("type"),
                "width": pj.get("width"), "height": pj.get("height"), "displayOption": pj.get("displayOption"),
                "visualCount": sum(1 for v in vjs.values() if "visual" in v), "groupCount": len(groups),
                "pageFilterCount": len(pj.get("filterConfig", {}).get("filters", [])),
                "drillthroughOrTooltipBinding": pb or None,
                "pageObjects": pj.get("objects") if include_raw else None,
                "visualInteractions": pj.get("visualInteractions") if include_raw else None}
        pages.append(page)
        filters += filter_rows("Page", pj.get("displayName"), pj.get("filterConfig", {}).get("filters"), mi)
        for vid, v in vjs.items():
            pos, vis = v.get("position", {}), v.get("visual", {})
            vtype = vis.get("visualType", "group" if "visualGroup" in v else None)
            fields = []
            for role, rd in (vis.get("query", {}).get("queryState", {}) or {}).items():
                for pr in rd.get("projections", []):
                    fi = field_info(pr.get("field"))
                    if fi:
                        fields.append((role, fi, pr.get("displayName") or pr.get("nativeQueryRef"), pr.get("active")))
            fmt = set()
            walk_fields({k: vis.get(k) for k in ("objects", "visualContainerObjects")}, fmt)
            rec = {"page": pj.get("displayName"), "pageId": pid, "id": vid, "type": vtype, "title": visual_title(vis),
                   "text": visual_text(vis) if vtype == "textbox" else None, "group": groups.get(v.get("parentGroupName")),
                   "hidden": bool(v.get("isHidden")),
                   "x": round(pos.get("x", 0)), "y": round(pos.get("y", 0)), "width": round(pos.get("width", 0)),
                   "height": round(pos.get("height", 0)), "z": pos.get("z"), "tabOrder": pos.get("tabOrder"),
                   "syncGroup": (vis.get("syncGroup") or {}).get("groupName"),
                   "fieldCount": len(fields), "fieldsUsed": "; ".join(dict.fromkeys(f"{fi[1]}[{fi[2]}]" for _, fi, _, _ in fields)) or None,
                   "visualFilterCount": len(v.get("filterConfig", {}).get("filters", []))}
            if include_raw:
                rec["definition"] = v
            visuals.append(rec)
            for role, fi, dn, active in fields:
                vfields.append({"page": pj.get("displayName"), "visualId": vid, "visualType": vtype, "title": rec["title"],
                                "role": role, "table": fi[1], "field": fi[2], "fieldKind": fi[0], "aggregationCode": fi[3],
                                "displayName": dn, "modelCheck": mi.status(fi[1], fi[2])})
            for path, fi in sorted(fmt, key=str):
                vfields.append({"page": pj.get("displayName"), "visualId": vid, "visualType": vtype, "title": rec["title"],
                                "role": f"Formatting: {path}", "table": fi[1], "field": fi[2], "fieldKind": fi[0],
                                "aggregationCode": fi[3], "displayName": None, "modelCheck": mi.status(fi[1], fi[2])})
            filters += filter_rows("Visual", f'{pj.get("displayName")} / {vtype} {vid}', v.get("filterConfig", {}).get("filters"), mi)

    bookmarks = []
    for bf in sorted((ddir / "bookmarks").glob("*.bookmark.json")) if (ddir / "bookmarks").exists() else []:
        b = load_json(bf)
        st = b.get("explorationState", {})
        bookmarks.append({"id": b.get("name"), "name": b.get("displayName"), "activePage": st.get("activeSection"),
                          "visualStates": sum(len(s.get("visualContainers", {})) for s in st.get("sections", {}).values()),
                          "definition": b if include_raw else None})
    ext_measures = []
    if (ddir / "reportExtensions.json").exists():
        for ent in load_json(ddir / "reportExtensions.json").get("entities", []):
            for m in ent.get("measures", []):
                ext_measures.append({"table": ent.get("name"), "name": m.get("name"), "dataType": m.get("dataType"),
                                     "formatString": m.get("formatString"), "expression": m.get("expression")})
    theme = rep.get("themeCollection", {})
    resources = [{"package": rp.get("name"), "name": i.get("name"), "path": i.get("path"), "type": i.get("type")}
                 for rp in rep.get("resourcePackages", []) for i in rp.get("items", [])]
    info = {"pbirFormat": "PBIR", "reportSchema": rep.get("$schema", "").split("/")[-2] if "/" in rep.get("$schema", "") else None,
            "baseTheme": theme.get("baseTheme", {}).get("name"), "customTheme": theme.get("customTheme", {}).get("name"),
            "settings": rep.get("settings"), "slowDataSourceSettings": rep.get("slowDataSourceSettings"),
            "reportObjects": rep.get("objects") if include_raw else None, "activePage": order.get("activePageName")}
    return {"info": info, "pages": pages, "visuals": visuals, "visualFields": vfields, "filters": filters,
            "bookmarks": bookmarks, "reportMeasures": ext_measures, "resources": resources,
            "pageOrderMismatch": sorted(set(order.get("pageOrder", [])) ^ set(page_dirs)) if order.get("pageOrder") else []}


def load_legacy_report(rdir: Path, mi: ModelIndex, include_raw: bool) -> dict:
    """Best-effort reader for report.json (PBIR-Legacy)."""
    rep = load_json(rdir / "report.json")
    js = lambda s: json.loads(s) if isinstance(s, str) and s.strip() else (s or {})
    cfg = js(rep.get("config"))
    pages, visuals, vfields, filters = [], [], [], []
    filters += filter_rows("Report", "(all pages)", js(rep.get("filters")) if isinstance(js(rep.get("filters")), list) else [], mi)
    for idx, sec in enumerate(rep.get("sections", []), 1):
        sfilters = js(sec.get("filters")) if isinstance(js(sec.get("filters")), list) else []
        pages.append({"order": idx, "id": sec.get("name"), "name": sec.get("displayName"), "activeOnOpen": idx == 1,
                      "visibility": "Visible", "pageType": "Standard", "width": sec.get("width"), "height": sec.get("height"),
                      "displayOption": sec.get("displayOption"), "visualCount": len(sec.get("visualContainers", [])),
                      "groupCount": 0, "pageFilterCount": len(sfilters), "drillthroughOrTooltipBinding": None,
                      "pageObjects": None, "visualInteractions": None})
        filters += filter_rows("Page", sec.get("displayName"), sfilters, mi)
        for vc in sec.get("visualContainers", []):
            c = js(vc.get("config"))
            sv = c.get("singleVisual", {})
            proto = sv.get("prototypeQuery", {})
            aliases = {f.get("Name"): f.get("Entity") for f in proto.get("From", [])}
            ref2role = {}
            for role, lst in (sv.get("projections") or {}).items():
                for pr in lst:
                    ref2role[pr.get("queryRef")] = role
            fields = []
            for sel in proto.get("Select", []):
                fi = field_info(sel, aliases)
                if fi:
                    fields.append((ref2role.get(sel.get("Name"), ""), fi))
            vtype = sv.get("visualType")
            rec = {"page": sec.get("displayName"), "pageId": sec.get("name"), "id": c.get("name"), "type": vtype,
                   "title": visual_title(sv), "text": visual_text(sv) if vtype == "textbox" else None, "group": None,
                   "hidden": bool(c.get("layouts", [{}])[0].get("position", {}).get("visibility") == 1) if False else False,
                   "x": round(vc.get("x", 0)), "y": round(vc.get("y", 0)), "width": round(vc.get("width", 0)),
                   "height": round(vc.get("height", 0)), "z": vc.get("z"), "tabOrder": None, "syncGroup": None,
                   "fieldCount": len(fields), "fieldsUsed": "; ".join(dict.fromkeys(f"{fi[1]}[{fi[2]}]" for _, fi in fields)) or None,
                   "visualFilterCount": 0}
            if include_raw:
                rec["definition"] = {k: (js(v) if k in ("config", "filters", "query", "dataTransforms") else v) for k, v in vc.items()}
            visuals.append(rec)
            for role, fi in fields:
                vfields.append({"page": sec.get("displayName"), "visualId": rec["id"], "visualType": vtype, "title": rec["title"],
                                "role": role, "table": fi[1], "field": fi[2], "fieldKind": fi[0], "aggregationCode": fi[3],
                                "displayName": None, "modelCheck": mi.status(fi[1], fi[2])})
            vf = js(vc.get("filters"))
            filters += filter_rows("Visual", f'{sec.get("displayName")} / {vtype} {rec["id"]}', vf if isinstance(vf, list) else [], mi, aliases)
    theme = cfg.get("themeCollection", {})
    return {"info": {"pbirFormat": "Legacy report.json (best effort)", "reportSchema": None,
                     "baseTheme": theme.get("baseTheme", {}).get("name"), "customTheme": theme.get("customTheme", {}).get("name"),
                     "settings": rep.get("settings"), "slowDataSourceSettings": None, "reportObjects": None, "activePage": None},
            "pages": pages, "visuals": visuals, "visualFields": vfields, "filters": filters, "bookmarks": [],
            "reportMeasures": [], "resources": [], "pageOrderMismatch": []}


def load_report(rdir: Path, model: dict | None, include_raw: bool) -> dict:
    mi = ModelIndex(model)
    if (rdir / "definition" / "report.json").exists():
        rep = load_pbir(rdir, mi, include_raw)
    elif (rdir / "report.json").exists():
        rep = load_legacy_report(rdir, mi, include_raw)
    else:
        raise SystemExit(f"No report definition found in {rdir}")
    try:
        p = load_json(rdir / ".platform")
        rep["info"]["displayName"] = p["metadata"]["displayName"]
        rep["info"]["logicalId"] = p["config"]["logicalId"]
    except Exception:
        pass
    try:
        rep["info"]["definitionPbirVersion"] = load_json(rdir / "definition.pbir").get("version")
    except Exception:
        pass
    static = [str(f.relative_to(rdir)).replace("\\", "/") for f in (rdir / "StaticResources").rglob("*") if f.is_file()] \
        if (rdir / "StaticResources").exists() else []
    rep["staticResources"] = static
    agg = defaultdict(lambda: {"n": 0, "pages": set(), "kind": None})
    for r in rep["visualFields"]:
        if r["modelCheck"] not in ("OK", "Model not available"):
            k = (r["table"], r["field"], r["modelCheck"])
            agg[k]["n"] += 1
            agg[k]["pages"].add(r["page"])
            agg[k]["kind"] = r["fieldKind"]
    rep["unresolvedReferences"] = [{"table": k[0], "field": k[1], "kind": v["kind"], "problem": k[2], "references": v["n"],
                                    "pages": sorted(v["pages"])} for k, v in sorted(agg.items(), key=lambda x: (-x[1]["n"], str(x[0])))]
    return rep


# =====================================================================================
# 6. VALIDATION + REBUILD KIT
# =====================================================================================
def validate(proj, model, report, skipped) -> list:
    v = []
    add = lambda c, s, d: v.append({"check": c, "status": s, "details": d})
    add("Semantic model found", "OK" if model else "WARN", model["format"] if model else "No local semantic model (live-connected report?)")
    if proj.get("dataset_ref"):
        ref = proj["dataset_ref"]
        if "byPath" in ref:
            add("Report -> model link (byPath)", "OK" if model else "FAIL", ref["byPath"].get("path"))
        else:
            add("Report -> model link", "INFO", f"byConnection: {json.dumps(ref.get('byConnection'))[:300]}")
    if model:
        names = {t["name"] for t in model["tables"]}
        cols = {(t["name"], c["name"]) for t in model["tables"] for c in t["columns"]}
        bad = [f'{r["fromTable"]}[{r["fromColumn"]}] -> {r["toTable"]}[{r["toColumn"]}]' for r in model["relationships"]
               if r["fromTable"] not in names or r["toTable"] not in names
               or (r["fromTable"], r["fromColumn"]) not in cols or (r["toTable"], r["toColumn"]) not in cols]
        add("Relationship endpoints exist", "OK" if not bad else "FAIL", f"{len(model['relationships'])} relationships" if not bad else "; ".join(bad))
        nomv = [t["name"] for t in model["tables"] if not t["partitions"]]
        add("Every table has a partition/source", "OK" if not nomv else "WARN", ", ".join(nomv) or "all tables")
        params = [e for e in model["expressions"] if e["kind"] == "Parameter"]
        add("Parameters that must be set before refresh", "INFO", ", ".join(e["name"] for e in params) or "none")
    if report:
        total = len(report["visualFields"])
        bad = [r for r in report["visualFields"] if r["modelCheck"] not in ("OK", "Model not available")]
        add("Report fields resolve against model", "OK" if not bad else "WARN",
            f"{total - len(bad)}/{total} resolve; {len(bad)} do not (see 'Unresolved Refs')")
        add("Page order matches page folders", "OK" if not report["pageOrderMismatch"] else "WARN",
            ", ".join(report["pageOrderMismatch"]) or f"{len(report['pages'])} pages")
    if skipped:
        add("Cached data skipped (not needed)", "INFO", "; ".join(f'{s["file"]} ({s["bytes"] / 1e6:.1f} MB)' for s in skipped[:5]))
    add("Data included in metadata", "INFO", "No. Refresh after opening the rebuilt project (credentials required).")
    return v


def _copy_tree(src: Path, dst: Path) -> None:
    for f in src.rglob("*"):
        if f.is_file() and not f.name.lower().endswith(CACHE_EXT):
            t = dst / f.relative_to(src)
            t.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(f, t)


def rebuild_project(proj: dict, out: Path) -> dict:
    dst = out / "rebuilt_project"
    if dst.exists():
        shutil.rmtree(dst)
    dst.mkdir(parents=True)
    base = proj["pbip"].parent if proj["pbip"] else None
    for key in ("report_dir", "model_dir"):
        d = proj[key]
        if d:
            rel = d.relative_to(base) if base and base in d.parents else Path(d.name)
            _copy_tree(d, dst / rel)
    pbip_name = proj["pbip"].name if proj["pbip"] else f'{proj["name"]}.pbip'
    if proj["pbip"]:
        shutil.copy2(proj["pbip"], dst / pbip_name)
    else:
        rp = proj["report_dir"].name if proj["report_dir"] else None
        (dst / pbip_name).write_text(json.dumps({
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
            "version": "1.0", "artifacts": [{"report": {"path": rp}}] if rp else [], "settings": {"enableAutoRecovery": True}}, indent=2))
    files = []
    for f in sorted(p for p in dst.rglob("*") if p.is_file()):
        files.append({"path": str(f.relative_to(dst)).replace("\\", "/"), "bytes": f.stat().st_size,
                      "sha256": hashlib.sha256(f.read_bytes()).hexdigest()})
    zpath = out / "rebuilt_project.zip"
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(p for p in dst.rglob("*") if p.is_file()):
            z.write(f, f.relative_to(dst))
    return {"folder": str(dst), "zip": str(zpath), "files": files, "pbip": pbip_name}
