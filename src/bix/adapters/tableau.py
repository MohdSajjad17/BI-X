from __future__ import annotations
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

def _xml(path: Path, include_raw: bool):
    root=ET.parse(path).getroot()
    worksheets=[]
    for ws in root.findall(".//worksheet"):
        worksheets.append({"name":ws.get("name",""),"metadata":{}})
    dashboards=[{"name":d.get("name",""),"metadata":{}} for d in root.findall(".//dashboard")]
    datasources=[{"name":d.get("name",""),"metadata":{}} for d in root.findall(".//datasource")]
    return {"identity":{"name":path.stem,"source":str(path)},"platform":"tableau","format":"twb","version":root.get("version"),"metadata":{"parser":"BI-X Tableau adapter"},"semantic_models":[{"name":path.stem,"tables":[],"relationships":[],"roles":[],"perspectives":[],"parameters":[],"expressions":[],"data_sources":datasources,"annotations":{},"extensions":{"worksheets":worksheets}}],"reports":[{"name":path.stem,"pages":dashboards,"visuals":worksheets,"filters":[],"bookmarks":[],"themes":[],"resources":[],"extensions":{"raw_xml":ET.tostring(root,encoding="unicode") if include_raw else None}}],"data_sources":datasources,"resources":[],"security":{},"lineage":{},"validation":{},"extensions":{}}

def extract(source: Path, include_raw=False):
    source=Path(source)
    if source.suffix.lower()==".twb":
        return _xml(source,include_raw)
    if source.suffix.lower()==".twbx":
        import tempfile
        with tempfile.TemporaryDirectory() as td:
            root=Path(td)
            with zipfile.ZipFile(source) as zf:
                zf.extractall(root)
            twbs=list(root.rglob("*.twb"))
            if not twbs: raise ValueError("No TWB file found inside TWBX")
            return _xml(twbs[0],include_raw)
    raise ValueError(f"Unsupported Tableau input: {source}")
