from __future__ import annotations

import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


def _xml(path: Path, include_raw: bool):
    root = ET.parse(path).getroot()

    worksheets = [
        {"name": ws.get("name", ""), "metadata": {}}
        for ws in root.findall(".//worksheet")
    ]
    dashboards = [
        {"name": d.get("name", ""), "metadata": {}}
        for d in root.findall(".//dashboard")
    ]
    datasources = [
        {"name": d.get("name", ""), "metadata": {}}
        for d in root.findall(".//datasource")
    ]

    raw_xml = None
    if include_raw:
        # This is intentionally opt-in because a large TWB can make the BI-IR
        # substantially larger than the source file.
        raw_xml = ET.tostring(root, encoding="unicode")

    return {
        "identity": {"name": path.stem, "source": str(path)},
        "platform": "tableau",
        "format": "twb",
        "version": root.get("version"),
        "metadata": {
            "parser": "BI-X Tableau adapter",
            "source_size_bytes": path.stat().st_size,
        },
        "semantic_models": [
            {
                "name": path.stem,
                "tables": [],
                "relationships": [],
                "roles": [],
                "perspectives": [],
                "parameters": [],
                "expressions": [],
                "data_sources": datasources,
                "annotations": {},
                "extensions": {"worksheets": worksheets},
            }
        ],
        "reports": [
            {
                "name": path.stem,
                "pages": dashboards,
                "visuals": worksheets,
                "filters": [],
                "bookmarks": [],
                "themes": [],
                "resources": [],
                "extensions": {"raw_xml": raw_xml},
            }
        ],
        "data_sources": datasources,
        "resources": [],
        "security": {},
        "lineage": {},
        "validation": {},
        "extensions": {},
    }


def _copy_zip_member_to_temp(zf: zipfile.ZipFile, member: zipfile.ZipInfo, destination: Path) -> Path:
    """Stream a single TWB member from a TWBX to disk.

    We deliberately do not call ZipFile.extractall(): TWBX files can contain
    very large Hyper extracts and other assets that are not needed to parse
    the workbook.
    """
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zf.open(member, "r") as source, destination.open("wb") as target:
        shutil.copyfileobj(source, target, length=1024 * 1024)
    return destination


def extract(source: Path, include_raw: bool = False):
    source = Path(source)
    suffix = source.suffix.lower()

    if suffix == ".twb":
        return _xml(source, include_raw)

    if suffix == ".twbx":
        with tempfile.TemporaryDirectory(prefix="bix-twbx-") as td:
            twb_path = None
            with zipfile.ZipFile(source, "r") as zf:
                # Prefer the smallest/most direct workbook member if there are
                # multiple TWBs, while never extracting unrelated large assets.
                members = [
                    info
                    for info in zf.infolist()
                    if not info.is_dir() and info.filename.lower().endswith(".twb")
                ]
                if not members:
                    raise ValueError("No TWB file found inside TWBX")

                # A TWBX normally contains one TWB. If there are multiple,
                # use the first workbook entry in archive order.
                member = members[0]
                safe_name = Path(member.filename).name or "workbook.twb"
                twb_path = _copy_zip_member_to_temp(
                    zf, member, Path(td) / safe_name
                )

            return _xml(twb_path, include_raw)

    raise ValueError(f"Unsupported Tableau input: {source}")
