from __future__ import annotations

import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


def _xml(path: Path, include_raw: bool):
    """Extract high-level Tableau metadata without loading the whole XML tree."""
    worksheets = []
    dashboards = []
    datasources = []

    # iterparse keeps memory bounded for very large TWB XML documents.
    for event, elem in ET.iterparse(path, events=("end",)):
        tag = elem.tag.rsplit("}", 1)[-1]
        if tag == "worksheet":
            worksheets.append({"name": elem.get("name", ""), "metadata": {}})
        elif tag == "dashboard":
            dashboards.append({"name": elem.get("name", ""), "metadata": {}})
        elif tag == "datasource":
            datasources.append({"name": elem.get("name", ""), "metadata": {}})
        elem.clear()

    raw_xml = None
    if include_raw:
        # Raw XML is intentionally opt-in. It can be enormous for large TWBs.
        raw_xml = path.read_text(encoding="utf-8", errors="replace")

    return {
        "identity": {"name": path.stem, "source": str(path)},
        "platform": "tableau",
        "format": "twb",
        "version": None,
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


def _copy_zip_member_to_temp(
    zf: zipfile.ZipFile, member: zipfile.ZipInfo, destination: Path
) -> Path:
    """Stream only the TWB member from a TWBX archive to disk."""
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
            with zipfile.ZipFile(source, "r") as zf:
                members = [
                    info
                    for info in zf.infolist()
                    if not info.is_dir()
                    and info.filename.lower().endswith(".twb")
                ]
                if not members:
                    raise ValueError("No TWB file found inside TWBX")

                # A normal TWBX contains one workbook definition. We extract
                # only that definition and deliberately ignore large .hyper
                # and other packaged assets during metadata parsing.
                member = members[0]
                twb_path = _copy_zip_member_to_temp(
                    zf, member, Path(td) / Path(member.filename).name
                )

            return _xml(twb_path, include_raw)

    raise ValueError(f"Unsupported Tableau input: {source}")
