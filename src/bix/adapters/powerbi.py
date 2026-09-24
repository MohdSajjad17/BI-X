from __future__ import annotations

import json
import tempfile
from pathlib import Path
from typing import Any

from . import powerbi_reference as reference


def _table_to_ir(table: dict[str, Any]) -> dict[str, Any]:
    return {
        "name": table.get("name"),
        "columns": [
            {
                "name": c.get("name"),
                "data_type": c.get("dataType"),
                "expression": c.get("expression"),
                "metadata": c,
            }
            for c in table.get("columns", []) or []
        ],
        "measures": [
            {
                "name": m.get("name"),
                "expression": m.get("expression"),
                "metadata": m,
            }
            for m in table.get("measures", []) or []
        ],
        "partitions": table.get("partitions", []) or [],
        "hierarchies": table.get("hierarchies", []) or [],
        "metadata": table,
    }


def _model_to_ir(model: dict[str, Any], name: str) -> dict[str, Any]:
    return {
        "name": name,
        "tables": [_table_to_ir(t) for t in model.get("tables", []) or []],
        "relationships": model.get("relationships", []) or [],
        "roles": model.get("roles", []) or [],
        "perspectives": model.get("perspectives", []) or [],
        "parameters": model.get("expressions", []) or [],
        "expressions": model.get("expressions", []) or [],
        "data_sources": model.get("otherObjects", []) or [],
        "annotations": model.get("annotations", {}) or {},
        "extensions": {
            "reference_model": model,
            "format": model.get("format"),
            "info": model.get("info", {}),
            "queryOrder": model.get("queryOrder", []),
            "queryGroups": model.get("queryGroups", []),
            "cultures": model.get("cultures", []),
            "daxQueries": model.get("daxQueries", []),
        },
    }


def _report_to_ir(report: dict[str, Any] | None, name: str) -> dict[str, Any]:
    if not report:
        return {
            "name": name,
            "pages": [],
            "visuals": [],
            "filters": [],
            "bookmarks": [],
            "themes": [],
            "resources": [],
            "extensions": {},
        }

    return {
        "name": name,
        "pages": report.get("pages", []) or [],
        "visuals": report.get("visuals", []) or [],
        "filters": report.get("filters", []) or [],
        "bookmarks": report.get("bookmarks", []) or [],
        "themes": report.get("themes", []) or [],
        "resources": report.get("staticResources", []) or [],
        "extensions": {
            "reference_report": report,
            "format": report.get("format"),
            "info": report.get("info", {}),
            "unresolvedReferences": report.get("unresolvedReferences", []),
        },
    }


def extract(source: Path, include_raw: bool = False) -> dict[str, Any]:
    """Extract Power BI metadata using the original BI-X reference implementation.

    The reference parser is the functional baseline. This adapter wraps its rich
    metadata in the canonical BI-IR rather than replacing its parsing behavior.
    ZIP/PBIP inputs are materialised to a temporary disk workspace, so large
    archives are not loaded into a second in-memory bytes buffer.
    """
    source = Path(source)

    with tempfile.TemporaryDirectory(prefix="bix-powerbi-") as td:
        work = Path(td)
        skipped: list[dict[str, Any]] = []

        root = reference.materialise(str(source), work, skipped)
        project = reference.locate_project(root)

        model = None
        if project.get("model_dir"):
            model = reference.load_model(project["model_dir"], full_culture=False)

        power_query, parameters = (
            reference.build_power_query(model) if model else ([], [])
        )

        report = None
        if project.get("report_dir"):
            report = reference.load_report(
                project["report_dir"],
                model,
                include_raw=include_raw,
            )

        validation = reference.validate(
            project,
            model,
            report,
            skipped,
        )

        project_name = project["name"]
        reference_metadata = {
            "tool": f"pbip_reverse_engineer.py v{reference.TOOL_VERSION}",
            "project": {
                "name": project_name,
                "pbip": project["pbip"].name if project.get("pbip") else None,
                "pbipContent": (
                    reference.load_json(project["pbip"])
                    if project.get("pbip")
                    else None
                ),
                "reportFolder": (
                    project["report_dir"].name
                    if project.get("report_dir")
                    else None
                ),
                "modelFolder": (
                    project["model_dir"].name
                    if project.get("model_dir")
                    else None
                ),
                "datasetReference": project.get("dataset_ref"),
            },
            "semanticModel": model,
            "powerQuery": power_query,
            "parameters": parameters,
            "report": report,
            "skippedCacheFiles": skipped,
            "validation": validation,
        }

        sm = _model_to_ir(model, project_name) if model else {
            "name": project_name,
            "tables": [],
            "relationships": [],
            "roles": [],
            "perspectives": [],
            "parameters": [],
            "expressions": [],
            "data_sources": [],
            "annotations": {},
            "extensions": {},
        }

        rp = _report_to_ir(report, project_name)

        return {
            "identity": {
                "name": project_name,
                "source": str(source),
            },
            "platform": "powerbi",
            "format": "pbip",
            "version": (
                model.get("info", {}).get("pbismVersion")
                if model
                else None
            ),
            "metadata": {
                "parser": "BI-X reference Power BI reverse engineer",
                "reference_implementation": "pbip_reverse_engineer.py",
                "reference_tool_version": reference.TOOL_VERSION,
            },
            "semantic_models": [sm],
            "reports": [rp],
            "data_sources": (
                model.get("otherObjects", []) if model else []
            ),
            "resources": report.get("staticResources", []) if report else [],
            "security": {
                "roles": model.get("roles", []) if model else [],
            },
            "lineage": {},
            "validation": {
                "reference_checks": validation,
            },
            "extensions": {
                "reference_metadata": reference_metadata,
                "skipped_cache_files": skipped,
                "source_type": source.suffix.lower() if source.is_file() else "folder",
            },
        }


def parse_tmdl_file(path: Path) -> dict[str, Any]:
    """Compatibility helper retained for callers of the previous adapter."""
    return reference._tmdl_table(
        reference.read_text(Path(path))
        if hasattr(reference, "read_text")
        else Path(path).read_text(encoding="utf-8-sig", errors="replace"),
        Path(path).stem,
    )
