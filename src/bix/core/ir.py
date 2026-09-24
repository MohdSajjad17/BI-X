from __future__ import annotations
from dataclasses import dataclass, field, asdict
from typing import Any
import json

IR_VERSION = "1.0"

@dataclass
class Expression:
    language: str
    code: str
    name: str | None = None
    metadata: dict[str, Any] = field(default_factory=dict)

@dataclass
class Column:
    name: str
    data_type: str | None = None
    expression: str | None = None
    metadata: dict[str, Any] = field(default_factory=dict)

@dataclass
class Measure:
    name: str
    expression: str | None = None
    metadata: dict[str, Any] = field(default_factory=dict)

@dataclass
class Table:
    name: str
    columns: list[Column] = field(default_factory=list)
    measures: list[Measure] = field(default_factory=list)
    partitions: list[dict[str, Any]] = field(default_factory=list)
    hierarchies: list[dict[str, Any]] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)

@dataclass
class SemanticModel:
    name: str
    tables: list[Table] = field(default_factory=list)
    relationships: list[dict[str, Any]] = field(default_factory=list)
    roles: list[dict[str, Any]] = field(default_factory=list)
    perspectives: list[dict[str, Any]] = field(default_factory=list)
    parameters: list[dict[str, Any]] = field(default_factory=list)
    expressions: list[Expression] = field(default_factory=list)
    data_sources: list[dict[str, Any]] = field(default_factory=list)
    annotations: dict[str, Any] = field(default_factory=dict)
    extensions: dict[str, Any] = field(default_factory=dict)

@dataclass
class Report:
    name: str
    pages: list[dict[str, Any]] = field(default_factory=list)
    visuals: list[dict[str, Any]] = field(default_factory=list)
    filters: list[dict[str, Any]] = field(default_factory=list)
    bookmarks: list[dict[str, Any]] = field(default_factory=list)
    themes: list[dict[str, Any]] = field(default_factory=list)
    resources: list[dict[str, Any]] = field(default_factory=list)
    extensions: dict[str, Any] = field(default_factory=dict)

@dataclass
class Project:
    identity: dict[str, Any]
    platform: str
    format: str
    version: str | None = None
    metadata: dict[str, Any] = field(default_factory=dict)
    semantic_models: list[SemanticModel] = field(default_factory=list)
    reports: list[Report] = field(default_factory=list)
    data_sources: list[dict[str, Any]] = field(default_factory=list)
    resources: list[dict[str, Any]] = field(default_factory=list)
    security: dict[str, Any] = field(default_factory=dict)
    lineage: dict[str, Any] = field(default_factory=dict)
    validation: dict[str, Any] = field(default_factory=dict)
    extensions: dict[str, Any] = field(default_factory=dict)

def to_dict(project):
    return asdict(project) if hasattr(project, "__dataclass_fields__") else project

def dumps(project):
    return json.dumps(to_dict(project), indent=2, ensure_ascii=False, default=str)
