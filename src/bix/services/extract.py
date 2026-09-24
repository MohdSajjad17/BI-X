from pathlib import Path
from bix.adapters.powerbi import extract as powerbi_extract
from bix.adapters.tableau import extract as tableau_extract

def extract_project(source: Path, include_raw=False):
    ext = source.suffix.lower()
    if ext in {".twb", ".twbx"}:
        return tableau_extract(source, include_raw)
    return powerbi_extract(source, include_raw)
