import re
import uuid
import datetime

def translate_expression(expression, source, target, ir=None):
    if not expression.strip():
        return {"translation_id": str(uuid.uuid4()), "status": "failed", "warnings": ["Empty expression"]}
    if source == target:
        return {"translation_id": str(uuid.uuid4()), "source_language": source, "target_language": target,
                "source_expression": expression, "target_expression": expression, "status": "exact",
                "confidence": "high", "warnings": [], "manual_review_required": False}
    translated = expression
    warnings = []
    if source == "TABLEAU_CALC" and target == "DAX":
        translated = re.sub(r"\bIF\s*\(", "IF(", expression, flags=re.I)
        translated = translated.replace(" THEN ", " , ").replace(" ELSE ", " , ")
        warnings.append("Tableau and DAX semantics can differ; review aggregation, filter and date behavior.")
    elif source == "DAX" and target == "TABLEAU_CALC":
        translated = expression
        warnings.append("Automatic DAX-to-Tableau conversion is heuristic and requires semantic review.")
    else:
        return {"translation_id": str(uuid.uuid4()), "status": "unsupported",
                "source_expression": expression, "target_expression": None,
                "warnings": [f"Unsupported conversion: {source}->{target}"]}
    return {
        "translation_id": str(uuid.uuid4()),
        "timestamp": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "source_language": source, "target_language": target,
        "source_expression": expression, "target_expression": translated,
        "status": "translated", "confidence": "low", "warnings": warnings,
        "manual_review_required": True,
    }
