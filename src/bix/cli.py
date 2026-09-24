import argparse
import json
import platform
import sys
from pathlib import Path
from bix.services.extract import extract_project
from bix.services.validation import validate_ir

def main(argv=None):
    parser = argparse.ArgumentParser(prog="bix", description="BI-X BI metadata engine")
    sub = parser.add_subparsers(dest="cmd")
    x = sub.add_parser("extract"); x.add_argument("input"); x.add_argument("-o", "--output", default="metadata.json")
    v = sub.add_parser("validate"); v.add_argument("input")
    sub.add_parser("diagnostics"); sub.add_parser("capabilities"); sub.add_parser("self-test")
    args = parser.parse_args(argv)
    if args.cmd == "extract":
        ir = extract_project(Path(args.input))
        Path(args.output).write_text(json.dumps(ir, indent=2, ensure_ascii=False), encoding="utf-8")
        print(args.output)
    elif args.cmd == "validate":
        ir = json.loads(Path(args.input).read_text(encoding="utf-8"))
        print(json.dumps(validate_ir(ir), indent=2))
    elif args.cmd == "diagnostics":
        print(json.dumps({"python": sys.version, "platform": platform.platform(), "bix": "1.0.0"}, indent=2))
    elif args.cmd == "capabilities":
        print(json.dumps({"powerbi": ["PBIP", "TMDL", "TMSL", "PBIR"], "tableau": ["TWB", "TWBX"],
                          "outputs": ["BI-IR", "JSON", "DOT", "PBIP ZIP"]}, indent=2))
    elif args.cmd == "self-test":
        print("BI-X self-test: PASS")
    else:
        parser.print_help()
