import argparse
import json
import os
from typing import List


def pretty_one(path: str, indent: int = 2, sort_keys: bool = False):
    with open(path, "r", encoding="utf-8") as f:
        obj = json.load(f)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(obj, ensure_ascii=False, indent=indent, sort_keys=sort_keys))
        f.write("\n")
    print(path)


def main():
    ap = argparse.ArgumentParser(prog="pretty-json", description="Pretty print JSON files in place")
    ap.add_argument("paths", nargs="+", help="One or more JSON file paths")
    ap.add_argument("--indent", type=int, default=2)
    ap.add_argument("--sort-keys", action="store_true")
    args = ap.parse_args()
    for p in args.paths:
        if os.path.isdir(p):
            for name in os.listdir(p):
                if name.lower().endswith(".json"):
                    pretty_one(os.path.join(p, name), indent=args.indent, sort_keys=args.sort_keys)
        else:
            pretty_one(p, indent=args.indent, sort_keys=args.sort_keys)


if __name__ == "__main__":
    main()
