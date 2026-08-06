# -*- coding: utf-8 -*-
"""Extract the program catalog from the demo pages into CAP-ready program-catalog.json.

Source of truth for demo + seed data: the CATALOG object in the generated demo HTML
(enrollment-center-demo/index.html, produced by tools/gen_demo_pages.py).
Output: Enrollment Center/program-catalog.json — array of program entries with
captureForm metadata matching the TDD ProgramCatalog entity.
"""
import io, json, re, sys, os

SRC = r"C:\Users\jnamm\OneDrive\Desktop\enrollment-center-demo\index.html"
OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "Enrollment Center", "program-catalog.json")

def js_object_to_json(src: str) -> str:
    """Convert the demo's JS object literal (unquoted keys, double-quoted strings) to JSON."""
    out = re.sub(r'([,{]\s*)([A-Za-z_][A-Za-z0-9_]*)\s*:', r'\1"\2":', src)
    out = out.replace("&mdash;", "\u2014").replace("&amp;", "&")
    return out

def main():
    with io.open(SRC, encoding="utf-8") as f:
        html = f.read()
    m = re.search(r'const CATALOG = (\{.*?\n\});', html, re.S)
    if not m:
        sys.exit("CATALOG object not found in " + SRC)
    catalog = json.loads(js_object_to_json(m.group(1)))

    entries = []
    for pid, c in catalog.items():
        entries.append({
            "programId": pid,
            "version": 1,
            "name": c["name"],
            "category": c["cat"],
            "recipeType": c["recipe"],
            "terms": c.get("terms", ""),
            "active": True,
            "captureForm": {
                "consentStatementId": f"{pid}_TERMS_V3",
                "fields": [
                    {
                        "key": f["k"],
                        "label": f["l"],
                        "type": f["t"],
                        **({"options": f["opts"]} if "opts" in f else {}),
                        **({"min": f["min"]} if "min" in f else {}),
                        **({"max": f["max"]} if "max" in f else {}),
                        "required": True,
                    }
                    for f in c.get("fields", [])
                ],
            },
        })

    with io.open(OUT, "w", encoding="utf-8") as f:
        json.dump({"programCatalog": entries,
                   "_source": "extracted from enrollment-center-demo (tools/extract_catalog.py); "
                              "seed data for the CAP ProgramCatalog entity per TDD section 8.1"},
                  f, indent=2, ensure_ascii=False)
    print(f"Wrote {OUT} with {len(entries)} programs")

if __name__ == "__main__":
    main()
