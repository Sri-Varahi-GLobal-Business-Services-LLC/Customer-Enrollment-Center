# Tools — Document & Demo Generators

Python scripts that generate (and regenerate) the Enrollment Center deliverables.
Requirements: Python 3.13+, `python-docx`, `openpyxl`. Run with `py -X utf8 <script>`.
Output paths are hardcoded inside each script — adjust if the folder structure moves.

## Generators (safe to re-run; they overwrite their output file)

| Script | Regenerates |
|---|---|
| `gen_demo_pages.py` | The three demo pages in `Desktop\enrollment-center-demo\` (index / workcenter / movein.html) — **single source of truth for the demo and the program catalog** |
| `gen_ec_tdd.py` | Technical Design (S4HANA Data Sources, APIs, CDS, BRFplus) |
| `gen_ec_config_guide.py` | BTP Deployment and Configuration Guide |
| `gen_ec_brfplus.py` | BRFplus Detailed Configuration |
| `gen_brf_workbook.py` | BRFplus Configuration Workbook (xlsx) |
| `gen_ec_runtime.py` | Runtime Service Interactions and Data Exchange |
| `gen_ui_eval.py` | UI Technology Evaluation |
| `gen_ec_ai_roadmap.py` | AI/ML Roadmap |
| `extract_catalog.py` | `Enrollment Center/program-catalog.json` from the demo CATALOG (CAP seed data) |

## One-shot patch scripts (already applied — do NOT re-run)

`add_demo_link.py`, `add_sdd_appendix.py`, `fix_xrefs.py`, `update_sdd_programs.py`,
`patch_deck_and_sdd.py`, `patch_sdd_ai_deck.py` performed in-place edits on the Solution Design v2.0 docx and the
Client Presentation pptx. Re-running would duplicate content.

## Not recoverable

Generators for the Solution Design v1.0/v2.0 base document, the CS-01/CS-02 BPDs and the
Client Presentation were lost to scratchpad wipes; those files are maintained directly in
Word/PowerPoint (or via one-shot patch scripts) from now on. **This tools/ folder is the canonical home for all generators — never keep the
only copy in a session scratchpad.**

## Workflow

1. Change demo/catalog: edit `gen_demo_pages.py` → run it → run `extract_catalog.py` →
   republish artifacts / push `enrollment-center-demo` repo.
2. Change a document: edit its generator → run → verify in Word → commit.
