# Service Cloud for Utilities

Working repository for SAP Service Cloud Version 2 (SC V2) utilities implementation assets:
solution designs, business process design (BPD) documents, legacy CRM reference material,
and the Customer Enrollment Center engineering set.

> **Confidentiality:** contains client-specific implementation documents and reference
> material. Keep this repository **private**.

## Enrollment Center — deliverable set (in `Enrollment Center/`)

| Document | Content |
|---|---|
| Solution Design v2.0 | Architecture, dual access points, flows, status model, user stories, UI technology appendix |
| Technical Design v1.2 | S/4HANA data sources (CDS/tables), API inventory, BRFplus design, per-utility device + DER context, custom developments D-01…D-20 |
| BTP Deployment & Configuration Guide v1.0 | BTP account/services, Cloud Connector, destinations, iFlows, Event Mesh, S/4HANA config + WRICEF specs |
| BRFplus Detailed Configuration v1.2 | ZEC_PROGRAM_ELIGIBILITY: data objects, expressions, 7 decision tables, 28 rulesets, 34 messages |
| BRFplus Configuration Workbook v1.0 (xlsx) | Decision-table content ready for BRFplus Excel import |
| Runtime Service Interactions v1.1 | Services per system and data exchange per user action (A1–A11), payloads incl. device/DER model |
| UI Technology Evaluation v1.0 | BTP vs RAP-on-S/4 vs Web Dynpro — weighted decision record |
| Client Presentation v1.0 (pptx) | 13-slide stakeholder deck |
| `program-catalog.json` | 28-program CAP-ready seed extracted from the demo |

## BPD documents (repo root)

- CS-01 Front Office Activities — Business Process Design v1.0
- CS-02 Business Master Data — Business Process Design v1.0

## Live demo (fictitious data, public)

- Repo: https://github.com/Sri-Varahi-GLobal-Business-Services-LLC/enrollment-center-demo
- Live: https://sri-varahi-global-business-services-llc.github.io/enrollment-center-demo/
  (also `/workcenter.html`, `/movein.html`) — custom domain `demo.svgbs.com` pending DNS
- Claude artifact mirrors: combined `d4ef964a…`, work center `d713d454…`, move-in `752b4f84…`

## CAP application scaffold

`Desktop\enrollment-center-app` (local; GitHub repo to be created) — CAP service implementing
EC-01…EC-09 with the seeded program catalog.

## Tools

`tools/` — Python generators for every document above (see `tools/README.md`). The demo pages
and program catalog are generated from `tools/gen_demo_pages.py` — edit there, never the HTML.

## Reference Material

- `Enrollment Center/` — legacy CRM Enrollment Center proposal, functional/technical specs (DPP, PTR, CPP, PrePay, CSDD), test case library, rate category eligibility function modules
- `SAP Standard Material` (Desktop) — S/4HANA Utilities CDS view catalog, API guides
- Excluded from git (size/sensitivity): `SO-Appointment/`, `FSM and EAM/`, `DRMS/`, `Implementation and Business Process/`, `DER Programs/` (contains real SMUD extracts — never push), `*.mp4`
