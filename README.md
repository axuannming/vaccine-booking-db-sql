# Vaccine Registration & Safety Monitoring Database (SQL & Access Automation)

A 17-table relational database design, DDL/DML SQL suite, 20 analytical/workflow queries, and PowerShell Access COM automation scripts for a **Vaccine Registration and Safety Monitoring System** (`SEHH2042` Database Systems Group Project).

The repository contains SQL, automation scripts, and documentation. The sample records use synthetic names, reserved email addresses, demonstration identity codes, and non-routable phone placeholders. Team member identities and identifying fields in the report screenshots have been redacted.

## Tech Stack
- **Database Engine Target:** Microsoft Access SQL / Jet / ACE DDL & DML
- **Automation:** Windows PowerShell (`Access.Application` COM automation for table/query creation, UI forms, reports, and form refresh patching)

## Schema Overview (17 Tables)
1. `Citizens` (40 synthetic records)
2. `Vaccines` (30 records)
3. `VaccinationCenters` (30 HK districts/centers)
4. `VaccineBatches` (40 batches)
5. `Users` (30 records)
6. `AppointmentSlots` (60 slots)
7. `Appointments` (40 records)
8. `VaccinationRecords` (36 records)
9. `PreVaccinationScreening` (40 records)
10. `ConsentRecords` (40 records)
11. `ObservationRecords` (36 records)
12. `SideEffects` (35 records)
13. `AdverseEffectFollowUps` (30 records)
14. `Notifications` (40 records)
15. `VaccineAllocation` (40 records)
16. `InventoryTransactions` (40 records)
17. `AuditLog` (40 records)

## Directory Structure
```text
.
├── sql/
│   ├── 01_create_tables.sql               # 17-table schema DDL
│   ├── 02_relationships_and_indexes.sql   # Foreign keys & indexes
│   ├── 03_insert_sample_data.sql          # Synthetic demo seed data (@example.com)
│   └── queries/                           # 20 saved workflow & analytics SQL queries
├── scripts/
│   ├── build_access_v2_data.ps1           # Populates schema, constraints, queries, and v2 sample data
│   ├── build_access_ui.ps1                # Generates portal, dashboard, and entry forms
│   ├── build_missing_v2_forms.ps1         # Adds screening/consent/observation/adverse-follow-up forms
│   ├── build_reports.ps1                  # Generates Access reports
│   ├── rebuild_appointment_and_screening.ps1
│   ├── rebuild_form_v2.ps1
│   └── patch_access_entry_save_refresh.ps1 # COM export/import patch for form save & refresh behavior
├── docs/
│   ├── Final_Report.docx                  # Anonymized database design report (ERD, data dictionary, forms)
│   ├── GenAI_Tools_Declaration.docx       # Anonymized project GenAI declaration
│   ├── CHAT_HANDOFF_2026-03-25.md         # Engineering handoff & blocker diagnosis
│   └── patch_log_sample.txt               # Verified COM form patch log (8 entry forms compiled OK)
├── .gitignore
└── README.md
```

## How to Rebuild Locally (Microsoft Access on Windows)
1. Create a blank `.accdb` file in Microsoft Access, or run the SQL scripts in order inside Access SQL View:
   1. `sql/01_create_tables.sql`
   2. `sql/02_relationships_and_indexes.sql`
   3. `sql/03_insert_sample_data.sql`
   4. Saved queries under `sql/queries/`
2. Optionally run the PowerShell COM builders under `scripts/` on Windows with Microsoft Access installed.

## Authorship & Attribution
- **Group Project:** Four-member team project for `SEHH2042` Database Systems. The report credits the team as `Member 1–4`. The archived material does not establish an individual split for the SQL and Access automation files.

## Development Difficulties, AI Prompting & Solutions
Summarized directly from `docs/CHAT_HANDOFF_2026-03-25.md`, `docs/GenAI_Tools_Declaration.docx`, and `docs/patch_log_sample.txt`:
1. **Unreliable VBA `[Event Procedure]` Button Bindings in Generated Forms:**
   - *Difficulty:* Programmatically created Access forms initially bound buttons to `OnClick = "[Event Procedure]"`, which failed to persist reliably across machines when opened without external scripts.
   - *Solution:* Documented in `CHAT_HANDOFF_2026-03-25.md` to migrate button logic to native Access embedded macros (`OpenForm`, `OpenQuery`, `GoToRecord`, `RunCommand SaveRecord`), and built `patch_access_entry_save_refresh.ps1` to export, update, re-import, and compile (`COMPILE OK`) all 8 entry forms cleanly.
2. **GenAI Usage (`docs/GenAI_Tools_Declaration.docx`):**
   - Used **Perplexity** to research domain references, **Gemini** to help generate realistic synthetic test rows, and **ChatGPT** to help characterize and verify the 17-table relational schema and foreign-key relationships.
