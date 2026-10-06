# SAP lead-gen (Actran Systems)

Purpose: generate Excel prospect lists of Thai companies for SAP Business One / S/4HANA Cloud Public Edition.

When the user asks for a new lead list (for example "generate 100 more leads"), follow `RUNBOOK.md` exactly.
Settings are in `config.yaml`. Companies already delivered are in `state/seen/`: never repeat them.
Apollo credits are scarce and the user does not want to top up: always confirm credit spend first.
After delivering the xlsx, run `leadgen.py commit` and push `state/seen/` and `runs/`.

The `/generate-leads` skill is the entry point for users. The Excel has a single sheet, "All Prospects": do not add per-industry tabs.

File naming: `output/SAP_Leads_Batch<NN>_<YYYY-MM-DD>.xlsx` (batch number + generation date). `leadgen.py build` picks the next batch number itself.
Batch history: Batch01 = 2026-10-06 (100 companies, runs/2026-10-06).
