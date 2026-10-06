# SAP lead-gen (Actran Systems)

Purpose: generate Excel prospect lists of Thai companies for SAP Business One / S/4HANA Cloud Public Edition.

When the user asks for a new lead list (for example "generate 100 more leads"), follow `RUNBOOK.md` exactly.
Settings are in `config.yaml`. Companies already delivered are in `state/seen.json`: never repeat them.
Apollo credits are scarce and the user does not want to top up: always confirm credit spend first.
After delivering the xlsx, run `leadgen.py commit` and push `state/seen.json` and `runs/`.
