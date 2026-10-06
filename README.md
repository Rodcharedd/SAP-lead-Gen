# SAP lead generator (Actran Systems)

Generates an Excel list of Thai companies that could buy SAP Business One or SAP S/4HANA Cloud Public Edition. Each run returns companies that have not been delivered before.

## How to generate a new list
1. Open this repo in Claude Code (claude.ai/code) with the **Apollo.io connector** connected.
2. Type: `/generate-leads` (or `/generate-leads 50` for a different size).
3. Approve the Apollo credit spend when asked (a normal run uses about 3 credits).
4. Download the Excel file. Claude records the companies as delivered, so the next run gives new names.

## Files
- `config.yaml`: revenue floor (THB), list size, excluded sectors
- `RUNBOOK.md`: the exact steps Claude follows
- `leadgen.py`: builds the Excel and tracks delivered companies
- `state/seen/`: one file per run listing delivered companies (do not edit by hand)
- `runs/`: raw Apollo data per run; `output/`: the Excel files

## Limits
- Apollo's free plan has no revenue filter, so revenue is checked after the search. Revenue is Apollo's estimate and can be rough.
- Contact emails and direct phones cost Apollo credits and are off by default. The list gives company switchboard numbers.
- Run one list at a time: two people generating at once can receive overlapping companies.
