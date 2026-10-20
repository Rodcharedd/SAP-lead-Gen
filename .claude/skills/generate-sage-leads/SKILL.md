---
name: generate-sage-leads
description: Generate a new Excel list of Thai companies (default 100, never repeating earlier Sage lists) as Sage Intacct prospects by segment (Financial Services, Hospitality, Business Services, Non-profit), using Apollo.io. Use when the user asks for new Sage or Sage Intacct leads.
---

# Generate new Sage Intacct leads

1. `git pull` so `state/seen_sage/` is current.
2. Follow `RUNBOOK_SAGE.md` step by step (single source of truth). Confirm the total Apollo credit cost before any search.
3. Send the xlsx, then run `sage_leadgen.py commit` and push `state/seen_sage/` and `runs_sage/`.
4. Reply in the user's language; the Excel is in English. Do not mix this with the SAP lists (`/generate-leads`).
