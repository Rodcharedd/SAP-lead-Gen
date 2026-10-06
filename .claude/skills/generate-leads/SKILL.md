---
name: generate-leads
description: Generate a new Excel list of Thai companies (default 100, never repeating earlier lists) as SAP Business One / S/4HANA Cloud prospects, using Apollo.io. Use when the user asks for new leads, a new prospect list, or "generate more leads".
---

# Generate new SAP leads

Arguments: optional number of companies (default `target_count` in `config.yaml`).

1. `git pull` first, so the list of delivered companies (`state/seen/`) is current.
2. Follow `RUNBOOK.md` step by step. It is the single source of truth. Do not skip the credit confirmations: Apollo credits are scarce and nobody wants to top up.
3. Send the xlsx to the user, then run `leadgen.py commit` and push `state/seen/` and `runs/` to the branch you are on, so the next person never gets the same companies.
4. Reply in the user's language. The Excel itself is in English.
