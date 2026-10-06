# Runbook: generate a new SAP prospect list (100 companies)

Trigger: the user asks for a new lead list. Output: an Excel file in `output/` with companies that have not been delivered before. Settings live in `config.yaml`.

## Credit rules (read first)
- The Apollo plan is small (about 75 credits at last check, 5 direct-dial). The user does **not** want to top up.
- Check `apollo_users_api_profile(include_credit_usage=true)` at the start. Never spend below `credit_floor`.
- `apollo_mixed_companies_search` costs 1 credit per request, up to 100 results per page. Before each call, say the exact confirmation the tool requires and wait for approval. Quote the total for the whole run up front.
- `apollo_mixed_people_api_search` is free but returns no email or phone.
- `apollo_people_bulk_match` (email/phone reveal) and `apollo_organizations_bulk_enrich` cost credits. Only run them if `enrich_contacts: true` and credits allow. Otherwise deliver company, revenue band and company phone only.

## Steps
1. Read `state/seen.json` (companies already delivered).
2. Search companies: `organization_locations: ["Thailand"]`, `revenue_range` from config (convert THB to USD with `fx_thb_per_usd`), `per_page: 100`. Pick the page so results differ from earlier runs. Pull a second page only if fewer than 100 new companies remain after dedup (each page is 1 more credit).
3. Save the `organizations` and `accounts` arrays together as `runs/<YYYY-MM-DD>/companies.json`.
4. Optional, free: people search per company for the personas in config, saved as `runs/<date>/contacts.json` (columns: Company, Name, Title, Email, Email status, Direct phone, Company phone, LinkedIn).
5. `python3 leadgen.py build runs/<date>/companies.json [--contacts runs/<date>/contacts.json]`
6. Send the xlsx to the user.
7. After the user has the file: `python3 leadgen.py commit runs/<date>/companies.json`, then commit and push `state/seen.json` and `runs/`. Without the push, the next session will repeat the same companies.

## Known limits
- Apollo revenue data for Thai companies is sparse and approximate. Rows without revenue are kept and marked "Revenue unknown".
- Phone numbers for contacts are often missing. The fallback is the company switchboard.
