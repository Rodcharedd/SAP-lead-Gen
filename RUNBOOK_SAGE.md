# Runbook: new Sage Intacct prospect list (100 companies)

Settings: `config_sage.yaml`. Builder: `sage_leadgen.py`. Delivered companies: `state/seen_sage/` (separate from the SAP lists; a company may appear in both).

## Credit rules
Same as `RUNBOOK.md`: check credits first, confirm the total cost before searching, never spend below `credit_floor`, no email/phone reveal unless asked. Credits are scarce; the user will not top up.

## Steps
1. `git pull`; look at `state/seen_sage/`.
2. Run the 4 searches in `config_sage.yaml` (`searches`): one per major segment. Each is `apollo_mixed_companies_search` with `organization_locations: ["Thailand"]`, `organization_naics_codes` from the search, `organization_num_employees_ranges` from `employee_ranges`, `per_page: 100`. The revenue filter is not available on the Free plan; revenue (>= 200M THB) is filtered by the script. If the NAICS filter is rejected, tell the user instead of falling back to guesses.
3. Total quote before searching: 4 credits for page 1 of each search. Extra pages (1 credit each) only after telling the user how many segments are short.
4. Save all results (`organizations` + `accounts`) together as `runs_sage/<YYYY-MM-DD>/companies.json`. Note the next page to fetch per search in `CLAUDE.md`.
5. `python3 sage_leadgen.py build runs_sage/<date>/companies.json`. It prints the count per segment and warns if fewer than 100.
6. Send the xlsx (`Sage_Leads_Batch<NN>_<date>.xlsx`, one sheet "All Prospects"). Say which segments were short and how many rows have unknown revenue.
7. After the user has the file: `python3 sage_leadgen.py commit runs_sage/<date>/delivered.json`, commit and push `state/seen_sage/` and `runs_sage/`.

## Existing customers
`state/sage_existing_customers.csv` (customers + prospects already worked, from the user's Sage customer list) is excluded by `build` using name matching; matches are written to `runs_sage/<date>/excluded_existing.csv` for checking. Short names (like "EY", "SW") must match exactly. When the user sends new customers, append them to the CSV and commit.

## Limits
- Profit, number of group entities, multi-currency and consolidation needs are not in Apollo: the xlsx has empty columns for the sales team. (DBD Datawarehouse blocks scripted access.)
- Apollo revenue for Thai companies is rough and often missing. Segment classification uses NAICS first, then company-name keywords (column "Matched by"): check the keyword matches.
- Banks and insurers are not in the target segments and are left out.
