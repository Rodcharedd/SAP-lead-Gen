#!/usr/bin/env python3
"""Turn raw Apollo results into a de-duplicated Excel prospect list.

Usage:
  python3 leadgen.py build  runs/<run>/companies.json [--contacts runs/<run>/contacts.json]
  python3 leadgen.py commit runs/<run>/companies.json   # mark the list as delivered (updates state/seen.json)

Apollo is queried from the Claude session (see RUNBOOK.md); this script handles
normalising, de-duplicating against earlier runs, tiering and writing the workbook.
"""
import argparse, json, re, sys
from datetime import date
from pathlib import Path

import yaml
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).parent
CFG = yaml.safe_load((ROOT / "config.yaml").read_text())
SEEN_DIR = ROOT / "state" / "seen"  # one file per run, so parallel users never conflict in git


def load_seen():
    seen = {}
    for f in sorted(SEEN_DIR.glob("*.json")):
        seen.update(json.loads(f.read_text()))
    return seen


def domain_key(c):
    d = c.get("primary_domain") or c.get("domain") or c.get("website_url") or ""
    d = re.sub(r"^https?://(www\.)?", "", d.strip().lower()).split("/")[0]
    return d or (c.get("name") or "").strip().lower()


def pick(c, *keys):
    for k in keys:
        v = c.get(k)
        if v not in (None, "", []):
            return v
    return None


def revenue_usd(c):
    v = pick(c, "organization_revenue", "annual_revenue", "revenue")
    if isinstance(v, (int, float)):
        return float(v)
    s = pick(c, "organization_revenue_printed", "annual_revenue_printed")
    if isinstance(s, str):
        m = re.match(r"([\d.]+)\s*([KMB]?)", s.replace(",", "").replace("$", "").strip(), re.I)
        if m:
            return float(m.group(1)) * {"": 1, "K": 1e3, "M": 1e6, "B": 1e9}[m.group(2).upper()]
    return None


NAICS = {"11": "Agriculture", "21": "Mining, Oil & Gas", "22": "Utilities & Energy", "23": "Construction",
         "31": "Manufacturing - Food & Textile", "32": "Manufacturing - Chemicals & Materials",
         "33": "Manufacturing - Machinery & Electronics", "42": "Wholesale", "44": "Retail", "45": "Retail",
         "48": "Transportation & Logistics", "49": "Transportation & Logistics", "51": "Information & Media",
         "52": "Finance & Insurance", "53": "Real Estate", "54": "Professional Services", "55": "Holding Companies",
         "56": "Business Services", "61": "Education", "62": "Healthcare", "71": "Arts & Entertainment",
         "72": "Hospitality & Food Service", "81": "Other Services", "92": "Public Administration"}


def industry(c):
    for code in c.get("naics_codes") or []:
        if code[:2] in NAICS:
            return NAICS[code[:2]]
    return pick(c, "industry") or "Other"


def excluded(c):
    pre = {str(x) for x in CFG.get("exclude_naics_prefixes") or []}
    return bool(pre) and any(code[:2] in pre for code in c.get("naics_codes") or [])


def pct(v):
    return round(v * 100, 1) if isinstance(v, (int, float)) else None


def tier(thb):
    if thb is None:
        return "Revenue unknown"
    for t in CFG["tiers"]:
        if thb >= t["min"] and (t["max"] is None or thb < t["max"]):
            return t["name"]
    return "Below range"


def normalise(c):
    usd = revenue_usd(c)
    thb = usd * CFG["fx_thb_per_usd"] if usd is not None else None
    phone = pick(c, "phone", "sanitized_phone")
    if not phone and isinstance(c.get("primary_phone"), dict):
        phone = c["primary_phone"].get("number")
    return {
        "Company": pick(c, "name"),
        "Industry": industry(c),
        "Website": domain_key(c),
        "City": pick(c, "city"),
        "Employees": pick(c, "estimated_num_employees", "employee_count"),
        "Revenue (THB M)": round(thb / 1e6, 1) if thb is not None else None,
        "Revenue range (as reported)": pick(c, "organization_revenue_printed", "annual_revenue_printed"),
        "SAP Fit": tier(thb),
        "Company phone": phone,
        "LinkedIn": pick(c, "linkedin_url"),
        "Headcount growth 12m (%)": pct(c.get("organization_headcount_twelve_month_growth")),
        "Parent company": (c.get("owned_by_organization") or {}).get("name"),
        "Founded": pick(c, "founded_year"),
        "Latest funding": pick(c, "latest_funding_stage"),
        "Apollo org ID": pick(c, "organization_id", "id"),
    }


def in_band(row):
    r = row["Revenue (THB M)"]
    if r is None:
        return True  # keep unknowns; flagged in SAP Fit
    lo = CFG["revenue_min_thb"] / 1e6
    hi = CFG["revenue_max_thb"] / 1e6 if CFG["revenue_max_thb"] else float("inf")
    return lo <= r <= hi


def sheet(wb, title, rows, cols):
    ws = wb.create_sheet(title[:31])
    ws.append(cols)
    for r in rows:
        ws.append([r.get(c) for c in cols])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    for i, col in enumerate(cols, 1):
        width = max([len(str(col))] + [len(str(r.get(col) or "")) for r in rows]) + 2
        ws.column_dimensions[get_column_letter(i)].width = min(width, 45)
    return ws


def build(companies_path, contacts_path, ignore_seen=False):
    raw = json.loads(Path(companies_path).read_text())
    seen = {} if ignore_seen else load_seen()
    fresh, dupes, keys = [], 0, set()
    for c in raw:
        k = domain_key(c)
        if not k or k in seen or k in keys or excluded(c):
            dupes += 1
            continue
        row = normalise(c)
        if not in_band(row):
            continue
        keys.add(k)
        fresh.append(row)
    # Prefer the sweet spot: B1 / mid-market first, then the smallest of the large companies. Unknown revenue last.
    order = {t["name"]: i for i, t in enumerate(CFG["tiers"])}
    fresh.sort(key=lambda r: (order.get(r["SAP Fit"], 99), r["Revenue (THB M)"] or 0))
    fresh = fresh[: CFG["target_count"]]

    cols = [k for k in (fresh[0].keys() if fresh else []) if any(r.get(k) not in (None, "") for r in fresh)]
    wb = Workbook()
    wb.remove(wb.active)

    sheet(wb, "All Prospects", fresh, cols)

    if contacts_path and Path(contacts_path).exists():
        contacts = json.loads(Path(contacts_path).read_text())
        ccols = ["Company", "Name", "Title", "Email", "Email status", "Direct phone", "Company phone", "LinkedIn"]
        sheet(wb, "Contacts", contacts, ccols)

    ids = {r["Apollo org ID"] for r in fresh}
    delivered = [c for c in raw if (c.get("organization_id") or c.get("id")) in ids]
    (Path(companies_path).parent / "delivered.json").write_text(json.dumps(delivered, ensure_ascii=False, indent=1))

    out = ROOT / "output" / f"SAP_Prospects_{date.today().isoformat()}_{Path(companies_path).parent.name}.xlsx"
    out.parent.mkdir(exist_ok=True)
    wb.save(out)
    print(f"Wrote {out}  ({len(fresh)} companies, {dupes} skipped)")
    if len(fresh) < CFG["target_count"]:
        print(f"WARNING: only {len(fresh)} of {CFG['target_count']} new companies. Fetch more pages or relax filters.")


def commit(companies_path):
    raw = json.loads(Path(companies_path).read_text())
    seen = load_seen()
    new = {}
    for c in raw:
        k = domain_key(c)
        if k and k not in seen:
            new[k] = {"name": c.get("name"), "first_listed": date.today().isoformat()}
    SEEN_DIR.mkdir(parents=True, exist_ok=True)
    (SEEN_DIR / f"{Path(companies_path).parent.name}.json").write_text(
        json.dumps(new, indent=1, ensure_ascii=False, sort_keys=True))
    print(f"state/seen/ now has {len(seen) + len(new)} companies (+{len(new)})")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["build", "commit"])
    ap.add_argument("companies")
    ap.add_argument("--contacts")
    ap.add_argument("--ignore-seen", action="store_true", help="rebuild a run without de-duplicating against earlier runs")
    a = ap.parse_args()
    build(a.companies, a.contacts, a.ignore_seen) if a.cmd == "build" else commit(a.companies)
