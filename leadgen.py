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
SEEN = ROOT / "state" / "seen.json"


def load_seen():
    return json.loads(SEEN.read_text()) if SEEN.exists() else {}


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
        "Industry": pick(c, "industry") or "Other",
        "Website": domain_key(c),
        "City": pick(c, "city"),
        "Employees": pick(c, "estimated_num_employees", "employee_count"),
        "Revenue (THB M)": round(thb / 1e6, 1) if thb is not None else None,
        "Revenue range (as reported)": pick(c, "organization_revenue_printed", "annual_revenue_printed"),
        "SAP Fit": tier(thb),
        "Company phone": phone,
        "LinkedIn": pick(c, "linkedin_url"),
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


def build(companies_path, contacts_path):
    raw = json.loads(Path(companies_path).read_text())
    seen = load_seen()
    fresh, dupes, keys = [], 0, set()
    for c in raw:
        k = domain_key(c)
        if not k or k in seen or k in keys:
            dupes += 1
            continue
        row = normalise(c)
        if not in_band(row):
            continue
        keys.add(k)
        fresh.append(row)
    fresh.sort(key=lambda r: (r["Revenue (THB M)"] is None, -(r["Revenue (THB M)"] or 0)))
    fresh = fresh[: CFG["target_count"]]

    cols = list(fresh[0].keys()) if fresh else []
    wb = Workbook()
    wb.remove(wb.active)

    by_ind = {}
    for r in fresh:
        by_ind.setdefault(r["Industry"], []).append(r)
    summary = wb.create_sheet("Summary")
    summary.append(["Actran Systems - SAP prospect list", date.today().isoformat()])
    summary.append(["Companies in this list", len(fresh)])
    summary.append(["Skipped (already delivered / duplicate)", dupes])
    summary.append([])
    summary.append(["Industry", "Companies"])
    for ind, rows in sorted(by_ind.items(), key=lambda x: -len(x[1])):
        summary.append([ind, len(rows)])
    summary.append([])
    summary.append(["SAP Fit", "Companies"])
    for t in [t["name"] for t in CFG["tiers"]] + ["Revenue unknown"]:
        summary.append([t, sum(1 for r in fresh if r["SAP Fit"] == t)])
    summary.column_dimensions["A"].width = 42
    summary["A1"].font = Font(bold=True, size=14)

    sheet(wb, "All Prospects", fresh, cols)
    for ind, rows in sorted(by_ind.items(), key=lambda x: -len(x[1])):
        sheet(wb, re.sub(r"[\[\]\*\?/\\:]", "-", ind), rows, cols)

    if contacts_path and Path(contacts_path).exists():
        contacts = json.loads(Path(contacts_path).read_text())
        ccols = ["Company", "Name", "Title", "Email", "Email status", "Direct phone", "Company phone", "LinkedIn"]
        sheet(wb, "Contacts", contacts, ccols)

    notes = wb.create_sheet("Methodology")
    for line in [
        "Source: Apollo.io company search (Thailand). Revenue is Apollo's estimate converted to THB at "
        f"{CFG['fx_thb_per_usd']} THB/USD; it is often missing or rough for Thai companies, so verify before outreach.",
        f"Revenue band: >= {CFG['revenue_min_thb']/1e6:,.0f}M THB, cap: {CFG['revenue_max_thb'] or 'none'}.",
        "SAP Fit is a rough sizing guide by revenue (SAP Business One vs S/4HANA Cloud Public Edition), not a qualification.",
        "Each run excludes companies delivered in earlier runs (state/seen.json).",
        "Existing ERP vendors are not excluded: users of other products may be upgrade or change-partner targets.",
        "Contact data is subject to PDPA: keep the source and honour opt-outs.",
    ]:
        notes.append([line])
    notes.column_dimensions["A"].width = 120

    out = ROOT / "output" / f"SAP_Prospects_{date.today().isoformat()}_{Path(companies_path).parent.name}.xlsx"
    out.parent.mkdir(exist_ok=True)
    wb.save(out)
    print(f"Wrote {out}  ({len(fresh)} companies, {dupes} skipped)")
    if len(fresh) < CFG["target_count"]:
        print(f"WARNING: only {len(fresh)} of {CFG['target_count']} new companies. Fetch more pages or relax filters.")


def commit(companies_path):
    raw = json.loads(Path(companies_path).read_text())
    seen = load_seen()
    n = 0
    for c in raw:
        k = domain_key(c)
        if k and k not in seen:
            seen[k] = {"name": c.get("name"), "first_listed": date.today().isoformat()}
            n += 1
    SEEN.parent.mkdir(exist_ok=True)
    SEEN.write_text(json.dumps(seen, indent=1, ensure_ascii=False, sort_keys=True))
    print(f"state/seen.json now has {len(seen)} companies (+{n})")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["build", "commit"])
    ap.add_argument("companies")
    ap.add_argument("--contacts")
    a = ap.parse_args()
    build(a.companies, a.contacts) if a.cmd == "build" else commit(a.companies)
