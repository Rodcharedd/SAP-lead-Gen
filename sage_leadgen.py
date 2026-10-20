#!/usr/bin/env python3
"""Sage Intacct prospect list: classify Apollo results by target segment and write the Excel.

Usage:
  python3 sage_leadgen.py build  runs_sage/<run>/companies.json [--ignore-seen] [--batch N]
  python3 sage_leadgen.py commit runs_sage/<run>/delivered.json

Reuses the helpers of leadgen.py (SAP). De-duplicates only against earlier Sage lists, never against SAP lists.
"""
import argparse, json
from datetime import date
from pathlib import Path

import yaml
from openpyxl import Workbook

import leadgen as lg

ROOT = Path(__file__).parent
CFG = yaml.safe_load((ROOT / "config_sage.yaml").read_text())
SEEN_DIR = ROOT / "state" / "seen_sage"


def load_seen():
    seen = {}
    for f in sorted(SEEN_DIR.glob("*.json")):
        seen.update(json.loads(f.read_text()))
    return seen


def classify(c):
    """Return the first matching segment (lowest priority number). NAICS first, then name keywords."""
    codes = c.get("naics_codes") or []
    name = (c.get("name") or "").lower()
    for seg in sorted(CFG["segments"], key=lambda s: s["priority"]):
        if any(code.startswith(p) for code in codes for p in seg["naics"]):
            return seg, "NAICS"
    for seg in sorted(CFG["segments"], key=lambda s: s["priority"]):
        if any(k in name for k in seg["keywords"]):
            return seg, "Name keyword"
    return None, None


def lead_tier(thb):
    if thb is None:
        return "Revenue unknown"
    for t in CFG["tiers"]:
        if thb >= t["min"] and (t["max"] is None or thb < t["max"]):
            return t["name"]
    return "Below 200M"


def normalise(c, seg, how):
    usd = lg.revenue_usd(c)
    thb = usd * CFG["fx_thb_per_usd"] if usd is not None else None
    phone = lg.pick(c, "phone", "sanitized_phone")
    if not phone and isinstance(c.get("primary_phone"), dict):
        phone = c["primary_phone"].get("number")
    return {
        "Company": lg.pick(c, "name"),
        "Major segment": seg["major"],
        "Target segment": seg["target"],
        "Priority": seg["priority"],
        "Matched by": how,
        "Revenue (THB M)": round(thb / 1e6, 1) if thb is not None else None,
        "Lead tier": lead_tier(thb),
        "Website": lg.domain_key(c),
        "Company phone": phone,
        "LinkedIn": lg.pick(c, "linkedin_url"),
        "Headcount growth 12m (%)": lg.pct(c.get("organization_headcount_twelve_month_growth")),
        "Locations": lg.pick(c, "num_retail_locations"),
        "Parent company": (c.get("owned_by_organization") or {}).get("name"),
        "Founded": lg.pick(c, "founded_year"),
        # Not available from Apollo: to be filled in by the sales team
        "Profit (THB M) - fill in": None,
        "Group entities - check": None,
        "Multi-currency / consolidation - check": None,
        "Apollo org ID": lg.pick(c, "organization_id", "id"),
    }


def select(rows):
    """Fill quota_per_major per major segment, then top up from the rest by priority."""
    def key(r):
        known = r["Revenue (THB M)"] is not None
        return (r["Priority"], not known, abs((r["Revenue (THB M)"] or 0) - 600))  # known first, near the core band
    rows = sorted(rows, key=key)
    chosen, left = [], []
    counts = {}
    for r in rows:
        m = r["Major segment"]
        if counts.get(m, 0) < CFG["quota_per_major"]:
            counts[m] = counts.get(m, 0) + 1
            chosen.append(r)
        else:
            left.append(r)
    chosen += left[: max(0, CFG["target_count"] - len(chosen))]
    chosen = chosen[: CFG["target_count"]]
    chosen.sort(key=lambda r: (r["Major segment"], r["Priority"], r["Revenue (THB M)"] is None, -(r["Revenue (THB M)"] or 0)))
    return chosen


def build(path, ignore_seen=False, batch=None):
    raw = json.loads(Path(path).read_text())
    seen = {} if ignore_seen else load_seen()
    rows, keys, skipped, no_seg = [], set(), 0, 0
    for c in raw:
        k = lg.domain_key(c)
        if not k or k in seen or k in keys:
            skipped += 1
            continue
        seg, how = classify(c)
        if not seg:
            no_seg += 1
            continue
        row = normalise(c, seg, how)
        r = row["Revenue (THB M)"]
        if r is None and not CFG["keep_unknown_revenue"]:
            continue
        if r is not None and r < CFG["revenue_min_thb"] / 1e6:
            continue
        keys.add(k)
        rows.append(row)
    chosen = select(rows)

    cols = list(chosen[0].keys()) if chosen else []
    wb = Workbook()
    wb.remove(wb.active)
    lg.sheet(wb, "All Prospects", chosen, cols)

    ids = {r["Apollo org ID"] for r in chosen}
    delivered = [c for c in raw if (c.get("organization_id") or c.get("id")) in ids]
    (Path(path).parent / "delivered.json").write_text(json.dumps(delivered, ensure_ascii=False, indent=1))

    batch = batch or len(list(SEEN_DIR.glob("*.json"))) + 1
    out = ROOT / "output" / f"Sage_Leads_Batch{batch:02d}_{date.today().isoformat()}.xlsx"
    out.parent.mkdir(exist_ok=True)
    wb.save(out)

    by = {}
    for r in chosen:
        by[(r["Major segment"], r["Target segment"])] = by.get((r["Major segment"], r["Target segment"]), 0) + 1
    print(f"Wrote {out}  ({len(chosen)} companies; skipped {skipped} duplicates/seen, {no_seg} outside target segments)")
    for (m, t), n in sorted(by.items()):
        print(f"  {m:20s} {t:55s} {n}")
    unknown = sum(1 for r in chosen if r["Revenue (THB M)"] is None)
    print(f"  revenue unknown: {unknown}")
    if len(chosen) < CFG["target_count"]:
        print(f"WARNING: only {len(chosen)} of {CFG['target_count']}. Fetch more pages or add searches.")


def commit(path):
    raw = json.loads(Path(path).read_text())
    seen = load_seen()
    new = {}
    for c in raw:
        k = lg.domain_key(c)
        if k and k not in seen:
            new[k] = {"name": c.get("name"), "first_listed": date.today().isoformat()}
    SEEN_DIR.mkdir(parents=True, exist_ok=True)
    (SEEN_DIR / f"{Path(path).parent.name}.json").write_text(json.dumps(new, indent=1, ensure_ascii=False, sort_keys=True))
    print(f"state/seen_sage/ now has {len(seen) + len(new)} companies (+{len(new)})")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["build", "commit"])
    ap.add_argument("companies")
    ap.add_argument("--batch", type=int)
    ap.add_argument("--ignore-seen", action="store_true")
    a = ap.parse_args()
    build(a.companies, a.ignore_seen, a.batch) if a.cmd == "build" else commit(a.companies)
