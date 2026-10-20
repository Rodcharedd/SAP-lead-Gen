#!/usr/bin/env python3
"""Sage Intacct prospect list: classify Apollo results by target segment and write the Excel.

Usage:
  python3 sage_leadgen.py build  runs_sage/<run>/companies.json [--ignore-seen] [--batch N]
  python3 sage_leadgen.py commit runs_sage/<run>/delivered.json

Reuses the helpers of leadgen.py (SAP). De-duplicates only against earlier Sage lists, never against SAP lists.
"""
import argparse, csv, json, re
from datetime import date
from pathlib import Path

import yaml
from openpyxl import Workbook

import leadgen as lg

ROOT = Path(__file__).parent
CFG = yaml.safe_load((ROOT / "config_sage.yaml").read_text())
SEEN_DIR = ROOT / "state" / "seen_sage"


EXISTING = ROOT / "state" / "sage_existing_customers.csv"   # customers and prospects Actran already has
STOP = {"co", "ltd", "limited", "company", "public", "plc", "pcl", "corporation", "corp", "inc", "thailand",
        "the", "group", "plant1", "plant2", "plant", "for", "and"}


def tokens(text):
    return [t for t in re.findall(r"[a-z0-9]+", (text or "").lower()) if t not in STOP]


def load_existing():
    if not EXISTING.exists():
        return []
    with open(EXISTING, encoding="utf-8") as f:
        return [(r["company"], tokens(r["company"]), r["kind"]) for r in csv.DictReader(f)]


def existing_match(c, existing):
    """Name match against Actran's existing customers/prospects. Short single words must match exactly."""
    cand = set(tokens(c.get("name")))
    cand_list = tokens(c.get("name"))
    dom = {t for t in re.findall(r"[a-z0-9]+", lg.domain_key(c).split(".")[0].lower())}
    for name, tk, kind in existing:
        if not tk:
            continue
        if len(tk) == 1 and len(tk[0]) < 5:
            hit = cand_list == tk
        else:
            hit = set(tk) <= cand or (len(tk) == 1 and tk[0] in dom)
        if hit:
            return name, kind
    return None


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
    counts, per_target = {}, {}
    mins = CFG.get("min_per_target", 0)
    for r in rows:  # first: a minimum for every target segment, so low-priority ones are not squeezed out
        t = r["Target segment"]
        if per_target.get(t, 0) < mins:
            per_target[t] = per_target.get(t, 0) + 1
            counts[r["Major segment"]] = counts.get(r["Major segment"], 0) + 1
            chosen.append(r)
    taken = {id(r) for r in chosen}
    rows = [r for r in rows if id(r) not in taken]
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
    existing, excluded = load_existing(), []
    for c in raw:
        k = lg.domain_key(c)
        if not k or k in seen or k in keys:
            skipped += 1
            continue
        hit = existing_match(c, existing)
        if hit and (CFG.get("exclude_existing_prospects", True) or hit[1] == "customer"):
            excluded.append({"candidate": c.get("name"), "website": k, "matches": hit[0], "kind": hit[1]})
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

    with open(Path(path).parent / "excluded_existing.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["candidate", "website", "matches", "kind"])
        w.writeheader()
        w.writerows(excluded)
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
    print(f"  excluded as existing customer/prospect: {len(excluded)} (see excluded_existing.csv in the run folder)")
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
