#!/usr/bin/env python3
"""
merch_planner.py  -  Automated store merchandise planning
==========================================================
Pipeline
  1. SKU master      : 1,000 invented SKUs (structured codes), prices, sizes, assumed performance tier
  2. Fixture inputs  : planner-editable CSV (type, number, capacity, footprint)  [or --interactive]
  3. Space check     : wall-ft and floor-sqft limits for a 1,000 sq ft store
  4. Fixture split   : fixtures allocated to Men / Women zones in proportion to sales
  5. Placement       : best sellers -> most visible fixtures and eye-level slots
  6. Quantities      : min. presentation per size, then fill fixture capacity by sales velocity
  7. Requirement     : how many fixtures the FULL 1,000-SKU range would need vs what you have
  8. Chain roll-out  : x 200 stores
  9. Lifecycle       : (demo) simulate N weeks of sales -> REPLENISH / MARKDOWN / SWAP-IN actions

Usage
  python merch_planner.py                          # defaults, writes ./merch_plan_output/
  python merch_planner.py --fixtures my_fixtures.csv
  python merch_planner.py --interactive            # type fixture counts/capacities at the prompt
  python merch_planner.py --assumptions my.json    # override any assumption (see assumptions.json)
  python merch_planner.py --sku-file real_skus.csv # use real SKUs / real performance tiers
  python merch_planner.py --simulate-weeks 0       # skip lifecycle demo

Requires: pandas, numpy, openpyxl
"""
import argparse
import copy
import heapq
import json
import math
import os

import numpy as np
import pandas as pd

# =====================================================================================
# 1. ASSUMPTIONS  (all can be overridden with --assumptions file.json)
# =====================================================================================
ASSUMPTIONS = {
    "seed": 42,
    "brand_prefix": "HM",
    "n_stores": 200,
    "store": {
        "total_sqft": 1000,
        "selling_pct": 0.70,            # rest = stockroom, billing, trial rooms
        "usable_wall_ft": 100,          # running feet of wall available for wall fixtures
        "max_floor_fixture_pct": 0.30,  # max share of selling area covered by floor fixtures
    },
    "sales_density_per_sqft_year": 35000,   # INR per sq ft of total store area per year
    # share of store sales
    "sales_mix": {"FW_M": 0.49, "FW_W": 0.21, "APP_M": 0.15, "APP_W": 0.15},
    # split of apparel sales (men and women each) into tops / bottoms / accessories
    "apparel_dept_mix": {"TOP": 0.50, "BOT": 0.35, "ACC": 0.15},
    # number of SKUs per department x gender (sum = 1000)
    "sku_counts": {"FW_M": 280, "FW_W": 120, "TOP_M": 130, "TOP_W": 130,
                   "BOT_M": 100, "BOT_W": 100, "ACC_M": 70, "ACC_W": 70},
    "price": {"APP": {"min": 500, "max": 7000, "asp": 1000},
              "FW": {"min": 2000, "max": 15000, "asp": 4000}},
    # performance tiers (random for demo; replace with real sell-through when available)
    "tiers": {"Best": {"share": 0.20, "velocity": 3.0},
              "Better": {"share": 0.30, "velocity": 1.4},
              "Good": {"share": 0.50, "velocity": 0.55}},
    "velocity_noise": 0.25,       # lognormal sigma of SKU-level noise within a tier
    "price_elasticity": 0.4,      # cheaper SKUs sell more units
    "target_weeks_cover": 10,     # used for the 'full range' fixture requirement
    "max_weeks_cover": 16,        # stop topping-up a size once it holds this many weeks of sales
    "demand_transfer": 0.5,       # share of un-ranged SKUs' demand captured by ranged SKUs
    "simulate_weeks": 6,
    "sim_noise": 0.50,            # how far 'actual' sales can deviate from assumed velocity
    "lifecycle": {
        "replenish_below_woc": 3.0,     # replenish when weeks-of-cover falls below this
        "replenish_target_woc": 8.0,
        "markdown_after_weeks": 6,
        "markdown1_if_sell_through_below": 0.20,   # -> 20% off
        "markdown2_if_sell_through_below": 0.08,   # -> 30% off and release the slot
    },
}

DEPT_FIXTURE_GROUP = {"FW": "FOOTWEAR", "TOP": "APPAREL", "BOT": "APPAREL", "ACC": "ACCESSORIES"}
DEPT_PRICE_GROUP = {"FW": "FW", "TOP": "APP", "BOT": "APP", "ACC": "APP"}
DEPT_NAME = {"FW": "Footwear", "TOP": "Tops", "BOT": "Bottoms", "ACC": "Accessories"}

SUBCATS = {  # sub-category mix per (dept, gender)
    ("FW", "M"): {"RUN": .30, "TRN": .20, "SNK": .25, "SLD": .10, "CAS": .15},
    ("FW", "W"): {"RUN": .28, "TRN": .20, "SNK": .28, "SLD": .14, "CAS": .10},
    ("TOP", "M"): {"TEE": .35, "POL": .20, "SHT": .15, "HOD": .15, "JKT": .15},
    ("TOP", "W"): {"TEE": .32, "TNK": .18, "SHT": .15, "HOD": .18, "JKT": .17},
    ("BOT", "M"): {"JNS": .30, "TRK": .30, "SHR": .20, "CHN": .20},
    ("BOT", "W"): {"JNS": .28, "LEG": .32, "SHR": .20, "TRK": .20},
    ("ACC", "M"): {"CAP": .25, "SOK": .25, "BAG": .20, "BLT": .15, "WLT": .15},
    ("ACC", "W"): {"CAP": .20, "SOK": .25, "BAG": .30, "SCF": .15, "HAC": .10},
}
SUB_NAME = {"RUN": "Running shoe", "TRN": "Training shoe", "SNK": "Sneaker", "SLD": "Slides/sandals",
            "CAS": "Casual shoe", "TEE": "T-shirt", "POL": "Polo", "SHT": "Shirt", "HOD": "Hoodie",
            "JKT": "Jacket", "TNK": "Tank/crop top", "JNS": "Jeans", "TRK": "Track pants",
            "SHR": "Shorts", "CHN": "Chinos", "LEG": "Leggings", "CAP": "Cap", "SOK": "Socks pack",
            "BAG": "Bag", "BLT": "Belt", "WLT": "Wallet", "SCF": "Scarf", "HAC": "Hair accessory"}
PRICE_FACTOR = {"RUN": 1.1, "TRN": 1.0, "SNK": 1.2, "SLD": 0.5, "CAS": 0.85, "TEE": 0.7, "POL": 0.9,
                "SHT": 1.0, "HOD": 1.6, "JKT": 2.6, "TNK": 0.6, "JNS": 1.5, "TRK": 1.0, "SHR": 0.8,
                "CHN": 1.3, "LEG": 0.9, "CAP": 0.6, "SOK": 0.45, "BAG": 1.8, "BLT": 0.7, "WLT": 0.8,
                "SCF": 0.7, "HAC": 0.45}
COLOURS = ["BLK", "WHT", "NVY", "GRY", "RED", "BLU", "GRN", "BGE", "OLV", "PNK", "MAR", "YLW"]

CURVES = {  # size curves (share of demand by size)
    ("FW", "M"): [.05, .14, .24, .26, .20, .11],     # UK6-UK11
    ("FW", "W"): [.07, .16, .26, .26, .17, .08],     # UK3-UK8
    ("APP", "M"): [.05, .20, .32, .28, .15],         # XS-XL
    ("APP", "W"): [.12, .28, .30, .20, .10],
    ("ACC", "M"): [1.0], ("ACC", "W"): [1.0],        # one size
}


def sizes_for(dept, g):
    if dept == "FW":
        return [f"UK{n}" for n in (range(6, 12) if g == "M" else range(3, 9))]
    if dept == "ACC":
        return ["OS"]
    return ["XS", "S", "M", "L", "XL"]


def curve_for(dept, g):
    key = (dept if dept in ("FW", "ACC") else "APP", g)
    c = np.array(CURVES[key], dtype=float)
    return c / c.sum()


# Default fixture inputs (planner edits the CSV / uses --interactive)
FIXTURE_COLUMNS = ["fixture_code", "fixture_name", "placement", "allowed_depts", "num_fixtures",
                   "wall_ft_each", "footprint_sqft_each", "sku_slots_each", "pieces_each", "levels",
                   "zone_rule"]
DEFAULT_FIXTURES = [
    ("WF-FW", "Footwear wall bay (shelved, 4ft)", "WALL", "FW", 9, 4, 0, 10, 120, 5, "SALES"),
    ("WF-APP", "Apparel wall bay (rails+shelf, 4ft)", "WALL", "TOP|BOT", 11, 4, 0, 8, 80, 3, "SALES"),
    ("WF-ACC", "Accessory wall bay (pegs, 4ft)", "WALL", "ACC", 2, 4, 0, 24, 96, 6, "SALES"),
    ("GOND-FW", "Footwear gondola (double-sided)", "FLOOR", "FW", 4, 0, 10, 14, 160, 4, "SALES"),
    ("GOND-ACC", "Accessory gondola (double-sided)", "FLOOR", "ACC", 1, 0, 10, 30, 150, 5, "SALES"),
    ("ISL-FW", "Footwear island / try-on bench", "FLOOR", "FW", 2, 0, 16, 8, 64, 2, "SALES"),
    ("TBL-FLD", "Folded-apparel table", "FLOOR", "TOP|BOT", 5, 0, 12, 6, 72, 2, "SALES"),
    ("RACK-4W", "4-way hanger rack (table hanger)", "FLOOR", "TOP|BOT", 5, 0, 9, 8, 64, 2, "SALES"),
    ("IMP-CW", "Cash-wrap impulse unit", "FLOOR", "ACC", 1, 0, 8, 12, 60, 3, "SALES"),
]


# =====================================================================================
# 2. SKU MASTER
# =====================================================================================
def gen_prices(n, factors, lo, hi, asp, rng):
    base = rng.lognormal(0, 0.6, n) * factors
    scale = asp / base.mean()
    for _ in range(40):  # iterate so that the clipped mean hits the target ASP
        scale *= asp / np.clip(base * scale, lo, hi).mean()
    p = np.clip(base * scale, lo, hi)
    p[np.argmin(p)], p[np.argmax(p)] = lo, hi       # make sure the stated range is spanned
    p = np.clip(np.round(p / 100) * 100 - 1, lo, hi)  # 599 / 999 style price points
    return p


def build_sku_master(A, rng):
    rows = []
    for key, n in A["sku_counts"].items():
        dept, g = key.split("_")
        subs = SUBCATS[(dept, g)]
        picks = rng.choice(list(subs), size=n, p=np.array(list(subs.values())) / sum(subs.values()))
        for s in picks:
            rows.append({"dept": dept, "gender": g, "subcat": s, "colour": rng.choice(COLOURS)})
    df = pd.DataFrame(rows)
    df["price"] = 0.0
    for grp, cfg in A["price"].items():
        m = df.dept.map(DEPT_PRICE_GROUP) == grp
        df.loc[m, "price"] = gen_prices(m.sum(), df.loc[m, "subcat"].map(PRICE_FACTOR).values,
                                        cfg["min"], cfg["max"], cfg["asp"], rng)
    # performance tier: stratified random within each dept x gender
    df["perf_tier"] = ""
    names = list(A["tiers"])
    for (_, _), idx in df.groupby(["dept", "gender"]).groups.items():
        n = len(idx)
        counts = [int(round(A["tiers"][t]["share"] * n)) for t in names[:-1]]
        counts.append(n - sum(counts))
        arr = np.repeat(names, counts)
        df.loc[idx, "perf_tier"] = rng.permutation(arr)
    # SKU codes:  <BRAND>-<DEPT><GENDER>-<SUBCAT>-<SEQ4>-<COLOUR>   e.g. HM-FWM-RUN-0001-BLK
    df["_d"] = df.dept.map({"FW": 0, "TOP": 1, "BOT": 2, "ACC": 3})
    df = df.sort_values(["_d", "gender", "subcat"]).drop(columns="_d").reset_index(drop=True)
    df["sku_code"] = [f"{A['brand_prefix']}-{d}{g}-{s}-{i + 1:04d}-{c}" for i, (d, g, s, c) in
                      enumerate(zip(df.dept, df.gender, df.subcat, df.colour))]
    return df


def finalize_sku(df, A, rng):
    """Adds sizes, velocity, priority.  Works for generated and user-supplied SKU files."""
    df = df.copy()
    df["description"] = [f"{SUB_NAME.get(s, s)} ({c})" for s, c in zip(df.subcat, df.colour)]
    df["dept_name"] = df.dept.map(DEPT_NAME)
    df["fixture_group"] = df.dept.map(DEPT_FIXTURE_GROUP)
    df["sizes"] = [sizes_for(d, g) for d, g in zip(df.dept, df.gender)]
    df["n_sizes"] = df.sizes.map(len)
    if "velocity_idx" not in df:
        asp = df.dept.map(DEPT_PRICE_GROUP).map(lambda k: A["price"][k]["asp"])
        tier_v = df.perf_tier.map(lambda t: A["tiers"][t]["velocity"])
        noise = rng.lognormal(0, A["velocity_noise"], len(df))
        df["velocity_idx"] = (tier_v * noise * (asp / df.price) ** A["price_elasticity"]).round(4)
    df["potential_rev_idx"] = df.velocity_idx * df.price
    # priority = percentile of revenue potential inside dept x gender (keeps tops/bottoms balanced)
    df["priority"] = df.groupby(["dept", "gender"]).potential_rev_idx.rank(pct=True, method="first").round(4)
    return df


# =====================================================================================
# 3. FIXTURE INPUTS / SPACE CHECK
# =====================================================================================
def load_fixtures(path):
    if path:
        fx = pd.read_csv(path)
        missing = set(FIXTURE_COLUMNS) - set(fx.columns)
        if missing:
            raise SystemExit(f"Fixture file missing columns: {missing}")
        return fx[FIXTURE_COLUMNS].copy()
    return pd.DataFrame(DEFAULT_FIXTURES, columns=FIXTURE_COLUMNS)


def interactive_edit(fx):
    print("\nEnter new values or press Enter to keep the default.\n")
    for i, r in fx.iterrows():
        print(f"{r.fixture_code}: {r.fixture_name} [{r.placement}] serves {r.allowed_depts}")
        for col, label in [("num_fixtures", "number of fixtures"), ("sku_slots_each", "SKU slots per fixture"),
                           ("pieces_each", "piece capacity per fixture")]:
            v = input(f"   {label} [{r[col]}]: ").strip()
            if v:
                fx.at[i, col] = int(float(v))
    return fx


def check_space(fx, A):
    st = A["store"]
    selling = st["total_sqft"] * st["selling_pct"]
    wall_used = float((fx.num_fixtures * fx.wall_ft_each)[fx.placement == "WALL"].sum())
    floor_used = float((fx.num_fixtures * fx.footprint_sqft_each)[fx.placement == "FLOOR"].sum())
    floor_max = selling * st["max_floor_fixture_pct"]
    res = {"selling_sqft": selling, "wall_ft_used": wall_used, "wall_ft_available": st["usable_wall_ft"],
           "floor_sqft_used": floor_used, "floor_sqft_max": floor_max,
           "wall_ok": wall_used <= st["usable_wall_ft"], "floor_ok": floor_used <= floor_max}
    if not res["wall_ok"]:
        print(f"  !! WARNING: wall fixtures need {wall_used:.0f} ft but only {st['usable_wall_ft']} ft available")
    if not res["floor_ok"]:
        print(f"  !! WARNING: floor fixtures cover {floor_used:.0f} sq ft; limit is {floor_max:.0f} sq ft")
    return res


# =====================================================================================
# 4. FIXTURE INSTANCES (zone split, visibility)
# =====================================================================================
def revenue_shares(A):
    R = {}
    for g in "MW":
        R[("FW", g)] = A["sales_mix"][f"FW_{g}"]
        for d, s in A["apparel_dept_mix"].items():
            R[(d, g)] = A["sales_mix"][f"APP_{g}"] * s
    return R


def build_instances(fx, R):
    rows, carry = [], {}
    for _, r in fx.iterrows():
        n = int(r.num_fixtures)
        if n <= 0:
            continue
        depts = r.allowed_depts.split("|")
        grp = DEPT_FIXTURE_GROUP[depts[0]]
        rule = str(r.zone_rule).upper()
        if rule == "M":
            share_m = 1.0
        elif rule == "W":
            share_m = 0.0
        elif rule == "EVEN":
            share_m = 0.5
        else:
            m = sum(R[(d, "M")] for d in depts)
            w = sum(R[(d, "W")] for d in depts)
            share_m = m / (m + w)
        exact = n * share_m + (carry.get(grp, 0.0) if rule in ("SALES", "EVEN") else 0.0)
        n_m = min(n, max(0, int(math.floor(exact + 0.5))))
        if rule in ("SALES", "EVEN"):
            carry[grp] = exact - n_m
        for g, cnt in (("M", n_m), ("W", n - n_m)):
            for i in range(cnt):
                pos = (i + 0.5) / cnt                      # 0 = entrance ... 1 = back of store
                vis = 1 - 0.5 * pos + (0.1 if r.placement == "FLOOR" else 0)
                depth = "Front" if pos < .34 else ("Mid" if pos < .67 else "Back")
                side = "Left" if g == "M" else "Right"
                rows.append({
                    "fixture_id": f"{r.fixture_code}-{g}-{i + 1:02d}", "fixture_code": r.fixture_code,
                    "fixture_name": r.fixture_name, "placement": r.placement, "gender": g,
                    "allowed_depts": r.allowed_depts, "fixture_group": grp,
                    "sku_slots": int(r.sku_slots_each), "pieces_cap": int(r.pieces_each),
                    "levels": int(r.levels), "visibility": round(vis, 3),
                    "location": f"{side} {'wall' if r.placement == 'WALL' else 'floor'} - {depth}",
                })
    return pd.DataFrame(rows)


def make_slots(S, L):
    """Slots per level; eye-level gets filled first. Returns [(level, position, tag)] best first."""
    eye = max(1, round(0.6 * L))
    order = sorted(range(1, L + 1), key=lambda k: (abs(k - eye), k))
    base, rem = divmod(S, L)
    slots = []
    for rank, k in enumerate(order):
        for p in range(base + (1 if rank < rem else 0)):
            tag = "Eye" if k == eye else ("High" if k > eye else "Low")
            slots.append((k, p + 1, tag))
    return slots


# =====================================================================================
# 5. PLACEMENT + QUANTITIES
# =====================================================================================
def place_skus(sku, inst):
    state = {r.fixture_id: {"slots": make_slots(r.sku_slots, r.levels), "ptr": 0, "pieces": r.pieces_cap}
             for r in inst.itertuples()}
    cand = {}
    for g in "MW":
        for d in DEPT_FIXTURE_GROUP:
            m = inst[(inst.gender == g) & inst.allowed_depts.map(lambda s: d in s.split("|"))]
            cand[(g, d)] = list(m.sort_values(["visibility", "fixture_id"], ascending=[False, True]).fixture_id)
    placed, unplaced = {}, []
    for r in sku.sort_values(["priority", "potential_rev_idx"], ascending=False).itertuples():
        for fid in cand.get((r.gender, r.dept), []):
            s = state[fid]
            if s["ptr"] < len(s["slots"]) and s["pieces"] >= r.n_sizes:
                lvl, pos, tag = s["slots"][s["ptr"]]
                s["ptr"] += 1
                s["pieces"] -= r.n_sizes
                placed[r.sku_code] = (fid, lvl, pos, tag)
                break
        else:
            unplaced.append(r.sku_code)
    return placed, unplaced, state


def calibrate_rates(sku, placed_codes, A, R):
    """Weekly units per store. Store revenue per segment is spread over the ranged SKUs by velocity."""
    weekly_rev = A["sales_density_per_sqft_year"] * A["store"]["total_sqft"] / 52
    sku = sku.copy()
    sku["rate_ranged"], sku["rate_full"] = 0.0, 0.0
    cover = {}
    for (d, g), idx in sku.groupby(["dept", "gender"]).groups.items():
        seg = sku.loc[idx]
        pot_all = seg.potential_rev_idx.sum()
        pl = seg[seg.sku_code.isin(placed_codes)]
        target = weekly_rev * R[(d, g)]
        sku.loc[idx, "rate_full"] = target / pot_all * seg.velocity_idx
        if len(pl):
            cov = pl.potential_rev_idx.sum() / pot_all
            cover[(d, g)] = cov
            seg_rev = target * (cov + (1 - cov) * A["demand_transfer"])
            sku.loc[pl.index, "rate_ranged"] = seg_rev / pl.potential_rev_idx.sum() * pl.velocity_idx
    return sku, cover, weekly_rev


def top_up(sku_idx, placed, state, A):
    """Fill remaining fixture capacity: always add a piece to the size with the lowest weeks-of-cover."""
    qty = {c: [1] * sku_idx.loc[c, "n_sizes"] for c in placed}
    by_fix = {}
    for c, (fid, *_rest) in placed.items():
        by_fix.setdefault(fid, []).append(c)
    maxw, cnt = A["max_weeks_cover"], 0
    for fid, codes in by_fix.items():
        left, heap = state[fid]["pieces"], []
        for c in codes:
            r = sku_idx.loc[c]
            cv = curve_for(r.dept, r.gender)
            for s in range(r.n_sizes):
                rate = max(r.rate_ranged * cv[s], 1e-9)
                if 1 / rate < maxw:
                    heap.append((1 / rate, cnt := cnt + 1, c, s, rate))
        heapq.heapify(heap)
        while left > 0 and heap:
            _, _, c, s, rate = heapq.heappop(heap)
            qty[c][s] += 1
            left -= 1
            if qty[c][s] / rate < maxw:
                heapq.heappush(heap, (qty[c][s] / rate, cnt := cnt + 1, c, s, rate))
        state[fid]["pieces_unfilled"] = left
    return qty


# =====================================================================================
# 6. FIXTURE REQUIREMENT FOR THE FULL RANGE
# =====================================================================================
def fixture_requirement(sku, inst, fx, A):
    grp_rows, code_rows = [], []
    scale_by_grp = {}
    for grp in ["FOOTWEAR", "APPAREL", "ACCESSORIES"]:
        s = sku[sku.fixture_group == grp]
        i = inst[inst.fixture_group == grp]
        slots_need = len(s)
        pieces_need = int(np.maximum(s.n_sizes, np.ceil(s.rate_full * A["target_weeks_cover"])).sum())
        slots_have, pieces_have = int(i.sku_slots.sum()), int(i.pieces_cap.sum())
        scale = max(slots_need / max(slots_have, 1), pieces_need / max(pieces_have, 1))
        scale_by_grp[grp] = scale
        grp_rows.append({"fixture_group": grp, "SKUs_in_range": slots_need, "slots_available": slots_have,
                         "pieces_needed_at_target_cover": pieces_need, "pieces_capacity": pieces_have,
                         "slots_fill_%": round(100 * slots_need / max(slots_have, 1)),
                         "pieces_fill_%": round(100 * pieces_need / max(pieces_have, 1)),
                         "capacity_multiple_needed": round(scale, 2)})
    for r in fx.itertuples():
        grp = DEPT_FIXTURE_GROUP[r.allowed_depts.split("|")[0]]
        need = math.ceil(r.num_fixtures * scale_by_grp[grp] - 1e-9) if r.num_fixtures else 0
        code_rows.append({"fixture_code": r.fixture_code, "fixture_name": r.fixture_name, "group": grp,
                          "current_count": r.num_fixtures, "required_for_full_range": need,
                          "gap": need - r.num_fixtures,
                          "wall_ft_required": need * r.wall_ft_each,
                          "floor_sqft_required": need * r.footprint_sqft_each})
    return pd.DataFrame(grp_rows), pd.DataFrame(code_rows)


# =====================================================================================
# 7. LIFECYCLE DEMO (replenish / markdown / swap-in)
# =====================================================================================
def lifecycle_actions(plan, sku, unplaced_codes, A, rng, weeks):
    L = A["lifecycle"]
    p = plan.copy()
    actual = p.weekly_rate_units * rng.lognormal(0, A["sim_noise"], len(p))   # 'real' sales differ from assumption
    p["units_sold"] = np.minimum(rng.poisson(actual * weeks), p.qty_total)
    p["on_hand"] = p.qty_total - p.units_sold
    p["sell_through_%"] = (100 * p.units_sold / p.qty_total).round(1)
    avg_wk = p.units_sold / weeks
    p["weeks_of_cover"] = np.where(avg_wk > 0, p.on_hand / avg_wk.replace(0, np.nan), 99).round(1)
    acts, qty, md = [], [], []
    for r in p.itertuples():
        st = r.units_sold / r.qty_total
        if r.weeks_of_cover < L["replenish_below_woc"]:
            need = math.ceil(r.units_sold / weeks * L["replenish_target_woc"]) - r.on_hand
            q = int(max(0, min(need, r.qty_total - r.on_hand)))
            acts.append("REPLENISH" if q > 0 else "HOLD"); qty.append(q); md.append(0)
        else:
            if weeks >= L["markdown_after_weeks"] and st < L["markdown2_if_sell_through_below"]:
                acts.append("MARKDOWN 30% + RELEASE SLOT"); qty.append(0); md.append(30)
            elif weeks >= L["markdown_after_weeks"] and st < L["markdown1_if_sell_through_below"]:
                acts.append("MARKDOWN 20%"); qty.append(0); md.append(20)
            else:
                acts.append("HOLD"); qty.append(0); md.append(0)
    p["action"], p["replenish_qty"], p["markdown_%"] = acts, qty, md
    # swap-in: best remaining un-ranged SKU of same dept x gender takes a released slot
    pool = sku[sku.sku_code.isin(unplaced_codes)].sort_values("priority", ascending=False)
    swaps = []
    for r in p[p.action.str.contains("RELEASE")].itertuples():
        m = pool[(pool.dept == r.dept) & (pool.gender == r.gender)]
        if len(m):
            swaps.append(m.iloc[0].sku_code)
            pool = pool.drop(m.index[0])
        else:
            swaps.append("")
    p.loc[p.action.str.contains("RELEASE"), "swap_in_sku"] = swaps
    p["swap_in_sku"] = p["swap_in_sku"].fillna("")
    return p


# =====================================================================================
# 8. EXCEL OUTPUT
# =====================================================================================
def write_excel(path, sheets):
    from openpyxl.styles import Alignment, Font, PatternFill
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name, index=False)
        for ws in xw.book.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    c.font = Font(name="Arial", size=10)
            for c in ws[1]:
                c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="CC0000")
                c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                w = max(len(str(c.value)) if c.value is not None else 0 for c in list(col)[:300])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, w + 2), 70)
            if ws.title == "README":
                for row in ws.iter_rows(min_row=2):
                    for c in row:
                        c.alignment = Alignment(wrap_text=True, vertical="top")


def readme_table(A, space):
    rows = [
        ("Purpose", "Fixture-driven assortment, placement and stock plan for one 1,000 sq ft store, rolled out to all stores."),
        ("Inputs you edit", "fixtures_input.csv (type, number, slots, pieces, footprint, levels). assumptions.json (everything else). Optional --sku-file with real SKUs & tiers."),
        ("Store", f"{A['store']['total_sqft']} sq ft total; {A['store']['selling_pct']:.0%} selling area; {A['store']['usable_wall_ft']} ft usable wall; floor fixtures limited to {A['store']['max_floor_fixture_pct']:.0%} of selling area."),
        ("Space check", f"Wall used {space['wall_ft_used']:.0f}/{space['wall_ft_available']} ft; floor fixtures {space['floor_sqft_used']:.0f}/{space['floor_sqft_max']:.0f} sq ft."),
        ("Sales", f"INR {A['sales_density_per_sqft_year']:,}/sq ft/yr. Footwear 70% (men 49%, women 21% of store); apparel men 15%, women 15% (apparel = tops 50%, bottoms 35%, accessories 15%)."),
        ("Pricing", "Apparel INR 500-7,000 (ASP ~1,000); footwear INR 2,000-15,000 (ASP ~4,000). Price points end in 99 (e.g. 999)."),
        ("SKU code", "<BRAND>-<DEPT><GENDER>-<SUBCAT>-<SEQ4>-<COLOUR>, e.g. HM-FWM-RUN-0001-BLK. Size level adds -<SIZE>, e.g. HM-FWM-RUN-0001-BLK-UK08."),
        ("Sizes", "Apparel XS-XL; men's footwear UK6-11; women's UK3-8; accessories one size. Size curves are assumptions (see code)."),
        ("Performance", "Best 20% / Better 30% / Good 50% assigned randomly within each department x gender (no history for new launches). Replace via --sku-file."),
        ("Placement logic", "Fixtures split Men (left) / Women (right) by sales share. SKUs ranked by revenue potential percentile; best SKUs go to the most visible fixtures (front, floor) and eye-level slots first."),
        ("Quantity logic", "1 piece per size minimum (presentation), then fixture capacity is filled by adding stock to the size with the lowest weeks-of-cover, capped at max weeks-of-cover."),
        ("Not everything fits", "Un-ranged SKUs are listed in Unranged_SKUs (candidates for swap-in, online-only, or a bigger store format)."),
        ("Lifecycle (demo)", "Simulated weekly sales -> REPLENISH when cover is low; MARKDOWN 20/30% when sell-through is low; released slots are offered to the best un-ranged SKU."),
        ("Layout", "Left/Right and Front/Mid/Back are schematic zones, not CAD coordinates."),
    ]
    return pd.DataFrame(rows, columns=["Topic", "Detail"])


# =====================================================================================
# MAIN
# =====================================================================================
def deep_update(d, u):
    for k, v in u.items():
        if isinstance(v, dict) and isinstance(d.get(k), dict):
            deep_update(d[k], v)
        else:
            d[k] = v
    return d


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--fixtures", help="fixture CSV (see fixtures_input.csv)")
    ap.add_argument("--assumptions", help="JSON file with assumption overrides")
    ap.add_argument("--sku-file", help="CSV with sku_code,dept,gender,subcat,colour,price,perf_tier")
    ap.add_argument("--interactive", action="store_true", help="prompt for fixture counts/capacities")
    ap.add_argument("--simulate-weeks", type=int, help="weeks of simulated sales for the lifecycle demo (0 = off)")
    ap.add_argument("--out", default="merch_plan_output")
    a = ap.parse_args()

    A = copy.deepcopy(ASSUMPTIONS)
    if a.assumptions:
        deep_update(A, json.load(open(a.assumptions)))
    if a.simulate_weeks is not None:
        A["simulate_weeks"] = a.simulate_weeks
    os.makedirs(a.out, exist_ok=True)
    rng = np.random.default_rng(A["seed"])
    R = revenue_shares(A)
    assert abs(sum(R.values()) - 1) < 1e-6, "sales mix must sum to 100%"

    # templates for the planner to edit
    pd.DataFrame(DEFAULT_FIXTURES, columns=FIXTURE_COLUMNS).to_csv(os.path.join(a.out, "fixtures_input.csv"), index=False)
    json.dump(ASSUMPTIONS, open(os.path.join(a.out, "assumptions.json"), "w"), indent=2)

    # 1. SKUs
    if a.sku_file:
        sku = pd.read_csv(a.sku_file)
    else:
        sku = build_sku_master(A, rng)
    sku = finalize_sku(sku, A, rng)
    print(f"SKUs: {len(sku)} | footwear ASP {sku[sku.dept == 'FW'].price.mean():,.0f} "
          f"(range {sku[sku.dept == 'FW'].price.min():,.0f}-{sku[sku.dept == 'FW'].price.max():,.0f}) | "
          f"apparel ASP {sku[sku.dept != 'FW'].price.mean():,.0f} "
          f"(range {sku[sku.dept != 'FW'].price.min():,.0f}-{sku[sku.dept != 'FW'].price.max():,.0f})")

    # 2-3. fixtures + space check
    fx = load_fixtures(a.fixtures)
    if a.interactive:
        fx = interactive_edit(fx)
    space = check_space(fx, A)
    print(f"Space: wall {space['wall_ft_used']:.0f}/{space['wall_ft_available']} ft | "
          f"floor fixtures {space['floor_sqft_used']:.0f}/{space['floor_sqft_max']:.0f} sq ft")

    # 4-6. instances, placement, rates, quantities
    inst = build_instances(fx, R)
    placed, unplaced, state = place_skus(sku, inst)
    sku, cover, weekly_rev = calibrate_rates(sku, set(placed), A, R)
    sku_idx = sku.set_index("sku_code")
    qty = top_up(sku_idx, placed, state, A)

    # planogram
    inst_idx = inst.set_index("fixture_id")
    rows, size_rows = [], []
    for c, (fid, lvl, pos, tag) in placed.items():
        r = sku_idx.loc[c]
        tot = sum(qty[c])
        rows.append({"sku_code": c, "description": r.description, "dept": r.dept_name, "gender": r.gender,
                     "price": r.price, "perf_tier": r.perf_tier, "priority": r.priority,
                     "fixture_id": fid, "fixture_code": inst_idx.loc[fid, "fixture_code"],
                     "store_location": inst_idx.loc[fid, "location"], "shelf_level": lvl, "level_tag": tag,
                     "slot": pos, "qty_total": tot, "weekly_rate_units": round(r.rate_ranged, 3),
                     "weeks_of_cover": round(tot / r.rate_ranged, 1), "retail_value": tot * r.price})
        for s, q in zip(r.sizes, qty[c]):
            size_rows.append({"sku_size_code": f"{c}-{s}", "sku_code": c, "size": s, "qty_per_store": q})
    plan = pd.DataFrame(rows).sort_values(["fixture_id", "shelf_level", "slot"]).reset_index(drop=True)
    size_df = pd.DataFrame(size_rows)

    # summaries
    used = plan.groupby("fixture_id").agg(skus=("sku_code", "count"), pieces=("qty_total", "sum"))
    inst_out = inst.join(used, on="fixture_id").fillna({"skus": 0, "pieces": 0})
    inst_out["slot_util_%"] = (100 * inst_out.skus / inst_out.sku_slots).round(0)
    inst_out["piece_util_%"] = (100 * inst_out.pieces / inst_out.pieces_cap).round(0)
    by_code = inst_out.groupby(["fixture_code", "fixture_name"]).agg(
        fixtures=("fixture_id", "count"), sku_slots=("sku_slots", "sum"), skus_placed=("skus", "sum"),
        piece_capacity=("pieces_cap", "sum"), pieces_stocked=("pieces", "sum")).reset_index()
    by_code["piece_util_%"] = (100 * by_code.pieces_stocked / by_code.piece_capacity).round(0)

    seg_rows = []
    tot_slots = inst.sku_slots.sum()
    for (d, g), seg in sku.groupby(["dept", "gender"]):
        pl = plan[(plan.dept == DEPT_NAME[d]) & (plan.gender == g)]
        seg_rows.append({"dept": DEPT_NAME[d], "gender": g, "SKUs_in_range": len(seg), "SKUs_ranged": len(pl),
                         "%_range_ranged": round(100 * len(pl) / len(seg), 1),
                         "%_potential_sales_covered": round(100 * cover.get((d, g), 0), 1),
                         "pieces_stocked": int(pl.qty_total.sum()),
                         "weekly_sales_INR": round(weekly_rev * R[(d, g)]),
                         "share_of_store_sales_%": round(100 * R[(d, g)], 1),
                         "share_of_SKU_slots_%": round(100 * len(pl) / max(tot_slots, 1), 1),
                         "stock_value_INR": int(pl.retail_value.sum())})
    seg_df = pd.DataFrame(seg_rows)

    req_grp, req_code = fixture_requirement(sku, inst, fx, A)
    req_code.loc[len(req_code)] = ["TOTAL", "", "", fx.num_fixtures.sum(), req_code.required_for_full_range.sum(),
                                   req_code.gap.sum(), req_code.wall_ft_required.sum(),
                                   req_code.floor_sqft_required.sum()]

    # chain roll-out
    n = A["n_stores"]
    chain = plan[["sku_code", "description", "dept", "gender", "price", "perf_tier"]].copy()
    chain["qty_per_store"] = plan.qty_total
    chain["stores"] = n
    chain["total_units"] = chain.qty_per_store * n
    chain["total_retail_value"] = chain.total_units * chain.price
    unr = sku[sku.sku_code.isin(unplaced)][["sku_code", "description", "dept_name", "gender", "price",
                                            "perf_tier", "priority"]].sort_values("priority", ascending=False)
    unr["reason"] = "No free slot / piece capacity in compatible fixtures"

    # export-friendly SKU master
    sku_out = sku.drop(columns=["sizes"]).copy()
    sku_out["sizes"] = sku.sizes.map("|".join)
    sku_out["ranged_in_store"] = sku_out.sku_code.isin(placed)
    sku_out = sku_out[["sku_code", "description", "dept_name", "gender", "subcat", "colour", "price", "sizes",
                       "perf_tier", "velocity_idx", "priority", "ranged_in_store", "rate_full", "rate_ranged"]]
    sku_out[["rate_full", "rate_ranged"]] = sku_out[["rate_full", "rate_ranged"]].round(3)

    sheets = {"README": readme_table(A, space),
              "Inputs_Fixtures": fx,
              "SKU_Master": sku_out,
              "Fixture_Layout": inst_out,
              "Planogram": plan,
              "SKU_Size_Qty": size_df,
              "Space_Summary": by_code,
              "Segment_Summary": seg_df,
              "Requirement_By_Group": req_grp,
              "Requirement_By_Fixture": req_code,
              "Chain_Rollout": chain,
              "Unranged_SKUs": unr}

    if A["simulate_weeks"] > 0:
        lc = lifecycle_actions(plan, sku, unplaced, A, rng, A["simulate_weeks"])
        sheets["Lifecycle_Actions"] = lc[["sku_code", "description", "perf_tier", "store_location", "qty_total",
                                          "units_sold", "on_hand", "sell_through_%", "weeks_of_cover", "action",
                                          "replenish_qty", "markdown_%", "swap_in_sku"]]
        print("Lifecycle actions after", A["simulate_weeks"], "weeks:\n",
              lc.action.value_counts().to_string())

    path = os.path.join(a.out, "merch_plan.xlsx")
    write_excel(path, sheets)
    plan.to_csv(os.path.join(a.out, "planogram.csv"), index=False)

    print(f"\nRanged in store: {len(plan)} of {len(sku)} SKUs | pieces per store: {int(plan.qty_total.sum()):,} "
          f"| stock at retail per store: INR {plan.retail_value.sum():,.0f} "
          f"| chain opening units ({n} stores): {int(plan.qty_total.sum()) * n:,}")
    print(seg_df[["dept", "gender", "SKUs_in_range", "SKUs_ranged", "%_potential_sales_covered",
                  "pieces_stocked"]].to_string(index=False))
    print("\nFull-range fixture requirement:\n", req_grp.to_string(index=False))
    print(f"\nWrote {path}")


if __name__ == "__main__":
    main()
