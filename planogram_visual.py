#!/usr/bin/env python3
"""
planogram_visual.py  -  Turns the output of merch_planner.py into a visual planogram
=====================================================================================
Reads  <plan_dir>/merch_plan.xlsx  (sheets: Fixture_Layout, Planogram, Inputs_Fixtures)
Writes <plan_dir>/visual/
    planogram_floor.png        store floor plan: Men left / Women right, fixtures colour-coded by department
    planogram_elevations.pdf   one card per fixture: shelf levels x slots, every cell = a SKU coloured by tier
    elevations_p1.png ...      same pages as PNGs (for slides)

Usage
    python planogram_visual.py                         # uses ./merch_plan_output
    python planogram_visual.py --plan-dir my_output --width-ft 24

Keep this file next to merch_planner.py (it re-uses make_slots so empty slots are drawn correctly).
Requires: pandas, numpy, matplotlib, openpyxl
"""
import argparse
import json
import math
import os
from collections import Counter

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import Patch, Rectangle

from merch_planner import make_slots

GROUP_COLOUR = {"FOOTWEAR": "#B3261E", "APPAREL": "#1F5FA8", "ACCESSORIES": "#2E7D5B"}
GROUP_ORDER = ["FOOTWEAR", "APPAREL", "ACCESSORIES"]
TIER_COLOUR = {"Best": "#4CAF50", "Better": "#FFD54F", "Good": "#CFD8DC"}
WALL_DEPTH = 2.0   # ft


# ------------------------------------------------------------------ helpers
def short_id(fid):
    return fid.replace("-M-", "-").replace("-W-", "-")


def sku_label(code):
    p = code.split("-")          # HM-FWM-RUN-0001-BLK
    return f"{p[2]}-{p[3]}", p[4]


def load(plan_dir):
    x = pd.read_excel(os.path.join(plan_dir, "merch_plan.xlsx"),
                      sheet_name=["Fixture_Layout", "Planogram", "Inputs_Fixtures"])
    inst, plan, fx = x["Fixture_Layout"], x["Planogram"], x["Inputs_Fixtures"]
    inst = inst.merge(fx[["fixture_code", "wall_ft_each", "footprint_sqft_each"]], on="fixture_code", how="left")
    a = {"store": {"total_sqft": 1000, "selling_pct": 0.70}}
    ap = os.path.join(plan_dir, "assumptions.json")
    if os.path.exists(ap):
        a = json.load(open(ap))
    return inst, plan, a["store"]


# ------------------------------------------------------------------ floor plan
def wall_path_point(s, W, Ds):
    """Distance s along the wall path: up left wall -> across back wall -> down right wall."""
    side = Ds - WALL_DEPTH
    if s <= side:
        return WALL_DEPTH / 2, s, 90
    s -= side
    if s <= W:
        return s, Ds - WALL_DEPTH / 2, 0
    s -= W
    return W - WALL_DEPTH / 2, max(side - s, 0), 90


def layout_wall(inst, W, Ds):
    """Returns list of (row, cx, cy, length, angle). Men start front-left, women end front-right."""
    perim = 2 * (Ds - WALL_DEPTH) + W
    walls = inst[inst.placement == "WALL"].copy()
    walls["g_ord"] = walls.fixture_group.map({g: i for i, g in enumerate(GROUP_ORDER)})
    men = walls[walls.gender == "M"].sort_values(["g_ord", "fixture_id"])
    wom = walls[walls.gender == "W"].sort_values(["g_ord", "fixture_id"], ascending=[False, False])
    total = walls.wall_ft_each.sum()
    scale = min(1.0, perim / total) if total else 1.0
    out = []
    s = 0.0
    for r in men.itertuples():
        L = r.wall_ft_each * scale
        out.append((r, *wall_path_point(s + L / 2, W, Ds)[:2], L, wall_path_point(s + L / 2, W, Ds)[2]))
        s += L
    s = perim - wom.wall_ft_each.sum() * scale
    for r in wom.itertuples():
        L = r.wall_ft_each * scale
        pt = wall_path_point(s + L / 2, W, Ds)
        out.append((r, pt[0], pt[1], L, pt[2]))
        s += L
    return out, scale


def pack_floor(items, x0, x1, y0, y1, fixed_scale=None):
    """Greedy row packing of floor fixtures into [x0,x1]x[y0,y1]; shrinks uniformly if it overflows."""
    scale, gap = (fixed_scale or 1.0), 0.6
    for _ in range(1 if fixed_scale else 8):
        rows, cur, cx, rh = [], [], x0, 0
        for r in items:
            s = math.sqrt(max(r.footprint_sqft_each, 1)) * scale
            w, h = s * 1.1, s / 1.1
            if cx + w > x1 + 1e-9 and cur:
                rows.append((cur, rh)); cur, cx, rh = [], x0, 0
            cur.append((r, cx, w, h)); cx += w + gap * scale; rh = max(rh, h)
        if cur:
            rows.append((cur, rh))
        depth = sum(rh for _, rh in rows) + gap * scale * max(len(rows) - 1, 0)
        if depth <= (y1 - y0) or not rows:
            break
        scale *= (y1 - y0) / depth * 0.97
    out, y = [], y0
    for cur, rh in rows:
        used = sum(w for _, _, w, _ in cur) + gap * scale * (len(cur) - 1)
        shift = (x1 - x0 - used) / 2          # centre each row inside its half
        for r, cx, w, h in cur:
            out.append((r, cx + shift, y, w, h))
        y += rh + gap * scale
    return out, scale


def draw_floor(inst, plan, store, width, path):
    total, selling = store["total_sqft"], store["total_sqft"] * store["selling_pct"]
    W = width
    Dt, Ds = total / W, selling / W
    fig = plt.figure(figsize=(15, 11))
    gs = fig.add_gridspec(1, 2, width_ratios=[2.3, 1.2], wspace=0.05)
    ax, side = fig.add_subplot(gs[0]), fig.add_subplot(gs[1])
    ax.set_aspect("equal"); ax.axis("off")
    ax.set_xlim(-2, W + 2); ax.set_ylim(-4.5, Dt + 2)

    ax.add_patch(Rectangle((0, 0), W, Ds, fc="#FAFAFA", ec="#222", lw=2))
    ax.add_patch(Rectangle((0, Ds), W, Dt - Ds, fc="#E8E8E8", ec="#222", lw=2, hatch="///", alpha=.6))
    ax.text(W / 2, Ds + (Dt - Ds) / 2, "STOCKROOM · TRIAL ROOMS · BILLING", ha="center", va="center",
            fontsize=9, color="#444", weight="bold",
            bbox=dict(fc="white", ec="none", pad=3, alpha=.85))
    ax.plot([W / 2, W / 2], [0, Ds], ls=":", c="#999", lw=1)
    ax.text(W * .25, -1.3, "MEN", ha="center", fontsize=13, weight="bold", color="#333")
    ax.text(W * .75, -1.3, "WOMEN", ha="center", fontsize=13, weight="bold", color="#333")
    ax.annotate("ENTRANCE", (W / 2, 0), (W / 2, -3.6), ha="center", fontsize=9, weight="bold",
                arrowprops=dict(arrowstyle="-|>", color="#222"))

    cnt = plan.groupby("fixture_id").sku_code.count().to_dict()

    # wall fixtures
    walls, wscale = layout_wall(inst, W, Ds)
    for r, cx, cy, L, ang in walls:
        w, h = (L, WALL_DEPTH) if ang == 0 else (WALL_DEPTH, L)
        ax.add_patch(Rectangle((cx - w / 2, cy - h / 2), w * .97, h * .97 if ang == 0 else h * .97,
                               fc=GROUP_COLOUR[r.fixture_group], ec="white", lw=.8, alpha=.92))
        ax.text(cx, cy, f"{short_id(r.fixture_id)}\n{cnt.get(r.fixture_id, 0)} SKUs", ha="center", va="center",
                rotation=ang if ang == 90 else 0, fontsize=4.6, color="white", weight="bold")

    # floor fixtures (men left half, women right half)
    fl = inst[inst.placement == "FLOOR"].copy()
    fl["g_ord"] = fl.fixture_group.map({g: i for i, g in enumerate(GROUP_ORDER)})
    pad = WALL_DEPTH + 1.0
    halves = {"M": (pad, W / 2 - .6), "W": (W / 2 + .6, W - pad)}
    sets = {g: list(fl[fl.gender == g].sort_values(["visibility", "g_ord", "fixture_id"],
                                                   ascending=[False, True, True]).itertuples()) for g in halves}
    common = min(pack_floor(sets[g], *halves[g], 2.2, Ds - pad)[1] for g in halves)   # same scale both sides
    for g, (x0, x1) in halves.items():
        placed, _ = pack_floor(sets[g], x0, x1, 2.2, Ds - pad, fixed_scale=common)
        for r, x, y, w, h in placed:
            ax.add_patch(Rectangle((x, y), w, h, fc=GROUP_COLOUR[r.fixture_group], ec="white", lw=1, alpha=.92))
            ax.text(x + w / 2, y + h / 2, f"{short_id(r.fixture_id)}\n{cnt.get(r.fixture_id, 0)} SKUs",
                    ha="center", va="center", fontsize=4.6, color="white", weight="bold")
    ax.set_title(f"Store planogram - floor plan  ({total:,.0f} sq ft, schematic, not to scale)",
                 fontsize=13, weight="bold", loc="left")

    # side panel
    side.axis("off")
    tiers = plan.perf_tier.value_counts()
    lines = [
        ("Store summary", True),
        (f"SKUs ranged: {len(plan)}", False),
        (f"Pieces in store: {int(plan.qty_total.sum()):,}", False),
        (f"Stock at retail: INR {plan.retail_value.sum() / 1e6:,.2f} mn", False),
        (f"Best / Better / Good: {tiers.get('Best', 0)} / {tiers.get('Better', 0)} / {tiers.get('Good', 0)}", False),
        (f"Fixtures: {len(inst)}  (wall {int((inst.placement == 'WALL').sum())}, "
         f"floor {int((inst.placement == 'FLOOR').sum())})", False),
        (f"Wall bays scaled to fit: {wscale:.0%}", False),
    ]
    y = .97
    for t, bold in lines:
        side.text(0, y, t, fontsize=12 if bold else 10, weight="bold" if bold else "normal", va="top",
                  transform=side.transAxes)
        y -= .045
    by_grp = plan.merge(inst[["fixture_id", "fixture_group"]], on="fixture_id").groupby("fixture_group").agg(
        skus=("sku_code", "count"), pcs=("qty_total", "sum"))
    y -= .02
    side.text(0, y, "By department", fontsize=12, weight="bold", va="top", transform=side.transAxes)
    y -= .05
    for g in GROUP_ORDER:
        if g in by_grp.index:
            side.text(0, y, f"{g.title()}: {by_grp.loc[g, 'skus']} SKUs · {by_grp.loc[g, 'pcs']:,} pcs",
                      fontsize=10, va="top", transform=side.transAxes)
            y -= .04
    side.legend(handles=[Patch(fc=GROUP_COLOUR[g], label=g.title()) for g in GROUP_ORDER],
                loc="lower left", title="Fixture colour = department", frameon=False, fontsize=9)
    fig.savefig(path, dpi=170, bbox_inches="tight")
    plt.close(fig)


# ------------------------------------------------------------------ fixture elevations
def draw_card(ax, r, items):
    slots = make_slots(int(r.sku_slots), int(r.levels))
    per_level = Counter(k for k, _, _ in slots)
    ncol = max(per_level.values())
    lookup = {(int(i.shelf_level), int(i.slot)): i for i in items.itertuples()}
    L = int(r.levels)
    fs = max(3.4, min(7.0, 24 / (ncol + 1)))
    for lvl in range(L, 0, -1):
        y = L - lvl
        n = per_level.get(lvl, 0)
        off = (ncol - n) / 2
        eye = (lvl == max(1, round(0.6 * L)))
        ax.text(-0.12, y + .5, "EYE" if eye else f"L{lvl}", ha="right", va="center", fontsize=5.5,
                weight="bold" if eye else "normal", color="#B3261E" if eye else "#666")
        for p in range(1, n + 1):
            x = off + p - 1
            it = lookup.get((lvl, p))
            if it is None:
                ax.add_patch(Rectangle((x + .04, y + .06), .92, .88, fc="white", ec="#BBB", ls="--", lw=.6))
                continue
            ax.add_patch(Rectangle((x + .04, y + .06), .92, .88, fc=TIER_COLOUR.get(it.perf_tier, "#EEE"),
                                   ec="#555", lw=.6))
            a, c = sku_label(it.sku_code)
            ax.text(x + .5, y + .60, a, ha="center", va="center", fontsize=fs, weight="bold")
            ax.text(x + .5, y + .30, f"{c} · {int(it.qty_total)}pc", ha="center", va="center", fontsize=fs * .9)
    ax.set_xlim(-0.7, ncol); ax.set_ylim(0, L)
    ax.axis("off")
    used = len(items); pcs = int(items.qty_total.sum()) if used else 0
    ax.set_title(f"{r.fixture_id}  |  {r.location}\n{used}/{int(r.sku_slots)} slots · {pcs}/{int(r.pieces_cap)} pcs",
                 fontsize=7.5, loc="left", weight="bold", color=GROUP_COLOUR[r.fixture_group])


def draw_elevations(inst, plan, out_dir, cols=3, rows=4):
    inst = inst.copy()
    inst["g_ord"] = inst.fixture_group.map({g: i for i, g in enumerate(GROUP_ORDER)})
    inst = inst.sort_values(["g_ord", "gender", "placement", "fixture_id"], ascending=[True, True, False, True])
    by_fix = {k: v for k, v in plan.groupby("fixture_id")}
    per_page = cols * rows
    pages = math.ceil(len(inst) / per_page)
    pdf = PdfPages(os.path.join(out_dir, "planogram_elevations.pdf"))
    for pg in range(pages):
        fig, axes = plt.subplots(rows, cols, figsize=(16.5, 11.7))
        fig.suptitle(f"Fixture planograms - page {pg + 1}/{pages}   (cell = SKU, colour = performance tier, "
                     f"'pc' = pieces per store)", fontsize=12, weight="bold", x=0.01, ha="left")
        fig.legend(handles=[Patch(fc=c, ec="#555", label=t) for t, c in TIER_COLOUR.items()] +
                          [Patch(fc="white", ec="#BBB", ls="--", label="Empty slot")],
                   loc="upper right", ncol=4, frameon=False, fontsize=9)
        chunk = inst.iloc[pg * per_page:(pg + 1) * per_page]
        for ax in axes.ravel():
            ax.axis("off")
        for ax, r in zip(axes.ravel(), chunk.itertuples()):
            draw_card(ax, r, by_fix.get(r.fixture_id, plan.iloc[0:0]))
        fig.tight_layout(rect=(0, 0, 1, 0.94), h_pad=2.2, w_pad=2)
        pdf.savefig(fig)
        fig.savefig(os.path.join(out_dir, f"elevations_p{pg + 1}.png"), dpi=130)
        plt.close(fig)
    pdf.close()
    return pages


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--plan-dir", default="merch_plan_output")
    ap.add_argument("--width-ft", type=float, default=24, help="store frontage used for the schematic")
    a = ap.parse_args()
    inst, plan, store = load(a.plan_dir)
    out = os.path.join(a.plan_dir, "visual")
    os.makedirs(out, exist_ok=True)
    draw_floor(inst, plan, store, a.width_ft, os.path.join(out, "planogram_floor.png"))
    pages = draw_elevations(inst, plan, out)
    print(f"Wrote floor plan and {pages} elevation pages to {out}/")


if __name__ == "__main__":
    main()
