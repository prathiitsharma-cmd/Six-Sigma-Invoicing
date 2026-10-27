"""Stage 2 - model the three solution options (A, B, C) on the case data.
Reads : outputs/results.json (or precomputed/results.json), data/original.xlsx
Writes: outputs/sol.json  (effectiveness, cost, payback and queue results for each solution)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Takes the Pareto of error reasons from stage 1 and, for each of the three solutions (A Digital Front Door, B Lean Flow,
#   C Supplier Partnership), estimates how much of each error reason it removes, what it costs and how fast it pays back.
#   It also re-runs the invoice queue simulation to see whether late (2+ day) invoices disappear. Result: sol.json.
#   NOTE: stage 3 re-uses the first half of this file (everything above the line that builds "out"), so keep them together.
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
import json, math, random
import numpy as np, pandas as pd
R = json.load(open(paths.find('results.json')))
SRC = paths.SRC
M = pd.read_excel(SRC, sheet_name='Measure', header=1, nrows=52); M.columns = ['Date', 'Inv', 'Samp', 'Err']
M['peak'] = (M.Inv >= 400).astype(int); M['dom'] = M.Date.dt.day; M['mon'] = M.Date.dt.strftime('%b')
M['dtype'] = np.where(M.Inv >= 400, np.where(M.dom >= 20, 'ME', 'MS'), 'N')
MA = M[~((M.Err == 0) & (M.peak == 1))]
pareto = R['pareto']; share = {p['reason']: p['share'] for p in pareto}
short = {'Mismatched invoice and PO amounts': 'PO amount mismatch', 'Unclear or missing line item descriptions': 'Unclear line items',
         'Missing supporting documents (e.g., delivery note)': 'Missing documents', 'System integration failure (e.g., ERP sync issues)': 'ERP sync failure',
         'Invoice submitted in wrong format': 'Wrong format', 'Duplicate invoice submission': 'Duplicate invoice', 'Missing invoice date': 'Missing date',
         'Incomplete invoice fields (e.g., missing bank details)': 'Incomplete fields', 'Incorrect vendor details': 'Wrong vendor details',
         'Missing Purchase Order (PO) reference': 'Missing PO ref', 'Currency mismatch': 'Currency mismatch', 'Unregistered vendor': 'Unregistered vendor',
         'Inconsistent invoice numbering': 'Inconsistent numbering', 'Wrong approval routing': 'Wrong routing', 'Incorrect tax codes': 'Wrong tax codes',
         'Late invoice submission': 'Late submission', 'Delayed approvals': 'Delayed approvals', 'Incorrect GST calculation': 'GST calc error', 'Manual data entry errors': 'Manual entry'}
order = [p['reason'] for p in pareto]
# effectiveness per reason (ANALYST ASSUMPTIONS) in Pareto order
EFF = {
 'S1': [.85, .70, .90, .50, .90, .95, .95, .90, .90, .95, .95, .95, .60, .70, .80, .20, .50, .85, .80],
 'S2': [.35, .40, .50, .30, .60, .40, .70, .60, .40, .60, .40, .50, .30, .60, .30, .50, .70, .30, .40],
 'S3': [.60, .50, .60, .10, .50, .30, .50, .60, .80, .70, .50, .90, .40, .10, .50, .60, .10, .50, .10],
}
# reorder EFF lists to match the pareto order produced by analysis (stable sort by count)
names_in_eff_order = ['Mismatched invoice and PO amounts', 'Unclear or missing line item descriptions', 'Missing supporting documents (e.g., delivery note)',
    'System integration failure (e.g., ERP sync issues)', 'Invoice submitted in wrong format', 'Duplicate invoice submission', 'Missing invoice date',
    'Incomplete invoice fields (e.g., missing bank details)', 'Incorrect vendor details', 'Missing Purchase Order (PO) reference', 'Currency mismatch',
    'Unregistered vendor', 'Inconsistent invoice numbering', 'Wrong approval routing', 'Incorrect tax codes', 'Late invoice submission', 'Delayed approvals',
    'Incorrect GST calculation', 'Manual data entry errors']
eff = {k: dict(zip(names_in_eff_order, v)) for k, v in EFF.items()}
# peak-excess errors: errors on peak days above what the non-peak rate would give
pk = MA[MA.peak == 1]; npk = MA[MA.peak == 0]
excess = pk.Err.sum() - pk.Samp.sum() * npk.Err.sum() / npk.Samp.sum()
excess_share = excess / MA.Err.sum()
PEAKFIX = {'S1': 0.2, 'S2': 0.6, 'S3': 0.1}   # ASSUMPTION: share of peak-excess errors removed
COPQ = R['copq']['base']; DPU0 = R['headline']['base_adj']['dpu']; DPMO0 = R['headline']['base_adj']['dpmo']
vol = R['copq']['vol_month']
COST = {'S1': dict(capex=45, opex=10, months=5), 'S2': dict(capex=6, opex=5, months=1.5), 'S3': dict(capex=8, opex=14, months=4)}
EXTRA = {'S1': 0, 'S2': R['lp']['level_saving_month'] * 12 / 1e5, 'S3': 0}   # S2 includes the submission calendar
# touch-time models (minutes) for steps 1-9, before exception share
TOUCH_BASE = {'S1': 0.5 + 0.2 + 1 + 0.5 + 1.5 + 1 + 1 + 0, 'S2': 1 + 1 + 5 + 3 + 5 + 2 + 3 + 0 + 2, 'S3': 20.0}
def des(vols, pks, touch_clean, exc_share, seed, staff=18, extra_peak=0):
    rnd = random.Random(seed); q = []; waits = []; late = 0; tot = 0; eb = []
    for day, (v, p) in enumerate(zip(vols, pks)):
        q += [day] * int(v); cap = (staff + (extra_peak if p else 0)) * (11 if p else 8) * 60; used = 0
        while q:
            st = (touch_clean + (12 if rnd.random() < exc_share else 0)) * rnd.uniform(0.8, 1.2)
            if used + st > cap: break
            used += st; a = q.pop(0); w = day - a; waits.append(w); tot += 1; late += w >= 2
        eb.append(len(q))
    return np.mean(waits), late / tot, max(eb)
def level(d, share=0.25):
    inv = d.Inv.astype(float).values.copy()
    for m in d.mon.unique():
        mi = (d.mon == m).values; pkm = mi & (d.dtype == 'ME').values; nm = mi & (d.dtype == 'N').values
        moved = np.clip(inv[pkm] - 400, 0, None) * share; inv[pkm] -= moved; inv[nm] += moved.sum() / nm.sum()
    return inv
out = {'excess_share': excess_share, 'sol': {}}
for s in ('S1', 'S2', 'S3'):
    cut_r = sum(share[r] * eff[s][r] for r in order)
    cut = 1 - (1 - cut_r) * (1 - excess_share * PEAKFIX[s])
    dpu1 = DPU0 * (1 - cut); dpmo1 = DPMO0 * (1 - cut)
    sigma1 = float(__import__('scipy.stats', fromlist=['norm']).norm.ppf(1 - dpu1 / 58) + 1.5)
    exc1 = 0.25 * (1 - cut)
    touch = TOUCH_BASE[s] + 12 * exc1
    vols = level(M) if s == 'S2' else M.Inv.values
    sims = [des(vols, M.peak.values, TOUCH_BASE[s], exc1, k, extra_peak=4 if s == 'S2' else 0) for k in range(20)]
    benefit = COPQ * cut / 1e5 + EXTRA[s]
    c = COST[s]; net = benefit - c['opex']
    npv = -c['capex'] + sum(net / 1.12 ** y for y in (1, 2, 3))
    # capacity released on normal days (FTE) vs 23-minute baseline
    fte_released = (23 - touch) * M.Inv.mean() / 60 / 8
    by_reason = {short[r]: share[r] * eff[s][r] for r in order}
    out['sol'][s] = dict(cut_reason=cut_r, cut=cut, dpmo=dpmo1, sigma=sigma1, pdef=1 - math.exp(-dpu1), exc=exc1, touch=touch,
                         wait=float(np.mean([x[0] for x in sims])), late2=float(np.mean([x[1] for x in sims])), maxbl=float(np.mean([x[2] for x in sims])),
                         copq_after=COPQ * (1 - cut), benefit=benefit, capex=c['capex'], opex=c['opex'], net=net,
                         payback=c['capex'] / net * 12 if net > 0 else None, npv=npv, roi=(net * 3 - c['capex']) / c['capex'], months=c['months'],
                         fte_released=fte_released, by_reason=by_reason,
                         cons_cut=1 - (1 - cut * 0.6), cons_net=COPQ * cut * 0.6 * (52.5 / 105.0) / 1e5 + EXTRA[s] * 0.6 - c['opex'] * 1.5)
    o = out['sol'][s]; o['cons_payback'] = (c['capex'] * 1.5) / o['cons_net'] * 12 if o['cons_net'] > 0 else None
# baseline DES for reference
b = [des(M.Inv.values, M.peak.values, 20, 0.25, k) for k in range(20)]
out['base_des'] = dict(wait=float(np.mean([x[0] for x in b])), late2=float(np.mean([x[1] for x in b])), maxbl=float(np.mean([x[2] for x in b])))
out['eff'] = {s: {short[r]: eff[s][r] for r in order} for s in eff}
out['peakfix'] = PEAKFIX; out['cost'] = COST
json.dump(out, open(paths.out('sol.json'), 'w'), indent=1, default=float)
print('excess share', round(excess_share, 3))
for s, o in out['sol'].items():
    print(s, {k: (round(v, 3) if isinstance(v, float) else v) for k, v in o.items() if k != 'by_reason'})
print(out['base_des'])
