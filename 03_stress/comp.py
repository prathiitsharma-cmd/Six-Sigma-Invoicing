"""Stage 3 core model - the 15 building blocks (A1..C5) of solutions A, B, C with costs, timing and effects.
Reads : 03_stress/base.py (exec-ed)
Writes: nothing (it is exec()-ed by every other stage-3 script and by 04_statistics)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Defines each of the 15 sub-components (what errors it removes, capex, opex, go-live month) and the function evaluate(),
#   which turns any chosen subset of components into error cut, DPMO, 3-year NPV and payback for many scenarios at once.
#   Also holds the queue helpers and mc_draws(), the seeded Monte Carlo scenario generator (default seed 11).
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
exec(open(paths.STRESS / 'base.py').read())
import itertools, numpy as np
RS = names_in_eff_order; NR = len(RS)
SH = np.array([share[r] for r in RS])
EORIG = {g: np.array(EFF[s]) for g, s in (('A', 'S1'), ('B', 'S2'), ('C', 'S3'))}
EXTRA_L = R['lp']['level_saving_month'] * 12 / 1e5          # Rs L/yr overtime avoided by the calendar (LP model)
EXC0 = 0.25; EXCMIN = 12.0; TOUCH0 = 20.0

# reason index shortcuts
PO, LI, DOC, ERP, FMT, DUP, DATE, INC, VEN, POREF, CUR, UNR, NUM, ROUTE, TAX, LATE, DELAY, GST, MAN = range(19)

# ---- components: group, name, weights {reason: share of the group's effect}, capex, opex, go-live month, peak-fix, touch delta (min), flags
COMPS = [
 # A - Digital Front Door
 ('A', 'A1 Supplier portal: mandatory fields & attachments', {PO: .1, LI: .4, DOC: 1, FMT: .6, DATE: .8, INC: .8, VEN: .2, POREF: .6, LATE: 1, MAN: .2}, 14, 4, 3, .10, -3.0, ''),
 ('A', 'A2 Automated 3-way PO/GRN match + PO-line picking', {PO: .8, LI: .6, POREF: .4}, 13, 2, 5, 0, -5.0, 'needs_digital'),
 ('A', 'A3 Rules engine: duplicates, vendor master, GSTIN, tax, currency', {PO: .1, DUP: .7, INC: .2, VEN: .8, CUR: .8, UNR: 1, NUM: .5, TAX: .6, GST: .6}, 7, 2, 4, 0, -2.3, 'needs_digital'),
 ('A', 'A4 GST e-invoice (IRN) feed for large vendors', {FMT: .4, DUP: .3, DATE: .2, CUR: .2, NUM: .5, TAX: .4, GST: .4, MAN: .3}, 4, 1, 2, 0, 0.0, ''),
 ('A', 'A5 Auto ERP posting + approval workflow with 24 h SLA', {ERP: 1, ROUTE: 1, DELAY: 1, MAN: .5}, 7, 1, 5, .11, -4.0, 'needs_digital'),
 # B - Lean Flow & Peak Control
 ('B', 'B1 Intake triage desk (2-min completeness checklist)', {PO: .4, LI: .2, DOC: .6, FMT: .3, DUP: .6, DATE: .6, INC: .6, VEN: .7, POREF: .7, CUR: .5, UNR: 1, TAX: .5, GST: .5}, 1, .5, 1, .05, +2.0, ''),
 ('B', 'B2 Standard invoice template + vendor checklist', {PO: .3, LI: .6, DOC: .4, FMT: .7, DATE: .4, INC: .4, POREF: .3, CUR: .5, NUM: .6, LATE: .3}, 2, .5, 1.5, 0, 0.0, ''),
 ('B', 'B3 Month-end submission calendar (level-load)', {LATE: .7}, 1, .5, 1.5, .30, 0.0, 'calendar'),
 ('B', 'B4 Flex pool: 4 cross-trained staff, last 7 days', {DELAY: .4, MAN: .4}, 1, 3, 1, .40, 0.0, 'flex'),
 ('B', 'B5 Daily u-chart huddle, Kaizen, standard work', {PO: .3, LI: .2, ERP: 1, DUP: .4, VEN: .3, NUM: .4, ROUTE: 1, TAX: .5, DELAY: .6, GST: .5, MAN: .6}, 1, .5, 1, 0, 0.0, ''),
 # C - Supplier Quality Partnership
 ('C', 'C1 Vendor scorecards, reviews, 2 coordinators', {PO: .2, LI: .5, DOC: .2, FMT: .5, DUP: .5, DATE: .5, INC: .2, CUR: .3, NUM: .6, TAX: .4, LATE: .4, GST: .5, MAN: 1}, 3, 8, 3, .10, 0.0, 'outscope'),
 ('C', 'C2 Internal PO-before-work + GRN within 48 h', {PO: .6, DOC: .4, POREF: .5, ROUTE: 1, DELAY: 1}, 1, 1, 2, 0, 0.0, 'outscope'),
 ('C', 'C3 Vendor onboarding gate + master-data clean-up', {ERP: 1, INC: .5, VEN: 1, CUR: .7, UNR: 1, TAX: .6, GST: .5}, 3, 1, 3, 0, 0.0, 'outscope'),
 ('C', 'C4 Certification + priority-pay incentive', {PO: .1, LI: .3, DOC: .1, FMT: .3, DUP: .5, NUM: .4, LATE: .6}, .5, 3, 4, 0, 0.0, ''),
 ('C', 'C5 "No PO, no pay" / return incomplete invoices', {PO: .1, LI: .2, DOC: .3, FMT: .2, DATE: .5, INC: .3, POREF: .5}, .5, 1, 1.5, 0, 0.0, 'legal'),
]
K = len(COMPS); NAMES = [c[1] for c in COMPS]; GRP = np.array([c[0] for c in COMPS]); CODE = [c[1].split()[0] for c in COMPS]
CAPEX = np.array([c[3] for c in COMPS], float); OPEX = np.array([c[4] for c in COMPS], float); MONTH = np.array([c[5] for c in COMPS], float)
PFX = np.array([c[6] for c in COMPS], float); TOUCHD = np.array([c[7] for c in COMPS], float); FLAG = [c[8] for c in COMPS]
IDX = {c: i for i, c in enumerate(CODE)}
# component effectiveness, calibrated so each full solution reproduces the original per-reason effectiveness exactly
WM = np.zeros((K, NR))
for i, c in enumerate(COMPS):
    for r, w in c[2].items(): WM[i, r] = w
ECOMP = np.zeros((K, NR))
for g in 'ABC':
    ids = [i for i in range(K) if GRP[i] == g]; tot = WM[ids].sum(0)
    for r in range(NR):
        if EORIG[g][r] > 0:
            assert tot[r] > 0, (g, r)
            for i in ids: ECOMP[i, r] = 1 - (1 - EORIG[g][r]) ** (WM[i, r] / tot[r])
for g in 'ABC':
    ids = [i for i in range(K) if GRP[i] == g]
    assert np.allclose(1 - np.prod(1 - ECOMP[ids], 0), EORIG[g]), g
    assert abs(sum(CAPEX[ids]) - {'A': 45, 'B': 6, 'C': 8}[g]) < 1e-9 and abs(sum(OPEX[ids]) - {'A': 10, 'B': 5, 'C': 14}[g]) < 1e-9
GMASK = {g: (GRP == g) for g in 'ABC'}

def valid(mask):
    if not mask.any(): return False
    digital = mask[IDX['A1']] or mask[IDX['A4']]
    for c in ('A2', 'A3', 'A5'):
        if mask[IDX[c]] and not digital: return False
    return True

def coverage(mask):
    return 1.0 if mask[IDX['A1']] else (0.5 if mask[IDX['A4']] else 0.0)   # IRN feed alone covers ~half the volume (vendors > Rs 5 Cr)

def evaluate(mask, P, months_eval=True):
    """mask: bool[K]; P: dict of arrays (N,) -> metrics arrays (N,)"""
    m, ov, copq, capm, opm, delay, comp, grn, legal = (P[k] for k in ('m', 'ov', 'copq', 'capm', 'opm', 'delay', 'comp', 'grn', 'legal'))
    N = len(m); ids = np.where(mask)[0]; cov = coverage(mask)
    e = np.minimum(1, ECOMP[ids][None, :, :] * m[:, None, None])               # N,k,R
    pf = np.minimum(1, PFX[ids][None, :] * m[:, None])                        # N,k
    for j, i in enumerate(ids):
        if FLAG[i] == 'needs_digital': e[:, j, :] *= cov
        if CODE[i] == 'A2' and not mask[IDX['C2']]: e[:, j, PO] *= grn       # 3-way match weaker without GRN discipline
        if FLAG[i] == 'legal': e[:, j, :] *= legal[:, None]
        if FLAG[i] == 'calendar': pf[:, j] *= comp
    golive = MONTH[ids][None, :] * (1 + delay[:, None] / 5)                   # N,k ; delays scale with build length
    def cut_at(live):  # live N,k bool
        el = e * live[:, :, None]; pl = pf * live
        G = []; PG = []
        for g in 'ABC':
            sel = [j for j, i in enumerate(ids) if GRP[i] == g]
            if sel:
                G.append(1 - np.prod(1 - el[:, sel, :], 1)); PG.append(1 - np.prod(1 - pl[:, sel], 1))
        G = np.stack(G); PG = np.stack(PG)                                    # g,N,R ; g,N
        comb = ov[:, None] * (1 - np.prod(1 - G, 0)) + (1 - ov[:, None]) * G.max(0)
        pfc = ov * (1 - np.prod(1 - PG, 0)) + (1 - ov) * PG.max(0)
        cr = (comb * SH).sum(1)
        return 1 - (1 - cr) * (1 - excess_share * pfc)
    full = cut_at(np.ones_like(golive, bool))
    capex = CAPEX[ids].sum() * capm; opex = OPEX[ids].sum() * opm
    has_cal = mask[IDX['B3']]
    extra = EXTRA_L * comp if has_cal else 0 * comp
    cal_live = (golive[:, list(ids).index(IDX['B3'])] if has_cal else np.full(N, 99.0))
    # monthly cash flow for 36 months
    cum = -capex.copy() if np.ndim(capex) else -capex * np.ones(N)
    yr = np.zeros((3, N)); pb = np.full(N, 99.0)
    for t in range(1, 37):
        if t <= 12:
            live = golive < t
            ct = cut_at(live) if live.any() else np.zeros(N)
        else: ct = full
        ben = (copq * ct + extra * (cal_live < t)) / 12
        cf = ben - opex / 12
        yr[(t - 1) // 12] += cf; cum = cum + cf
        pb = np.where((pb == 99) & (cum >= 0), t, pb)
    npv = -capex + yr[0] / 1.12 + yr[1] / 1.12 ** 2 + yr[2] / 1.12 ** 3
    lastlive = golive.max(1)
    return dict(cut=full, dpmo=DPMO0 * (1 - full), npv=npv, payback=pb, capex=capex * np.ones(N), opex=opex * np.ones(N), lastlive=lastlive,
                benefit=copq * full + extra)

def touch_clean(mask):
    cov = coverage(mask); t = TOUCH0
    for i in np.where(mask)[0]:
        t += TOUCHD[i] * (cov if FLAG[i] == 'needs_digital' else 1)
    return max(t, 4.0)

def fluid_late2(mask, cut, vm=1.0, staff=18, comp=1.0, flex=None):
    """deterministic FIFO fluid queue: share of invoices processed >=2 days after arrival, and max backlog"""
    vols = (level(M, 0.25 * comp) if mask[IDX['B3']] else M.Inv.values.astype(float)) * vm
    extra = (4 if mask[IDX['B4']] else 0) if flex is None else flex
    t = touch_clean(mask) + EXCMIN * EXC0 * (1 - cut)
    q = []; late = 0; tot = 0; bl = 0
    for d, (v, p) in enumerate(zip(vols, M.peak.values)):
        q.append([d, v]); capn = (staff + (extra if p else 0)) * (11 if p else 8) * 60 / t
        while q and capn > 1e-9:
            a, n = q[0]; take = min(n, capn); capn -= take; tot += take; late += take * (d - a >= 2)
            if take >= n - 1e-9: q.pop(0)
            else: q[0][1] -= take
        bl = max(bl, sum(x[1] for x in q))
    return late / tot, bl

def des_late2(mask, cut, vm=1.0, staff=18, comp=1.0, runs=10):
    vols = (level(M, 0.25 * comp) if mask[IDX['B3']] else M.Inv.values.astype(float)) * vm
    extra = 4 if mask[IDX['B4']] else 0
    sims = [des(vols, M.peak.values, touch_clean(mask), EXC0 * (1 - cut), k, staff=staff, extra_peak=extra) for k in range(runs)]
    return float(np.mean([x[1] for x in sims])), float(np.mean([x[2] for x in sims]))

def scen(N=1, **kw):
    base = dict(m=1.0, ov=1.0, copq=105.0, capm=1.0, opm=1.0, delay=0.0, comp=1.0, grn=1.0, legal=1.0); base.update(kw)
    return {k: np.full(N, float(v)) for k, v in base.items()}

def mc_draws(N, seed=11):
    r = np.random.default_rng(seed)
    return dict(m=r.uniform(.5, 1.1, N), ov=r.uniform(0, 1, N), copq=r.triangular(52.5, 105, 158, N), capm=r.uniform(1, 1.8, N), opm=r.uniform(1, 1.5, N),
                delay=r.uniform(0, 4, N), comp=r.uniform(.4, 1.0, N), grn=r.uniform(.7, 1.0, N), legal=(r.uniform(0, 1, N) > .3).astype(float))

def mk(codes):
    mask = np.zeros(K, bool)
    for c in codes: mask[IDX[c]] = True
    return mask
ORIG = {'A (original)': mk(['A1', 'A2', 'A3', 'A4', 'A5']), 'B (original)': mk(['B1', 'B2', 'B3', 'B4', 'B5']),
        'C (original)': mk(['C1', 'C2', 'C3', 'C4', 'C5']), 'A+B+C (original)': np.ones(K, bool)}
