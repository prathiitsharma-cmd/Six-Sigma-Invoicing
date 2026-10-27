"""Stage 4, step 3 - conditional decision rule: infer how well Hybrid H is working from the first audited month and escalate to H+ or A+B+C if it is not.
Reads : outputs/stage3.json (or precomputed/), 03_stress/comp.py (exec-ed)
Writes: outputs/rule.json
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Builds a table of "if the measured October cut is X, the true effectiveness is probably Y", sets escalation thresholds,
#   and tests the rule on 20,000 simulated worlds (scenario seed 5, audit-noise seed 7).
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
exec(open(paths.STRESS / 'comp.py').read())
import json, math
from scipy.stats import poisson
H = mk(['A1','A2','A3','A4','B1','B2','B3','B4','B5','C2']); HP = mk(['A1','A2','A3','A4','A5','B1','B2','B3','B4','B5','C2','C4']); ABC = np.ones(K, bool)
def cut_live(mask, t, **kw):
    """cut in cash month t (t=1 Sep) with base scenario except kw; parts live if go-live < t"""
    P = scen(**kw); ids = np.where(mask)[0]
    m2 = np.zeros(K, bool); m2[[i for i in ids if MONTH[i] * (1 + P['delay'][0] / 5) < t]] = True
    return float(evaluate(m2, P)['cut'][0]) if m2.any() else 0.0
ms = np.round(np.arange(0.40, 1.21, 0.01), 2)
tab = [dict(m=float(m), oct=cut_live(H, 2, m=m), dec=cut_live(H, 4, m=m), H=float(evaluate(H, scen(m=m))['cut'][0]), HP=float(evaluate(HP, scen(m=m))['cut'][0]), ABC=float(evaluate(ABC, scen(m=m))['cut'][0])) for m in ms]
ST = json.load(open(paths.find('stage3.json')))
mH, mHP, mABC = ST['Hybrid H (recommended)']['m75'], ST['Hybrid H+ (quality variant)']['m75'], ST['A+B+C (original)']['m75']
def at(m, key):  # interpolate
    xs = [r['m'] for r in tab]; ys = [r[key] for r in tab]; return float(np.interp(m, xs, ys))
thr = {k: dict(m=v, oct=at(v, 'oct'), dec=at(v, 'dec')) for k, v in (('H', mH), ('HP', mHP), ('ABC', mABC))}
# sampling noise of one audited month: ~10% of 8,332 invoices; DPU0 baseline
from_base = DPU0; n_m = round(0.1 * R['copq']['vol_month'])
def obs_band(c, n=n_m):
    lam = n * from_base * (1 - c); lo, hi = poisson.ppf(0.05, lam), poisson.ppf(0.95, lam)
    return 1 - hi / (n * from_base), 1 - lo / (n * from_base)
bands = {k: obs_band(v['oct']) for k, v in thr.items()}; bands20 = {k: obs_band(v['oct'], 2 * n_m) for k, v in thr.items()}
# Monte Carlo evaluation of the rule (decide at end-Oct on the observed October cut)
Pm = mc_draws(20000, seed=5); N = 20000
def run_masks(mask, extra_delay_codes=()):
    saved = MONTH.copy()
    for c in extra_delay_codes: MONTH[IDX[c]] += 2      # escalation parts start building only after the end-Oct decision
    r = evaluate(mask, Pm); MONTH[:] = saved; return r
rH = evaluate(H, Pm); rHP0 = evaluate(HP, Pm)
rHPe = run_masks(HP, ('A5', 'C4')); rABCe = run_masks(ABC, ('A5', 'C1', 'C3', 'C4', 'C5'))
# October cut per run (B live, delays apply)
ids = np.where(H)[0]
def oct_cut_runs():
    out = np.zeros(N)
    for j in range(N):
        pass
    return out
# vectorised: evaluate only B parts with each run's draws, then mask runs where B is not yet live in Oct
Bmask = mk(['B1','B2','B3','B4','B5']); rB = evaluate(Bmask, Pm)
golive_B = np.outer(1 + Pm['delay'] / 5, MONTH[Bmask])          # N x 5
live_frac = (golive_B < 2).all(1)
oct_true = rB['cut']   # measured in the first full month with all of B live (planned October; later if B slips)
rng = np.random.default_rng(7); n_dec = 2 * n_m   # audit sample doubled to 20% in the decision month
obs_err = rng.poisson(n_dec * DPU0 * (1 - oct_true)); oct_obs = 1 - obs_err / (n_dec * DPU0)
obs_err1 = rng.poisson(n_m * DPU0 * (1 - oct_true)); oct_obs1 = 1 - obs_err1 / (n_m * DPU0)
def policy(c_H, c_HP, c_ABC):
    cut = np.where(oct_obs >= c_H, rH['cut'], np.where(oct_obs >= c_HP, rHPe['cut'], np.where(oct_obs >= c_ABC, rABCe['cut'], rH['cut'])))
    npv = np.where(oct_obs >= c_H, rH['npv'], np.where(oct_obs >= c_HP, rHPe['npv'], np.where(oct_obs >= c_ABC, rABCe['npv'], rH['npv'])))
    esc = np.where(oct_obs >= c_H, 0, np.where(oct_obs >= c_HP, 1, np.where(oct_obs >= c_ABC, 2, 3)))
    return dict(p75=float(np.mean(cut >= .75)), pneg=float(np.mean(npv < 0)), npv50=float(np.median(npv)), npv5=float(np.percentile(npv, 5)), esc_HP=float(np.mean(esc == 1)), esc_ABC=float(np.mean(esc == 2)), redesign=float(np.mean(esc == 3)))
pol = policy(thr['H']['oct'], thr['HP']['oct'], thr['ABC']['oct'])
def bucket(v): return np.where(v >= thr['H']['oct'], 0, np.where(v >= thr['HP']['oct'], 1, np.where(v >= thr['ABC']['oct'], 2, 3)))
truth = np.where(Pm['m'] >= mH, 0, np.where(Pm['m'] >= mHP, 1, np.where(Pm['m'] >= mABC, 2, 3)))
acc20 = float(np.mean(bucket(oct_obs) == truth)); acc10 = float(np.mean(bucket(oct_obs1) == truth))
near = float(np.mean(np.abs(bucket(oct_obs) - truth) <= 1))
# value when the right call is to stay with H
stayH_ok = float(np.mean(rH['cut'][bucket(oct_obs) == 0] >= .75)); buckets = [float(np.mean(bucket(oct_obs) == k)) for k in range(4)]
p75_by_bucket = {k: dict(H=float(np.mean(rH['cut'][bucket(oct_obs) == k] >= .75)), HPe=float(np.mean(rHPe['cut'][bucket(oct_obs) == k] >= .75)), ABCe=float(np.mean(rABCe['cut'][bucket(oct_obs) == k] >= .75))) for k in range(4)}
base = dict(alwaysH=dict(p75=float(np.mean(rH['cut'] >= .75)), pneg=float(np.mean(rH['npv'] < 0)), npv50=float(np.median(rH['npv'])), npv5=float(np.percentile(rH['npv'], 5))),
            alwaysHP=dict(p75=float(np.mean(rHP0['cut'] >= .75)), pneg=float(np.mean(rHP0['npv'] < 0)), npv50=float(np.median(rHP0['npv'])), npv5=float(np.percentile(rHP0['npv'], 5))),
            alwaysABC=dict(p75=float(np.mean(evaluate(ABC, Pm)['cut'] >= .75))))
# how often the rule picks the "right" option (true m vs threshold)
trueH = Pm['m'] >= mH
acc = float(np.mean((oct_obs >= thr['H']['oct']) == trueH))
out = dict(thr=thr, bands=bands, bands20=bands20, n_month=n_m, policy=pol, base=base, accuracy_stayH=acc, acc20=acc20, acc10=acc10, within_one=near, buckets=buckets, p75_by_bucket=p75_by_bucket, stayH_ok=stayH_ok, oct_live_share=float(live_frac.mean()),
           tab=[t for t in tab if abs(t['m'] * 100 % 5) < 1e-6], HP_golive=dict(A5=MONTH[IDX['A5']] + 2, C4=MONTH[IDX['C4']] + 2))
json.dump(out, open(paths.out('rule.json'), 'w'), indent=1, default=float)
print(json.dumps({k: v for k, v in out.items() if k != 'tab'}, indent=1, default=float))
for t in out['tab']: print({k: round(v, 3) for k, v in t.items()})
