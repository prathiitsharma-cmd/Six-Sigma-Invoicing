"""Stage 3, step 3 - stress-test the candidate solutions (A, B, C, A+B+C, Hybrid H, H+, H-lite).
Reads : outputs/stage2.pkl (or precomputed/), 03_stress/comp.py
Writes: outputs/stage3.json  (base case + ~15 what-if tests + 20,000-run Monte Carlo per candidate)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   For each candidate set of components it reruns the model with weaker effects, higher costs, delays, no calendar, and so on,
#   solves for the break-even effectiveness, runs a 20,000-scenario Monte Carlo (seed 5) and a 10-run queue simulation.
#   The queue numbers here are later replaced by 200-run versions in des200.py (run that next).
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
exec(open(paths.STRESS / 'comp.py').read())
import pickle, pandas as pd, json
S = pickle.load(open(paths.find('stage2.pkl'), 'rb'))
best = S.loc[S['Balanced (main)'].idxmax()]; qbest = S.loc[S['Quality-first'].idxmax()]
def bm(b): return np.array([(int(b) >> k) & 1 for k in range(K)], bool)
CANDS = dict(ORIG); CANDS['Hybrid H (recommended)'] = bm(best.bits); CANDS['Hybrid H+ (quality variant)'] = bm(qbest.bits); CANDS['H-lite (money variant)'] = bm(S.loc[S['Money-first'].idxmax()].bits)
print({k: ' + '.join(CODE[i] for i in range(K) if v[i]) for k, v in CANDS.items()})
OUT = {}
def one(mask, **kw): return {k: v[0] for k, v in evaluate(mask, scen(**kw)).items()}
def solve(f, lo=0.0, hi=3.0):
    if not f(hi): return None
    for _ in range(50):
        mid = (lo + hi) / 2
        if f(mid): hi = mid
        else: lo = mid
    return hi
Pm = mc_draws(20000, seed=5)
for name, mask in CANDS.items():
    o = {}
    b = one(mask); o['base'] = b
    c0 = b['cut']
    o['des_base'] = des_late2(mask, c0)
    o['eff07'] = one(mask, m=.7)
    o['overlap0'] = one(mask, ov=0)
    o['copq52'] = one(mask, copq=52.5)
    o['cost'] = one(mask, capm=2, opm=1.5)
    o['delay3'] = one(mask, delay=3)
    o['grn'] = one(mask, grn=.7)
    o['cal0'] = one(mask, comp=0)
    o['legal0'] = one(mask, legal=0)
    o['des_cal0'] = des_late2(mask, c0, comp=0) if mask[IDX['B3']] else o['des_base']
    o['des_v12'] = des_late2(mask, c0, vm=1.2); o['des_v14'] = des_late2(mask, c0, vm=1.4)
    o['des_s16'] = des_late2(mask, c0, staff=16); o['des_v12s16'] = des_late2(mask, c0, vm=1.2, staff=16)
    o['m75'] = solve(lambda m: evaluate(mask, scen(m=m))['cut'][0] >= .75)
    o['m500'] = solve(lambda m: evaluate(mask, scen(m=m))['dpmo'][0] <= 500)
    o['mnpv'] = solve(lambda m: evaluate(mask, scen(m=m))['npv'][0] >= 0, 0, 3)
    o['copq_be'] = solve(lambda q: evaluate(mask, scen(copq=q))['npv'][0] >= 0, 0, 400)
    r = evaluate(mask, Pm)
    o['mc'] = dict(pneg=np.mean(r['npv'] < 0), npv5=np.percentile(r['npv'], 5), npv50=np.median(r['npv']), npv95=np.percentile(r['npv'], 95),
                   p75=np.mean(r['cut'] >= .75), p500=np.mean(r['dpmo'] <= 500), dpmo50=np.median(r['dpmo']), dpmo90=np.percentile(r['dpmo'], 90),
                   pb50=np.median(r['payback']), pb90=np.percentile(r['payback'], 90), ontime=np.mean(r['lastlive'] <= 6))
    o['flags'] = dict(outscope=int(sum(FLAG[i] == 'outscope' for i in np.where(mask)[0])), legal=int(mask[IDX['C5']]), n=int(mask.sum()),
                      parts=[CODE[i] for i in range(K) if mask[i]], touch=touch_clean(mask), lastmonth=float(MONTH[mask].max()))
    OUT[name] = o
    print(name, 'done')
json.dump(OUT, open(paths.out('stage3.json'), 'w'), default=float, indent=1)
