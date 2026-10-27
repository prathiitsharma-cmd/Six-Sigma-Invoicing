"""Stage 3, step 6 - "one change at a time" test around Hybrid H (add or remove each single component).
Reads : outputs/stage2.pkl (or precomputed/), 03_stress/comp.py
Writes: outputs/marginal.json (used by the workbook), outputs/marginal_raw.json (text version); also prints the H-lite checks
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Starts from the recommended Hybrid H, flips each of the 15 components in or out one at a time, rescoring with a 3,000-scenario
#   Monte Carlo (seed 11), and reports how the overall score changes. Shows which parts of H really matter.
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
exec(open(paths.STRESS / 'comp.py').read())
import pickle, pandas as pd
S = pickle.load(open(paths.find('stage2.pkl'), 'rb'))
def bm(b): return np.array([(int(b) >> k) & 1 for k in range(K)], bool)
H = bm(S.loc[S['Balanced (main)'].idxmax()].bits)
Pm = mc_draws(3000)
def metrics(mask):
    r = evaluate(mask, Pm); b = evaluate(mask, scen())
    ls, _ = fluid_late2(mask, evaluate(mask, scen(m=.8, ov=.5))['cut'][0], vm=1.2, staff=16, comp=.7)
    return dict(cut_b=b['cut'][0], npv_b=b['npv'][0], p75=np.mean(r['cut'] >= .75), med_dpmo=np.median(r['dpmo']), pneg=np.mean(r['npv'] < 0),
                npv5=np.percentile(r['npv'], 5), late_s=ls, ontime=np.mean(r['lastlive'] <= 6),
                outscope=sum(FLAG[i] == 'outscope' for i in np.where(mask)[0]), legal=int(mask[IDX['C5']]))
def score(d, w=(.25, .15, .2, .1, .15, .1, .05)):
    q2 = min(1, max(0, (DPMO0 - d['med_dpmo']) / (DPMO0 - 500))); f2 = min(1, max(0, (d['npv5'] + 50) / 150)); cap = min(1, max(0, 1 - d['late_s'] / .3))
    return 100 * (w[0] * d['p75'] + w[1] * q2 + w[2] * (1 - d['pneg']) + w[3] * f2 + w[4] * cap + w[5] * d['ontime'] + w[6] * (1 - (d['outscope'] + d['legal']) / 4))
h = metrics(H); hs = score(h); print('H', {k: round(v, 3) for k, v in h.items()}, round(hs, 1))
rows = []; MR = []
for i in range(K):
    m2 = H.copy(); m2[i] = not m2[i]
    if not valid(m2): rows.append((('remove ' if H[i] else 'add ') + CODE[i], 'invalid (needs digital intake)')); continue
    d = metrics(m2); MR.append(dict(change=(('remove ' if H[i] else 'add ') + CODE[i]), dscore=score(d)-hs, **{k: float(v) for k, v in d.items()})); rows.append((('remove ' if H[i] else 'add ') + CODE[i], f"score {score(d)-hs:+5.1f} | base cut {d['cut_b']*100:.1f} | base NPV {d['npv_b']:.0f} | P(cut>=75%) {d['p75']*100:.0f}% | P(NPV<0) {d['pneg']*100:.1f}% | NPV P5 {d['npv5']:.0f} | med DPMO {d['med_dpmo']:.0f} | late2 stress {d['late_s']*100:.1f}%"))
for r in rows: print(r)
import json
json.dump([dict(change=a, txt=b) for a, b in rows], open(paths.out('marginal_raw.json'),'w'))
lite = S.loc[S['Money-first'].idxmax()]; L = bm(lite.bits)
d = metrics(L); print('LITE', ' + '.join(CODE[i] for i in range(K) if L[i]), {k: round(v, 3) for k, v in d.items()}, round(score(d), 1))
r = evaluate(L, mc_draws(20000, 5)); b = evaluate(L, scen())
print('LITE base', b['cut'][0], b['dpmo'][0], b['npv'][0], b['payback'][0], b['capex'][0], b['opex'][0], 'MC pneg', np.mean(r['npv'] < 0), 'p75', np.mean(r['cut'] >= .75), 'dpmo50', np.median(r['dpmo']), 'npv5', np.percentile(r['npv'], 5), 'npv50', np.median(r['npv']), 'pb50', np.median(r['payback']), 'ontime', np.mean(r['lastlive'] <= 6))
print('LITE DES v1.2', des_late2(L, b['cut'][0], vm=1.2), 'v1.4', des_late2(L, b['cut'][0], vm=1.4), 'v1.2 s16', des_late2(L, b['cut'][0], vm=1.2, staff=16))
# how many combos evaluated / how many beat original A+B+C on both P(NPV<0) and P(cut>=75%)
o = S[S.label == ' + '.join(CODE)]
print('orig ABC in S', o[['p75', 'pneg', 'npv5']].values)
print('shortlisted combos better than A+B+C on P(NPV<0) and NPV P5 with P(cut>=75%) >= 50%:', ((S.pneg < .256) & (S.npv5 > -43.7) & (S.p75 >= .5)).sum())

json.dump(dict(H=dict(score=hs, **{k: float(v) for k, v in h.items()}), rows=MR), open(paths.out('marginal.json'),'w'), indent=1)
