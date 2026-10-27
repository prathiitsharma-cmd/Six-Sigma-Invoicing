"""Stage 4, step 4 - grid search for the best escalation thresholds of the decision rule from rule.py.
Reads : 04_statistics/rule.py (the part before "pol = policy(" is exec-ed), outputs/stage3.json (via rule.py)
Writes: outputs/rule_final.json
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Tries many threshold pairs for "escalate to H+" and "escalate to A+B+C", picks the final rule and records its performance
#   against always-H and always-H+. Uses the same seeds as rule.py, so results are repeatable.
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
exec(open(paths.STATS / 'rule.py').read().split("pol = policy(")[0])
res = []
def summarize(cut, npv): return dict(p75=float(np.mean(cut >= .75)), npv50=float(np.median(npv)), pneg=float(np.mean(npv < 0)), npv5=float(np.percentile(npv, 5)), mean=float(np.mean(npv)))
for c in [0.30, 0.32, 0.34, 0.36, 0.38, 0.40, 0.42]:
    esc = oct_obs < c
    res.append(('H+ if c1<%.2f' % c, float(esc.mean()), summarize(np.where(esc, rHPe['cut'], rH['cut']), np.where(esc, rHPe['npv'], rH['npv']))))
for c in [0.34, 0.36, 0.38]:
    esc = oct_obs < c
    res.append(('A+B+C if c1<%.2f' % c, float(esc.mean()), summarize(np.where(esc, rABCe['cut'], rH['cut']), np.where(esc, rABCe['npv'], rH['npv']))))
res.append(('always H', 0, summarize(rH['cut'], rH['npv']))); res.append(('always H+ (from start)', 1, summarize(rHP0['cut'], rHP0['npv'])))
rA = evaluate(ABC, Pm); res.append(('always A+B+C', 1, summarize(rA['cut'], rA['npv'])))
res.append(('always H+ escalated late', 1, summarize(rHPe['cut'], rHPe['npv'])))
for r in res: print(f"{r[0]:28s} esc {r[1]:.3f}  p75 {r[2]['p75']:.3f}  npv50 {r[2]['npv50']:6.1f}  mean {r[2]['mean']:6.1f}  pneg {r[2]['pneg']:.3f}  npv5 {r[2]['npv5']:6.1f}")
# incremental value of H+ conditional on c1 bucket
for lo, hi in [(0, .30), (.30, .34), (.34, .36), (.36, .40), (.40, 1)]:
    sel = (oct_obs >= lo) & (oct_obs < hi)
    print(f'c1 in [{lo:.2f},{hi:.2f}) share {sel.mean():.3f}  P75 H {np.mean(rH["cut"][sel]>=.75):.3f}  H+ {np.mean(rHPe["cut"][sel]>=.75):.3f}  ABC {np.mean(rABCe["cut"][sel]>=.75):.3f}  dNPV H+ {np.median(rHPe["npv"][sel]-rH["npv"][sel]):.1f}  dNPV ABC {np.median(rABCe["npv"][sel]-rH["npv"][sel]):.1f}  med m {np.median(Pm["m"][sel]):.2f}')

def rule(lo, hi):
    esc = (oct_obs >= lo) & (oct_obs < hi)
    cut = np.where(esc, rHPe['cut'], rH['cut']); npv = np.where(esc, rHPe['npv'], rH['npv'])
    return dict(esc=float(esc.mean()), **summarize(cut, npv), redesign=float(np.mean(oct_obs < lo)), stay=float(np.mean(oct_obs >= hi)))
FINAL = rule(0.34, 0.40)
print('FINAL targeted rule', FINAL)
# pass chance by zone for the slide
zones = {}
for nm, lo, hi in [('stay', 0.40, 1.01), ('escalate', 0.34, 0.40), ('rediagnose', -1, 0.34)]:
    sel = (oct_obs >= lo) & (oct_obs < hi)
    zones[nm] = dict(share=float(sel.mean()), p75_H=float(np.mean(rH['cut'][sel] >= .75)), p75_HP=float(np.mean(rHPe['cut'][sel] >= .75)), p75_ABC=float(np.mean(rABCe['cut'][sel] >= .75)),
                     dnpv_HP=float(np.median(rHPe['npv'][sel] - rH['npv'][sel])), dnpv_ABC=float(np.median(rABCe['npv'][sel] - rH['npv'][sel])), med_m=float(np.median(Pm['m'][sel])))
    print(nm, zones[nm])
# measurement precision of the decision month
from scipy.stats import poisson
def band(c, n):
    lam = n * DPU0 * (1 - c); return 1 - poisson.ppf(0.95, lam) / (n * DPU0), 1 - poisson.ppf(0.05, lam) / (n * DPU0)
prec = {n: band(0.37, n) for n in (833, 1666)}
print('precision at 37% cut', prec)
# Oct cut implied by m (base case) for the thresholds
print('oct cut at m=0.75', at(0.75, 'oct'), 'm=0.608', at(mHP, 'oct'), 'dec at 0.75', at(0.75, 'dec'), 'dec at 0.608', at(mHP, 'dec'))
json.dump(dict(final=FINAL, zones=zones, precision={str(k): v for k, v in prec.items()}, alwaysH=summarize(rH['cut'], rH['npv']), alwaysHP=summarize(rHP0['cut'], rHP0['npv']),
               thresholds=dict(stay=0.40, esc=0.34, m_stay=0.75, m_esc=float(mHP), dec_stay=at(0.75, 'dec'), dec_esc=at(mHP, 'dec')),
               HP_cost=dict(capex=7.5, opex=4.0)), open(paths.out('rule_final.json'), 'w'), indent=1, default=float)
