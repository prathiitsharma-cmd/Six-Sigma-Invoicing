"""Stage 3, step 2 - shortlist the best combinations and rank them with a 3,000-scenario Monte Carlo.
Reads : outputs/stage1.pkl (or precomputed/), 03_stress/comp.py
Writes: outputs/stage2.pkl  (the shortlist with probabilities and four weighted scores)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Keeps the Pareto-efficient combinations plus the top performers on several criteria, then re-scores each one under 3,000
#   random scenarios (seed 11) to get chance of reaching a 75% cut, chance of negative NPV, and so on. Four weightings
#   (Balanced, Quality-first, Money-first, Ops-first) turn these into scores; the Balanced winner is "Hybrid H".
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
exec(open(paths.STRESS / 'comp.py').read())
import pickle, pandas as pd, time
D = pd.DataFrame(pickle.load(open(paths.find('stage1.pkl'), 'rb')))
X = np.c_[D.cut_c, D.npv_c, D.npv_h, -D.late_s, D.cut_h]
# Pareto-efficient set
order_ = np.argsort(-X[:, 0]); eff_ = []
for i in order_:
    dom = False
    for j in eff_:
        if np.all(X[j] >= X[i]) and np.any(X[j] > X[i]): dom = True; break
    if not dom: eff_.append(i)
print('pareto', len(eff_))
rk = D[['cut_c', 'npv_c', 'npv_h', 'cut_h']].rank(ascending=False).sum(1) + D.late_s.rank()
short_ = set(eff_) | set(D.nlargest(150, 'cut_c').index) | set(D.nlargest(100, 'npv_c').index) | set(rk.nsmallest(200).index)
orig_bits = [int(sum(1 << i for i in range(K) if m_[i])) for m_ in ORIG.values()]
short_ |= set(D.index[D.bits.isin(orig_bits)])
print('shortlist', len(short_))
Pm = mc_draws(3000)
res = []; t0 = time.time()
for i in short_:
    bits = int(D.bits[i]); mask = np.array([(bits >> k) & 1 for k in range(K)], bool)
    r = evaluate(mask, Pm)
    res.append(dict(bits=bits, p75=np.mean(r['cut'] >= .75), p500=np.mean(r['dpmo'] <= 500), med_dpmo=np.median(r['dpmo']), pneg=np.mean(r['npv'] < 0),
                    npv5=np.percentile(r['npv'], 5), npv50=np.median(r['npv']), pb50=np.median(r['payback']), pb90=np.percentile(r['payback'], 90),
                    ontime=np.mean(r['lastlive'] <= 6), late_s=D.late_s[i], late_b=D.late_b[i], outscope=D.outscope[i], legal=D.legal[i], n=D.n[i],
                    capex=D.capex[i], opex=D.opex[i], cut_b=D.cut_b[i], npv_b=D.npv_b[i]))
S = pd.DataFrame(res); print('mc done', round(time.time() - t0), 's')
def comp_score(S, w):
    q1 = S.p75; q2 = ((DPMO0 - S.med_dpmo) / (DPMO0 - 500)).clip(0, 1); f1 = 1 - S.pneg; f2 = ((S.npv5 + 50) / 150).clip(0, 1)
    cap = (1 - S.late_s / .3).clip(0, 1); tl = S.ontime; sc = 1 - (S.outscope + S.legal) / 4
    return 100 * (w[0] * q1 + w[1] * q2 + w[2] * f1 + w[3] * f2 + w[4] * cap + w[5] * tl + w[6] * sc)
W = {'Balanced (main)': [.25, .15, .2, .1, .15, .1, .05], 'Quality-first': [.35, .25, .1, .05, .1, .1, .05],
     'Money-first': [.15, .05, .35, .2, .1, .1, .05], 'Ops-first': [.2, .1, .15, .05, .3, .15, .05]}
for k, w in W.items(): S[k] = comp_score(S, w)
S['label'] = S.bits.apply(lambda b: ' + '.join(CODE[k] for k in range(K) if (b >> k) & 1))
pickle.dump(S, open(paths.out('stage2.pkl'), 'wb'))
pd.set_option('display.width', 250); pd.set_option('display.max_colwidth', 80)
cols = ['label', 'Balanced (main)', 'Quality-first', 'Money-first', 'Ops-first', 'p75', 'med_dpmo', 'pneg', 'npv5', 'npv50', 'pb50', 'late_s', 'ontime', 'capex', 'opex']
print(S.sort_values('Balanced (main)', ascending=False)[cols].head(15).round(3).to_string())
for k in W: print(k, 'winner:', S.loc[S[k].idxmax(), 'label'], round(S[k].max(), 1))
print(S[S.bits.isin(orig_bits)][cols].round(3).to_string())
