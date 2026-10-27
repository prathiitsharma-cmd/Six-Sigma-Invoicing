"""Stage 3, step 1 - evaluate every valid combination of the 15 components (about 25,599 of the 32,767 subsets).
Reads : 03_stress/comp.py (exec-ed)
Writes: outputs/stage1.pkl  (one row of results per combination, ~7 MB)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Loops over every non-empty subset of the 15 components, skips combinations that are not allowed (e.g. A2 without a digital
#   front door), and scores each under three fixed scenarios: base, cautious and harsh. Slow-ish but fully deterministic.
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
exec(open(paths.STRESS / 'comp.py').read())
import time, pickle
P = {k: np.array([b, c, h], float) for k, (b, c, h) in dict(
    m=(1, .8, .6), ov=(1, .5, 0), copq=(105, 79, 52.5), capm=(1, 1.4, 1.8), opm=(1, 1.25, 1.5), delay=(0, 2, 4), comp=(1, .7, .4), grn=(1, .85, .7), legal=(1, .7, 0)).items()}
rows = []; t0 = time.time()
for bits in range(1, 2 ** K):
    mask = np.array([(bits >> i) & 1 for i in range(K)], bool)
    if not valid(mask): continue
    r = evaluate(mask, P)
    lb, bb = fluid_late2(mask, r['cut'][0])
    ls, bs = fluid_late2(mask, r['cut'][1], vm=1.2, staff=16, comp=.7)
    rows.append(dict(bits=bits, n=int(mask.sum()), cut_b=r['cut'][0], cut_c=r['cut'][1], cut_h=r['cut'][2], npv_b=r['npv'][0], npv_c=r['npv'][1], npv_h=r['npv'][2],
                     pb_b=r['payback'][0], last_c=r['lastlive'][1], capex=r['capex'][0], opex=r['opex'][0], late_b=lb, late_s=ls, bl_s=bs,
                     outscope=sum(FLAG[i] == 'outscope' for i in np.where(mask)[0]), legal=int(mask[IDX['C5']])))
print(len(rows), 'valid subsets', round(time.time() - t0), 's')
pickle.dump(rows, open(paths.out('stage1.pkl'), 'wb'))
