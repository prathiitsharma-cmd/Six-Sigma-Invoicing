"""Stage 3, step 7 - month-by-month error cut of Hybrid H while its parts go live (months 1-8).
Reads : 03_stress/comp.py
Writes: outputs/ramp_H.json  (8 numbers, % cut by month; not used by the workbook, kept for the slide deck)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Switches on only the parts of Hybrid H whose go-live month is before month t and reports the resulting % error cut.
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
exec(open(paths.STRESS / 'comp.py').read())
import json
H = mk(['A1','A2','A3','A4','B1','B2','B3','B4','B5','C2'])
P = scen(); ids = np.where(H)[0]
out = []
for t in range(1, 9):
    live = np.array([[MONTH[i] < t for i in ids]])
    # reuse evaluate internals via a tiny wrapper: evaluate with a mask of only the live parts gives the same cut (no cross-part dependencies change: A1 live by t=4 before A2/A3)
    m2 = np.zeros(K, bool); m2[[i for i in ids if MONTH[i] < t]] = True
    cut = float(evaluate(m2, P)['cut'][0]) if m2.any() else 0.0
    out.append(round(cut * 100, 1))
print(out)
json.dump(out, open(paths.out('ramp_H.json'), 'w'))
