"""EXTRA (not needed for the workbook) - quick sanity check of the four original option sets.
Reads : 03_stress/comp.py
Writes: nothing (console output only)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Prints cut, DPMO, NPV, payback and queue results for A, B, C and A+B+C using the component model, as a smoke test.
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[2]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
exec(open(paths.STRESS / 'comp.py').read())
b = scen()
for n, mask in ORIG.items():
    r = evaluate(mask, b)
    c = r['cut'][0]
    fl = fluid_late2(mask, c); ds = des_late2(mask, c, runs=5)
    print(n, f"cut {c*100:.1f} dpmo {r['dpmo'][0]:.0f} npv(ramp) {r['npv'][0]:.1f} pb {r['payback'][0]:.0f} capex {r['capex'][0]} opex {r['opex'][0]} touch {touch_clean(mask):.1f} fluid {fl[0]*100:.1f}%/{fl[1]:.0f} DES {ds[0]*100:.1f}%/{ds[1]:.0f}")
base_mask = np.zeros(K,bool)
print('fluid baseline (no comps): ', end='')
t = TOUCH0 + 3
# baseline check via fake: use des directly
print(np.mean([des(M.Inv.values, M.peak.values, 20, .25, k)[1] for k in range(5)]))
