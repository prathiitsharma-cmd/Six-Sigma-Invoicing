"""Stage 3 helper - loads the "definitions" half of 02_solutions/solutions_calc.py and a few shared constants.
Reads : 02_solutions/solutions_calc.py (everything before the line that starts building "out")
Writes: nothing (it is exec()-ed by comp.py)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Not run on its own. It executes the first part of the stage-2 script so that the data, the reason names and the
#   queue simulator are available, then defines baseline DPMO, penalty and rework constants and small helper functions.
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
import json, math, random, sys
import numpy as np, pandas as pd
sys.argv=['x']
src = open(paths.SOLUTIONS / 'solutions_calc.py').read()
src = src.split("out = {'excess_share'")[0]  # definitions only
exec(src)
from scipy.stats import norm
DPMO0 = R['headline']['base_adj']['dpmo']; DPU0 = R['headline']['base_adj']['dpu']
vol = R['copq']['vol_month']
PEN = vol*12*0.35*0.05*5000/1e5; RW = vol*12*0.25*12/60*350/1e5
EXTRA_L = R['lp']['level_saving_month']*12/1e5
def cut_of(effd, peakfix, mult=1.0, pmult=1.0):
    cr = sum(share[r]*min(1, effd[r]*mult) for r in order)
    return 1-(1-cr)*(1-excess_share*min(1,peakfix*pmult))
def comb_eff(mode='indep', w=0.5):
    d={}
    for r in order:
        a,b,c = eff['S1'][r], eff['S2'][r], eff['S3'][r]
        ind = 1-(1-a)*(1-b)*(1-c); mx = max(a,b,c)
        d[r] = ind if mode=='indep' else (mx if mode=='max' else w*ind+(1-w)*mx)
    return d
PF_C = 1-(1-0.2)*(1-0.6)*(1-0.1)
