"""EXTRA (not needed for the workbook) - extra statistics on the post-improvement audit (exact intervals, Bayesian check, 25-Aug sensitivity).
Reads : data/original.xlsx, 03_stress/base.py (exec-ed)
Writes: nothing (console output only)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Exploratory script. Prints confidence intervals for the after-improvement error rate and checks sensitivity to the 25-Aug day.
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[2]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
exec(open(paths.STRESS / 'base.py').read())
from scipy.stats import norm, chi2, beta
I = pd.read_excel(SRC, sheet_name='Improve', header=1); I = I.iloc[:, :4]; I.columns=['Date','Inv','Samp','Err']; I=I.dropna()
E, Sm = I.Err.sum(), I.Samp.sum()
lo, hi = chi2.ppf(0.025, 2*E)/2, chi2.ppf(0.975, 2*E+2)/2
print('after: errors', E, 'samples', Sm, 'DPMO', E/Sm/58*1e6, 'CI', lo/Sm/58*1e6, hi/Sm/58*1e6)
# probability true DPMO > 500 (one-sided, Poisson, Bayesian flat prior gamma)
from scipy.stats import gamma
post = gamma(E+1, scale=1/(Sm*58)); print('P(true DPMO>500)=', 1-post.cdf(500/1e6))
# invoices with any error: share of audited invoices — Poisson estimate 1-exp(-dpu); with overdispersion less
dpu = E/Sm; print('pdef poisson', 1-np.exp(-dpu), 'upper', 1-np.exp(-hi/Sm))
# peak vs normal after
I['peak']=I.Inv>=400
for k,g in I.groupby('peak'): print('after peak' if k else 'after normal', g.Err.sum()/g.Samp.sum(), g.Samp.sum(), g.Err.sum())
print('after days', len(I), 'last date', I.Date.max())
# 25-Aug sensitivity
Mrec = M; print('baseline recorded DPU', M.Err.sum()/M.Samp.sum(), 'adj', MA.Err.sum()/MA.Samp.sum())
print('reduction recorded', 1-dpu/(M.Err.sum()/M.Samp.sum()), 'adj', 1-dpu/(MA.Err.sum()/MA.Samp.sum()))
# peak threshold sensitivity for excess share
for th in (350,400,450,500):
    pk_=MA[MA.Inv>=th]; np_=MA[MA.Inv<th]
    ex=(pk_.Err.sum()-pk_.Samp.sum()*np_.Err.sum()/np_.Samp.sum())/MA.Err.sum(); print('peak threshold',th,'excess share',round(ex,3), 'peak days', len(pk_))
# DES stress
print('== DES capacity stress (mean of 10 runs): late2 share, max backlog')
TB = TOUCH_BASE
def run(s, vm=1.0, staff=18, extra=None):
    vols = (level(M) if s=='S2' else M.Inv.values)*vm
    cut = cut_of(eff[s],PEAKFIX[s]) if s!='base' else 0
    exc1 = 0.25*(1-cut)
    tc = TB[s] if s!='base' else 20
    ep = (4 if s=='S2' else 0) if extra is None else extra
    sims=[des(vols, M.peak.values, tc, exc1, k, staff=staff, extra_peak=ep) for k in range(10)]
    return round(np.mean([x[1] for x in sims])*100,1), round(np.mean([x[2] for x in sims]))
for s in ('base','S1','S2','S3'):
    print(s, 'base', run(s), '| vol+20%', run(s,1.2), '| vol+40%', run(s,1.4), '| 2 analysts absent', run(s,1.0,16), '| vol+20% & 16 staff', run(s,1.2,16))
print('B without flex pool', run('S2',1.0,18,0), ' B calendar 0% moved?')
# B with no calendar compliance
def runB_nocal(vm=1.0):
    sims=[des(M.Inv.values*vm, M.peak.values, TB['S2'], 0.25*(1-cut_of(eff['S2'],PEAKFIX['S2'])), k, extra_peak=4) for k in range(10)]
    return round(np.mean([x[1] for x in sims])*100,1), round(np.mean([x[2] for x in sims]))
print('B, vendors ignore calendar', runB_nocal(), 'vol+20%', runB_nocal(1.2))
# decision matrix weights
sc = {'A':[8.2,2.5,4.0,6,9],'B':[5.0,7.9,8.2,8,4],'C':[5.5,5.0,5.2,5,7]}
w0=[.35,.2,.2,.15,.1]
def score(w): return {k: sum(a*b for a,b in zip(v,w)) for k,v in sc.items()}
print('base', score(w0))
for wi in np.arange(0.35,0.66,0.05):
    rest=1-wi; w=[wi]+[x/0.65*rest for x in w0[1:]]; s_=score(w); print('impact weight', round(wi,2), {k:round(v,2) for k,v in s_.items()}, max(s_,key=s_.get))
