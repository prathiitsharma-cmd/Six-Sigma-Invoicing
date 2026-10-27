"""Stage 4, step 1 - Laney u-prime control chart (control limits corrected for over-dispersion).
Reads : data/original.xlsx
Writes: outputs/laney.json  (centre line, sigma-z, limits and out-of-control days, baseline and after)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   A normal u-chart assumes errors are Poisson; this data is over-dispersed, so the normal limits are too tight.
#   Laney widens them by the spread of the standardised daily points. Deterministic - no random numbers.
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
import pandas as pd, numpy as np, json
SRC=paths.SRC
def load(s,n):
    d=pd.read_excel(SRC,sheet_name=s,header=1,nrows=n); d.columns=['Date','Inv','Samp','Err']; return d
out={}
B=load('Measure',52); I=load('Improve',51)
adj=~((B.Date.dt.month==8)&(B.Date.dt.day==25))
def lan(d, mask=None):
    dd=d if mask is None else d[mask]
    ub=dd.Err.sum()/dd.Samp.sum(); u=dd.Err/dd.Samp; s=np.sqrt(ub/dd.Samp); z=(u-ub)/s
    mr=np.abs(np.diff(z.values)).mean(); sz=mr/1.128
    return ub, sz
ub,sz=lan(B,adj)
u=B.Err/B.Samp; s=np.sqrt(ub/B.Samp)
ucl=ub+3*sz*s; lcl=np.maximum(0,ub-3*sz*s); uclc=ub+3*s; lclc=np.maximum(0,ub-3*s)
out['base']=dict(ubar=ub, sz=sz, d=list(B.Date.dt.strftime('%d-%b')), u=list(u), ucl=list(ucl), lcl=list(lcl), ucl_std=list(uclc), lcl_std=list(lclc),
                 out=[B.Date.dt.strftime('%d-%b')[i] for i in range(52) if u[i]>ucl[i] or u[i]<lcl[i]],
                 out_std_adj=int(((u>uclc)|(u<lclc))[adj].sum()), out_lan_adj=int(((u>ucl)|(u<lcl))[adj].sum()),
                 hi_std=int((u>uclc)[adj].sum()), lo_std=int((u<lclc)[adj].sum()))
ui,szi=lan(I); uu=I.Err/I.Samp; si=np.sqrt(ui/I.Samp); szu=max(1,szi)
out['imp']=dict(ubar=ui, sz=szi, sz_used=szu, u=list(uu), ucl=list(ui+3*szu*si), out=int((uu>ui+3*szu*si).sum()))
# month-end clustering of the 7 standard-limit signals (adjusted)
sig=[B.Date.dt.strftime('%d-%b')[i] for i in range(52) if adj[i] and (u[i]>uclc[i] or u[i]<lclc[i])]
out['base']['std_signals']=sig; out['base']['std_signals_inv']=[int(B.Inv[i]) for i in range(52) if adj[i] and (u[i]>uclc[i] or u[i]<lclc[i])]
# share of over-dispersion explained by peak: dispersion within non-peak / peak
for nm,m in (('nonpeak',adj&(B.Inv<400)),('peak',adj&(B.Inv>=400))):
    dd=B[m]; ubb=dd.Err.sum()/dd.Samp.sum(); phi=(((dd.Err-dd.Samp*ubb)**2)/(dd.Samp*ubb)).sum()/(len(dd)-1); out['phi_'+nm]=phi
print({k:(v if not isinstance(v,dict) else {a:b for a,b in v.items() if not isinstance(b,list) or len(b)<10}) for k,v in out.items()})
json.dump(out,open(paths.out('laney.json'),'w'),default=float)
