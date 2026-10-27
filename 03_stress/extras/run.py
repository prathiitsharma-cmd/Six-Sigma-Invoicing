"""EXTRA (not needed for the workbook) - early stress tests of the original A/B/C options (haircut, overlap, cost, delay ...).
Reads : 03_stress/base.py (exec-ed)
Writes: nothing (console output only)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Exploratory script written before comp.py existed. Prints one-at-a-time sensitivity tables for solutions A, B, C and A+B+C.
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[2]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
exec(open(paths.STRESS / 'base.py').read())
OUT = {}
COST2 = dict(S1=(45,10,5), S2=(6,5,1.5), S3=(8,14,4), C=(59,29,5))
def econ(s, cut, copq=105, capm=1, opm=1, ramp=True, extra=None, delay=0):
    cap, opx, mo = COST2[s]; cap*=capm; opx*=opm
    ex = (EXTRA_L if s in ('S2','C') else 0) if extra is None else extra
    ben = copq*cut + ex
    net = ben - opx
    m0 = (mo+delay) if ramp else 0
    # year 1 only gets benefit after go-live month m0 (opex assumed from month 0)
    y1 = ben*max(0,12-m0)/12 - opx
    yrs = [y1, net, net] if ramp else [net]*3
    npv = -cap + sum(v/1.12**(i+1) for i,v in enumerate(yrs))
    # payback month with ramp
    cum=-cap; pb=None
    for m in range(1,61):
        cum += (ben if m>m0 else 0)/12 - opx/12
        if cum>=0: pb=m; break
    return dict(benefit=ben, net=net, npv=npv, payback=pb)
def dpmo(c): return DPMO0*(1-c)
def sig(c): return norm.ppf(1-DPU0*(1-c)/58)+1.5
EFFS = {'S1':eff['S1'],'S2':eff['S2'],'S3':eff['S3'],'C':comb_eff('indep')}
PFS = {'S1':0.2,'S2':0.6,'S3':0.1,'C':PF_C}
NM = {'S1':'A Digital Front Door','S2':'B Lean Flow','S3':'C Supplier','C':'A+B+C'}

# 1. effectiveness haircut
print('== 1 effectiveness haircut (multiplier on every per-reason effectiveness & peak fix)')
for s in EFFS:
    row=[]
    for m in (1.0,0.9,0.8,0.7,0.6,0.5):
        c=cut_of(EFFS[s],PFS[s],m,m); e=econ(s,c,ramp=False)
        row.append(f"x{m}: cut {c*100:4.1f}% DPMO {dpmo(c):6.0f} NPV {e['npv']:6.1f}")
    print(NM[s]); [print('   ',r) for r in row]
# breakeven multipliers
def solve(f, lo=0.0, hi=3.0):
    for _ in range(60):
        mid=(lo+hi)/2
        if f(mid): hi=mid
        else: lo=mid
    return hi
print('== breakevens')
for s in EFFS:
    m75 = solve(lambda m: cut_of(EFFS[s],PFS[s],m,m)>=0.75)
    m500 = solve(lambda m: dpmo(cut_of(EFFS[s],PFS[s],m,m))<=500)
    mnpv = solve(lambda m: econ(s,cut_of(EFFS[s],PFS[s],m,m),ramp=False)['npv']>=0)
    mnpv_r = solve(lambda m: econ(s,cut_of(EFFS[s],PFS[s],m,m),ramp=True)['npv']>=0)
    print(NM[s], f"mult for >=75% cut {m75:.2f} | for DPMO<=500 {m500:.2f} (cap 1.0 eff) | NPV>=0 {mnpv:.2f} | NPV>=0 with ramp {mnpv_r:.2f}")
# max achievable with all eff=1 for remaining? DPMO if combined eff capped at 1
# 2 overlap
print('== 2 combination rule')
for mode,pf in (('indep',PF_C),('blend',0.5*PF_C+0.5*0.6),('max',0.6)):
    c=cut_of(comb_eff(mode),pf); e=econ('C',c,ramp=True)
    print(mode, f"cut {c*100:.1f}% DPMO {dpmo(c):.0f} sigma {sig(c):.2f} NPV {e['npv']:.1f} payback {e['payback']}")
# 3 COPQ / penalty assumptions
print('== 3 COPQ level (ramp on)')
for copq,lab in ((52.5,'penalty trigger 2%'),(65,'MC P5'),(105,'base'),(147.4,'46% defective'),(158,'trigger 8%')):
    print(lab, ' | '.join(f"{NM[s].split()[0]} NPV {econ(s,cut_of(EFFS[s],PFS[s]),copq)['npv']:6.1f} pb {econ(s,cut_of(EFFS[s],PFS[s]),copq)['payback']}" for s in EFFS))
# breakeven COPQ for NPV>=0 with ramp
for s in EFFS:
    c=cut_of(EFFS[s],PFS[s]); b=solve(lambda q: econ(s,c,q)['npv']>=0,0,300)
    print('breakeven COPQ', NM[s], round(b,1), 'L/yr  (penalty part alone =',round(PEN,1),')')
# 4 penalties tied to tax-relevant reasons only
print('== 4 penalty only from tax-relevant reasons')
taxr = ['Incorrect GST calculation','Incorrect tax codes','Mismatched invoice and PO amounts','Duplicate invoice submission','Unregistered vendor','Incorrect vendor details']
ts = sum(share[r] for r in taxr)
for s in EFFS:
    c=cut_of(EFFS[s],PFS[s]); ctax = sum(share[r]*EFFS[s][r] for r in taxr)/ts
    copq_cut = (PEN*ctax + RW*c)/(PEN+RW)
    e = econ(s, copq_cut, 105)
    print(NM[s], f"overall cut {c*100:.0f}% vs tax-relevant cut {ctax*100:.0f}% -> NPV {e['npv']:.1f}")
# 5 cost overruns
print('== 5 cost overruns (ramp on)')
for capm,opm in ((1,1),(1.5,1),(2,1),(1,1.5),(2,1.5)):
    print(f'capex x{capm} opex x{opm}:', ' | '.join(f"{NM[s].split()[0]} NPV {econ(s,cut_of(EFFS[s],PFS[s]),105,capm,opm)['npv']:6.1f} pb {econ(s,cut_of(EFFS[s],PFS[s]),105,capm,opm)['payback']}" for s in EFFS))
# 6 delays
print('== 6 go-live delay (months)')
for d in (0,3,6):
    print(f'+{d} mo:', ' | '.join(f"{NM[s].split()[0]} NPV {econ(s,cut_of(EFFS[s],PFS[s]),delay=d)['npv']:6.1f} pb {econ(s,cut_of(EFFS[s],PFS[s]),delay=d)['payback']}" for s in EFFS))
print('model NPV (no ramp) vs ramp:', {NM[s]: (round(econ(s,cut_of(EFFS[s],PFS[s]),ramp=False)['npv'],1), round(econ(s,cut_of(EFFS[s],PFS[s]))['npv'],1)) for s in EFFS})
# 9 joint Monte Carlo
print('== 9 joint Monte Carlo 20k')
rng=np.random.default_rng(7); N=20000
res={s:dict(npv=[],dp=[],pb=[]) for s in EFFS}
for i in range(N):
    m=rng.uniform(0.5,1.1); q=rng.triangular(52.5,105,158); cm=rng.uniform(1,1.8); om=rng.uniform(1,1.5); dl=rng.uniform(0,4); ov=rng.uniform(0,1)
    for s in EFFS:
        if s=='C':
            e_=comb_eff('blend',ov); pf=ov*PF_C+(1-ov)*0.6
        else: e_=EFFS[s]; pf=PFS[s]
        c=cut_of(e_,pf,m,m); e=econ(s,c,q,cm,om,True,None,dl)
        res[s]['npv'].append(e['npv']); res[s]['dp'].append(dpmo(c)); res[s]['pb'].append(e['payback'] or 99)
for s in EFFS:
    n=np.array(res[s]['npv']); d=np.array(res[s]['dp']); p=np.array(res[s]['pb'])
    print(NM[s], f"P(NPV<0) {np.mean(n<0)*100:.1f}% | NPV P5/P50/P95 {np.percentile(n,5):.0f}/{np.percentile(n,50):.0f}/{np.percentile(n,95):.0f} | P(DPMO<=500) {np.mean(d<=500)*100:.1f}% | P(cut>=75%) {np.mean(d<=DPMO0*0.25)*100:.1f}% | payback P50/P90 {np.percentile(p,50):.0f}/{np.percentile(p,90):.0f} mo")
