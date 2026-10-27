"""Stage 1 - statistical analysis of the case data (baseline, improvement, root causes, queueing, cost of poor quality).
Reads : data/original.xlsx  (sheets Measure, Improve, Error Reasons, Process Steps)
Writes: outputs/results.json  (every statistic the later stages need)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Loads the 52 baseline days and 51 post-improvement days of invoice audits and computes error rates (DPMO, sigma level),
#   confidence intervals, control-chart limits, a regression, the Pareto of error reasons, the invoice-processing queue,
#   a cost-of-poor-quality model with Monte Carlo ranges, and early cost/benefit numbers for the solution ideas.
#   Everything is collected in one dictionary R and saved as results.json. Random draws use fixed seeds, so reruns match.
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
import json, math, heapq, random
import numpy as np, pandas as pd
from scipy import stats, optimize
import statsmodels.api as sm
import statsmodels.formula.api as smf

SRC = paths.SRC
def load(sh, n):
    d = pd.read_excel(SRC, sheet_name=sh, header=1, nrows=n)
    d.columns = ['Date', 'Inv', 'Samp', 'Err']; return d
M = load('Measure', 52); I = load('Improve', 51)
OPP = 58; PEAK = 400; CUT = 20
for d, per in ((M, 'base'), (I, 'imp')):
    d['period'] = per
    d['dom'] = d.Date.dt.day; d['mon'] = d.Date.dt.strftime('%b'); d['wd'] = d.Date.dt.strftime('%a')
    d['dpu'] = d.Err / d.Samp; d['dpmo'] = d.dpu / OPP * 1e6
    d['peak'] = (d.Inv >= PEAK).astype(int)
    d['dtype'] = np.where(d.Inv >= PEAK, np.where(d.dom >= CUT, 'Month-end peak', 'Month-start carry-over'), 'Normal')
    d['qend'] = ((d.mon == 'Sep') & (d.dtype == 'Month-end peak')).astype(int)
    d['idx'] = range(1, len(d) + 1)
M['anom'] = ((M.Err == 0) & (M.peak == 1)).astype(int); I['anom'] = 0
MA = M[M.anom == 0].copy()            # adjusted baseline
R = {}

def dpmo(e, s): return e / (s * OPP) * 1e6
def sigma(e, s): return stats.norm.ppf(1 - e / (s * OPP)) + 1.5
R['headline'] = dict(
    base_rec=dict(dpmo=dpmo(M.Err.sum(), M.Samp.sum()), sigma=sigma(M.Err.sum(), M.Samp.sum()), dpu=M.Err.sum() / M.Samp.sum()),
    base_adj=dict(dpmo=dpmo(MA.Err.sum(), MA.Samp.sum()), sigma=sigma(MA.Err.sum(), MA.Samp.sum()), dpu=MA.Err.sum() / MA.Samp.sum()),
    imp=dict(dpmo=dpmo(I.Err.sum(), I.Samp.sum()), sigma=sigma(I.Err.sum(), I.Samp.sum()), dpu=I.Err.sum() / I.Samp.sum()),
)
H = R['headline']
H['reduction'] = 1 - H['imp']['dpu'] / H['base_adj']['dpu']
H['pdef_base'] = 1 - math.exp(-H['base_adj']['dpu']); H['pdef_imp'] = 1 - math.exp(-H['imp']['dpu'])

# ---------------- descriptive
def desc(x):
    q1, q3 = np.percentile(x, [25, 75])
    return dict(n=len(x), mean=float(np.mean(x)), median=float(np.median(x)), sd=float(np.std(x, ddof=1)), min=float(np.min(x)),
                q1=float(q1), q3=float(q3), max=float(np.max(x)), cv=float(np.std(x, ddof=1) / np.mean(x)) if np.mean(x) else None,
                skew=float(stats.skew(x, bias=False)))
R['desc'] = {f'{p}_{k}': desc(d[k].values) for p, d in (('base', M), ('imp', I)) for k in ('Inv', 'Err', 'dpu')}

# ---------------- rate normalisation demo
R['norm'] = dict(
    corr_raw_err_vs_inv=float(np.corrcoef(M.Inv, M.Err)[0, 1]),
    corr_dpu_vs_inv=float(np.corrcoef(M.Inv, M.dpu)[0, 1]),
    example=[dict(date=r.Date.strftime('%d-%b'), inv=int(r.Inv), samp=int(r.Samp), err=int(r.Err), dpu=float(r.dpu))
             for _, r in M[M.Date.dt.strftime('%d-%b').isin(['08-Aug', '22-Aug', '26-Sep', '28-Aug'])].iterrows()])

# ---------------- trend
def trend(d):
    s = stats.linregress(d.idx, d.dpu)
    return dict(slope_per_day=s.slope, p=s.pvalue, r2=s.rvalue ** 2)
R['trend'] = dict(base=trend(MA), imp=trend(I))
def weekly(d):
    d = d.copy(); d['wk'] = ((d.Date - (d.Date.iloc[0] - pd.to_timedelta(d.Date.iloc[0].weekday(), 'D'))).dt.days // 7) + 1
    g = d.groupby('wk').agg(e=('Err', 'sum'), s=('Samp', 'sum'), inv=('Inv', 'sum'), days=('Date', 'count'))
    return [dict(wk=int(k), dpmo=dpmo(r.e, r.s), inv=int(r.inv), days=int(r.days)) for k, r in g.iterrows()]
R['weekly'] = dict(base=weekly(M), imp=weekly(I))

# ---------------- stratification
def strat(d):
    out = {}
    for t in ['Normal', 'Month-start carry-over', 'Month-end peak']:
        x = d[d.dtype == t]
        if len(x): out[t] = dict(days=len(x), inv=int(x.Inv.sum()), err=int(x.Err.sum()), samp=int(x.Samp.sum()),
                                  dpu=x.Err.sum() / x.Samp.sum(), dpmo=dpmo(x.Err.sum(), x.Samp.sum()), sigma=sigma(x.Err.sum(), x.Samp.sum()) if x.Err.sum() else None)
    return out
R['strat'] = dict(base=strat(MA), imp=strat(I))
R['month'] = {m: dict(inv=int(x.Inv.sum()), days=len(x), dpmo=dpmo(x.Err.sum(), x.Samp.sum())) for m, x in pd.concat([M, I]).groupby('mon')}
R['aug_me'] = dpmo(MA[(MA.mon == 'Aug') & (MA.dtype == 'Month-end peak')].Err.sum(), MA[(MA.mon == 'Aug') & (MA.dtype == 'Month-end peak')].Samp.sum())
R['sep_me'] = dpmo(MA[(MA.mon == 'Sep') & (MA.dtype == 'Month-end peak')].Err.sum(), MA[(MA.mon == 'Sep') & (MA.dtype == 'Month-end peak')].Samp.sum())

# ---------------- pareto
ER = pd.read_excel(SRC, sheet_name='Error Reasons', nrows=19)
ER.columns = ['cat', 'reason', 'n', 'src']
ER = ER.sort_values('n', ascending=False, kind='stable').reset_index(drop=True)
ER['share'] = ER.n / ER.n.sum(); ER['cum'] = ER.share.cumsum()
R['pareto'] = ER[['reason', 'n', 'src', 'share', 'cum']].to_dict('records')
R['src'] = ER.groupby('src').n.sum().to_dict()

# ---------------- sampling analysis
pop = M.Inv.sum(); n = M.Samp.sum()
fpc = math.sqrt((pop - n) / (pop - 1))
e = MA.Err.sum(); s = MA.Samp.sum()
se_dpu = math.sqrt(e) / s
R['sampling'] = dict(pop=int(pop), n=int(n), rate=n / pop, fpc=fpc, se_dpu=se_dpu, se_dpu_fpc=se_dpu * fpc,
                     moe_dpu=1.96 * se_dpu * fpc, rel_moe=1.96 * se_dpu * fpc / (e / s),
                     daily_n_min=int(M.Samp.min()), daily_n_max=int(M.Samp.max()),
                     daily_rel_moe_small=1.96 * math.sqrt(H['base_adj']['dpu'] / 10) / H['base_adj']['dpu'],
                     daily_rel_moe_large=1.96 * math.sqrt(H['base_adj']['dpu'] / 88) / H['base_adj']['dpu'],
                     share_samples_peak=float(M[M.peak == 1].Samp.sum() / M.Samp.sum()), share_inv_peak=float(M[M.peak == 1].Inv.sum() / M.Inv.sum()),
                     rate_peak=float(M[M.peak == 1].Samp.sum() / M[M.peak == 1].Inv.sum()), rate_nonpeak=float(M[M.peak == 0].Samp.sum() / M[M.peak == 0].Inv.sum()))
# invoices with >=1 error in sample if errors cluster: bounds
R['sampling']['defective_bounds'] = dict(min_share=None)

# ---------------- CI
def ci(e, s):
    z = 1.96; lo = max(0, e / s - z * math.sqrt(e) / s); hi = e / s + z * math.sqrt(e) / s
    return dict(dpu=e / s, lo=lo, hi=hi, dpmo_lo=lo / OPP * 1e6, dpmo_hi=hi / OPP * 1e6)
R['ci'] = dict(base=ci(e, s), imp=ci(I.Err.sum(), I.Samp.sum()))
rr = (I.Err.sum() / I.Samp.sum()) / (e / s); se = math.sqrt(1 / e + 1 / I.Err.sum())
R['ci']['rr'] = dict(rr=rr, lo=rr * math.exp(-1.96 * se), hi=rr * math.exp(1.96 * se))

# ---------------- correlation
R['corr'] = {}
for k, d in (('base', M), ('imp', I)):
    pr = stats.pearsonr(d.Inv, d.dpu); sp = stats.spearmanr(d.Inv, d.dpu)
    R['corr'][k] = dict(r=pr.statistic, p=pr.pvalue, rho=sp.statistic, p_s=sp.pvalue)

# ---------------- hypothesis tests
def ztest(e1, n1, e2, n2):
    p1, p2 = e1 / n1, e2 / n2; p = (e1 + e2) / (n1 + n2)
    z = (p1 - p2) / math.sqrt(p * (1 - p) * (1 / n1 + 1 / n2)); return dict(z=z, p=2 * (1 - stats.norm.cdf(abs(z))), ratio=p1 / p2 if p2 else None)
o = lambda d: (int(d.Err.sum()), int(d.Samp.sum() * OPP))
R['tests'] = dict(
    before_after=ztest(*o(MA), *o(I)),
    peak_base=ztest(*o(MA[MA.peak == 1]), *o(MA[MA.peak == 0])),
    peak_imp=ztest(*o(I[I.peak == 1]), *o(I[I.peak == 0])),
    sep_vs_aug=ztest(*o(MA[(MA.mon == 'Sep') & (MA.dtype == 'Month-end peak')]), *o(MA[(MA.mon == 'Aug') & (MA.dtype == 'Month-end peak')])),
    mannwhitney_peak=dict(zip(['U', 'p'], stats.mannwhitneyu(MA[MA.peak == 1].dpu, MA[MA.peak == 0].dpu))),
)
# weekday chi-square (errors vs expected by samples)
g = MA.groupby('wd').agg(e=('Err', 'sum'), s=('Samp', 'sum'))
exp_ = g.s / g.s.sum() * g.e.sum()
chi = stats.chisquare(g.e, exp_)
R['tests']['weekday_chi2'] = dict(chi2=chi.statistic, p=chi.pvalue)

# ---------------- regression (Poisson + NB GLM with offset)
A = pd.concat([MA, I], ignore_index=True)
A['post'] = (A.period == 'imp').astype(int)
A['logs'] = np.log(A.Samp)
pois = smf.glm('Err ~ peak + qend + post', data=A, family=sm.families.Poisson(), offset=A.logs).fit()
disp = pois.pearson_chi2 / pois.df_resid
nb = smf.glm('Err ~ peak + qend + post', data=A, family=sm.families.NegativeBinomial(alpha=max(0.05, (disp - 1) / A.Err.mean())), offset=A.logs).fit()
R['regression'] = dict(dispersion=disp, nb_alpha=max(0.05, (disp - 1) / A.Err.mean()),
                       coefs={k: dict(irr=math.exp(v), lo=math.exp(nb.conf_int().loc[k, 0]), hi=math.exp(nb.conf_int().loc[k, 1]), p=nb.pvalues[k])
                              for k, v in nb.params.items()},
                       pois_coefs={k: dict(irr=math.exp(v), p=pois.pvalues[k]) for k, v in pois.params.items()})
lin = smf.ols('dpu ~ Inv + qend', data=MA).fit()
R['regression']['ols_base'] = dict(b_inv100=lin.params['Inv'] * 100, p_inv=lin.pvalues['Inv'], b_qend=lin.params['qend'], p_qend=lin.pvalues['qend'], r2=lin.rsquared)

# ---------------- control charts
def uchart(d, ub=None):
    ub = ub or d.Err.sum() / d.Samp.sum()
    ucl = ub + 3 * np.sqrt(ub / d.Samp); lcl = np.clip(ub - 3 * np.sqrt(ub / d.Samp), 0, None)
    sig = np.where(d.dpu > ucl, 'high', np.where(d.dpu < lcl, 'low', ''))
    return dict(ubar=ub, points=[dict(d=r.Date.strftime('%d-%b'), u=float(r.dpu), ucl=float(ucl[i]), lcl=float(lcl[i]), s=sig[i]) for i, (_, r) in enumerate(d.iterrows())],
                out=int((sig != '').sum()))
R['uchart'] = dict(base=uchart(M), imp=uchart(I))
# run rule: longest run on one side
def longest_run(x, c):
    best = cur = 0; prev = None
    for v in x:
        side = v > c
        cur = cur + 1 if side == prev else 1; prev = side; best = max(best, cur)
    return best
R['uchart']['base_run'] = longest_run(M.dpu, R['uchart']['base']['ubar']); R['uchart']['imp_run'] = longest_run(I.dpu, R['uchart']['imp']['ubar'])
# control limits for ongoing monitoring at typical sample sizes
ubi = R['uchart']['imp']['ubar']
R['uchart']['monitor'] = {str(nn): dict(ucl=ubi + 3 * math.sqrt(ubi / nn), errors_trigger=math.floor((ubi + 3 * math.sqrt(ubi / nn)) * nn) + 1) for nn in (15, 20, 30, 50, 80)}

# ---------------- capability (attribute)
def cap(dpu_):
    dpo = dpu_ / OPP; zlt = stats.norm.ppf(1 - dpo); return dict(z_lt=zlt, z_st=zlt + 1.5, ppk=zlt / 3, pdef=1 - math.exp(-dpu_))
R['capability'] = dict(base=cap(H['base_adj']['dpu']), imp=cap(H['imp']['dpu']))
# against a CTQ spec: <=2% invoices with any error
spec = 0.02
R['capability']['spec_pdef'] = spec
R['capability']['dpu_needed'] = -math.log(1 - spec)

# ---------------- time series
def acf1(x): x = np.asarray(x) - np.mean(x); return float(np.sum(x[1:] * x[:-1]) / np.sum(x * x))
R['ts'] = dict(acf1_inv_base=acf1(M.Inv), acf1_dpu_base=acf1(MA.dpu), acf1_dpu_imp=acf1(I.dpu),
               peak_index=float(M[M.peak == 1].Inv.mean() / M[M.peak == 0].Inv.mean()),
               peak_index_imp=float(I[I.peak == 1].Inv.mean() / I[I.peak == 0].Inv.mean()))
# working days before month end profile
def wd_to_me(d):
    d = d.copy(); out = []
    for m, x in d.groupby('mon', sort=False):
        x = x.sort_values('Date'); k = len(x)
        for j, (_, r) in enumerate(x.iterrows()): out.append((k - j, r.Inv, r.dpu, r.Date.strftime('%d-%b')))
    return out
prof = pd.DataFrame(wd_to_me(M) + wd_to_me(I), columns=['wdme', 'inv', 'dpu', 'd'])
R['ts']['profile'] = prof.groupby('wdme').inv.mean().loc[1:10].round(0).to_dict()

# ---------------- DOE power: audit days needed per run to detect a 50% cut at base rate
def days_needed(base_dpu, cut, samp_per_day=32, alpha=0.05, power=0.8):
    za, zb = stats.norm.ppf(1 - alpha / 2), stats.norm.ppf(power)
    l1, l2 = base_dpu, base_dpu * (1 - cut)
    n_inv = ((za * math.sqrt(2 * (l1 + l2) / 2) + zb * math.sqrt(l1 + l2)) / (l1 - l2)) ** 2
    return n_inv, math.ceil(n_inv / samp_per_day)
R['doe'] = dict(inv_needed_50=days_needed(H['base_adj']['dpu'], .5)[0], days_needed_50=days_needed(H['base_adj']['dpu'], .5)[1],
                inv_needed_20=days_needed(H['base_adj']['dpu'], .2)[0], days_needed_20=days_needed(H['base_adj']['dpu'], .2)[1],
                inv_needed_imp50=days_needed(H['imp']['dpu'], .5)[0], days_needed_imp50=days_needed(H['imp']['dpu'], .5)[1])

# ---------------- process / VSM / ABC
steps = pd.read_excel(SRC, sheet_name='Process Steps', nrows=12)
steps.columns = ['n', 'desc', 'el', 'tt', 'pct', 'scope']
steps['w_el'] = steps.el * steps.pct; steps['w_tt'] = steps.tt * steps.pct; steps['w_wait'] = (steps.el - steps.tt) * steps.pct
ins = steps[steps.scope == 'Yes']
R['vsm'] = dict(steps=steps[['n', 'desc', 'el', 'tt', 'pct', 'w_el', 'w_tt', 'w_wait', 'scope']].to_dict('records'),
                el=float(ins.w_el.sum()), tt=float(ins.w_tt.sum()), wait=float(ins.w_wait.sum()), pce=float(ins.w_tt.sum() / ins.w_el.sum()))
RATE = 350
steps['cost'] = steps.w_tt / 60 * RATE
R['abc'] = dict(rows=[dict(n=int(r.n), desc=r.desc, cost=float(r.cost), scope=r.scope) for _, r in steps.iterrows()],
                per_invoice=float(steps[steps.scope == 'Yes'].cost.sum()),
                exception_per_exc_invoice=12 / 60 * RATE)
R['abc']['monthly_inscope'] = R['abc']['per_invoice'] * M.Inv.mean() * 26

# ---------------- capacity / queueing / Little
TT = R['vsm']['tt']; mu = 60 / TT; c = 18
def erlang_c(lam, mu, c):
    a = lam / mu; rho = a / c
    if rho >= 1: return dict(rho=rho, pw=1.0, wq_min=None)
    s = sum(a ** k / math.factorial(k) for k in range(c)); last = a ** c / math.factorial(c) / (1 - rho)
    pw = last / (s + last); wq = pw / (c * mu - lam) * 60
    return dict(rho=rho, pw=pw, wq_min=wq)
R['queue'] = {}
for label, inv, hrs in [('Median normal day (186)', 186, 8), ('Average day (320)', 320, 8), ('Busy normal day (376)', 376, 8),
                        ('Peak day 600 (with OT)', 600, 11), ('Peak day 880 (with OT)', 880, 11)]:
    lam = inv / hrs; q = erlang_c(lam, mu, c); q.update(lam=lam, inv=inv); R['queue'][label] = q
R['little'] = dict(lam_avg_hr=M.Inv.mean() / 8, W_hr=R['vsm']['el'] / 60, L=M.Inv.mean() / 8 * R['vsm']['el'] / 60,
                   lam_peak_hr=880 / 11, L_peak=880 / 11 * R['vsm']['el'] / 60)
cap_n = c * 8 * 60 / TT; cap_p = c * 11 * 60 / TT
R['capacity'] = dict(cap_normal=cap_n, cap_peak=cap_p, over_days=int(((M.Inv * TT / 60) > c * np.where(M.peak == 1, 11, 8)).sum()),
                     max_util=float((M.Inv * TT / 60 / (c * np.where(M.peak == 1, 11, 8))).max()))

# ---------------- discrete-event simulation (baseline volumes)
def des(vols, peaks, seed, extra=None, c=18):
    rnd = random.Random(seed); backlog = []  # (arrival_day)
    waits = []; done_late = 0; total = 0; end_backlog = []
    queue = []
    for day, (v, pk) in enumerate(zip(vols, peaks)):
        for _ in range(int(v)): queue.append(day)
        hours = 11 if pk else 8
        staff = c + (extra[day] if extra else 0)
        cap_min = staff * hours * 60
        used = 0.0
        while queue:
            st = rnd.choice([20.0] * 3 + [32.0]) * rnd.uniform(0.8, 1.2)
            if used + st > cap_min: break
            used += st; a = queue.pop(0); w = day - a; waits.append(w); total += 1
            if w >= 2: done_late += 1
        end_backlog.append(len(queue))
    return dict(mean_wait_days=float(np.mean(waits)), p95_wait=float(np.percentile(waits, 95)), share_2plus=done_late / total,
                max_backlog=int(max(end_backlog)), backlog=end_backlog, end=int(end_backlog[-1]))
vols = list(M.Inv); pks = list(M.peak)
sims = [des(vols, pks, s) for s in range(30)]
R['des'] = dict(base=dict(mean_wait=np.mean([x['mean_wait_days'] for x in sims]), share_2plus=np.mean([x['share_2plus'] for x in sims]),
                          max_backlog=np.mean([x['max_backlog'] for x in sims]), backlog=sims[0]['backlog']))
# scenario: +6 flex staff on peak days
extra = [6 if p else 0 for p in pks]
sims2 = [des(vols, pks, s, extra) for s in range(30)]
R['des']['flex6'] = dict(mean_wait=np.mean([x['mean_wait_days'] for x in sims2]), share_2plus=np.mean([x['share_2plus'] for x in sims2]),
                         max_backlog=np.mean([x['max_backlog'] for x in sims2]), backlog=sims2[0]['backlog'])
# scenario: exceptions cut (touch 20.12) via service distribution change
def des_low(vols, pks, seed):
    rnd = random.Random(seed); queue = []; waits = []; late = 0; tot = 0; eb = []
    for day, (v, pk) in enumerate(zip(vols, pks)):
        queue += [day] * int(v); cap_min = 18 * (11 if pk else 8) * 60; used = 0
        while queue:
            st = (20.0 if rnd.random() > 0.01 else 32.0) * rnd.uniform(0.8, 1.2)
            if used + st > cap_min: break
            used += st; a = queue.pop(0); w = day - a; waits.append(w); tot += 1; late += w >= 2
        eb.append(len(queue))
    return dict(mean_wait_days=float(np.mean(waits)), share_2plus=late / tot, max_backlog=max(eb), backlog=eb)
sims3 = [des_low(vols, pks, s) for s in range(30)]
R['des']['fewer_exc'] = dict(mean_wait=np.mean([x['mean_wait_days'] for x in sims3]), share_2plus=np.mean([x['share_2plus'] for x in sims3]),
                             max_backlog=np.mean([x['max_backlog'] for x in sims3]))

# ---------------- LP: cheapest mix of overtime + flex hours so backlog never exceeds half a normal day
T = len(M); base_h = 18 * 8
c_ot = RATE * 1.5; c_flex = 600  # ASSUMPTION ₹600/h agency/temp analyst
def solve_lp(inv):
    req = np.asarray(inv) * TT / 60; nv = 3 * T
    cvec = np.zeros(nv); cvec[0:T] = c_ot; cvec[T:2 * T] = c_flex
    A = np.zeros((T, nv)); b = np.zeros(T)
    for t in range(T):
        A[t, 2 * T + t] = 1
        if t > 0: A[t, 2 * T + t - 1] = -1
        A[t, t] = 1; A[t, T + t] = 1; b[t] = req[t] - base_h
    bounds = [(0, 18 * 3)] * T + [(0, 8 * 20)] * T + [(0, base_h * 0.5)] * T
    lp = optimize.linprog(cvec, A_ub=-A, b_ub=-b, bounds=bounds, method='highs')
    ot = lp.x[:T]; fl = lp.x[T:2 * T]
    return dict(status=lp.message, cost_month=lp.fun / 2, ot_hours_month=float(ot.sum()) / 2, flex_hours_month=float(fl.sum()) / 2,
                max_flex_people=float(fl.max() / 8), days_with_flex=int((fl > 0.5).sum()), flex=fl.tolist(), ot=ot.tolist())
# level-loading: ASSUMPTION 25% of volume above 400 on month-end peak days is pulled forward, spread over that month's normal days
def level(d, share=0.25):
    inv = d.Inv.astype(float).values.copy()
    for m in d.mon.unique():
        mi = (d.mon == m).values; pk = mi & (d.dtype == 'Month-end peak').values; nm = mi & (d.dtype == 'Normal').values
        moved = (np.clip(inv[pk] - 400, 0, None) * share)
        inv[pk] -= moved; inv[nm] += moved.sum() / nm.sum()
    return inv
R['lp'] = dict(current_ot_cost_month=18 * 3 * 6 * c_ot, base=solve_lp(M.Inv.values), level=solve_lp(level(M)))
R['lp']['level_saving_month'] = R['lp']['base']['cost_month'] - R['lp']['level']['cost_month']
sims4 = [des(list(level(M)), pks, s) for s in range(30)]
R['des']['level'] = dict(mean_wait=np.mean([x['mean_wait_days'] for x in sims4]), share_2plus=np.mean([x['share_2plus'] for x in sims4]), max_backlog=np.mean([x['max_backlog'] for x in sims4]))
# staffing rule at 85% utilisation
R['staffing'] = {t: dict(avg_inv=float(x.Inv.mean()), max_inv=float(x.Inv.max()),
                         fte_avg=float(x.Inv.mean() * TT / 60 / 8 / 0.85), fte_max=float(x.Inv.max() * TT / 60 / 8 / 0.85))
                 for t, x in M.groupby('dtype')}

# ---------------- COPQ + Monte Carlo + sensitivity
vol = M.Inv.mean() * 26
def copq(defshare=0.35, pen=0.05, amt=5000, exc=0.25, excmin=12, rate=350):
    return vol * 12 * (defshare * pen * amt + exc * excmin / 60 * rate)
base_copq = copq()
rng = np.random.default_rng(7); N = 20000
ds = rng.uniform(0.35, 0.46, N); pr = rng.triangular(0.02, 0.05, 0.08, N); am = rng.uniform(2500, 5000, N)
ex = rng.uniform(0.25, 0.46, N); em = rng.triangular(8, 12, 18, N); rt = rng.triangular(300, 350, 450, N)
mc = vol * 12 * (ds * pr * am + ex * em / 60 * rt)
R['copq'] = dict(vol_month=vol, base=base_copq, data=copq(defshare=H['pdef_base'], exc=H['pdef_base']),
                 after=copq(defshare=0.35 * H['imp']['dpu'] / H['base_adj']['dpu'], exc=0.25 * H['imp']['dpu'] / H['base_adj']['dpu']))
R['mc'] = dict(p5=float(np.percentile(mc, 5)), p50=float(np.percentile(mc, 50)), p95=float(np.percentile(mc, 95)), mean=float(mc.mean()),
               hist=np.histogram(mc / 1e5, bins=24)[0].tolist(), edges=np.histogram(mc / 1e5, bins=24)[1].round(1).tolist())
tor = []
for name, lo, hi, key in [('Penalty trigger rate (2%–8%)', 0.02, 0.08, 'pen'), ('Penalty per invoice (₹2,500–₹5,000)', 2500, 5000, 'amt'),
                          ('Share of invoices with errors (35%–46%)', 0.35, 0.46, 'defshare'), ('Exception share (20%–46%)', 0.20, 0.46, 'exc'),
                          ('Exception minutes (8–18)', 8, 18, 'excmin'), ('Analyst cost/hour (₹300–₹450)', 300, 450, 'rate')]:
    tor.append(dict(name=name, lo=copq(**{key: lo}), hi=copq(**{key: hi})))
R['tornado'] = sorted(tor, key=lambda x: -(abs(x['hi'] - x['lo'])))

# ---------------- solution options: cost-benefit / ROI / NPV / payback (ILLUSTRATIVE COST ASSUMPTIONS)
# expected reduction of baseline COPQ via preventability mapping (Yes 80%, Partly 40%, No 0%)
prev = {'Mismatched invoice and PO amounts': ('Y', 'S1'), 'Unclear or missing line item descriptions': ('P', 'S1'),
        'Missing supporting documents (e.g., delivery note)': ('Y', 'S1'), 'System integration failure (e.g., ERP sync issues)': ('N', 'S4'),
        'Invoice submitted in wrong format': ('Y', 'S2'), 'Duplicate invoice submission': ('Y', 'S1'), 'Missing invoice date': ('Y', 'S1'),
        'Incomplete invoice fields (e.g., missing bank details)': ('Y', 'S1'), 'Incorrect vendor details': ('Y', 'S1'),
        'Missing Purchase Order (PO) reference': ('Y', 'S1'), 'Currency mismatch': ('Y', 'S1'), 'Unregistered vendor': ('Y', 'S1'),
        'Inconsistent invoice numbering': ('P', 'S2'), 'Wrong approval routing': ('N', 'S4'), 'Incorrect tax codes': ('Y', 'S2'),
        'Late invoice submission': ('P', 'S3'), 'Delayed approvals': ('N', 'S4'), 'Incorrect GST calculation': ('Y', 'S2'), 'Manual data entry errors': ('N', 'S2')}
eff = {'Y': 0.8, 'P': 0.4, 'N': 0.0}
share = dict(zip(ER.reason, ER.share))
R['preventable'] = dict(yes=sum(share[r] for r, (p, _) in prev.items() if p == 'Y'), partly=sum(share[r] for r, (p, _) in prev.items() if p == 'P'),
                        no=sum(share[r] for r, (p, _) in prev.items() if p == 'N'),
                        expected_cut=sum(share[r] * eff[p] for r, (p, _) in prev.items()))
opts = [  # name, capex (₹L), opex/yr (₹L), error-share addressed (from mapping/assumption), effectiveness, extra benefit/yr (₹L)
    dict(id='S1', name='Smart vendor portal: mandatory fields, PO/GRN auto-match, document upload, duplicate & vendor checks', capex=25, opex=5),
    dict(id='S2', name='E-invoice (IRN) data intake + OCR/IDP capture for non-e-invoice vendors', capex=30, opex=10),
    dict(id='S3', name='Vendor onboarding, invoice templates, training & monthly scorecards', capex=4, opex=3),
    dict(id='S4', name='Approval workflow rules + SLAs, ERP sync monitoring', capex=12, opex=3),
    dict(id='S5', name='Peak-period flex staffing & submission cut-off calendar', capex=1, opex=None),
]
cut_by = {k: 0 for k in ['S1', 'S2', 'S3', 'S4']}
for r, (p, sid) in prev.items(): cut_by[sid] += share[r] * eff[p]
cut_by['S3'] += 0.10  # ASSUMPTION: training/scorecards lift compliance across vendor errors by 10% of total errors
cut_by['S4'] += 0.5 * (share['Wrong approval routing'] + share['Delayed approvals'] + share['System integration failure (e.g., ERP sync issues)'])
penalty_part = vol * 12 * 0.35 * 0.05 * 5000; rework_part = vol * 12 * 0.25 * 12 / 60 * 350
ot_cost_year = 18 * 3 * 6 * c_ot * 12
lp_year = R['lp']['base']['cost_month'] * 12
r_disc = 0.12
for o_ in opts:
    if o_['id'] == 'S5':
        benefit = R['lp']['level_saving_month'] * 12 / 1e5; o_['opex'] = 0; o_['cut'] = 0
    else:
        o_['cut'] = min(cut_by[o_['id']], 0.95)
        benefit = (penalty_part + rework_part) * o_['cut'] / 1e5
    o_['benefit'] = benefit; net = benefit - o_['opex']
    o_['net'] = net
    o_['payback_months'] = (o_['capex'] / net * 12) if net > 0 else None
    o_['roi_3y'] = (net * 3 - o_['capex']) / o_['capex'] if o_['capex'] else None
    o_['npv_3y'] = -o_['capex'] + sum(net / (1 + r_disc) ** y for y in (1, 2, 3))
    o_['breakeven_capex_1y'] = net
# package S1+S3+S4+S5 (+S2 optional) – overlap handled multiplicatively
def package(ids):
    rem = 1.0
    for i in ids:
        if i != 'S5': rem *= (1 - next(x for x in opts if x['id'] == i)['cut'])
    cut = 1 - rem
    ben = (penalty_part + rework_part) * cut / 1e5 + (R['lp']['level_saving_month'] * 12 / 1e5 if 'S5' in ids else 0)
    capex = sum(x['capex'] for x in opts if x['id'] in ids); opex = sum(x['opex'] for x in opts if x['id'] in ids)
    net = ben - opex
    return dict(ids=ids, cut=cut, benefit=ben, capex=capex, opex=opex, net=net, payback_months=capex / net * 12 if net > 0 else None,
                npv_3y=-capex + sum(net / (1 + r_disc) ** y for y in (1, 2, 3)), roi_3y=(net * 3 - capex) / capex)
R['options'] = opts
R['package'] = dict(core=package(['S1', 'S3', 'S5']), plus_s4=package(['S1', 'S3', 'S4', 'S5']), full=package(['S1', 'S2', 'S3', 'S4', 'S5']))
# scenarios for core package
def scen(mult_cut, mult_cost, pen):
    p = package(['S1', 'S3', 'S5'])
    ben = ((vol * 12 * 0.35 * pen * 5000 + rework_part) * min(p['cut'] * mult_cut, 0.97) / 1e5) + R['lp']['level_saving_month'] * 12 / 1e5 * mult_cut
    capex = p['capex'] * mult_cost; opex = p['opex'] * mult_cost; net = ben - opex
    return dict(benefit=ben, capex=capex, net=net, payback_months=capex / net * 12 if net > 0 else None, npv_3y=-capex + sum(net / 1.12 ** y for y in (1, 2, 3)))
R['scenarios'] = dict(conservative=scen(0.6, 1.5, 0.02), central=scen(1.0, 1.0, 0.05), optimistic=scen(1.15, 0.8, 0.08))
# expected value of penalty exposure per erroneous invoice
R['ev'] = dict(per_err_invoice=0.05 * 5000, per_err_invoice_range=(0.02 * 2500, 0.08 * 5000),
               annual_exposure=vol * 12 * 0.35 * 0.05 * 5000, annual_exposure_data=vol * 12 * H['pdef_base'] * 0.05 * 5000)
# false accept / false reject threshold
fa = 0.05 * 5000 + 0; fr = 12 / 60 * 350
R['fafr'] = dict(cost_fa=fa, cost_fr=fr, threshold=fr / (fa + fr))
# cost of delay (MSME interest; ASSUMPTIONS: bank rate 6.5%? share MSME 30%, avg invoice ₹5 lakh)
R['cod'] = dict(note='illustrative')
json.dump(R, open(paths.out('results.json'), 'w'), indent=1, default=lambda x: float(x) if isinstance(x, (np.floating, np.integer)) else str(x))
print('ok')
R['cap_sens'] = [(t, int(((M.Inv * t / 60) > 18 * np.where(M.peak == 1, 11, 8)).sum())) for t in (18, 20, 23, 26)]
json.dump(R, open(paths.out('results.json'), 'w'), indent=1, default=lambda x: float(x) if isinstance(x, (np.floating, np.integer)) else str(x))
print(R['cap_sens'])
