"""Workbook extension (exec-ed by build_workbook.py) - "Uncertainty & Decision" sheet.
Reads : uncert.json, rule_final.json, rule.json, results.json, laney.json, data/original.xlsx (sheet Improve)
Writes: outputs/uref.json (cell map used by the reconciliation sheet); adds a sheet to the workbook object `wb`
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Not run on its own. Lays out confidence intervals, bootstrap and Bayesian results, the Monte Carlo for Hybrid H and the
#   escalation decision rule. Live formulas where Excel can do the maths; simulation results are pasted as numbers.
# ---------------------------------------------------------------------------
# ============================================================ UNCERTAINTY & DECISION  (limit_2: rigorous intervals + the G1 decision rule)
# exec'd from build_model2.py after ext_hybrid.py; uses its helpers (put, inp, link, calc, hdr, title, section, widths, name) and ST.
import json as _j
_U = _j.load(open(paths.find('uncert.json')))
_RF = _j.load(open(paths.find('rule_final.json')))
_RL = _j.load(open(paths.find('rule.json')))
_R = _j.load(open(paths.find('results.json')))
_LAN = _j.load(open(paths.find('laney.json')))
_IMP = pd.read_excel(SRC, sheet_name='Improve', header=1, nrows=51); _IMP.columns = ['Date', 'Inv', 'Samp', 'Err']
UN = 'Uncertainty & Decision'
UD = wb.create_sheet(UN)
UREF = {}
def ur(key, cell): UREF[key] = f"'{UN}'!{cell}"; return cell
title(UD, 'Uncertainty & Decision · how sure we are, and the end-October decision rule',
      'Live formulas on the case data wherever Excel can do the maths; bootstrap, regression and Monte Carlo results come from the Python model and are pasted (method, B and seed stated). Every limit_2 deck number maps to a cell on the Deck Reconciliation sheet.')
widths(UD, {'A': 58, 'B': 15, 'C': 15, 'D': 15, 'E': 15, 'F': 15, 'G': 15, 'H': 15, 'I': 15, 'J': 4, 'K': 12, 'L': 10, 'M': 10, 'N': 10})
PASTE = PatternFill('solid', fgColor='F2F2F2')
def pv(ref, v, fmt):   # pasted value: grey fill, black text
    return put(UD, ref, v, fmt, al='center', fill=PASTE)
def note(ref, t): return put(UD, ref, t, sz=8, i=True, border=False)

# ---- improve-audit day data (case Improve tab) so the audit dispersion is live
hdr(UD, 4, 11, ['Audit day', 'Invoices', 'Audited', 'Errors'])
put(UD, 'K3', 'Supplied audit, Mar–Apr 2027 (case Improve tab)', b=True, sz=9, border=False)
for i, rr in _IMP.iterrows():
    row = 5 + i
    put(UD, f'K{row}', rr.Date.to_pydatetime(), 'dd-mmm-yy', al='center', sz=8)
    for col, v in (('L', rr.Inv), ('M', rr.Samp), ('N', rr.Err)): put(UD, f'{col}{row}', int(v), '0', al='center', sz=8)
IR1 = 5 + len(_IMP) - 1
IMS, IME = f'$M$5:$M${IR1}', f'$N$5:$N${IR1}'

# ---- 1. DPMO intervals
section(UD, 'A4', '1. Before and after: DPMO with four kinds of 95% interval')
hdr(UD, 5, 1, ['Measure', 'Baseline Aug–Sep 2026 (adjusted)', 'Supplied audit Mar–Apr 2027', 'Method'])
BSI = "'Baseline'!$I$5:$I$56"; BSF = "'Baseline'!$F$5:$F$56"; BSG = "'Baseline'!$G$5:$G$56"
rows1 = [
 ('Errors found, x', "='Baseline'!B8", f'=SUM({IME})', '0', 'x', 'Baseline excludes 25-Aug (recording anomaly)'),
 ('Invoices audited, n', "='Baseline'!B9", f'=SUM({IMS})', '#,##0', 'n', ''),
 ('DPMO = x ÷ (58 n) × 10⁶', '=B6/B7/OPP*1000000', '=C6/C7/OPP*1000000', '#,##0', 'dpmo', '58 fields per invoice'),
 ('Sigma level = NORMSINV(1 − DPMO/10⁶) + 1.5', '=NORMSINV(1-B8/1000000)+1.5', '=NORMSINV(1-C8/1000000)+1.5', '0.00', 'sig', 'Short-term sigma with the 1.5 shift'),
 ('Exact Poisson 95% CI, low = χ²(0.025; 2x) ÷ 2', '=CHIINV(0.975,2*B6)/2/B7/OPP*1000000', '=CHIINV(0.975,2*C6)/2/C7/OPP*1000000', '#,##0', 'ex_lo', 'Garwood interval on the error count'),
 ('Exact Poisson 95% CI, high = χ²(0.975; 2x+2) ÷ 2', '=CHIINV(0.025,2*(B6+1))/2/B7/OPP*1000000', '=CHIINV(0.025,2*(C6+1))/2/C7/OPP*1000000', '#,##0', 'ex_hi', ''),
 ('Pearson dispersion φ̂ = Σ (e − n·u)² ÷ (n·u) ÷ (days − 1)', f'=SUMPRODUCT({BSI},({BSG}-{BSF}*DPU0)^2/({BSF}*DPU0))/(SUM({BSI})-1)',
  f'=SUMPRODUCT(({IME}-{IMS}*C6/C7)^2/({IMS}*C6/C7))/(COUNT({IME})-1)', '0.00', 'phi', '1 = pure Poisson; baseline errors cluster by day'),
 ('Quasi-Poisson 95% CI, low = x − 1.96 √(φ̂ x)', '=(B6-1.96*SQRT(B12*B6))/B7/OPP*1000000', '=(C6-1.96*SQRT(C12*C6))/C7/OPP*1000000', '#,##0', 'qp_lo', 'Widens the Poisson interval by √φ̂'),
 ('Quasi-Poisson 95% CI, high = x + 1.96 √(φ̂ x)', '=(B6+1.96*SQRT(B12*B6))/B7/OPP*1000000', '=(C6+1.96*SQRT(C12*C6))/C7/OPP*1000000', '#,##0', 'qp_hi', '')]
for k, (lab, fb, fc, nf, key, mt) in enumerate(rows1):
    rr = 6 + k
    put(UD, f'A{rr}', lab, sz=9); calc(UD, f'B{rr}', fb, nf); calc(UD, f'C{rr}', fc, nf); put(UD, f'D{rr}', mt, sz=8)
    ur(f'{key}_b', f'B{rr}'); ur(f'{key}_i', f'C{rr}')
put(UD, 'A15', 'Day-level bootstrap 95% CI, low (pasted)', sz=9); pv('B15', _U['boot']['base'][0], '#,##0'); pv('C15', _U['boot']['imp'][0], '#,##0')
put(UD, 'A16', 'Day-level bootstrap 95% CI, high (pasted)', sz=9); pv('B16', _U['boot']['base'][1], '#,##0'); pv('C16', _U['boot']['imp'][1], '#,##0')
put(UD, 'D15', f'Resample the audit days with replacement, B = {_U["boot"]["B"]:,}, seed 2026; DPMO of each resample', sz=8)
ur('bt_lo_b', 'B15'); ur('bt_hi_b', 'B16'); ur('bt_lo_i', 'C15'); ur('bt_hi_i', 'C16')
put(UD, 'A17', 'Bootstrap range below the estimate (1 − low ÷ DPMO)', sz=9); calc(UD, 'B17', '=1-B15/B8', '0%'); ur('bt_dn', 'B17')
put(UD, 'A18', 'Bootstrap range above the estimate (high ÷ DPMO − 1)', sz=9); calc(UD, 'B18', '=B16/B8-1', '0%'); ur('bt_up', 'B18')
put(UD, 'A19', 'Textbook ± on DPU (simple random sample of 1,669 invoices, pasted)', sz=9); pv('B19', _R['sampling']['rel_moe'], '0%'); ur('moe', 'B19')
put(UD, 'D19', 'Too narrow: it ignores day-to-day clustering (φ̂ above)', sz=8)

# ---- 2. Size of the reduction
section(UD, 'A21', '2. Size of the error reduction (supplied audit vs adjusted baseline)')
hdr(UD, 22, 1, ['Method', 'Point', '95% CI low', '95% CI high', 'Note'])
rr_ = _R['ci']['rr']
put(UD, 'A23', 'Ratio of rates, exact conditional Poisson CI', sz=9); calc(UD, 'B23', '=1-C8/B8', '0.0%'); pv('C23', 1 - rr_['hi'], '0.0%'); pv('D23', 1 - rr_['lo'], '0.0%'); put(UD, 'E23', 'Assumes pure Poisson counts', sz=8)
put(UD, 'A24', 'Day-level bootstrap (pasted)', sz=9); calc(UD, 'B24', '=1-C8/B8', '0.0%'); pv('C24', _U['boot']['red'][0], '0.0%'); pv('D24', _U['boot']['red'][1], '0.0%'); put(UD, 'E24', 'Same resamples as row 15', sz=8)
put(UD, 'A25', 'Negative-binomial regression (pasted) — used in the deck', sz=9, b=True); pv('B25', _U['nb']['red_point'], '0.0%'); pv('C25', _U['nb']['red'][0], '0.0%'); pv('D25', _U['nb']['red'][1], '0.0%')
put(UD, 'E25', 'ln E[e_d] = ln n_d + β0 + β1·peak + β2·quarter-end + β3·post; reduction = 1 − e^β3', sz=8)
put(UD, 'A26', 'NB rate ratio for the audit period, e^β3 (pasted)', sz=9); pv('B26', _U['nb']['irr'], '0.000'); pv('C26', _U['nb']['lo'], '0.000'); pv('D26', _U['nb']['hi'], '0.000'); put(UD, 'E26', f'Dispersion α = {_R["regression"]["nb_alpha"]:.3f}', sz=8)
ur('red_simple', 'B23'); ur('rr_lo', 'C23'); ur('nb_pt', 'B25'); ur('nb_lo', 'C25'); ur('nb_hi', 'D25'); ur('nb_irr', 'B26')

# ---- 3. Stretch target
section(UD, 'A28', '3. Is the ≤ 500 DPMO stretch met in the supplied audit?')
hdr(UD, 29, 1, ['Measure', 'Value', 'Method'])
r3 = [('Expected errors at exactly 500 DPMO, λ* = 500 × 58 × n ÷ 10⁶', '=500*OPP*C7/1000000', '0.0', 'lam', 'Audit n = 1,505'),
      ('Posterior shape with a flat prior, x + 1', '=C6+1', '0', 'shape', 'Gamma(x + 1, 1) for the expected count'),
      ('P(true DPMO > 500), flat prior = 1 − F_Γ(x+1, 1)(λ*)', '=1-GAMMADIST(B30,B31,1,TRUE)', '0.0%', 'p500', 'Bayesian; shown on slide 14'),
      ('P(true DPMO > 500), Jeffreys prior (x + 0.5)', '=1-GAMMADIST(B30,C6+0.5,1,TRUE)', '0.0%', 'p500j', 'Sensitivity to the prior'),
      ('Share of day-bootstrap resamples above 500 (pasted)', _U['boot']['p_imp_gt500'], '0.0%', 'p500b', 'Frequentist check'),
      ('Note: row 34 is a share of bootstrap estimates, not P(true > 500); rows 32–33 are Bayesian', '', '@', 'p500r', 'Different measures: quoted separately in the deck')]
for k, (lab, fm, nf, key, mt) in enumerate(r3):
    rr = 30 + k; put(UD, f'A{rr}', lab, sz=9)
    if isinstance(fm, float): pv(f'B{rr}', fm, nf)
    else: calc(UD, f'B{rr}', fm, nf)
    put(UD, f'C{rr}', mt, sz=8); ur(key, f'B{rr}')

# ---- 4. Monte Carlo intervals for Hybrid H
section(UD, 'A37', '4. Monte Carlo intervals for Hybrid H (20,000 runs, seed 5)')
hdr(UD, 38, 1, ['Output', 'P5', 'Median', 'P95', 'Note'])
for k, (lab, key, nf) in enumerate([('Error cut', 'cut', '0.0%'), ('DPMO', 'dpmo', '#,##0'), ('NPV incl. build time (₹ L)', 'npv', '0.0')]):
    rr = 39 + k; put(UD, f'A{rr}', lab, sz=9)
    for j, c in enumerate('BCD'): pv(f'{c}{rr}', _U['mcH'][key][j], nf); ur(f'mc_{key}_{j}', f'{c}{rr}')
put(UD, 'E39', 'Base case 90.4% vs median 76.7%: the base case assumes full effectiveness', sz=8)
hdr(UD, 43, 1, ['Probability', 'Estimate', 'Std. error = √(p(1−p)/N)', '95% CI low', '95% CI high'])
put(UD, 'A44', 'Runs, N', sz=9); inp(UD, 'B44', 20000, '#,##0'); ur('N', 'B44')
for k, (lab, f_, key) in enumerate([('P(error cut ≥ 75%), Hybrid H', "='Stress Test'!F40", 'p75'), ('P(NPV < 0), Hybrid H', "='Stress Test'!F36", 'pneg'),
                                    ('P(NPV < 0), A + B + C', "='Stress Test'!E36", 'pnegabc'), ('P(error cut ≥ 75%), A + B + C', "='Stress Test'!E40", 'p75abc')]):
    rr = 45 + k; put(UD, f'A{rr}', lab, sz=9); link(UD, f'B{rr}', f_, '0.0%')
    calc(UD, f'C{rr}', f'=SQRT(B{rr}*(1-B{rr})/$B$44)', '0.00%'); calc(UD, f'D{rr}', f'=MAX(0,B{rr}-1.96*C{rr})', '0.0%'); calc(UD, f'E{rr}', f'=B{rr}+1.96*C{rr}', '0.0%')
    ur(key, f'B{rr}'); ur(key + '_se', f'C{rr}')
put(UD, 'A49', 'Largest possible std. error at N = 20,000 (p = 0.5), in points', sz=9); calc(UD, 'B49', '=SQRT(0.25/B44)*100', '0.00'); ur('se_max', 'B49')

# ---- 5. Decision rule at G1
section(UD, 'A51', '5. Decision rule at gate G1 (end-October audit, first full month with all of B live)')
hdr(UD, 52, 1, ['Item', 'Value', 'Note'])
r5 = [('Baseline DPU (adjusted), the denominator of c₁', '=DPU0', '0.000', 'dpu0', 'c₁ = 1 − DPU_Oct ÷ DPU0', 'c'),
      ('Monthly invoice volume', '=VOL', '#,##0', 'vol', '', 'c'),
      ('Audit share, normal month', 0.10, '0%', 'sh10', 'Case: 10% daily sample', 'i'),
      ('Audit share, decision month (October)', 0.20, '0%', 'sh20', 'Doubled for one month', 'i'),
      ('Invoices audited, normal month', '=ROUND(B54*B55,0)', '#,##0', 'n10', '', 'c'),
      ('Invoices audited, decision month', '=ROUND(B54*B56,0)', '#,##0', 'n20', 'Slide 13: ≈1,666', 'c'),
      ('Stay-with-H threshold: c₁ ≥', 0.40, '0%', 'thr_stay', 'Rounded from the model value below', 'i'),
      ('Escalate-to-H+ floor: c₁ ≥', 0.34, '0%', 'thr_esc', 'Below it: re-diagnose', 'i'),
      ('Effectiveness m at which H just reaches 75% (base case)', "='Stress Test'!F31", '0.000', 'mH', 'Linked', 'l'),
      ('Effectiveness m at which H+ just reaches 75% (base case)', "='Stress Test'!G31", '0.000', 'mHP', 'Linked', 'l'),
      ('October cut c₁ when m = 0.75 (pasted, ramp model)', _RL['tab'][7]['oct'], '0.0%', 'oct75', 'Stay threshold: m ≥ 0.75 keeps H well above 75% (89% of runs)', 'p'),
      ('October cut c₁ when m = m(H+) (pasted, ramp model)', _RL['thr']['HP']['oct'], '0.0%', 'octHP', 'Escalation floor: below it even H+ misses 75% in the base case', 'p'),
      ('True cut used for the precision check', 0.37, '0%', 'cprec', 'Middle of the escalate band', 'i'),
      ('Std. error of c₁, normal month = √(n·DPU0·(1−c)) ÷ (n·DPU0)', '=SQRT(B57*B53*(1-B65))/(B57*B53)', '0.0%', 'se10', 'Poisson noise only', 'c'),
      ('Std. error of c₁, decision month', '=SQRT(B58*B53*(1-B65))/(B58*B53)', '0.0%', 'se20', '', 'c'),
      ('90% band half-width, normal month = 1.645 × SE (points)', '=1.645*B66*100', '0.0', 'hw10', 'Too wide to separate 34% from 40%', 'c'),
      ('90% band half-width, decision month (points)', '=1.645*B67*100', '0.0', 'hw20', 'Slide 13: "known to ±4 pts"', 'c'),
      ('Exact Poisson 90% band for c₁, decision month: low (pasted)', _RF['precision']['1666'][0], '0.0%', 'pb20lo', 'Poisson quantiles, true cut 37%', 'p'),
      ('Exact Poisson 90% band for c₁, decision month: high (pasted)', _RF['precision']['1666'][1], '0.0%', 'pb20hi', '', 'p'),
      ('H+ add-on: extra capex (₹ L)', "='Hybrid Plan'!D25", '0.0', 'hp_capex', 'A5 + C4', 'l'),
      ('H+ add-on: extra opex (₹ L/yr)', "='Hybrid Plan'!D26", '0.0', 'hp_opex', '', 'l'),
      ('Minutes per 58-field audit (team assumption)', 10, '0', 'audmin', 'Not given in the case', 'i'),
      ('Extra audits in the decision month', '=B58-B57', '#,##0', 'audx', '20% minus the normal 10%', 'c'),
      ('Extra audit hours in October', '=B75*B74/60', '0', 'audh', 'Scheduled on normal (non-peak) days', 'c'),
      ('Extra audit cost (₹ L, at ₹350/h)', '=B76*RATE/100000', '0.00', 'audcost', 'Not counted in H\'s NPV (below ₹0.5 L)', 'c'),
      ('Extra audit load, analysts per working day', '=B76/WD/8', '0.00', 'audfte', '26 working days × 8 h', 'c')]
for k, (lab, fm, nf, key, nt, kind) in enumerate(r5):
    rr = 53 + k; put(UD, f'A{rr}', lab, sz=9)
    {'c': lambda r_, f_: calc(UD, r_, f_, nf), 'i': lambda r_, f_: inp(UD, r_, f_, nf), 'l': lambda r_, f_: link(UD, r_, f_, nf), 'p': lambda r_, f_: pv(r_, f_, nf)}[kind](f'B{rr}', fm)
    put(UD, f'C{rr}', nt, sz=8); ur(key, f'B{rr}')
# 5b mapping table
R5B = 53 + len(r5) + 1
put(UD, f'A{R5B}', '5b. What each October cut implies (base case, pasted from the ramp model)', b=True, border=False)
hdr(UD, R5B + 1, 1, ['Effectiveness m (× plan)', 'October cut c₁', 'December cut', 'Final cut, H', 'Final cut, H+', 'Final cut, A+B+C'])
for k, t in enumerate(_RL['tab']):
    rr = R5B + 2 + k
    pv(f'A{rr}', t['m'], '0.00'); [pv(f'{c}{rr}', t[key], '0.0%') for c, key in zip('BCDEF', ['oct', 'dec', 'H', 'HP', 'ABC'])]
# 5c zones
R5C = R5B + 2 + len(_RL['tab']) + 1
put(UD, f'A{R5C}', '5c. Outcome by zone (20,000 Monte Carlo runs, October audit at 20% with Poisson noise; pasted)', b=True, border=False)
hdr(UD, R5C + 1, 1, ['Zone (action)', 'Share of runs', 'P(≥75%) if H', 'P(≥75%) if H+', 'P(≥75%) all 15 parts', 'Median ΔNPV H+ vs H (₹ L)', 'Median ΔNPV all 15 (₹ L)', 'Median m in zone'], h=42)
for k, (zn, lab) in enumerate([('stay', 'c₁ ≥ 40%: stay with H'), ('escalate', '34% ≤ c₁ < 40%: escalate to H+'), ('rediagnose', 'c₁ < 34%: re-diagnose, keep H')]):
    rr = R5C + 2 + k; z = _RF['zones'][zn]; put(UD, f'A{rr}', lab, sz=9, b=True)
    for c, key, nf in zip('BCDEFGH', ['share', 'p75_H', 'p75_HP', 'p75_ABC', 'dnpv_HP', 'dnpv_ABC', 'med_m'], ['0.0%', '0.0%', '0.0%', '0.0%', '0.0', '0.0', '0.00']):
        pv(f'{c}{rr}', z[key], nf); ur(f'z_{zn}_{key}', f'{c}{rr}')
calc(UD, f'B{R5C + 5}', f'=SUM(B{R5C + 2}:B{R5C + 4})', '0.0%'); put(UD, f'A{R5C + 5}', 'Check: shares add to 100%', sz=8, i=True)
# 5d policies
R5D = R5C + 7
put(UD, f'A{R5D}', '5d. Policy comparison over the same 20,000 runs (pasted) and what the rule buys (live)', b=True, border=False)
hdr(UD, R5D + 1, 1, ['Policy', 'P(≥ 75% cut)', 'Median NPV (₹ L)', 'P(NPV < 0)', 'Worst-5% NPV (₹ L)', 'Mean NPV (₹ L)'])
for k, (lab, d, key) in enumerate([('Always Hybrid H', _RF['alwaysH'], 'polH'), ('G1 rule (stay / H+ / re-diagnose)', _RF['final'], 'polR'), ('Always H+', _RF['alwaysHP'], 'polHP')]):
    rr = R5D + 2 + k; put(UD, f'A{rr}', lab, sz=9, b=k == 1)
    for c, kk, nf in zip('BCDEF', ['p75', 'npv50', 'pneg', 'npv5', 'mean'], ['0.0%', '0.0', '0.0%', '0.0', '0.0']):
        pv(f'{c}{rr}', d[kk], nf); ur(f'{key}_{kk}', f'{c}{rr}')
a_, b_, c_ = R5D + 2, R5D + 3, R5D + 4
lv = [('Rule vs always H: extra pass chance (points)', f'=(B{b_}-B{a_})*100', '0.0', 'gain_rule'), ('Rule vs always H: median NPV given up (₹ L)', f'=C{a_}-C{b_}', '0.0', 'cost_rule'),
      ('Always H+ vs always H: extra pass chance (points)', f'=(B{c_}-B{a_})*100', '0.0', 'gain_hp'), ('Always H+ vs always H: median NPV given up (₹ L)', f'=C{a_}-C{c_}', '0.0', 'cost_hp'),
      ('Rule vs always H: change in loss chance (points)', f'=(D{b_}-D{a_})*100', '0.0', 'dloss_rule'), ('G2 (end-Dec) re-check: stay if December cut ≥ (pasted)', _RF['thresholds']['dec_stay'], '0.0%', 'g2_stay'), ('G2 re-check: H+ floor for December cut (pasted)', _RF['thresholds']['dec_esc'], '0.0%', 'g2_esc')]
for k, (lab, fm, nf, key) in enumerate(lv):
    rr = R5D + 6 + k; put(UD, f'A{rr}', lab, sz=9)
    (pv(f'B{rr}', fm, nf) if isinstance(fm, float) else calc(UD, f'B{rr}', fm, nf)); ur(key, f'B{rr}')

# ---- 6. Other statistics quoted in the deck
R6 = R5D + 6 + len(lv) + 1
section(UD, f'A{R6}', '6. Other statistics quoted in the deck (case data; tests run in Python and pasted unless a formula is shown)')
hdr(UD, R6 + 1, 1, ['Statistic', 'Value', 'p-value', 'Slide', 'Method'])
T = _R['tests']; RG = _R['regression']['coefs']; ST_ = _R['strat']['base']; Q = _R['queue']
def _vsm(k): return sum(s[k] for s in _R['vsm']['steps'] if s['n'] <= 9)
r6 = [('DPMO, normal days (raw)', ST_['Normal']['dpmo'], None, '#,##0', '04', 'Stratified by day type', 'st_norm'),
      ('DPMO, month-start carry-over days', ST_['Month-start carry-over']['dpmo'], None, '#,##0', '04', '', 'st_ms'),
      ('DPMO, month-end peak days', ST_['Month-end peak']['dpmo'], None, '#,##0', '04', '', 'st_me'),
      ('Month-end volume ÷ normal-day volume', _R['ts']['peak_index'], None, '0.0', '04', '', 'pk_idx'),
      ('Peak vs normal DPU ratio (raw)', T['peak_base']['ratio'], T['peak_base']['p'], '0.00', '04', 'Two-rate z-test', 'pk_ratio'),
      ('NB regression: peak-day rate ratio e^β1', RG['peak']['irr'], RG['peak']['p'], '0.00', '01, 04', 'Controls for quarter-end and period', 'nb_peak'),
      ('NB regression: quarter-end rate ratio e^β2', RG['qend']['irr'], RG['qend']['p'], '0.00', '—', '', 'nb_qend'),
      ('September vs August peak-day DPU ratio', T['sep_vs_aug']['ratio'], T['sep_vs_aug']['p'], '0.00', '04', 'Two-rate z-test', 'sep_aug'),
      ('Month-start vs normal-day DPU ratio', 57 / 141 / (304 / 624), 0.19546466166888476, '0.00', '04', 'Two-rate z-test: not significant', 'ms_ratio'),
      ('Correlation, daily volume vs DPU (r)', _R['corr']['base']['r'], _R['corr']['base']['p'], '0.00', '03, 04', 'Pearson', 'r_dpu'),
      ('Correlation, daily volume vs raw error count (r)', _R['norm']['corr_raw_err_vs_inv'], None, '0.00', '03', 'Why raw counts mislead', 'r_raw'),
      ('Trend in daily DPU, slope per day', _R['trend']['base']['slope_per_day'], _R['trend']['base']['p'], '0.0000', '04', 'OLS on day index', 'trend'),
      ('u-chart: baseline days above the upper limit', 5, None, '0', '04', 'Limits ū ± 3√(ū/n)', 'uc_hi'),
      ('u-chart: baseline days below the lower limit', 3, None, '0', '04', '', 'uc_lo'),
      ('u-chart: supplied-audit days outside limits', 0, None, '0', '14', '', 'uc_imp'),
      ('u-chart centre line ū (all 52 days) = raw DPMO × 58 ÷ 10⁶', "='Baseline'!B24*OPP/1000000", None, '0.000', '04', 'Live', 'ubar'),
      ('Worst baseline day (28-Sep), errors per audited invoice', 1.0574712643678161, None, '0.00', '04', '', 'worst'),
      ('Before/after test, z', T['before_after']['z'], T['before_after']['p'], '0.0', '14', 'Two-rate z-test (p < 0.0001)', 'z_ba'),
      ('Supplied audit: peak vs normal DPU ratio', T['peak_imp']['ratio'], T['peak_imp']['p'], '0.0', '14', 'Borderline', 'pk_imp'),
      ('Errors from vendor submissions (count)', _R['src']['Vendor submission'], None, '0', '05', 'Error Reasons tab', 'src_v'),
      ('Errors from systems (count)', _R['src']['System'], None, '0', '05', '', 'src_s'),
      ('Errors from internal process (count)', _R['src']['Internal process'], None, '0', '05', '', 'src_i'),
      ('Share from vendor submissions', f"=B{R6 + 21}/'Baseline'!B8", None, '0.0%', '01, 05', 'Live', 'sh_v'),
      ('Share from systems', f"=B{R6 + 22}/'Baseline'!B8", None, '0.0%', '05', 'Live', 'sh_s'),
      ('Share from internal process', f"=B{R6 + 23}/'Baseline'!B8", None, '0.0%', '05', 'Live', 'sh_i'),
      ('Top three reasons, share of all errors', "=SUM('Baseline'!M5:M7)", None, '0.0%', '01, 05', 'Live', 'top3'),
      ('Invoices per day, highest day', "=MAX('Baseline'!E5:E56)", None, '#,##0', '01', 'Live', 'maxday'),
      ('August invoices', "=SUMPRODUCT((MONTH('Baseline'!D5:D56)=8)*'Baseline'!E5:E56)", None, '#,##0', '02', 'Live', 'aug'),
      ('September invoices', "=SUMPRODUCT((MONTH('Baseline'!D5:D56)=9)*'Baseline'!E5:E56)", None, '#,##0', '02', 'Live', 'sep'),
      ('August invoices per audit day', f"=B{R6 + 29}/SUMPRODUCT(--(MONTH('Baseline'!D5:D56)=8))", None, '0', '03', 'Live', 'aug_d'),
      ('September invoices per audit day', f"=B{R6 + 30}/SUMPRODUCT(--(MONTH('Baseline'!D5:D56)=9))", None, '0', '03', 'Live', 'sep_d'),
      ('Peak load: 880 invoices ÷ peak-day capacity', "=MAX('Baseline'!E5:E56)/'Baseline'!B42", None, '0%', '01, 04, 06', 'Live', 'maxutil'),
      ('Days over capacity (of 52)', _R['capacity']['over_days'], None, '0', '04', '', 'overdays'),
      ('Queue (M/M/c, c = 18, μ = 60/23 per h): load at 186 per day', Q['Median normal day (186)']['rho'], None, '0%', '06', 'Erlang C', 'q186'),
      ('… load at 320 per day', Q['Average day (320)']['rho'], None, '0%', '06', '', 'q320'),
      ('… mean wait in queue at 320 per day (min)', Q['Average day (320)']['wq_min'], None, '0.0', '06', 'W_q = C(c, a) ÷ (cμ − λ)', 'wq320'),
      ('… load at 376 per day', Q['Busy normal day (376)']['rho'], None, '0%', '06', '', 'q376'),
      ('… load at 600 per day (with overtime)', Q['Peak day 600 (with OT)']['rho'], None, '0%', '06', '', 'q600'),
      ('… load at 880 per day (with overtime)', Q['Peak day 880 (with OT)']['rho'], None, '0%', '06', '', 'q880'),
      ('Queue simulation (52 baseline days): invoices waiting 2+ days', _R['des']['base']['share_2plus'], None, '0%', '02, 06', 'Discrete-event simulation', 'des2'),
      ('Queue simulation: largest month-end backlog (invoices)', _R['des']['base']['max_backlog'], None, '#,##0', '06', 'Shown as ~1,800', 'backlog'),
      ('Value stream: touch minutes per invoice (steps 1–9, weighted)', _vsm('w_tt'), None, '0', '06', 'Case Process Steps tab', 'vsm_tt'),
      ('Value stream: elapsed minutes per invoice', _vsm('w_el'), None, '0', '06', '', 'vsm_el'),
      ('Process cycle efficiency = touch ÷ elapsed', f'=B{R6 + 43}/B{R6 + 44}', None, '0%', '06', 'Live', 'pce'),
      ('Approval wait (steps 6 and 9), minutes', 48, None, '0', '06', '50 elapsed − 2 touch', 'appr'),
      ('Approval wait share of lead time', f'=B{R6 + 46}/B{R6 + 44}', None, '0%', '06', 'Live', 'appr_sh'),
      ('Exception rework share of touch time (25% × 12 min ÷ 23)', f'=0.25*12/B{R6 + 43}', None, '0%', '06', 'Live', 'exc_sh'),
      ('25-Aug invoices (the zero-error peak day)', "=SUMIFS('Baseline'!E5:E56,'Baseline'!D5:D56,DATE(2026,8,25))", None, '#,##0', '03', 'Live', 'aug25'),
      ('Busiest day below 400 invoices', "=_xlfn.MAXIFS('Baseline'!E5:E56,'Baseline'!E5:E56,\"<400\")", None, '#,##0', '03, 15', 'Live: no day between this and the next row', 'gap_lo'),
      ('Quietest day at or above 400 invoices', "=_xlfn.MINIFS('Baseline'!E5:E56,'Baseline'!E5:E56,\">=400\")", None, '#,##0', '03, 15', 'Live: peak day = ≥ 400', 'gap_hi')]
r6 += [('Month-end vs normal-day DPMO ratio (calendar groups)', f'=B{R6 + 4}/B{R6 + 2}', None, '0.00', '04', 'Live; differs from the ≥ 400-invoice ratio (1.44×)', 'me_ratio'),
       ('Dispersion φ on peak (≥ 400) days, baseline', _LAN['phi_peak'], None, '0.0', '04', 'Pasted', 'phi_pk'), ('Dispersion φ on other days, baseline', _LAN['phi_nonpeak'], None, '0.0', '04', 'Pasted', 'phi_np')]
r6 += [(f'Average invoices, {n}{"st" if n == 1 else "nd" if n == 2 else "rd" if n == 3 else "th"} last working day of the month', v, None, '#,##0', '04', 'Daily volume chart' if n == 1 else '', f'prof{n}')
       for n, v in sorted(((int(k_), v_) for k_, v_ in _R['ts']['profile'].items()))]
for k, (lab, v, p, nf, sl_, mt, key) in enumerate(r6):
    rr = R6 + 2 + k; put(UD, f'A{rr}', lab, sz=9)
    (calc(UD, f'B{rr}', v, nf) if isinstance(v, str) else pv(f'B{rr}', v, nf))
    if p is not None: pv(f'C{rr}', p, '0.000' if p >= 0.001 else '0.0E+00'); ur(key + '_p', f'C{rr}')
    put(UD, f'D{rr}', sl_, sz=8, al='center'); put(UD, f'E{rr}', mt, sz=8); ur(key, f'B{rr}')
assert UD[f'A{R6 + 21}'].value.startswith('Errors from vendor') and UD[f'A{R6 + 43}'].value.startswith('Value stream: touch') and UD[f'A{R6 + 29}'].value == 'August invoices'
# FMEA (team scores) and COPQ sensitivity
R7 = R6 + 2 + len(r6) + 1
section(UD, f'A{R7}', '7. FMEA scores (team judgement) and the COPQ sensitivity on slide 15')
hdr(UD, R7 + 1, 1, ['Failure mode', 'Severity', 'Occurrence', 'Detection', 'RPN = S × O × D'])
for k, (fm_, s_, o_, d_) in enumerate([('Missing documents', 7, 9, 5), ('Unclear line items', 5, 10, 6), ('PO amount mismatch', 7, 10, 4), ('Duplicate invoice', 9, 4, 5), ('ERP sync failure', 6, 4, 7)]):
    rr = R7 + 2 + k; put(UD, f'A{rr}', fm_, sz=9)
    for c, v in zip('BCD', (s_, o_, d_)): inp(UD, f'{c}{rr}', v, '0')
    calc(UD, f'E{rr}', f'=B{rr}*C{rr}*D{rr}', '0'); ur(f'rpn{k}', f'E{rr}')
R8 = R7 + 8
hdr(UD, R8, 1, ['Annual COPQ when one driver changes (₹ L)', 'Low', 'High', 'Formula'])
r8 = [('Penalty trigger rate 2% → 8%', '=COPQ_PEN*0.02/PEN+COPQ_RW', '=COPQ_PEN*0.08/PEN+COPQ_RW', 'Penalty part scales with the trigger rate', 'cs_pen'),
      ('Monte Carlo P5 → P95 (pasted)', _R['mc']['p5'] / 1e5, _R['mc']['p95'] / 1e5, 'All cost drivers drawn together', 'cs_mc'),
      ('Penalty ₹2,500 → ₹5,000 per invoice', '=COPQ_PEN*2500/PEN_AMT+COPQ_RW', '=COPQ', '', 'cs_amt'),
      ('Share of invoices with errors 35% → 46% (Poisson)', '=COPQ', "='Baseline'!B37", 'Case basis vs Poisson basis', 'cs_share')]
for k, (lab, lo_, hi_, nt, key) in enumerate(r8):
    rr = R8 + 1 + k; put(UD, f'A{rr}', lab, sz=9)
    for c, v in (('B', lo_), ('C', hi_)): (calc(UD, f'{c}{rr}', v, '0.0') if isinstance(v, str) else pv(f'{c}{rr}', v, '0.0'))
    put(UD, f'D{rr}', nt, sz=8); ur(key + '_lo', f'B{rr}'); ur(key + '_hi', f'C{rr}')
note(f'A{R8 + 6}', 'Grey cells are pasted from the Python model (method stated alongside); yellow cells are inputs; black cells are live formulas; green cells link to other sheets.')
UD.freeze_panes = 'A4'
_j.dump(UREF, open(paths.out('uref.json'), 'w'), indent=0)
