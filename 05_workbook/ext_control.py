"""Workbook extension (exec-ed by build_workbook.py) - "Control Chart" sheet (standard u-chart vs Laney u-prime, live formulas) and a few notes.
Reads : nothing directly (uses cells defined by earlier sheets)
Writes: outputs/uref.json (final version of the cell map); adds a sheet to the workbook object `wb`
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Not run on its own. Builds the live control-chart table (formulas) and appends a few explanatory notes to other sheets.
# ---------------------------------------------------------------------------
# ============================================================ CONTROL CHART (standard u-chart vs Laney u′, live)
CCN = 'Control Chart'
CC = wb.create_sheet(CCN)
title(CC, 'Control Chart · standard u-chart vs Laney u′ (dispersion-adjusted), live',
      'Baseline errors are over-dispersed (φ ≈ 4.4), so the standard limits ū ± 3√(ū/n) are too narrow. The Laney u′ chart widens them by σz, the day-to-day spread of the standardised points. σz is only applied when it exceeds 1.')
widths(CC, {'A': 12, 'B': 10, 'C': 9, 'D': 9, 'E': 9, 'F': 9, 'G': 9, 'H': 9, 'I': 9, 'J': 9, 'K': 9, 'L': 10, 'M': 10, 'N': 3, 'O': 46, 'P': 14, 'Q': 14})
def cref(key, cell): UREF[key] = f"'{CCN}'!{cell}"
HD = ['Day', 'Audited n', 'Errors', 'u = errors ÷ n', '√(ū ÷ n)', 'z = (u − ū) ÷ √(ū/n)', 'MR = |z − z prev|', 'UCL standard', 'LCL standard', 'UCL Laney', 'LCL Laney', 'Out, standard', 'Out, Laney']
def block(r0, label, rows, ub_ref, sz_ref):
    put(CC, f'A{r0}', label, b=True, border=False); hdr(CC, r0 + 1, 1, HD, h=42)
    first = r0 + 2
    for k, (dref, nref, eref) in enumerate(rows):
        r = first + k
        link(CC, f'A{r}', dref, 'dd-mmm'); link(CC, f'B{r}', nref, '0'); link(CC, f'C{r}', eref, '0')
        calc(CC, f'D{r}', f'=C{r}/B{r}', '0.000'); calc(CC, f'E{r}', f'=SQRT({ub_ref}/B{r})', '0.000'); calc(CC, f'F{r}', f'=(D{r}-{ub_ref})/E{r}', '0.00')
        calc(CC, f'G{r}', f'=ABS(F{r}-F{r-1})' if k else '', '0.00')
        calc(CC, f'H{r}', f'={ub_ref}+3*E{r}', '0.000'); calc(CC, f'I{r}', f'=MAX(0,{ub_ref}-3*E{r})', '0.000')
        calc(CC, f'J{r}', f'={ub_ref}+3*{sz_ref}*E{r}', '0.000'); calc(CC, f'K{r}', f'=MAX(0,{ub_ref}-3*{sz_ref}*E{r})', '0.000')
        calc(CC, f'L{r}', f'=IF(OR(D{r}>H{r},D{r}<I{r}),1,0)', '0'); calc(CC, f'M{r}', f'=IF(OR(D{r}>J{r},D{r}<K{r}),1,0)', '0')
        for c in 'ABCDEFGHIJKLM': CC[f'{c}{r}'].font = f(8, c=GREEN if c in 'ABC' else BLACK)
    return first, first + len(rows) - 1
# baseline rows (adjusted: 25-Aug excluded), linked to the Baseline sheet
B_ROWS = [(f"='Baseline'!D{5+i}", f"='Baseline'!F{5+i}", f"='Baseline'!G{5+i}") for i in range(52) if i != 20]
I_ROWS = [(f"='{UN}'!K{5+i}", f"='{UN}'!M{5+i}", f"='{UN}'!N{5+i}") for i in range(51)]
# summary block
section(CC, 'O4', 'Summary')
hdr(CC, 5, 15, ['Measure', 'Baseline (adj.)', 'Supplied audit'])
b0, b1 = 40, 40 + len(B_ROWS) - 1
i0, i1 = b1 + 5, b1 + 5 + len(I_ROWS) - 1
summ = [('Centre line ū = Σ errors ÷ Σ audited', f'=SUM(C{b0}:C{b1})/SUM(B{b0}:B{b1})', f'=SUM(C{i0}:C{i1})/SUM(B{i0}:B{i1})', '0.0000', 'ubar'),
        ('Average moving range of z (MR̄)', f'=AVERAGE(G{b0+1}:G{b1})', f'=AVERAGE(G{i0+1}:G{i1})', '0.00', 'mrbar'),
        ('σz = MR̄ ÷ 1.128 (1 = Poisson-like)', '=P7/1.128', '=Q7/1.128', '0.00', 'sz'),
        ('σz applied to the limits = MAX(1, σz)', '=MAX(1,P8)', '=MAX(1,Q8)', '0.00', 'szu'),
        ('Days outside standard u-chart limits', f'=SUM(L{b0}:L{b1})', f'=SUM(L{i0}:L{i1})', '0', 'out_std'),
        ('… of which above the upper limit', f'=SUMPRODUCT(--(D{b0}:D{b1}>H{b0}:H{b1}))', f'=SUMPRODUCT(--(D{i0}:D{i1}>H{i0}:H{i1}))', '0', 'hi_std'),
        ('Days outside Laney u′ limits', f'=SUM(M{b0}:M{b1})', f'=SUM(M{i0}:M{i1})', '0', 'out_lan'),
        ('Standard-limit signals on days with ≥ 400 invoices', f'=SUMPRODUCT(L{b0}:L{b1},--(Q{b0}:Q{b1}>=400))', '', '0', 'sig_peak')]
for k, (lab, fb, fi, nf, key) in enumerate(summ):
    r = 6 + k; put(CC, f'O{r}', lab, sz=9); calc(CC, f'P{r}', fb, nf); cref(key + '_b', f'P{r}')
    if fi: calc(CC, f'Q{r}', fi, nf); cref(key + '_i', f'Q{r}')
put(CC, 'O15', '25-Aug (excluded day): audited invoices', sz=9); link(CC, 'P15', "='Baseline'!F25", '0')
put(CC, 'O16', '25-Aug: errors found', sz=9); link(CC, 'P16', "='Baseline'!G25", '0')
put(CC, 'O17', '25-Aug: Laney lower limit at its sample size', sz=9); calc(CC, 'P17', '=MAX(0,P6-3*P9*SQRT(P6/P15))', '0.000'); cref('aug25_lcl', 'P17')
put(CC, 'O18', '25-Aug outside the Laney limits?', sz=9); calc(CC, 'P18', '=IF(P16/P15<P17,"Yes – suspected recording anomaly","No")', '@')
note('O20', 'Reading: with dispersion-adjusted limits no adjusted-baseline day is a special cause; the variation is systematic (month-end), which calls for process redesign rather than day-by-day firefighting. After the improvement σz < 1, so standard limits apply and the daily trigger in the control plan (≥ 4 errors in 30 audited) stands. Re-estimate σz monthly.')
CC['O20'].alignment = Alignment(wrap_text=True, vertical='top'); CC.merge_cells('O20:Q26')
first_b, last_b = block(38, 'Baseline Aug–Sep 2026 (adjusted: 25-Aug excluded), linked to the Baseline sheet', B_ROWS, '$P$6', '$P$9')
assert (first_b, last_b) == (b0, b1)
first_i, last_i = block(b1 + 3, 'Supplied audit Mar–Apr 2027, linked to the Uncertainty & Decision sheet', I_ROWS, '$Q$6', '$Q$9')
assert (first_i, last_i) == (i0, i1), (first_i, i0)
# invoices received per day (for the ≥ 400 check), column Q next to the baseline block
put(CC, f'Q{b0-1}', 'Invoices', b=True, sz=8)
bi = 0
for i in range(52):
    if i == 20: continue
    link(CC, f'Q{b0 + bi}', f"='Baseline'!E{5+i}", '0'); CC[f'Q{b0 + bi}'].font = f(8, c=GREEN); bi += 1
CC.freeze_panes = 'A4'
CC.sheet_properties.tabColor = '7030A0'
_j.dump(UREF, open(paths.out('uref.json'), 'w'), indent=0)
for _sh in ('Stress Test', 'Combination Search'):
    put(wb[_sh], 'A3', 'Pasted simulation outputs (grey in the Uncertainty sheet): change them by rerunning the Python model. Base-case cut, DPMO, NPV and payback rows link to live sheets. Probabilities are shares of simulated scenarios under assumed ranges, not real-world odds.', sz=8, i=True, border=False)
_na = AS.max_row + 1
for _a, _b, _c in [('H14', 'Control limits: Laney u′ (limits widened by σz, the moving-range spread of standardised daily points) whenever σz > 1; standard u-chart otherwise. Baseline σz = 1.86; after improvement 0.69.', 'Control Chart'),
                   ('H15', 'A full 58-field audit of one invoice takes 10 minutes; the doubled October audit (≈833 extra invoices, ≈139 h) is done on normal days by the audit team, ≈₹0.5 L, not counted in NPV.', 'Uncertainty & Decision')]:
    put(AS, f'A{_na}', _a, b=True, al='center'); put(AS, f'B{_na}', _b, wrap=True); put(AS, f'C{_na}', _c, sz=9); AS.row_dimensions[_na].height = 42; _na += 1
