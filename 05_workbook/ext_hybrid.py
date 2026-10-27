"""Workbook extension (exec-ed by build_workbook.py) - Parts Library, Hybrid H, Hybrid Plan, Part Builder, Stress Test, Combination Search sheets.
Reads : parts.json, stage3.json, marginal.json (outputs/ if present, else precomputed/)
Writes: nothing directly (adds sheets to the workbook object `wb` built by build_workbook.py)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Not run on its own. Uses the helper functions defined in build_workbook.py (put, calc, link, hdr ...) to lay out the sheets that
#   show the 15 components, the recommended Hybrid H plan, a "build your own" selector with live formulas, and pasted stress-test tables.
# ---------------------------------------------------------------------------
# Extension: Parts Library, Hybrid H, Part Builder, Stress Test, Combination Search sheets (exec'd inside build_model2.py)
import json as _json
PJ = _json.load(open(paths.find('parts.json')))
ST = _json.load(open(paths.find('stage3.json')))
MG = _json.load(open(paths.find('marginal.json')))
PARTS = PJ['parts']; PRES = PJ['presets']; NPARTS = len(PARTS); NRS = len(REASONS)
PRESET_NAMES = ['A', 'B', 'C', 'A+B+C', 'Hybrid H', 'H+ (quality)', 'H-lite']
F_KEEP = PatternFill('solid', fgColor='E2EFDA'); F_DROP = PatternFill('solid', fgColor='FBE5D6'); F_WARN = PatternFill('solid', fgColor='FFF2CC')
SHORTR = PJ['reasons']
BOTHER = f"'{SOLREF['B']['sheet']}'!C{int(SOLREF['B']['BEN'][1:]) - 1}"     # Sol B other saving (scenario-adjusted)

# ============================================================ PARTS LIBRARY
PL = wb.create_sheet('Parts Library')
title(PL, 'Parts library – the 15 building blocks of Solutions A, B and C',
      'Each solution is split into five parts. Part effectiveness is calibrated so the five parts of a solution, combined, reproduce that solution sheet\'s per-reason effectiveness exactly. Edit yellow cells to test other assumptions.')
PCOL = {'code': 'A', 'sol': 'B', 'name': 'C', 'capex': 'D', 'opex': 'E', 'month': 'F', 'peak': 'G', 'touch': 'H', 'digital': 'I', 'scope': 'J', 'legal': 'K'}
PRESCOL = {p: CL(12 + i) for i, p in enumerate(PRESET_NAMES)}          # L..R
EFFC0 = 12 + len(PRESET_NAMES)                                          # S
widths(PL, {'A': 7, 'B': 9, 'C': 46, 'D': 9, 'E': 9, 'F': 9, 'G': 10, 'H': 9, 'I': 10, 'J': 10, 'K': 9})
for i in range(len(PRESET_NAMES)): PL.column_dimensions[CL(12 + i)].width = 9
for k in range(NRS): PL.column_dimensions[CL(EFFC0 + k)].width = 9
section(PL, 'A4', 'Parts, costs, timing and flags'); PL.cell(row=4, column=12, value='Membership (1 = part included)').font = f(10, True, NAVY)
PL.cell(row=4, column=EFFC0, value='Share of each error reason removed by the part (calibrated input)').font = f(10, True, NAVY)
hdr(PL, 5, 1, ['Code', 'Solution', 'Part', 'Capex ₹ L', 'Opex ₹ L/yr', 'Go-live month', 'Peak-excess errors removed', 'Touch-time change (min)',
               'Needs digital intake', 'Outside steps 1–9', 'Legal risk'] + PRESET_NAMES + SHORTR, h=48)
PROW0 = 6
for i, p in enumerate(PARTS):
    r = PROW0 + i
    put(PL, f'A{r}', p['code'], b=True, al='center'); put(PL, f'B{r}', p['sol'], al='center'); put(PL, f'C{r}', p['name'], sz=9, wrap=True)
    inp(PL, f'D{r}', p['capex'], '0.0'); inp(PL, f'E{r}', p['opex'], '0.0'); inp(PL, f'F{r}', p['month'], '0.0'); inp(PL, f'G{r}', p['peak'], '0%'); inp(PL, f'H{r}', p['touch'], '+0.0;-0.0;0.0')
    put(PL, f'I{r}', 'Yes' if p['flag'] == 'needs_digital' else 'No', al='center', sz=9)
    put(PL, f'J{r}', 'Yes' if p['flag'] == 'outscope' else 'No', al='center', sz=9, fill=F_WARN if p['flag'] == 'outscope' else None)
    put(PL, f'K{r}', 'Yes' if p['flag'] == 'legal' else 'No', al='center', sz=9, fill=F_WARN if p['flag'] == 'legal' else None)
    for pn in PRESET_NAMES:
        v = PRES[pn][i]; c = put(PL, f'{PRESCOL[pn]}{r}', v, '0', al='center', fill=(F_KEEP if v else None))
        if pn == 'Hybrid H': c.font = f(10, True, '006100' if v else 'C00000')
    for k in range(NRS):
        inp(PL, f'{CL(EFFC0 + k)}{r}', round(p['eff'][k], 4), '0%')
    PL.row_dimensions[r].height = 26
PROW1 = PROW0 + NPARTS - 1
r = PROW1 + 1
put(PL, f'C{r}', 'Totals (all 15 parts)', b=True, fill=F_TOT)
for col in 'DE': calc(PL, f'{col}{r}', f'=SUM({col}{PROW0}:{col}{PROW1})', '0.0', True, F_TOT)
for pn in PRESET_NAMES: calc(PL, f'{PRESCOL[pn]}{r}', f'=SUM({PRESCOL[pn]}{PROW0}:{PRESCOL[pn]}{PROW1})', '0', True, F_TOT)
r += 2
notes = ['How parts combine: within a solution, parts act independently: 1 − Π(1 − part effectiveness). Across solutions, an overlap setting blends independent effects with full overlap (only the strongest lever counts).',
         'Dependencies: A2, A3 and A5 need digital intake – full effect with A1 (portal), half effect with only A4 (IRN feed covers vendors above ₹5 Cr), none without either.',
         'A2\'s 3-way match relies on timely GRN data: without C2 (PO before work, GRN in 48 h) its PO-mismatch effect is multiplied by the GRN factor (stress input).',
         'Costs of the five parts add up to each solution\'s capex/opex (A 45/10, B 6/5, C 8/14). Part splits, go-live months and effectiveness splits are team assumptions.',
         'Note: editing effectiveness on the Sol A/B/C sheets does not update this library (the library is a separate calibrated copy).']
for k, t in enumerate(notes):
    put(PL, f'A{r+k}', t, sz=9, wrap=True, border=False); PL.merge_cells(f'A{r+k}:K{r+k}'); PL.row_dimensions[r + k].height = 28
PL.freeze_panes = 'D6'

def plref(col, i, absolute=True): return f"'Parts Library'!${col}${PROW0 + i}" if absolute else f"'Parts Library'!{col}{PROW0 + i}"
IDXP = {p['code']: i for i, p in enumerate(PARTS)}
GRPP = {g: [i for i, p in enumerate(PARTS) if p['sol'] == g] for g in 'ABC'}
PO_K = SHORTR.index('PO amount mismatch')

# ============================================================ OPTION CALCULATOR (used by Hybrid H and Part Builder)
def build_option(ws, sheet_title, sub, mode, fixed='Hybrid H'):
    title(ws, sheet_title, sub)
    widths(ws, {'A': 44, 'B': 16, 'C': 13, 'D': 11, 'E': 11, 'F': 12, 'G': 12, 'H': 12, 'I': 11, 'J': 13, 'K': 13})
    section(ws, 'A4', 'Stress settings (1 = base case; lower = harsher)')
    st = [('Overlap between solutions (1 = independent, 0 = full overlap)', 1.0), ('GRN timeliness factor (applies to A2 when C2 is absent)', 1.0),
          ('Legal factor for C5 "No PO, no pay" (0 = blocked)', 1.0), ('Vendor compliance with month-end calendar (B3)', 1.0)]
    for k, (lab, v) in enumerate(st):
        put(ws, f'A{5+k}', lab, sz=9); inp(ws, f'B{5+k}', v, '0.00')
    OV, GRN, LEG, CAL = '$B$5', '$B$6', '$B$7', '$B$8'
    put(ws, 'A10', 'Option shown', b=True)
    if mode == 'H':
        put(ws, 'B10', fixed, b=True, al='center', fill=F_KPI)
        put(ws, 'C10', f'Fixed to the "{fixed}" column of the Parts Library. Use Part Builder to try other mixes.', sz=9, c=GREY, border=False)
    else:
        inp(ws, 'B10', 'Hybrid H')
        dv2 = DataValidation(type='list', formula1='"' + ','.join(PRESET_NAMES + ['Custom']) + '"'); ws.add_data_validation(dv2); dv2.add('B10')
        put(ws, 'C10', 'Pick a preset, or choose "Custom" and set 1/0 in column E of the parts table.', sz=9, c=GREY, border=False)
    # ---- parts table
    P0 = 37
    section(ws, f'A{P0-2}', 'Parts in this option')
    hdr(ws, P0 - 1, 1, ['Part', 'Code', 'Solution', 'Selected (1/0)', 'Custom (1/0)', 'Effect factor', 'Capex ₹ L', 'Opex ₹ L/yr', 'Go-live month', 'Peak removed', 'Touch Δ min'], h=34)
    for i, p in enumerate(PARTS):
        r = P0 + i
        put(ws, f'A{r}', p['name'], sz=9); put(ws, f'B{r}', p['code'], al='center', b=True); put(ws, f'C{r}', p['sol'], al='center')
        if mode == 'H':
            link(ws, f'D{r}', f"={plref(PRESCOL[fixed], i)}", '0'); put(ws, f'E{r}', '–', al='center', c=GREY)
        else:
            inp(ws, f'E{r}', PRES['Hybrid H'][i], '0')
            calc(ws, f'D{r}', f"=IF($B$10=\"Custom\",E{r},INDEX('Parts Library'!$L${PROW0+i}:$R${PROW0+i},MATCH($B$10,'Parts Library'!$L$5:$R$5,0)))", '0')
        fl = p['flag']
        fac = '$D$54' if fl == 'needs_digital' else (LEG if fl == 'legal' else '1')
        calc(ws, f'F{r}', f'={fac}', '0.00')
        calc(ws, f'G{r}', f"=D{r}*{plref('D', i)}*M_COST", '0.0'); calc(ws, f'H{r}', f"=D{r}*{plref('E', i)}*M_COST", '0.0')
        link(ws, f'I{r}', f"={plref('F', i)}", '0.0')
        calc(ws, f'J{r}', f"=MIN(1,{plref('G', i)}*M_EFF)*D{r}" + (f'*{CAL}' if p['code'] == 'B3' else ''), '0%')
        calc(ws, f'K{r}', f"=D{r}*{plref('H', i)}*F{r}" if fl == 'needs_digital' else f"=D{r}*{plref('H', i)}", '0.0')
    P1 = P0 + NPARTS - 1
    put(ws, f'A{P1+1}', 'Totals', b=True, fill=F_TOT)
    for col, nf in (('D', '0'), ('G', '0.0'), ('H', '0.0'), ('K', '0.0')): calc(ws, f'{col}{P1+1}', f'=SUM({col}{P0}:{col}{P1})', nf, True, F_TOT)
    sel = lambda code: f'$D${P0 + IDXP[code]}'
    put(ws, 'A54', 'Digital coverage for A2/A3/A5 (A1 → 1, only A4 → 0.5, neither → 0)', sz=9)
    calc(ws, 'D54', f"=IF({sel('A1')}=1,1,IF({sel('A4')}=1,0.5,0))", '0.0')
    put(ws, 'A55', 'Dependency check', sz=9)
    calc(ws, 'D55', f"=IF(AND({sel('A1')}+{sel('A4')}=0,{sel('A2')}+{sel('A3')}+{sel('A5')}>0),\"A2/A3/A5 need A1 or A4\",\"OK\")")
    # ---- adjusted effectiveness matrix (parts × reasons) in columns M..AE, rows aligned with parts table
    AM0 = 13
    ws.cell(row=P0 - 2, column=AM0, value='Adjusted part effectiveness = MIN(1, library value × effectiveness multiplier) × selected × factor').font = f(10, True, NAVY)
    hdr(ws, P0 - 1, AM0, SHORTR, h=34)
    for k in range(NRS): ws.column_dimensions[CL(AM0 + k)].width = 8
    for i, p in enumerate(PARTS):
        r = P0 + i
        for k in range(NRS):
            extra = ''
            if p['code'] == 'A2' and k == PO_K: extra = f"*IF({sel('C2')}=1,1,{GRN})"
            c = calc(ws, f'{CL(AM0+k)}{r}', f"=MIN(1,{plref(CL(EFFC0+k), i)}*M_EFF)*$D{r}*$F{r}{extra}", '0%'); c.font = f(8)
    def adj(i, k): return f'${CL(AM0+k)}${P0+i}'
    # ---- per-reason model (all parts live)
    R0 = 60
    section(ws, f'A{R0-2}', 'Error model at full run-rate (all selected parts live)')
    hdr(ws, R0 - 1, 1, ['Error reason', 'Share', 'A parts', 'B parts', 'C parts', 'Combined', 'Share removed'])
    for k, rs in enumerate(REASONS):
        r = R0 + k
        put(ws, f'A{r}', rs, sz=9); link(ws, f'B{r}', f"='Baseline'!M{5+k}", '0.0%')
        for col, g in zip('CDE', 'ABC'):
            calc(ws, f'{col}{r}', '=1-' + '*'.join(f'(1-{adj(i, k)})' for i in GRPP[g]), '0%')
        calc(ws, f'F{r}', f'={OV}*(1-(1-C{r})*(1-D{r})*(1-E{r}))+(1-{OV})*MAX(C{r}:E{r})', '0%'); calc(ws, f'G{r}', f'=B{r}*F{r}', '0.0%')
    R1 = R0 + NRS - 1; RT = R1 + 1
    put(ws, f'A{RT}', 'Reason-level cut', b=True, fill=F_TOT); calc(ws, f'G{RT}', f'=SUM(G{R0}:G{R1})', '0.0%', True, F_TOT)
    put(ws, f'A{RT+1}', 'Peak-excess removal by solution group → combined')
    for col, g in zip('CDE', 'ABC'):
        calc(ws, f'{col}{RT+1}', '=1-' + '*'.join(f'(1-$J${P0+i})' for i in GRPP[g]), '0%')
    calc(ws, f'F{RT+1}', f'={OV}*(1-(1-C{RT+1})*(1-D{RT+1})*(1-E{RT+1}))+(1-{OV})*MAX(C{RT+1}:E{RT+1})', '0%')
    put(ws, f'A{RT+2}', 'Peak cut = peak-excess share × combined removal'); link(ws, f'B{RT+2}', '=EXCESS', '0.0%'); calc(ws, f'G{RT+2}', f'=B{RT+2}*F{RT+1}', '0.0%')
    put(ws, f'A{RT+3}', 'TOTAL ERROR CUT', b=True, fill=F_KPI); calc(ws, f'G{RT+3}', f'=1-(1-G{RT})*(1-G{RT+2})', '0.0%', True, F_KPI)
    CUTc = f'$G${RT+3}'
    ws.conditional_formatting.add(f'G{R0}:G{R1}', DataBarRule(start_type='num', start_value=0, end_type='max', color='5B8BD6'))
    # ---- monthly ramp (36 months) – helper block
    M0 = RT + 8
    section(ws, f'A{M0-2}', 'Monthly ramp: each part starts paying back the month after it goes live (helper block). Month 1 = Sep 2026 (build only), 2 = Oct, … 6 = Feb')
    labels = ['Month']
    for k in range(NRS): labels += [f'{SHORTR[k]} A', 'B', 'C', 'comb.']
    labels += ['Reason cut', 'Peak A', 'Peak B', 'Peak C', 'Peak comb.', 'Total cut', 'Benefit ₹ L/yr rate', 'Cash flow ₹ L', 'Cumulative ₹ L']
    hdr(ws, M0 - 1, 1, labels, h=40)
    ncol = len(labels)
    for mth in range(1, 37):
        r = M0 + mth - 1
        put(ws, f'A{r}', mth, '0', al='center')
        combs = []
        for k in range(NRS):
            cA, cB, cC, cM = (CL(2 + 4 * k + j) for j in range(4))
            for col, g in zip((cA, cB, cC), 'ABC'):
                fm = '=1-' + '*'.join(f'(1-{adj(i, k)}*($I${P0+i}<$A{r}))' for i in GRPP[g])
                c = calc(ws, f'{col}{r}', fm, '0%'); c.font = f(8)
            c = calc(ws, f'{cM}{r}', f'={OV}*(1-(1-{cA}{r})*(1-{cB}{r})*(1-{cC}{r}))+(1-{OV})*MAX({cA}{r},{cB}{r},{cC}{r})', '0%'); c.font = f(8)
            combs.append(f'{cM}{r}')
        b0 = 2 + 4 * NRS
        cRC, cPA, cPB, cPC, cPM, cCUT, cBEN, cCF, cCUM = (CL(b0 + j) for j in range(9))
        calc(ws, f'{cRC}{r}', '=' + '+'.join(f'$B${R0+k}*{combs[k]}' for k in range(NRS)), '0.0%')
        for col, g in zip((cPA, cPB, cPC), 'ABC'):
            calc(ws, f'{col}{r}', '=1-' + '*'.join(f'(1-$J${P0+i}*($I${P0+i}<$A{r}))' for i in GRPP[g]), '0%')
        calc(ws, f'{cPM}{r}', f'={OV}*(1-(1-{cPA}{r})*(1-{cPB}{r})*(1-{cPC}{r}))+(1-{OV})*MAX({cPA}{r},{cPB}{r},{cPC}{r})', '0%')
        calc(ws, f'{cCUT}{r}', f'=1-(1-{cRC}{r})*(1-EXCESS*{cPM}{r})', '0.0%')
        b3 = P0 + IDXP['B3']
        calc(ws, f'{cBEN}{r}', f'=COPQ*{cCUT}{r}+{BOTHER}*{CAL}*$D${b3}*($I${b3}<$A{r})', '0.0')
        calc(ws, f'{cCF}{r}', f'={cBEN}{r}/12-$H${P1+1}/12', '0.00')
        calc(ws, f'{cCUM}{r}', (f'=-$G${P1+1}+{cCF}{r}' if mth == 1 else f'={cCUM}{r-1}+{cCF}{r}'), '0.0')
    MEND = M0 + 35
    CF = lambda a, b: f'SUM({cCF}{M0+a}:{cCF}{M0+b})'
    # ---- KPIs
    section(ws, 'A12', 'Results (live)')
    hdr(ws, 13, 1, ['Metric', 'Value'])
    b3 = P0 + IDXP['B3']
    k_ = [('Total error cut (pass mark ≥ 75%)', f'={CUTc}', '0.0%', True), ('DPMO after (stretch ≤ 500)', f'=DPMO0*(1-{CUTc})', '#,##0', True),
          ('Sigma level after', f'=NORMSINV(1-DPU0*(1-{CUTc})/OPP)+1.5', '0.00', False), ('% invoices with ≥ 1 error after', f'=1-EXP(-DPU0*(1-{CUTc}))', '0.0%', False),
          ('Meets ≥ 75% error cut?', f'=IF({CUTc}>=0.75,"Yes","No")', None, False), ('Meets ≤ 500 DPMO stretch?', f'=IF(DPMO0*(1-{CUTc})<=500,"Yes","No")', None, False),
          ('Touch time per invoice incl. exceptions (min)', f'=MAX(4,BASE_TOUCH-EXC_SHARE*EXC_MIN+K{P1+1})+EXC_SHARE*(1-{CUTc})*EXC_MIN', '0.0', False),
          ('Analyst-equivalents freed (average day)', f"=(BASE_TOUCH-B20)*'Baseline'!B7/60/8", '0.0', False),
          ('Capex (₹ L)', f'=G{P1+1}', '0.0', True), ('Opex per year (₹ L)', f'=H{P1+1}', '0.0', False),
          ('Benefit per year at full run-rate (₹ L)', f'=COPQ*{CUTc}+{BOTHER}*{CAL}*$D${b3}', '0.0', False), ('Net benefit per year (₹ L)', '=B24-B23', '0.0', False),
          ('Simple payback, months (capex ÷ year-1 net, benefits from day 1)', '=IF(B25>0,B22/B25*12,"No payback")', '0.0', False),
          ('Simple NPV 3 yrs, ₹ L (benefits from day 1)', '=-B22+B25/(1+DISC)+B25/(1+DISC)^2+B25/(1+DISC)^3', '0.0', False),
          ('NPV 3 yrs, ₹ L, WITH build-time ramp (each part from its go-live)', f'=-B22+{CF(0,11)}/(1+DISC)+{CF(12,23)}/(1+DISC)^2+{CF(24,35)}/(1+DISC)^3', '0.0', True),
          ('Payback from project start, months (with ramp)', f'=IF({cCUM}{MEND}>=0,COUNTIF({cCUM}{M0}:{cCUM}{MEND},"<0")+1,"> 36")', '0', True),
          ('Last part goes live (month)', f'=SUMPRODUCT(MAX(I{P0}:I{P1}*D{P0}:D{P1}))', '0.0', False),
          ('First impact (month)', f'=SUMPRODUCT(MIN(I{P0}:I{P1}+(1-D{P0}:D{P1})*99))', '0.0', False),
          ('Parts selected', f'=D{P1+1}', '0', False),
          ('Parts outside steps 1–9 scope', f"=SUMPRODUCT((D{P0}:D{P1}=1)*('Parts Library'!J{PROW0}:J{PROW1}=\"Yes\"))", '0', False),
          ('Parts with legal risk', f"=SUMPRODUCT((D{P0}:D{P1}=1)*('Parts Library'!K{PROW0}:K{PROW1}=\"Yes\"))", '0', False)]
    for k, (lab, fm, nf, bold) in enumerate(k_):
        r = 14 + k
        put(ws, f'A{r}', lab, b=bold, sz=9 if not bold else 10)
        calc(ws, f'B{r}', fm, nf, bold, F_KPI if bold else None)
    ws.conditional_formatting.add('B18:B19', CellIsRule(operator='equal', formula=['"Yes"'], fill=PatternFill('solid', fgColor='C6EFCE'), font=Font(name=FONT, bold=True, color='006100')))
    ws.conditional_formatting.add('B18:B19', CellIsRule(operator='equal', formula=['"No"'], fill=PatternFill('solid', fgColor='FFC7CE'), font=Font(name=FONT, bold=True, color='9C0006')))
    ws.freeze_panes = 'A4'
    return dict(cut='B14', dpmo='B15', touch='B20', fte='B21', capex='B22', opex='B23', ben='B24', net='B25', pb='B26', npv='B27', npvr='B28', pbr='B29', first='B31', P0=P0, CUT=CUTc, M0=M0, cCUT=cCUT)

HY = wb.create_sheet('Hybrid H')
HREF = build_option(HY, 'Hybrid H · the stress-tested best mix of A, B and C',
                    'Keeps A1–A4, all of B and C2; drops A5, C1, C3, C4, C5. Chosen from 25,599 valid combinations of the 15 parts (see Combination Search). Every number below is a live formula.', 'H')
# keep/drop reasons on the right of KPIs
section(HY, 'D12', 'Why each part is in or out')
hdr(HY, 13, 4, ['Part', 'In H?', 'Reason'], h=24); HY.merge_cells('F13:K13')
why = {'A1': 'Core: removing it drops the cut to 72% and P(≥75% cut) to 0%', 'A2': 'Biggest quality lever after A1; P(≥75% cut) 55% → 12% without it',
       'A3': 'Rules engine: +2.9 pts cut for ₹7 L', 'A4': 'Cheap structured data for large vendors (₹4 L)', 'A5': 'Dropped: +1.2 pts cut for ₹7 L capex; keep 24 h approval SLA as a B5 policy',
       'B1': 'Low-cost completeness check while A builds', 'B2': 'Template fixes line items at source', 'B3': 'Calendar protects month-end and the downside (₹12.7 L/yr saving)',
       'B4': 'Flex pool for the last 7 days', 'B5': 'Daily huddle + approval SLA policy', 'C1': 'Dropped: ₹8 L/yr coordinators for +1.9 pts cut',
       'C2': 'Kept: ₹1 L; makes A2\'s 3-way match reliable (GRN on time)', 'C3': 'Dropped: outside scope; do a one-off master clean-up while setting up A3',
       'C4': 'Dropped: ₹3 L/yr early-payment cost for +1.1 pts', 'C5': 'Dropped: MSME legal risk for +0.8 pts'}
for i, p in enumerate(PARTS):
    r = 14 + i; inn = PRES['Hybrid H'][i]
    put(HY, f'D{r}', p['code'], b=True, al='center'); put(HY, f'E{r}', 'Keep' if inn else 'Drop', al='center', fill=F_KEEP if inn else F_DROP, b=True)
    put(HY, f'F{r}', why[p['code']], sz=9); HY.merge_cells(f'F{r}:K{r}')
for _k, _p in (('A', 'A'), ('B', 'B'), ('C', 'C'), ('A+B+C', 'A+B+C')):
    _ws = wb.create_sheet(f'Calc {_k}')
    build_option(_ws, f'Calc {_k} · option {_p} rebuilt from its parts (helper)', 'Hidden helper: gives the build-time NPV and payback used on the solution, Combined and Comparison sheets.', 'H', fixed=_p)
    _ws.sheet_state = 'hidden'
PB_ = wb.create_sheet('Part Builder')
BREF = build_option(PB_, 'Part Builder · try any mix of the 15 parts', 'Pick a preset (A, B, C, A+B+C, Hybrid H, H+, H-lite) or "Custom", then change the stress settings to see how the option holds up.', 'B')

# ============================================================ COMPARISON: Hybrid H column
hy = lambda c: f"='Hybrid H'!{c}"
for row, key in zip(range(5, 17), ['cut', 'dpmo', 'touch', 'fte', None, 'capex', 'opex', 'ben', 'pbr', 'npvr', None, 'first']):
    if key: link(CP, f'F{row}', hy(HREF[key]), CP[f'B{row}'].number_format)
    elif row == 9: put(CP, f'F{row}', ST['Hybrid H (recommended)']['des_base'][0], '0%', al='center')
    else: put(CP, f'F{row}', '–', al='center')
section(CP, 'A33', 'Simple basis, for reference only (benefits assumed from day 1)')
hdr(CP, 34, 1, ['Metric', 'A Digital Front Door', 'B Lean Flow', 'C Supplier Partnership', 'A + B + C', 'Hybrid H'])
put(CP, 'A35', 'Simple NPV 3 yrs (₹ L)', b=True); put(CP, 'A36', 'Simple payback (months)', b=True)
for col, s_ in zip('BCD', 'ABC'):
    link(CP, f'{col}35', f"='{SOLREF[s_]['sheet']}'!{SOLREF[s_]['NPV']}", '0.0'); link(CP, f'{col}36', f"='{SOLREF[s_]['sheet']}'!{SOLREF[s_]['PB']}", '0.0')
link(CP, 'E35', f"={COMB['npv']}", '0.0'); link(CP, 'E36', f"={COMB['pb']}", '0.0'); link(CP, 'F35', hy(HREF['npv']), '0.0'); link(CP, 'F36', hy(HREF['pb']), '0.0')
put(CP, 'A37', 'The deck and all headline rows above use build time. These simple figures understate payback because nothing pays back before go-live.', sz=8, c=GREY, border=False, i=True)
# decision matrix moved down? (rows 20-28 already used) – note only
put(CP, 'A31', 'Hybrid H is not scored in the original decision matrix; it was selected with the stress-test robustness score (see Combination Search).', sz=9, i=True, border=False)
CP.merge_cells('A31:E31')

# ============================================================ STRESS TEST (values from the Python model)
SS = wb.create_sheet('Stress Test')
title(SS, 'Stress test · how each option holds up', 'Values pasted from the Python stress model (component model + discrete-event queue simulation, 200 runs per test + 20,000-run Monte Carlo). NPVs include build time (each part pays from its go-live).')
COLS = ['A (original)', 'B (original)', 'C (original)', 'A+B+C (original)', 'Hybrid H (recommended)', 'Hybrid H+ (quality variant)', 'H-lite (money variant)']
HD = ['A', 'B', 'C', 'A+B+C', 'Hybrid H', 'H+ (quality)', 'H-lite']
widths(SS, {'A': 46, 'B': 14, 'C': 14, 'D': 14, 'E': 14, 'F': 14, 'G': 14, 'H': 14, 'I': 40})
section(SS, 'A4', '1. Comparison across all stress tests')
hdr(SS, 5, 1, ['Stress factor / metric'] + HD + ['Better is'], h=30)
def g(o, *ks):
    x = o
    for k in ks: x = x[k]
    return x
rows_ = [('BASE CASE', None, None, None),
         ('Capex (₹ L)', lambda o: o['base']['capex'], '0.0', 'lower'), ('Opex per year (₹ L)', lambda o: o['base']['opex'], '0.0', 'lower'),
         ('Error cut (pass mark ≥ 75%)', lambda o: o['base']['cut'], '0.0%', 'higher'), ('DPMO (stretch ≤ 500)', lambda o: o['base']['dpmo'], '#,##0', 'lower'),
         ('NPV 3 yrs with build time (₹ L)', lambda o: o['base']['npv'], '0.0', 'higher'), ('Payback from project start (months)', lambda o: o['base']['payback'], '0', 'lower'),
         ('Invoices waiting 2+ days (simulation)', lambda o: o['des_base'][0], '0.0%', 'lower'),
         ('SINGLE-FACTOR STRESS TESTS', None, None, None),
         ('1. Effectiveness 30% lower: error cut', lambda o: o['eff07']['cut'], '0.0%', 'higher'), ('    … NPV (₹ L)', lambda o: o['eff07']['npv'], '0.0', 'higher'),
         ('2. Parts fully overlap: error cut', lambda o: o['overlap0']['cut'], '0.0%', 'higher'), ('    … DPMO', lambda o: o['overlap0']['dpmo'], '#,##0', 'lower'),
         ('3. Penalty trigger 2% (COPQ ₹52.5 L): NPV', lambda o: o['copq52']['npv'], '0.0', 'higher'),
         ('4. Capex ×2 and opex ×1.5: NPV', lambda o: o['cost']['npv'], '0.0', 'higher'),
         ('5. 3-month go-live delay: NPV', lambda o: o['delay3']['npv'], '0.0', 'higher'),
         ('6. GRN data late (A2 at 70% on PO mismatch): error cut', lambda o: o['grn']['cut'], '0.0%', 'higher'),
         ('7. Vendors ignore month-end calendar: NPV', lambda o: o['cal0']['npv'], '0.0', 'higher'), ('    … waiting 2+ days', lambda o: o['des_cal0'][0], '0.0%', 'lower'),
         ('8. "No PO, no pay" blocked legally: error cut', lambda o: o['legal0']['cut'], '0.0%', 'higher'),
         ('9. Volume +20%: waiting 2+ days', lambda o: o['des_v12'][0], '0.0%', 'lower'), ('10. Volume +40%: waiting 2+ days', lambda o: o['des_v14'][0], '0.0%', 'lower'),
         ('11. 2 analysts absent: waiting 2+ days', lambda o: o['des_s16'][0], '0.0%', 'lower'), ('12. Volume +20% and 2 absent: waiting 2+ days', lambda o: o['des_v12s16'][0], '0.0%', 'lower'),
         ('BREAK-EVEN POINTS', None, None, None),
         ('Effectiveness needed for ≥ 75% cut (× assumed)', lambda o: o['m75'] if o['m75'] is not None else 'never', '0.00', 'lower'),
         ('Effectiveness needed for DPMO ≤ 500 (× assumed)', lambda o: o['m500'] if o['m500'] is not None else 'never', '0.00', 'lower'),
         ('Effectiveness needed for NPV ≥ 0 (× assumed)', lambda o: o['mnpv'], '0.00', 'lower'),
         ('Break-even cost of poor quality (₹ L/yr)', lambda o: o['copq_be'], '0.0', 'lower'),
         ('JOINT MONTE CARLO (20,000 runs, all risks at once)', None, None, None),
         ('P(NPV < 0)', lambda o: o['mc']['pneg'], '0.0%', 'lower'), ('NPV – worst 5% (₹ L)', lambda o: o['mc']['npv5'], '0.0', 'higher'),
         ('NPV – median (₹ L)', lambda o: o['mc']['npv50'], '0.0', 'higher'), ('NPV – best 5% (₹ L)', lambda o: o['mc']['npv95'], '0.0', 'higher'),
         ('P(error cut ≥ 75%)', lambda o: o['mc']['p75'], '0.0%', 'higher'), ('P(DPMO ≤ 500)', lambda o: o['mc']['p500'], '0.0%', 'higher'),
         ('Median DPMO', lambda o: o['mc']['dpmo50'], '#,##0', 'lower'), ('DPMO – worst 10%', lambda o: o['mc']['dpmo90'], '#,##0', 'lower'),
         ('Median payback (months)', lambda o: o['mc']['pb50'], '0', 'lower'), ('Payback – worst 10% (months; 99 = not within 5 yrs)', lambda o: o['mc']['pb90'], '0', 'lower'),
         ('P(all parts live within 6-month window)', lambda o: o['mc']['ontime'], '0%', 'higher'),
         ('DELIVERY RISK', None, None, None),
         ('Parts outside CFO scope (steps 1–9)', lambda o: o['flags']['outscope'], '0', 'lower'), ('Parts with legal risk', lambda o: o['flags']['legal'], '0', 'lower'),
         ('Last part goes live (month, no delay)', lambda o: o['flags']['lastmonth'], '0.0', 'lower')]
r = 6; first_data = None; ROWPOS = {}
for lab, fn, nf, better in rows_:
    if fn is None:
        put(SS, f'A{r}', lab, b=True, fill=F_SUB, c=NAVY)
        for j in range(len(COLS) + 1): put(SS, f'{CL(2+j)}{r}', None, fill=F_SUB)
        r += 1; continue
    put(SS, f'A{r}', lab, sz=9, b=not lab.startswith('    '))
    vals = []
    for j, cn in enumerate(COLS):
        v = fn(ST[cn]); vals.append(v)
        put(SS, f'{CL(2+j)}{r}', v, nf if not isinstance(v, str) else None, al='center', sz=9, b=(j == 4), fill=(F_KPI if j == 4 else None))
    put(SS, f'I{r}', better, sz=8, c=GREY, al='center'); ROWPOS[lab] = r
    num = [v for v in vals[:5] if not isinstance(v, str)]
    if num and len(set(num)) > 1 and better in ('higher', 'lower'):
        best = max(num) if better == 'higher' else min(num)
        for j, v in enumerate(vals[:5]):
            if not isinstance(v, str) and abs(v - best) < 1e-9: SS[f'{CL(2+j)}{r}'].font = f(9, True, '006100')
    r += 1
for _lab, _src, _nf in (('Error cut (pass mark ≥ 75%)', 5, '0.0%'), ('DPMO (stretch ≤ 500)', 6, '#,##0'), ('NPV 3 yrs with build time (₹ L)', 14, '0.0'), ('Payback from project start (months)', 13, '0')):
    for _col in 'BCDEF': link(SS, f'{_col}{ROWPOS[_lab]}', f"='Comparison'!{_col}{_src}", _nf)
put(SS, f'A{r}', 'Green = best of A, B, C, A+B+C and Hybrid H on that row. H+ and H-lite shown for reference. Base error cut, DPMO, NPV and payback for A–H link to the Comparison sheet (live models); the stress model reproduces them to within rounding.', sz=8, c=GREY, i=True, border=False); r += 2
# --- flags with status for Hybrid H
section(SS, f'A{r}', '2. Where the original solutions fail – and what Hybrid H changes'); r += 1
hdr(SS, r, 1, ['Flag', 'Severity', 'Evidence', '', '', '', 'Status in Hybrid H', '', ''], h=24)
SS.merge_cells(f'C{r}:F{r}'); SS.merge_cells(f'G{r}:I{r}'); r += 1
FLAGS = [('Combined model misses DPMO ≤ 500', 'Critical', 'A+B+C model 515 DPMO; needs ≥ 101% of assumed effectiveness; 1.3% chance in Monte Carlo', 'Reframed – ≤ 500 DPMO is a team-set stretch (not in the brief). H clears the ≥ 75% pass mark (90% cut, 1,026 DPMO); gate G3 escalates to H+ (765 DPMO). The audit\'s 435 is a point estimate whose range (308–598) crosses 500.'),
         ('Verified audit does not prove the target or our solution', 'Critical', '95% CI for post DPMO ≈ 308–598 (crosses 500); P(true DPMO > 500) ≈ 22%; Improve data is what the company did', 'Wording fixed – decks now say "consistent with" and show the exact 308–598 range; the 22% risk itself remains'),
         ('Paybacks ignored build time', 'Critical', 'A+B+C simple payback 8.5 months becomes 11 months once each part pays only after go-live', 'Fixed – every headline payback and NPV in the model and both decks now counts build time'),
         ('B alone beats A+B+C financially', 'Critical', 'A+B+C loses money in 25.6% of runs; B never does', 'Fixed – H: NPV ₹162 L, P(NPV<0) 1.8%, worst-5% NPV +₹15 L'),
         ('A only just clears the 75% cut', 'Major', 'A needs ≥ 89% of assumed effectiveness; P(≥75%) 29.4%', 'Improved – H needs ≥ 66%; P(≥75%) 54.5%'),
         ('Effectiveness rates are team judgement', 'Major', 'At 70% of assumed effectiveness A+B+C reaches only 1,550 DPMO', 'Still open – H stays ≥ 75% cut down to 66% effectiveness'),
         ('Savings rest on assumed penalties', 'Major', '83% of COPQ is "up to ₹5,000" penalties at a 5% trigger; A+B+C breaks even at ₹46 L', 'Improved – H breaks even at ₹25 L COPQ'),
         ('B collapses with volume growth', 'Major', 'B: 6% wait 2+ days at +20% volume, 35% at +40% (200 runs)', 'Fixed – H: 0% in all 200 runs, even at +40% or 2 analysts absent'),
         ('C alone leaves the backlog', 'Major', 'C: 8% wait 2+ days today, 29% at +20% volume (200 runs)', 'Fixed – H uses A1–A4 capacity'),
         ('Post-improvement targets still missed', 'Major', '2.5% invoices with an error vs ≤ 2%; peak days 2.0× normal error rate', 'Still open – needs SPC + Kaizen on peak days'),
         ('Scope creep beyond steps 1–9', 'Major', 'C1, C2, C3 rely on procurement / site teams', 'Improved – only C2 remains (needs a Procurement sponsor)'),
         ('6-month window is tight', 'Major', 'A2 needs 5 months; 25% chance all parts live by month 6 with delays', 'Managed – gates G1–G3 (Hybrid Plan); Phase 1 (B, A1, A4, C2) live by month 3 gives a 77% cut by December'),
         ('Cost overruns hit the combination hard', 'Major', 'A+B+C NPV ₹33 L at capex ×2 / opex ×1.5', 'Improved – H keeps ₹99 L'),
         ('Decision-matrix ranking is weight-sensitive', 'Major', 'A overtakes B if impact weight ≥ ~46%', 'Replaced by robustness score: H ranks #1 under balanced and ops-first, #5 quality-first, #20 of 524 money-first; A1 + all of B in every winner'),
         ('Peak-day definition drives the peak benefit', 'Major', 'Peak-excess share 12–28% for thresholds 350–500 invoices', 'Still open – affects B3/B4 benefit'),
         ('MSME legal risk of "No PO, no pay"', 'Minor', 'Returning invoices may clash with the 45-day MSME payment rule', 'Removed – C5 not in H'),
         ('Benefit above total COPQ', 'Minor', 'Benefit includes ₹12.7 L/yr staffing saving from the calendar', 'Explained – H benefit ₹107.5 L = penalties ₹79.0 L + rework ₹15.8 L + calendar ₹12.7 L (Hybrid Plan bridge)'),
         ('Triage desk under-staffed at month-end', 'Major', '880 invoices × 2 min = 29 analyst-hours vs 16 h from 2 analysts', 'Fixed – 3 analysts at month-end from the flex pool (~25 h after the calendar); queue model already charges triage time to the whole pool'),
         ('Error-cut timeline shown one month early', 'Major', 'Decks put the first month of benefit in Sep, a baseline month', 'Fixed – Sep = build only; first gains Oct (50%, all of B), 77% by Dec, 90% by Feb; verify Mar–Apr'),
         ('Queue results rested on 10 runs', 'Minor', 'Small number of replications per stress case', 'Fixed – 200 runs per test; H shows no 2-day waits in any run, also with calendar volume piled into the 5 pre-peak days'),
         ('Standalone benefits do not add up', 'Minor', 'A + B + C standalone ₹209 L vs combined ₹113 L', 'Explained – the same errors can only be removed once (bridge on Hybrid Plan sheet)'),
         ('COPQ bases mixed (35% case vs 46% Poisson)', 'Minor', '₹4 L residual COPQ uses the case 35% basis', 'Explained – ₹105 L × (1 − 95.9%) = ₹4.3 L; Poisson basis ₹133 L → ₹7 L (Baseline sheet)'),
         ('25-Aug exclusion needs evidence', 'Minor', 'Deck quoted ~46 expected errors without the basis', 'Fixed – 46 at the peak-day rate, 41 at average; P(0) < 10⁻¹⁷; raw 10,227 and adjusted 10,648 DPMO both reported'),
         ('Ppk-equivalent overstated', 'Minor', 'DPMO-derived sigma is not a capability study', 'Removed from the deck – sigma level only, labelled DPMO-based'),
         ('Before/after design', 'Minor', 'No control group; April audit ends 28-Apr; March is FY-end', 'Still open – state as a limitation')]
for fl, sev, ev, stt in FLAGS:
    put(SS, f'A{r}', fl, sz=9, b=True, wrap=True)
    put(SS, f'B{r}', sev, sz=9, al='center', b=True, fill=PatternFill('solid', fgColor={'Critical': 'FFC7CE', 'Major': 'FFEB9C', 'Minor': 'E2EFDA'}[sev]))
    put(SS, f'C{r}', ev, sz=9, wrap=True); SS.merge_cells(f'C{r}:F{r}')
    st_fill = F_KEEP if stt.startswith(('Fixed', 'Removed', 'Improved', 'Replaced')) else F_DROP
    put(SS, f'G{r}', stt, sz=9, wrap=True, fill=st_fill); SS.merge_cells(f'G{r}:I{r}')
    SS.row_dimensions[r].height = 36; r += 1
r += 1
section(SS, f'A{r}', '3. How the stress test was run'); r += 1
meth = ['Component model: A, B, C split into 15 parts (Parts Library); within a solution parts act independently; across solutions an overlap setting blends independent and fully-overlapping effects.',
        'Money: each part pays back from the month after it goes live; opex from month 1; 3-year NPV at 12%; capex in month 0.',
        'Capacity: discrete-event simulation of the 52 baseline days (FIFO, 18 analysts, 3 h overtime on peak days, flex pool if B4, calendar levelling if B3), 200 runs per test (seeds 0–199). Also re-run with the calendar\'s moved volume piled into the 5 working days before month-end: H, H+ and A+B+C still show no 2-day waits.',
        'Monte Carlo (20,000 runs): effectiveness ×0.5–1.1; overlap 0–1; COPQ ₹52.5–158 L (triangular, mode 105); capex ×1–1.8; opex ×1–1.5; delay 0–4 months (scaled by build length); calendar compliance 40–100%; GRN factor 0.7–1.0 when C2 absent; 30% chance C5 is blocked.',
        'Earlier quick stress test (same day) used one go-live date for all parts of a combination; this version staggers go-lives, so A+B+C numbers differ slightly.']
for t in meth:
    put(SS, f'A{r}', t, sz=9, wrap=True, border=False); SS.merge_cells(f'A{r}:I{r}'); SS.row_dimensions[r].height = 30; r += 1
# charts
CR = r + 1
section(SS, f'A{CR}', 'Chart data'); hdr(SS, CR + 1, 1, ['Option', 'P(NPV < 0)', 'P(cut ≥ 75%)', 'Worst-5% NPV'])
for j, (cn, hd) in enumerate(zip(COLS, HD)):
    rr = CR + 2 + j
    put(SS, f'A{rr}', hd); put(SS, f'B{rr}', ST[cn]['mc']['pneg'] * 100, '0.0', al='center'); put(SS, f'C{rr}', ST[cn]['mc']['p75'] * 100, '0.0', al='center')
    put(SS, f'D{rr}', ST[cn]['mc']['npv5'], '0.0', al='center')
def sbar(t, col, anchor, color):
    ch = BarChart(); ch.type = 'bar'; ch.gapWidth = 50; ch.title = t; ch.legend = None; ch.width = 14; ch.height = 7
    ch.add_data(Reference(SS, min_col=col, min_row=CR + 2, max_row=CR + 1 + len(COLS)), titles_from_data=False)
    ch.set_categories(Reference(SS, min_col=1, min_row=CR + 2, max_row=CR + 1 + len(COLS)))
    ch.series[0].graphicalProperties.solidFill = color; ch.series[0].graphicalProperties.line.solidFill = color
    ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True; ch.dataLabels.showSerName = False; ch.dataLabels.showCatName = False; ch.dataLabels.showLegendKey = False
    ch.x_axis.delete = False; ch.y_axis.delete = False; ch.x_axis.scaling.orientation = 'maxMin'
    SS.add_chart(ch, anchor)
sbar('Chance of losing money, % of runs', 2, f'F{CR}', 'C0392B'); sbar('Chance of ≥ 75% error cut, % of runs', 3, f'F{CR+16}', '2E8B57')
SS.freeze_panes = 'B6'

# ============================================================ COMBINATION SEARCH
CS = wb.create_sheet('Combination Search')
title(CS, 'Combination search · finding the best mix of parts', f'All {2**NPARTS - 1:,} combinations of the 15 parts were checked; {PJ["n_valid"]:,} are valid (A2/A3/A5 need A1 or A4). Values pasted from the Python search.')
widths(CS, {'A': 62, 'B': 12, 'C': 12, 'D': 12, 'E': 12, 'F': 12, 'G': 12, 'H': 12, 'I': 12, 'J': 12, 'K': 12, 'L': 12, 'M': 10, 'N': 10})
section(CS, 'A4', '1. Method')
steps = [f'Stage 1 – every valid combination ({PJ["n_valid"]:,}) scored under a base, a moderate-stress and a harsh-stress scenario, plus a queue model at +20% volume with 2 analysts absent.',
         f'Stage 2 – {PJ["n_short"]} shortlisted (Pareto-efficient + top ranks) run through 3,000 Monte Carlo draws each and scored with a robustness score.',
         'Stage 3 – finalists and the original options re-run with 20,000 draws, the discrete-event simulation and every single-factor stress test (Stress Test sheet).',
         'Order of building only changes when benefits start; that is captured by each part\'s go-live month, so combinations (not orderings) are what is searched.']
for k, t in enumerate(steps): put(CS, f'A{5+k}', t, sz=9, wrap=True, border=False); CS.merge_cells(f'A{5+k}:L{5+k}'); CS.row_dimensions[5 + k].height = 26
section(CS, 'A10', '2. Robustness score weights (judgement)')
hdr(CS, 11, 1, ['Component', 'Balanced (main)', 'Quality-first', 'Money-first', 'Ops-first'])
WTS = [('P(error cut ≥ 75%)', .25, .35, .15, .2), ('Median DPMO (scaled 10,648 → 500)', .15, .25, .05, .1), ('P(NPV ≥ 0)', .2, .1, .35, .15), ('Worst-5% NPV (scaled −50 → +100)', .1, .05, .2, .05),
       ('Capacity: waiting 2+ days at +20% volume & 2 absent', .15, .1, .1, .3), ('All parts live within 6 months', .1, .1, .1, .15), ('Scope & legal (out-of-scope + legal parts)', .05, .05, .05, .05)]
for k, row in enumerate(WTS):
    put(CS, f'A{12+k}', row[0], sz=9)
    for j, v in enumerate(row[1:]): put(CS, f'{CL(2+j)}{12+k}', v, '0%', al='center', sz=9)
put(CS, 'A19', 'Winner', b=True, fill=F_TOT)
for j, k in enumerate(['Balanced (main)', 'Quality-first', 'Money-first', 'Ops-first']):
    put(CS, f'{CL(2+j)}19', PJ['winners'][k].replace(' + ', '+'), sz=8, wrap=True, fill=F_TOT, b=True)
CS.row_dimensions[19].height = 64
section(CS, 'A21', '3. Top 15 combinations (balanced score, 3,000-run Monte Carlo)')
hdr(CS, 22, 1, ['Parts', 'Balanced score', 'Quality-first', 'Money-first', 'Ops-first', 'P(cut ≥ 75%)', 'Median DPMO', 'P(NPV < 0)', 'Worst-5% NPV', 'Median NPV', 'Median payback (mo)', 'Capex ₹ L', 'Opex ₹ L'], h=34)
def put_combo(rw, d, hl=False):
    put(CS, f'A{rw}', d['label'], sz=9, b=hl, fill=F_KEEP if hl else None)
    for j, (k, nf) in enumerate([('Balanced (main)', '0.0'), ('Quality-first', '0.0'), ('Money-first', '0.0'), ('Ops-first', '0.0'), ('p75', '0%'), ('med_dpmo', '#,##0'),
                                 ('pneg', '0.0%'), ('npv5', '0'), ('npv50', '0'), ('pb50', '0'), ('capex', '0.0'), ('opex', '0.0')]):
        put(CS, f'{CL(2+j)}{rw}', d[k], nf, al='center', sz=9, b=hl, fill=F_KEEP if hl else None)
for k, d in enumerate(PJ['top']): put_combo(23 + k, d, k == 0)
r = 23 + len(PJ['top']) + 1
put(CS, f'A{r}', 'Original options on the same basis', b=True, c=NAVY); r += 1
for d in PJ['orig_scores']: put_combo(r, d); r += 1
r += 1
section(CS, f'A{r}', '4. Is Hybrid H a local optimum? Effect of adding or removing one part'); r += 1
hdr(CS, r, 1, ['Change to Hybrid H', 'Score change', 'Base error cut', 'Base NPV ₹ L', 'P(cut ≥ 75%)', 'Median DPMO', 'P(NPV < 0)', 'Worst-5% NPV'], h=30); r += 1
Hm = MG['H']
put(CS, f'A{r}', 'Hybrid H as chosen', b=True, fill=F_KEEP)
for j, (v, nf) in enumerate([(0, '+0.0;-0.0;0.0'), (Hm['cut_b'], '0.0%'), (Hm['npv_b'], '0'), (Hm['p75'], '0%'), (Hm['med_dpmo'], '#,##0'), (Hm['pneg'], '0.0%'), (Hm['npv5'], '0')]):
    put(CS, f'{CL(2+j)}{r}', v, nf, al='center', sz=9, b=True, fill=F_KEEP)
r += 1
for d in MG['rows']:
    put(CS, f'A{r}', d['change'], sz=9)
    for j, (v, nf) in enumerate([(d['dscore'], '+0.0;-0.0;0.0'), (d['cut_b'], '0.0%'), (d['npv_b'], '0'), (d['p75'], '0%'), (d['med_dpmo'], '#,##0'), (d['pneg'], '0.0%'), (d['npv5'], '0')]):
        c = put(CS, f'{CL(2+j)}{r}', v, nf, al='center', sz=9)
    CS[f'B{r}'].font = f(9, True, 'C00000' if d['dscore'] < 0 else '006100'); r += 1
put(CS, f'A{r}', 'Every single change lowers the score, so Hybrid H is a local optimum. Closest calls: add C4 (−0.3) and add A5 (−0.6).', sz=9, i=True, border=False)
CS.freeze_panes = 'A4'


# ============================================================ HYBRID PLAN (timeline, gates, benefit bridge, weight sensitivity)
HPS = wb.create_sheet('Hybrid Plan')
title(HPS, 'Hybrid Plan · timeline, gates and benefit bridge', 'Live links to the Hybrid H, Comparison and Inputs sheets; stress values pasted from the Python model. Month 1 = September 2026 (build only, nothing live in the baseline window).')
widths(HPS, {'A': 46, 'B': 16, 'C': 30, 'D': 16, 'E': 34, 'F': 14})
hh = lambda c: f"'Hybrid H'!{c}"
section(HPS, 'A4', '1. Model error cut by month (Hybrid H)')
hdr(HPS, 5, 1, ['Month', 'Model month', 'Parts live (go-live before the month)', 'Error cut'])
_live = ['none – build only', 'B1–B5', '+ A4, C2', '+ A1 (Phase 1 complete)', '+ A3', '+ A2 (all parts live)', 'all · supplied audit month 1', 'all · supplied audit month 2']
for k, (mn, lv) in enumerate(zip(['Sep 2026', 'Oct 2026', 'Nov 2026', 'Dec 2026', 'Jan 2027', 'Feb 2027', 'Mar 2027', 'Apr 2027'], _live)):
    put(HPS, f'A{6+k}', mn, b=True); put(HPS, f'B{6+k}', k + 1, '0', al='center'); put(HPS, f'C{6+k}', lv, sz=9)
    link(HPS, f'D{6+k}', '=' + hh(f"{HREF['cCUT']}{HREF['M0'] + k}"), '0.0%')
section(HPS, 'A15', '2. Go / no-go gates (thresholds are team proposals)')
hdr(HPS, 16, 1, ['Gate', 'When', 'Go criteria', 'Model forecast (cut)', 'If not met'])
_gates = [('G1', 'end-Oct (month 2)', 'B live; ≥ 60% of vendors on the calendar; cut ≥ 40%', '=D7', 'Fix B discipline before Phase 1 go-lives'),
          ('G2', 'end-Dec (month 4)', 'Phase 1 live; cut ≥ 70%; portal false rejects ≤ 5%', '=D9', 'Hold A2/A3 go-live; tune rules'),
          ('G3', 'end-Feb (month 6)', 'All parts live; cut ≥ 75% (pass mark)', '=D11', 'Escalate to H+: add A5 + C4'),
          ('Verify', 'Mar–Apr 2027', 'Supplied 2-month audit; SPC on the ≤ 500 DPMO stretch', '=D12', 'Kaizen on top reason; peak-day plan')]
for k, (g_, w_, c_, fc, nm_) in enumerate(_gates):
    rr = 17 + k
    put(HPS, f'A{rr}', g_, b=True); put(HPS, f'B{rr}', w_, sz=9); put(HPS, f'C{rr}', c_, sz=9, wrap=True); link(HPS, f'D{rr}', fc, '0.0%'); put(HPS, f'E{rr}', nm_, sz=9, wrap=True)
    HPS.row_dimensions[rr].height = 30
put(HPS, 'A21', 'Timeline risk: chance all parts are live by end-Feb (month 6) with 0–4 month delays', sz=9); put(HPS, 'B21', ST['Hybrid H (recommended)']['mc']['ontime'], '0%', al='center')
section(HPS, 'A23', '3. Escalation path if G3 fails: Hybrid H → H+ (pasted from the stress model)')
hdr(HPS, 24, 1, ['Metric', 'Hybrid H', 'H+ (adds A5 + C4)', 'Change'])
_esc = [('Capex (₹ L)', 'capex', '0.0'), ('Opex per year (₹ L)', 'opex', '0.0'), ('Error cut', 'cut', '0.0%'), ('DPMO', 'dpmo', '#,##0'), ('NPV incl. build time (₹ L)', 'npv', '0.0')]
for k, (lab, key, nf) in enumerate(_esc):
    rr = 25 + k; b_ = ST['Hybrid H+ (quality variant)']['base'][key]
    put(HPS, f'A{rr}', lab); link(HPS, f'B{rr}', '=' + hh(HREF[{'capex': 'capex', 'opex': 'opex', 'cut': 'cut', 'dpmo': 'dpmo', 'npv': 'npvr'}[key]]), nf); put(HPS, f'C{rr}', b_, nf, al='center'); calc(HPS, f'D{rr}', f'=C{rr}-B{rr}', nf)
put(HPS, 'A30', 'Chance of a loss (20,000 runs)'); put(HPS, 'B30', ST['Hybrid H (recommended)']['mc']['pneg'], '0.0%', al='center'); put(HPS, 'C30', ST['Hybrid H+ (quality variant)']['mc']['pneg'], '0.0%', al='center'); calc(HPS, 'D30', '=C30-B30', '0.0%')
section(HPS, 'A32', "4. Where Hybrid H's benefit comes from (₹ L/yr, live)")
hdr(HPS, 33, 1, ['Component', '₹ L/yr', 'Type of benefit'])
_br = [('GST penalties avoided', "=COPQ_PEN*" + hh(HREF['cut']), 'Expected risk reduction'), ('Exception rework avoided', "=COPQ_RW*" + hh(HREF['cut']), 'Analyst time released (₹350/h)'),
       ('Calendar overtime saving', '=' + hh(HREF['ben']) + '-COPQ*' + hh(HREF['cut']), 'Cash saving (LP staffing model)'), ('Total benefit per year', '=' + hh(HREF['ben']), ''),
       ('Analyst-equivalents freed (not counted above)', '=' + hh(HREF['fte']), 'Released capacity')]
for k, (lab, fm, ty) in enumerate(_br):
    rr = 34 + k; put(HPS, f'A{rr}', lab, b=(k == 3)); link(HPS, f'B{rr}', fm, '0.0', k == 3); put(HPS, f'C{rr}', ty, sz=9)
section(HPS, 'A40', '5. Why standalone benefits do not add up (₹ L/yr)')
hdr(HPS, 41, 1, ['Item', '₹ L/yr', 'Note'])
_sb = [('A standalone', "='Comparison'!B12", ''), ('B standalone', "='Comparison'!C12", 'incl. ₹12.7 L calendar saving'), ('C standalone', "='Comparison'!D12", ''),
       ('Sum of standalone benefits', '=SUM(B42:B44)', 'counts the same errors up to three times'), ('A + B + C combined', "='Comparison'!E12", 'each error removed once'),
       ('Overlap removed', '=B45-B46', ''), ('Hybrid H', "='Comparison'!F12", '10 of the 15 parts')]
for k, (lab, fm, nt) in enumerate(_sb):
    rr = 42 + k; put(HPS, f'A{rr}', lab, b=k in (3, 4, 6)); link(HPS, f'B{rr}', fm, '0.0'); put(HPS, f'C{rr}', nt, sz=9)
section(HPS, 'A50', '6. Business case: A + B + C vs Hybrid H (incl. build time)')
hdr(HPS, 51, 1, ['Metric', 'A + B + C', 'Hybrid H'])
_bc = [('3-yr NPV incl. build time (₹ L)', "='Comparison'!E14", "='Comparison'!F14", '0.0'), ('Payback from project start (months)', "='Comparison'!E13", "='Comparison'!F13", '0'),
       ('Capex (₹ L)', "='Comparison'!E10", "='Comparison'!F10", '0.0'), ('Opex per year (₹ L)', "='Comparison'!E11", "='Comparison'!F11", '0.0'),
       ('Benefit per year (₹ L)', "='Comparison'!E12", "='Comparison'!F12", '0.0'), ('Model DPMO (stretch ≤ 500)', "='Comparison'!E6", "='Comparison'!F6", '#,##0')]
for k, (lab, a_, b_, nf) in enumerate(_bc):
    rr = 52 + k; put(HPS, f'A{rr}', lab); link(HPS, f'B{rr}', a_, nf); link(HPS, f'C{rr}', b_, nf)
for k, (lab, key) in enumerate([('Chance of ≥ 75% cut (20,000 runs)', 'p75'), ('Chance of a loss (20,000 runs)', 'pneg')]):
    rr = 58 + k; put(HPS, f'A{rr}', lab); put(HPS, f'B{rr}', ST['A+B+C (original)']['mc'][key], '0.0%', al='center'); put(HPS, f'C{rr}', ST['Hybrid H (recommended)']['mc'][key], '0.0%', al='center')
section(HPS, 'A61', '7. Does Hybrid H survive a change of weights? (524 shortlisted mixes re-ranked)')
hdr(HPS, 62, 1, ['Weighting', 'Winner', 'Winner score', 'Hybrid H score', 'Hybrid H rank (of 524)'])
for k, (wn, win, ws_, hs_, hr_) in enumerate([('Balanced (main)', 'Hybrid H', 70.8, 70.8, 1), ('Quality-first', 'H+ (adds A5, C4)', 68.2, 67.3, 5),
                                               ('Money-first', 'H-lite (A1 + all of B)', 72.8, 71.3, 20), ('Ops-first', 'Hybrid H', 73.3, 73.3, 1)]):
    rr = 63 + k; put(HPS, f'A{rr}', wn, b=True); put(HPS, f'B{rr}', win, sz=9); put(HPS, f'C{rr}', ws_, '0.0', al='center'); put(HPS, f'D{rr}', hs_, '0.0', al='center'); put(HPS, f'E{rr}', hr_, '0', al='center')
HPS.freeze_panes = 'A4'
put(CS, 'A20', 'Hybrid H rank under this weighting (of 524)', b=True)
for j, v in enumerate([1, 5, 20, 1]): put(CS, f'{CL(2+j)}20', v, '0', al='center', b=True)

# ============================================================ README / ASSUMPTIONS additions
nr = RD.max_row + 1
for a, b in [('New section (Hybrid & stress test)', 'Parts Library (15 parts of A, B, C) · Hybrid H (live model of the recommended mix) · Part Builder (try any mix + stress settings) · Stress Test (all tests, flags, status) · Combination Search (how H was chosen).'),
             ('Hybrid H', 'A1 portal + A2 3-way match + A3 rules engine + A4 IRN feed + all of B + C2 PO/GRN discipline. Capex ₹45 L, opex ₹15 L/yr. Drops A5, C1, C3, C4, C5.'),
             ('Basis note', 'Every headline NPV and payback (Sol A/B/C, Combined, Comparison, Hybrid H, both decks) counts build time: each part pays from the month after it goes live. Simple versions that assume benefits from day 1 are kept, labelled "Simple", for reference only.'),
             ('Hybrid Plan', 'Month-by-month error cut (Sep 2026 = build only), go/no-go gates G1–G3, the H → H+ escalation, the benefit bridge, standalone-vs-combined benefits and the weight sensitivity behind the limit_3 deck.'),
             ('Reproducibility', 'Stress Test and Combination Search values come from the Python stress model and are pasted as values; every other figure in this workbook is a live formula.')]:
    put(RD, f'A{nr}', a, b=True, fill=F_SUB); put(RD, f'B{nr}', b, wrap=True); RD.row_dimensions[nr].height = 32; nr += 1
na = AS.max_row + 1
for a, b, c in [('H1', 'Each solution is split into 5 parts; part effectiveness calibrated so the 5 parts combined reproduce the solution exactly. Split weights, part costs and go-live months are team assumptions.', 'Parts Library'),
                ('H2', 'Across solutions, an overlap setting (1 = independent … 0 = only the strongest lever counts) replaces the plain independence assumption.', 'Hybrid H, Part Builder'),
                ('H3', 'A2/A3/A5 need digital intake (A1 full, A4 alone half). A2\'s PO-mismatch effect needs timely GRN data (C2).', 'Hybrid H, Part Builder'),
                ('H4', 'Each part pays from the month after go-live; opex from month 1; capex in month 0.', 'Hybrid H ramp NPV'),
                ('H5', 'Stress distributions: effectiveness ×0.5–1.1, overlap 0–1, COPQ ₹52.5–158 L, capex ×1–1.8, opex ×1–1.5, delay 0–4 months, calendar compliance 40–100%, GRN 0.7–1.0, 30% chance C5 blocked.', 'Stress Test'),
                ('H6', 'Robustness score weights are judgement; the winner is reported under four weightings.', 'Combination Search'),
                ('H7', 'Timeline: builds start in Sep 2026 with nothing live in the baseline window; go-lives Oct–Jan; the supplied verification audit covers Mar–Apr 2027.', 'Hybrid Plan'),
                ('H8', 'Targets: ≥ 75% cut in errors per invoice is the pass mark; ≤ 500 DPMO is a team-set stretch goal, not part of the case brief.', 'Comparison, Hybrid H'),
                ('H9', 'Gate thresholds (G1 cut ≥ 40%, G2 ≥ 70%, G3 ≥ 75%; ≥ 60% calendar compliance; ≤ 5% false rejects) are team proposals.', 'Hybrid Plan'),
                ('H10', 'Residual COPQ uses the case basis (35% of invoices with errors) scaled by the verified cut; the Poisson basis (≈46%) is shown alongside.', 'Baseline'),
                ('H11', 'Regulation: e-invoicing (IRN) mandatory above ₹5 Cr turnover from 1 Aug 2023 (CBIC Notification 10/2023–Central Tax); MSMED Act 2006 s.15–16 (pay MSMEs within 45 days; interest at 3× RBI bank rate); Income-tax Act s.43B(h).', 'Sol A, Sol C')]:
    put(AS, f'A{na}', a, b=True, al='center'); put(AS, f'B{na}', b, wrap=True); put(AS, f'C{na}', c, sz=9); AS.row_dimensions[na].height = 30; na += 1
for ws in wb.worksheets:
    ws.sheet_properties.tabColor = (NAVY if ws.title.startswith('Sol') else '2E5C8A' if ws.title in ('Comparison', 'Combined A+B+C') else
                                    '2E8B57' if ws.title in ('Hybrid H', 'Hybrid Plan', 'Part Builder', 'Parts Library') else 'C0392B' if ws.title in ('Stress Test', 'Combination Search') else 'A6A6A6')
