"""Workbook extension (exec-ed by build_workbook.py) - "Deck Reconciliation" sheet.
Reads : 05_workbook/recon_spec.py (exec-ed)
Writes: adds a sheet to the workbook object `wb`
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Not run on its own. For each number quoted in the earlier slide deck it writes a live link to the workbook cell holding that
#   number, a rounded copy and a match check. The list of numbers lives in recon_spec.py.
# ---------------------------------------------------------------------------
# ============================================================ DECK RECONCILIATION (every limit_2 deck number -> a live workbook cell, with a rounding check)
import re as _re
exec(open(paths.WORKBOOK / 'recon_spec.py').read())
RC = wb.create_sheet('Deck Reconciliation')
title(RC, 'Deck Reconciliation · every number on the limit_2 deck traced to a cell',
      'Column F links live to the cell named in column E; column G rounds it the way the deck prints it; column I checks it against the printed value. A ✗ means the workbook has moved since the deck was built.')
widths(RC, {'A': 6, 'B': 9, 'C': 30, 'D': 52, 'E': 34, 'F': 13, 'G': 12, 'H': 12, 'I': 8})
put(RC, 'A4', 'Matches', b=True, border=False); put(RC, 'D4', 'Numbers checked', b=True, border=False)
hdr(RC, 6, 1, ['#', 'Slide', 'As printed on the deck', 'What it is', 'Workbook cell', 'Live value', 'Rounded as on the deck', 'Deck value', 'Match'])
def _xl(ref):
    if ref.startswith('U:'): return '=' + UREF[ref[2:]]
    if ref.startswith('P:'): return float(ref[2:])
    if ref.startswith('='):
        def tok(m):
            t = m.group(1)
            if t.startswith('U:'): return UREF[t[2:]]
            sh, c = t.rsplit('!', 1); return f"'{sh}'!{c}"
        return '=' + _re.sub(r'\{([^}]+)\}', tok, ref[1:])
    sh, c = ref.rsplit('!', 1); return f"='{sh}'!{c}"
def _shown(ref):
    if ref.startswith('U:'): return UREF[ref[2:]].replace("'", '')
    if ref.startswith('P:'): return 'pasted: Combination Search log'
    if ref.startswith('='): return 'calc: ' + _re.sub(r'\{U:([^}]+)\}', lambda m: UREF[m.group(1)].replace("'", ''), ref[1:]).replace('{', '').replace('}', '')
    return ref
for i, (sl, txt, what, ref, v, k, dp) in enumerate(E):
    rr = 7 + i
    put(RC, f'A{rr}', i + 1, '0', al='center', sz=8)
    put(RC, f'B{rr}', ', '.join(f'{s:02d}' for s in sl), sz=8, al='center')
    put(RC, f'C{rr}', txt, sz=8); RC[f'C{rr}'].data_type = 's'; put(RC, f'D{rr}', what, sz=8); put(RC, f'E{rr}', _shown(ref), sz=8)
    fx = _xl(ref)
    if isinstance(fx, float): pv_ = put(RC, f'F{rr}', fx, 'General', al='center', sz=8, fill=PatternFill('solid', fgColor='F2F2F2'))
    else: link(RC, f'F{rr}', fx, 'General'); RC[f'F{rr}'].font = f(8, c=GREEN)
    nf = '#,##0' if dp <= 0 else '0.' + '0' * dp
    calc(RC, f'G{rr}', f'=ROUND(F{rr}*{k},{dp})', nf); RC[f'G{rr}'].font = f(8)
    put(RC, f'H{rr}', v, nf, al='center', sz=8)
    calc(RC, f'I{rr}', f'=IF(ABS(G{rr}-H{rr})<0.000001,"✓","✗")', None); RC[f'I{rr}'].font = f(9, True)
last = 7 + len(E) - 1
calc(RC, 'B4', f'=COUNTIF(I7:I{last},"✓")', '0', True); calc(RC, 'E4', f'=COUNTA(I7:I{last})', '0', True)
RC.conditional_formatting.add(f'I7:I{last}', CellIsRule(operator='equal', formula=['"✗"'], fill=PatternFill('solid', fgColor='F8CBAD')))
RC.freeze_panes = 'A7'
RC.auto_filter.ref = f'A6:I{last}'
