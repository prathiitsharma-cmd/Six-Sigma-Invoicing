"""Stage 3, step 5 - export the component library and shortlist results for the workbook.
Reads : outputs/stage2.pkl (or precomputed/), 03_stress/comp.py
Writes: outputs/parts.json  (15 components with costs/effects, preset combinations, top-15 table, winners)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Converts the Python component definitions and the stage-2 ranking into a plain JSON file that the Excel build reads to fill
#   the "Parts Library" and "Combination Search" sheets. Note: the count 25,599 valid combinations is typed in here as a constant.
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
exec(open(paths.STRESS / 'comp.py').read())
import json, pickle
S2 = pickle.load(open(paths.find('stage2.pkl'), 'rb'))
def bits_of(m): return int(sum(1 << i for i in range(K) if m[i]))
H = S2.loc[S2['Balanced (main)'].idxmax()]; HP = S2.loc[S2['Quality-first'].idxmax()]; HL = S2.loc[S2['Money-first'].idxmax()]
pres = {'A': ORIG['A (original)'], 'B': ORIG['B (original)'], 'C': ORIG['C (original)'], 'A+B+C': ORIG['A+B+C (original)'],
        'Hybrid H': np.array([(int(H.bits) >> k) & 1 for k in range(K)], bool), 'H+ (quality)': np.array([(int(HP.bits) >> k) & 1 for k in range(K)], bool),
        'H-lite': np.array([(int(HL.bits) >> k) & 1 for k in range(K)], bool)}
parts = [dict(code=CODE[i], sol=GRP[i], name=NAMES[i].split(' ', 1)[1], capex=CAPEX[i], opex=OPEX[i], month=MONTH[i], peak=PFX[i], touch=TOUCHD[i],
              flag=FLAG[i], eff=list(ECOMP[i])) for i in range(K)]
top = S2.sort_values('Balanced (main)', ascending=False).head(15)
cols = ['label', 'Balanced (main)', 'Quality-first', 'Money-first', 'Ops-first', 'p75', 'med_dpmo', 'pneg', 'npv5', 'npv50', 'pb50', 'late_s', 'capex', 'opex']
json.dump(dict(parts=parts, presets={k: [int(x) for x in v] for k, v in pres.items()}, reasons=[short[r] for r in RS], extra=EXTRA_L,
               top=top[cols].to_dict('records'), orig_scores=S2[S2.bits.isin([bits_of(m) for m in ORIG.values()])][cols].to_dict('records'),
               winners={k: S2.loc[S2[k].idxmax(), 'label'] for k in ['Balanced (main)', 'Quality-first', 'Money-first', 'Ops-first']},
               n_valid=25599, n_short=len(S2)), open(paths.out('parts.json'), 'w'), default=float, indent=1)
print('ok', {k: int(v.sum()) for k, v in pres.items()})
