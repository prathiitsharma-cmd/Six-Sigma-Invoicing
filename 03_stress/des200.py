"""Stage 3, step 4 - re-run every queue (discrete-event simulation) test with 200 replications instead of 10; also test a concentrated calendar.
Reads : outputs/stage2.pkl and outputs/stage3.json (or the precomputed/ copies), 03_stress/comp.py
Writes: outputs/des200.json (mean + max late share per test); UPDATES outputs/stage3.json in place (des_* entries);
        outputs/stage3_backup_10runs.json keeps the 10-run version when stage3.json came from outputs/
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   The queue simulation is random, so one run is noisy. This script repeats every queue test 200 times (seeds 0-199) and
#   overwrites the 10-run averages in stage3.json with the 200-run averages. This is the slowest script in the project.
#   It is safe to re-run: it only replaces the des_* entries, and does so deterministically.
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
exec(open(paths.STRESS / 'comp.py').read())
import json, shutil, pickle
S = pickle.load(open(paths.find('stage2.pkl'), 'rb'))
def bm(b): return np.array([(int(b) >> k) & 1 for k in range(K)], bool)
CANDS = dict(ORIG)
CANDS['Hybrid H (recommended)'] = bm(S.loc[S['Balanced (main)'].idxmax()].bits)
CANDS['Hybrid H+ (quality variant)'] = bm(S.loc[S['Quality-first'].idxmax()].bits)
CANDS['H-lite (money variant)'] = bm(S.loc[S['Money-first'].idxmax()].bits)
RUNS = 200

def level_conc(d, share=0.25, ndays=5):
    """calendar variant: moved volume lands on the 5 working days just before the month-end window"""
    inv = d.Inv.astype(float).values.copy()
    for m in d.mon.unique():
        idx = np.where((d.mon == m).values)[0]; pk = [i for i in idx if d.dtype.values[i] == 'ME']
        if not pk: continue
        moved = np.clip(inv[pk] - 400, 0, None) * share; inv[pk] -= moved
        pre = [i for i in idx if i < min(pk) and d.dtype.values[i] == 'N'][-ndays:]
        inv[pre] += moved.sum() / len(pre)
    return inv

def run(mask, cut, vm=1.0, staff=18, comp=1.0, conc=False):
    if mask[IDX['B3']]:
        vols = level_conc(M, 0.25 * comp) if conc else level(M, 0.25 * comp)
    else:
        vols = M.Inv.values.astype(float)
    vols = vols * vm
    extra = 4 if mask[IDX['B4']] else 0
    sims = [des(vols, M.peak.values, touch_clean(mask), EXC0 * (1 - cut), k, staff=staff, extra_peak=extra) for k in range(RUNS)]
    late = np.array([s_[1] for s_ in sims]); bl = np.array([s_[2] for s_ in sims])
    return dict(mean=float(late.mean()), max=float(late.max()), runs_with_late=int((late > 0).sum()), backlog=float(bl.mean()), maxday=float(vols.max()))

ST = json.load(open(paths.find('stage3.json')))
if pathlib.Path(paths.find('stage3.json')).parent == paths.OUT:   # packaging glue: only back up a genuine 10-run file
    shutil.copy(paths.find('stage3.json'), paths.out('stage3_backup_10runs.json'))
OUT = {}
for name, mask in CANDS.items():
    c0 = ST[name]['base']['cut']
    o = dict(base=run(mask, c0), cal0=run(mask, c0, comp=0), v12=run(mask, c0, vm=1.2), v14=run(mask, c0, vm=1.4),
             s16=run(mask, c0, staff=16), v12s16=run(mask, c0, vm=1.2, staff=16))
    if mask[IDX['B3']]:
        o['conc'] = run(mask, c0, conc=True); o['conc_v14'] = run(mask, c0, vm=1.4, conc=True)
    OUT[name] = o
    for k, key in (('base', 'des_base'), ('cal0', 'des_cal0'), ('v12', 'des_v12'), ('v14', 'des_v14'), ('s16', 'des_s16'), ('v12s16', 'des_v12s16')):
        if key == 'des_cal0' and not mask[IDX['B3']]:
            ST[name][key] = [o['base']['mean'], o['base']['backlog']]
        else:
            ST[name][key] = [o[k]['mean'], o[k]['backlog']]
    print(name, {k: (round(v['mean'], 4), round(v['max'], 4), v['runs_with_late']) for k, v in o.items()})
json.dump(OUT, open(paths.out('des200.json'), 'w'), indent=1)
json.dump(ST, open(paths.out('stage3.json'), 'w'), indent=1, default=float)
print('saved')
