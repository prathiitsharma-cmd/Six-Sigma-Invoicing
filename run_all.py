"""Run the whole pipeline (or part of it) in the right order.

Usage (from the project folder):
    python run_all.py                    # everything, stage 1 to 5 (slow stages included)
    python run_all.py --skip-slow        # skip the slow simulations and use the saved results in precomputed/ instead
    python run_all.py --from 4           # start at stage 4 (earlier results are taken from precomputed/)
    python run_all.py --from 3 --to 4    # only stages 3 and 4
    python run_all.py --list             # show the steps and their expected run times, run nothing
    python run_all.py --no-recalc        # stage 5: do not try the optional LibreOffice step
    python run_all.py --verbose          # show every script's printed output (otherwise it goes to outputs/logs/)

Reads : data/original.xlsx and (for any stage you start after) the files in precomputed/
Writes: outputs/*.json, outputs/*.pkl, outputs/model2.xlsx (and outputs/final_model5.xlsx if LibreOffice is installed)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   Runs each script in order as a separate Python process, times it, checks that the files it should have produced exist,
#   and stops at the first failure. Each script finds its inputs in outputs/ first and precomputed/ second, which is why
#   --from N and --skip-slow work: whatever you do not recompute is read from the shipped copies.
# ---------------------------------------------------------------------------
import argparse, pathlib, subprocess, sys, time

ROOT = pathlib.Path(__file__).resolve().parent
OUT = ROOT / 'outputs'

# (stage, script, is_slow, approx. minutes on a 2020s laptop-class CPU, files it must produce in outputs/)
STEPS = [
    (1, '01_analysis/analysis.py',        False, '0.1',  ['results.json']),
    (2, '02_solutions/solutions_calc.py', False, '0.05', ['sol.json']),
    (3, '03_stress/stage1.py',            True,  '1.5', ['stage1.pkl']),
    (3, '03_stress/stage2.py',            True,  '0.5', ['stage2.pkl']),
    (3, '03_stress/stage3.py',            True,  '0.2', ['stage3.json']),
    (3, '03_stress/des200.py',            True,  '2', ['des200.json', 'stage3.json']),
    (3, '03_stress/dump_parts.py',        False, '0.1',  ['parts.json']),
    (3, '03_stress/stage4b.py',           False, '0.05', ['marginal.json', 'marginal_raw.json']),
    (3, '03_stress/ramp.py',              False, '0.1',  ['ramp_H.json']),
    (4, '04_statistics/laney.py',         False, '0.05', ['laney.json']),
    (4, '04_statistics/uncert.py',        False, '0.05', ['uncert.json']),
    (4, '04_statistics/rule.py',          False, '0.1', ['rule.json']),
    (4, '04_statistics/rule_grid.py',     False, '0.1', ['rule_final.json']),
    (5, '05_workbook/build_workbook.py',  False, '0.2',  ['model2.xlsx']),
]


def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n\n')[0])
    ap.add_argument('--from', dest='first', type=int, default=1, choices=range(1, 6), metavar='N', help='first stage to run (1-5)')
    ap.add_argument('--to', dest='last', type=int, default=5, choices=range(1, 6), metavar='N', help='last stage to run (1-5)')
    ap.add_argument('--skip-slow', action='store_true', help='skip the slow simulations (stage 3 steps 1-4); use precomputed/ results')
    ap.add_argument('--no-recalc', action='store_true', help='stage 5: skip the optional LibreOffice recalculation')
    ap.add_argument('--verbose', action='store_true', help='print each script output instead of logging it')
    ap.add_argument('--list', action='store_true', help='list the steps and exit')
    a = ap.parse_args()
    todo = [s for s in STEPS if a.first <= s[0] <= a.last]
    if a.list:
        for st, script, slow, mins, outs in STEPS:
            print(f'stage {st}  {script:38s} ~{mins:>6} min{"  (slow)" if slow else ""}  -> {", ".join(outs)}')
        return 0
    OUT.mkdir(exist_ok=True); (OUT / 'logs').mkdir(exist_ok=True)
    t_all = time.time()
    for st, script, slow, mins, outs in todo:
        if slow and a.skip_slow:
            print(f'[skip]  stage {st}  {script}  (--skip-slow: using precomputed/{" ".join(outs)})')
            continue
        print(f'[run ]  stage {st}  {script}  (about {mins} min) ...', flush=True)
        cmd = [sys.executable, str(ROOT / script)] + (['--no-recalc'] if (script.endswith('build_workbook.py') and a.no_recalc) else [])
        t0 = time.time()
        log = OUT / 'logs' / (pathlib.Path(script).stem + '.log')
        if a.verbose:
            rc = subprocess.run(cmd, cwd=ROOT).returncode
        else:
            with open(log, 'w') as fh:
                rc = subprocess.run(cmd, cwd=ROOT, stdout=fh, stderr=subprocess.STDOUT).returncode
        dt = time.time() - t0
        missing = [o for o in outs if not (OUT / o).exists()]
        if rc != 0 or missing:
            print(f'[FAIL]  {script}: exit code {rc}, missing outputs {missing}.' + ('' if a.verbose else f' See {log}'))
            return 1
        print(f'[done]  {script}  {dt:.0f} s')
    print(f'All requested steps finished in {time.time() - t_all:.0f} s. Results are in {OUT}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
