"""Stage 5 (OPTIONAL) - let LibreOffice calculate model2.xlsx and store the results inside the file.
Reads : outputs/model2.xlsx  (made by build_workbook.py)
Writes: outputs/final_model5.xlsx  (same workbook, but every formula cell also carries its calculated value)
        outputs/model2_recalculated.xlsx  (LibreOffice's own copy; only used as a source of numbers, can be deleted)
"""
# ---------------------------------------------------------------------------
# WHAT THIS FILE DOES (plain English)
#   You do NOT need this to use the model: Excel recalculates model2.xlsx by itself when you open it.
#   This helper is only for people who want numbers to be visible in programs that do not calculate formulas
#   (file previews, pandas/openpyxl with data_only=True, some phone viewers).
#   It looks for LibreOffice ("soffice") on your computer; if it is not installed, it says so and exits cleanly (exit code 0).
#   If found: LibreOffice opens model2.xlsx headless, calculates everything, and inject.py copies the numbers into the original file.
# ---------------------------------------------------------------------------
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))   # packaging glue: lets us "import paths"
import paths                                                            # packaging glue: every file location lives in paths.py
import os, shutil, subprocess, tempfile


def find_soffice():
    for name in ('soffice', 'libreoffice'):
        p = shutil.which(name)
        if p:
            return p
    for p in (r'C:\Program Files\LibreOffice\program\soffice.exe', r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
              '/Applications/LibreOffice.app/Contents/MacOS/soffice'):
        if os.path.exists(p):
            return p
    return None


def main():
    soffice = find_soffice()
    if soffice is None:
        print('LibreOffice not found on this computer -> skipping the optional recalculation step.')
        print('That is fine: open outputs/model2.xlsx in Excel and it will calculate every formula on opening.')
        return 0
    src = paths.OUT / 'model2.xlsx'
    if not src.exists():
        print('outputs/model2.xlsx not found - run 05_workbook/build_workbook.py first.')
        return 1
    final = paths.OUT / 'final_model5.xlsx'
    recalc = paths.OUT / 'model2_recalculated.xlsx'
    with tempfile.TemporaryDirectory() as tmp:
        tmp = pathlib.Path(tmp)
        profile = (tmp / 'lo_profile').as_uri()       # private profile so your own LibreOffice settings are untouched
        shutil.copy(src, tmp / 'model2.xlsx')
        cmd = [soffice, '--headless', f'-env:UserInstallation={profile}', '--convert-to', 'xlsx:Calc MS Excel 2007 XML',
               '--outdir', str(tmp / 'out'), str(tmp / 'model2.xlsx')]
        print('running:', ' '.join(cmd))
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
        made = tmp / 'out' / 'model2.xlsx'
        if r.returncode != 0 or not made.exists():
            print('LibreOffice did not produce a file; skipping. Output was:\n', r.stdout, r.stderr)
            return 0
        shutil.copy(made, recalc)
    r = subprocess.run([sys.executable, '-W', 'ignore', str(pathlib.Path(__file__).with_name('inject.py')), str(src), str(recalc), str(final)],
                       capture_output=True, text=True)
    print(r.stdout.strip(), r.stderr.strip()[-500:])
    if r.returncode != 0:
        return 1
    print('wrote', final)
    return 0


if __name__ == '__main__':
    sys.exit(main())
