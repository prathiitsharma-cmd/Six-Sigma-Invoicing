"""Shared file locations for the whole project (one place, so nothing is hard-coded elsewhere).

Folders
  ROOT         the project folder (the one that contains this file)
  DATA         data/            the case workbook original.xlsx (never modified)
  OUT          outputs/         where every script WRITES its results
  PRECOMPUTED  precomputed/     a frozen copy of every intermediate result, so you can start at any stage

Two helpers
  out(name)   -> path inside outputs/ to write to (the folder is created if missing)
  find(name)  -> path to READ: outputs/<name> if you have produced it yourself,
                 otherwise precomputed/<name>. This is what lets you run any single stage on its own.
"""
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent
DATA = ROOT / 'data'
OUT = ROOT / 'outputs'
PRECOMPUTED = ROOT / 'precomputed'
SRC = str(DATA / 'original.xlsx')          # the supplied case workbook (sheets Measure, Improve, Error Reasons, Process Steps ...)


def out(name):
    """Path (as str) of a file to write inside outputs/."""
    OUT.mkdir(exist_ok=True)
    return str(OUT / name)


def find(name):
    """Path (as str) of a file to read: your own outputs/ copy if present, else the shipped precomputed/ copy."""
    p = OUT / name
    if p.exists():
        return str(p)
    q = PRECOMPUTED / name
    if q.exists():
        return str(q)
    raise FileNotFoundError(f'{name} not found in outputs/ or precomputed/ - run the earlier stage that produces it')

# Locations of the code folders (used by scripts that run other scripts)
ANALYSIS = ROOT / '01_analysis'
SOLUTIONS = ROOT / '02_solutions'
STRESS = ROOT / '03_stress'
STATS = ROOT / '04_statistics'
WORKBOOK = ROOT / '05_workbook'
