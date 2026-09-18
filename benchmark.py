"""How long one company's run takes, and how often it reads its workbook (Task 17).

For each company in data/: build every output (Excel, deck, memo, manifest) into a temporary
folder with main.run_company, the way `python main.py --all --resume` does. The AI text comes
from the saved analyses in tests/golden/analysis/, so nothing calls the API.

Two things are measured:
- seconds: the median of REPEATS runs, each started from empty caches (clear_caches), so a run
  can't borrow work an earlier run did. The first run is a warm-up and isn't counted: it also pays
  for Python loading its libraries.
- reads: how many times the workbook was opened and parsed (clean.find_kpi_sheet runs), and how
  many times the metric table was worked out (metrics.nrr runs, once per metric table).

    python benchmark.py
"""

import contextlib
import io
import shutil
import statistics
import sys
import tempfile
import time
from pathlib import Path

import main
from golden import companies, fixture_path
from metrics import load_config

PROJECT_DIR = Path(__file__).parent
REPEATS = 5
COUNTED = {("clean.py", "find_kpi_sheet"): "reads", ("metrics.py", "nrr"): "metric tables"}


def clear_caches():
    """Forget every cached clean and metric result, if this version of the code has caches."""
    for module in ("clean", "metrics"):
        clear = getattr(sys.modules[module], "clear_cache", None)
        if clear:
            clear()


def one_run(company, config):
    """Build one company's outputs from empty caches in a fresh folder. Returns the seconds it took."""
    clear_caches()
    with tempfile.TemporaryDirectory() as folder:
        shutil.copy(fixture_path(company), folder)   # the saved analysis, reused: no API call
        started = time.perf_counter()
        with contextlib.redirect_stdout(io.StringIO()):   # run_company prints a line per step
            main.run_company(PROJECT_DIR / "data" / f"{company}.xlsx", config, skip_ai=True,
                             output_dir=folder, reuse_saved=True)
        return time.perf_counter() - started


def counted_run(company, config):
    """One run, counting the calls named in COUNTED. Returns {"reads": n, "metric tables": n}."""
    counts = dict.fromkeys(COUNTED.values(), 0)

    def count(frame, event, arg):
        key = (Path(frame.f_code.co_filename).name, frame.f_code.co_name)
        if event == "call" and key in COUNTED:
            counts[COUNTED[key]] += 1

    sys.setprofile(count)
    try:
        one_run(company, config)
    finally:
        sys.setprofile(None)
    return counts


def measure(company, config):
    """{"company", "seconds" (median of REPEATS), "reads", "metric tables"} for one company."""
    one_run(company, config)   # warm-up
    seconds = statistics.median(one_run(company, config) for _ in range(REPEATS))
    return {"company": company, "seconds": seconds, **counted_run(company, config)}


if __name__ == "__main__":
    config = load_config()
    print(f"{'Company':<12} {'Seconds (median of ' + str(REPEATS) + ')':>24} {'Reads':>6} {'Metric tables':>14}")
    for row in (measure(company, config) for company in companies()):
        print(f"{row['company']:<12} {row['seconds']:>24.3f} {row['reads']:>6} {row['metric tables']:>14}")
