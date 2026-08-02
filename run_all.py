import subprocess
import sys
from pathlib import Path

BASE = Path(__file__).parent

EXTRACT_OUTPUTS = [
    BASE / "gaji.csv",
    BASE / "lembar3.csv",
    BASE / "perhitungan.csv",
    BASE / "perhitungan_ringkasan.csv",
]

MERGE_OUTPUT = BASE / "gabung.xlsx"

ALL_OLD_FILES = EXTRACT_OUTPUTS + [MERGE_OUTPUT]

SCRIPTS = [
    (BASE / "gaji.py", [str(BASE / "gaji.pdf")]),
    (BASE / "lem.py", [str(BASE / "lembar3.pdf")]),
    (BASE / "per.py", [str(BASE / "perhitungan.pdf")]),
]


def clean_old():
    for f in ALL_OLD_FILES:
        if f.exists():
            f.unlink()
            print(f"Removed old file: {f.name}")


def run_step(name, script, args):
    print(f"\n{'=' * 60}")
    print(f"Running: {name}")
    print(f"{'=' * 60}")
    cmd = [sys.executable, str(script)] + args
    result = subprocess.run(cmd, cwd=str(BASE))
    if result.returncode != 0:
        print(f"ERROR: {name} failed with exit code {result.returncode}")
        return False
    print(f"OK: {name} completed")
    return True


def main():
    clean_old()

    for script, args in SCRIPTS:
        if not run_step(script.stem, script, args):
            print(f"\nPipeline stopped at {script.stem}.")
            sys.exit(1)

    print(f"\n{'=' * 60}")
    print("Running: gabung.py (merge)")
    print(f"{'=' * 60}")
    if not run_step("gabung", BASE / "gabung.py", []):
        print("\nPipeline stopped at gabung.py.")
        sys.exit(1)

    print(f"\n{'=' * 60}")
    print("Pipeline finished successfully.")
    if MERGE_OUTPUT.exists():
        print(f"Output: {MERGE_OUTPUT.name}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()