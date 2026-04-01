"""
Run pdf_to_excel.py over a contrast grid (default 1.0 .. 2.5 step 0.2) and write
one Excel per value: <pdf_stem>_ocr_contrast_<value>.xlsx in the PDF directory.

Requires: pdf_to_excel.py with --ocr-contrast and --output-excel support.

Example:
  python scripts/ocr_contrast_sweep.py ^
    "D:\\path\\to\\statement.pdf"

Optional: scan each output for VIVAAEROBUS / amount substrings and print a table.
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys


def _repo_root() -> str:
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _contrast_values(start: float, end: float, step: float) -> list[float]:
    """1.0 .. 2.5 step 0.2 yields 1.0..2.4 plus 2.5 (step does not land on 2.5)."""
    out: list[float] = []
    v = round(start, 6)
    end_r = round(end, 6)
    step_r = round(step, 6)
    while v <= end_r + 1e-9:
        out.append(round(v, 2))
        nv = round(v + step_r, 2)
        if nv > end_r + 1e-9:
            break
        v = nv
    if out and abs(out[-1] - end_r) > 1e-6 and end_r > out[-1]:
        out.append(end_r)
    return out


def _flatten_xlsx_text(path: str) -> str:
    try:
        import openpyxl
    except ImportError:
        return ""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        parts: list[str] = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                parts.append(
                    "\t".join("" if c is None else str(c) for c in row)
                )
        wb.close()
        return "\n".join(parts)
    except Exception as e:
        return f"<read error: {e}>"


def main() -> int:
    root = _repo_root()
    script = os.path.join(root, "pdf_to_excel.py")
    ap = argparse.ArgumentParser(description="OCR contrast sweep for pdf_to_excel.py")
    ap.add_argument(
        "pdf",
        nargs="?",
        default=os.path.join(
            root,
            "Test",
            "Bank Statement",
            "Final Test",
            "OCR",
            "3a7c8e9c-cf71-4a38-8079-51b6f1cbf689 (1).pdf",
        ),
        help="Input PDF path",
    )
    ap.add_argument("--start", type=float, default=1.0)
    ap.add_argument("--end", type=float, default=2.5)
    ap.add_argument("--step", type=float, default=0.2)
    ap.add_argument(
        "--extra-args",
        nargs="*",
        default=[],
        help="Extra args forwarded to pdf_to_excel.py (e.g. --debug)",
    )
    ap.add_argument(
        "--no-scan",
        action="store_true",
        help="Skip scanning Excel for VIVAAEROBUS / amount markers",
    )
    args = ap.parse_args()

    pdf = os.path.normpath(os.path.abspath(args.pdf))
    if not os.path.isfile(pdf):
        print(f"PDF not found: {pdf}", file=sys.stderr)
        return 1
    if not os.path.isfile(script):
        print(f"pdf_to_excel.py not found: {script}", file=sys.stderr)
        return 1

    pdf_dir = os.path.dirname(pdf)
    stem = os.path.splitext(os.path.basename(pdf))[0]
    contrasts = _contrast_values(args.start, args.end, args.step)
    print(f"Contrasts: {contrasts}", flush=True)

    results: list[tuple[float, str, str]] = []
    py = sys.executable

    for c in contrasts:
        out_xlsx = os.path.join(pdf_dir, f"{stem}_ocr_contrast_{c:.2f}.xlsx")
        cmd = [
            py,
            script,
            pdf,
            "--ocr-contrast",
            str(c),
            "--output-excel",
            out_xlsx,
        ] + list(args.extra_args)
        print(f"\n=== contrast {c} -> {out_xlsx} ===", flush=True)
        r = subprocess.run(cmd, cwd=root)
        if r.returncode != 0:
            print(f"[FAIL] exit {r.returncode} for contrast {c}", flush=True)
            results.append((c, out_xlsx, f"exit_{r.returncode}"))
            continue
        if args.no_scan:
            results.append((c, out_xlsx, "ok"))
            continue
        blob = _flatten_xlsx_text(out_xlsx)
        has_v = "VIVAAEROBUS" in blob.upper()
        has_23 = "23/AGO" in blob
        has_wrong = "2,189.27" in blob
        has_right = "2,789.27" in blob
        note = []
        if has_v:
            note.append("VIVAAEROBUS")
        if has_23:
            note.append("23/AGO")
        if has_wrong:
            note.append("amt_2189_wrong")
        if has_right:
            note.append("amt_2789")
        flag = ", ".join(note) if note else "(no markers)"
        print(f"  scan: {flag}", flush=True)
        results.append((c, out_xlsx, flag))

    print("\n--- summary ---", flush=True)
    for c, path, flag in results:
        print(f"  {c:>4}: {flag}  |  {path}", flush=True)

    return 0


if __name__ == "__main__":
    sys.exit(main())
