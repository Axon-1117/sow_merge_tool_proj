"""GUI benchmark: measure overwrite performance (row overwrite) without desktop automation.

Run:
  .venv\Scripts\python.exe _gui_benchmark_overwrite.py

Outputs timing for:
- A2B row overwrite
- B2A row overwrite

This is a best-effort benchmark; results vary by machine and file size.
"""

import os
import tempfile
import time
from openpyxl import Workbook


def _make_xlsx(path: str, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, v in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx).value = v
    wb.save(path)


def main():
    # Create a moderately wide row to simulate real data
    cols = 120
    header = [f"H{i}" for i in range(1, cols + 1)]

    base = [f"v{i}" for i in range(1, cols + 1)]
    a_row = base.copy()
    b_row = base.copy()

    # Introduce diffs across many cells to stress overwrite loop
    for i in range(10, 90, 3):
        b_row[i] = f"DIFF{i}"

    a_rows = [header, a_row]
    b_rows = [header, b_row]

    td1 = tempfile.mkdtemp(prefix="sow_merge_bench_a_")
    td2 = tempfile.mkdtemp(prefix="sow_merge_bench_b_")
    fa = os.path.join(td1, "same.xlsx")
    fb = os.path.join(td2, "same.xlsx")
    _make_xlsx(fa, a_rows)
    _make_xlsx(fb, b_rows)

    import sys
    sys.path.insert(0, r"D:\Tools\sow_merge_tool")
    import sow_merge_tool as mod

    app = mod.SowMergeApp(fa, fb)

    sheet = app.common_sheets[0]
    # trigger lazy view creation
    app.nb.select(app._sheet_containers[sheet])
    try:
        app.root.update_idletasks(); app.root.update()
    except Exception:
        pass

    view = app.sheet_views[sheet]
    view.only_diff_var.set(0)
    view.refresh(row_only=None, rescan=True)

    # Wait for background edit preload (if enabled)
    try:
        if getattr(app, "_edit_loading_started", False):
            app._edit_loaded_event.wait(timeout=10)
    except Exception:
        pass

    # Select excel row 2
    view.selected_excel_row = 2

    # Measure A2B
    t0 = time.perf_counter()
    view._copy_selected_row("A2B")
    t1 = time.perf_counter()

    # Recompute diff map now (benchmark includes any per-row work done by overwrite)
    view.refresh(row_only=2, rescan=False)

    # Measure B2A
    t2 = time.perf_counter()
    view._copy_selected_row("B2A")
    t3 = time.perf_counter()

    a2b_ms = (t1 - t0) * 1000.0
    b2a_ms = (t3 - t2) * 1000.0

    try:
        app.root.destroy()
    except Exception:
        pass

    print(f"BENCH_OVERWRITE_A2B_MS={a2b_ms:.1f}")
    print(f"BENCH_OVERWRITE_B2A_MS={b2a_ms:.1f}")


if __name__ == "__main__":
    main()
