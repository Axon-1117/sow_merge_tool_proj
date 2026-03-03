"""GUI self-test: verifies horizontal scrolling is synchronized between panes.

Run:
  .venv\\Scripts\\python.exe _gui_self_test_hscroll_sync.py

No desktop automation required.
"""

import os
import tempfile
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
    # Create wide text so xscroll is meaningful
    long = "x" * 300
    a_rows = [["h1"], [long]]
    b_rows = [["h1"], [long]]

    td1 = tempfile.mkdtemp(prefix="sow_merge_gui_test_hscroll_a_")
    td2 = tempfile.mkdtemp(prefix="sow_merge_gui_test_hscroll_b_")
    fa = os.path.join(td1, "same.xlsx")
    fb = os.path.join(td2, "same.xlsx")
    _make_xlsx(fa, a_rows)
    _make_xlsx(fb, b_rows)

    import sow_merge_tool as mod

    app = mod.SowMergeApp(fa, fb)
    sheet = app.common_sheets[0]
    # Ensure view is created (lazy)
    view = app.sheet_views.get(sheet)
    if view is None:
        app.nb.select(app._sheet_containers[sheet])
        try:
            app.root.update_idletasks(); app.root.update()
        except Exception:
            pass
        view = app.sheet_views[sheet]

    # Ensure geometry calculated
    try:
        app.root.update_idletasks()
        app.root.update()
    except Exception:
        pass

    # Move left horizontally
    view.left.xview_moveto(0.5)
    try:
        app.root.update_idletasks()
        app.root.update()
    except Exception:
        pass

    lf, ll = view.left.xview()
    rf, rl = view.right.xview()
    # Allow tiny differences due to widget rounding
    assert abs(lf - rf) < 0.02, f"Expected xscroll synced; left={lf, ll} right={rf, rl}"

    # Move right horizontally
    view.right.xview_moveto(0.2)
    try:
        app.root.update_idletasks()
        app.root.update()
    except Exception:
        pass

    lf2, ll2 = view.left.xview()
    rf2, rl2 = view.right.xview()
    assert abs(lf2 - rf2) < 0.02, f"Expected xscroll synced (reverse); left={lf2, ll2} right={rf2, rl2}"

    try:
        app.root.destroy()
    except Exception:
        pass

    print("GUI_SELF_TEST_HSCROLL_SYNC_OK")


if __name__ == "__main__":
    main()
