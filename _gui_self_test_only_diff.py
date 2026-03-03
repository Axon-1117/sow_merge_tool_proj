"""GUI self-test: verifies 'Only show diffs' actually filters rows.

Run:
  .venv\\Scripts\\python.exe _gui_self_test_only_diff.py

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
    # 5 rows, only row 3 differs
    a_rows = [
        ["h1", "h2"],
        [1, 1],
        [2, 2],
        [3, 3],
        [4, 4],
    ]
    b_rows = [
        ["h1", "h2"],
        [1, 1],
        [2, 999],  # diff at row 3
        [3, 3],
        [4, 4],
    ]

    td1 = tempfile.mkdtemp(prefix="sow_merge_gui_test_onlydiff_a_")
    td2 = tempfile.mkdtemp(prefix="sow_merge_gui_test_onlydiff_b_")
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
        # simulate tab selection to trigger lazy creation
        app.nb.select(app._sheet_containers[sheet])
        try:
            app.root.update_idletasks(); app.root.update()
        except Exception:
            pass
        view = app.sheet_views[sheet]

    # Full mode
    view.only_diff_var.set(0)
    view.refresh(row_only=None, rescan=True)
    full_count = len(view.display_rows)
    assert full_count == 5, f"Expected 5 rows shown, got {full_count}"

    # Only diff
    view.only_diff_var.set(1)
    view._toggle_only_diff()
    diff_count = len(view.display_rows)
    assert diff_count == 1, f"Expected 1 diff row shown, got {diff_count}; display_rows={view.display_rows}"
    pair_idx_row3 = view.row_a_to_pair_idx.get(3)
    assert pair_idx_row3 is not None, "Expected row 3 to map to a pair index"
    assert view.display_rows[0] == pair_idx_row3, \
        f"Expected diff row pair index {pair_idx_row3}, got {view.display_rows}"

    try:
        app.root.destroy()
    except Exception:
        pass

    print("GUI_SELF_TEST_ONLY_DIFF_OK")


if __name__ == "__main__":
    main()
