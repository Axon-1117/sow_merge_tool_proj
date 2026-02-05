"""GUI self-test (no pywinauto): verifies cursor compare block highlights diff cells.

Run with:
  .venv\\Scripts\\python.exe _gui_self_test_cursor_highlight.py

This test:
- Creates two temporary xlsx files with one sheet.
- Introduces a diff in a specific column.
- Opens the app (Tk) without mainloop.
- Forces selection on the first displayed row.
- Calls _update_cursor_lines() and asserts that the diffcell tag is present on both lines.

NOTE: This is a programmatic Tk test; it does not require an interactive desktop.
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
    # Prepare data: include header row and a diff on row 2, col 3
    a_rows = [
        ["h1", "h2", "h3"],
        ["a", "b", "X"],
    ]
    b_rows = [
        ["h1", "h2", "h3"],
        ["a", "b", "Y"],
    ]

    td = tempfile.mkdtemp(prefix="sow_merge_gui_test_")
    fa = os.path.join(td, "same.xlsx")
    fb = os.path.join(td, "same.xlsx.copy.xlsx")
    # The app requires same filename; create in separate dirs
    td2 = tempfile.mkdtemp(prefix="sow_merge_gui_test2_")
    fb = os.path.join(td2, "same.xlsx")

    _make_xlsx(fa, a_rows)
    _make_xlsx(fb, b_rows)

    import sow_merge_tool as mod

    app = mod.SowMergeApp(fa, fb)

    # Get the only sheet view (lazy)
    sheet = app.common_sheets[0]
    view = app.sheet_views.get(sheet)
    if view is None:
        app.nb.select(app._sheet_containers[sheet])
        try:
            app.root.update_idletasks(); app.root.update()
        except Exception:
            pass
        view = app.sheet_views[sheet]

    # Force show all rows
    view.only_diff_var.set(False)
    view.refresh(row_only=None, rescan=True)

    # Select line 2 (excel row 2) and update cursor lines
    # Put insert marks on both panes
    view.left.mark_set("insert", "2.0")
    view.right.mark_set("insert", "2.0")
    view._highlight_selected_line(2)
    view._update_cursor_lines()

    # Inspect cursor_cmp tag ranges for 'diffcell'
    ranges = view.cursor_cmp.tag_ranges("diffcell")
    assert ranges, "Expected diffcell tag ranges in cursor compare block, but found none"

    # We expect at least one range on line 1 and one on line 2
    has_line1 = any(str(r).startswith("1.") for r in ranges)
    has_line2 = any(str(r).startswith("2.") for r in ranges)
    assert has_line1, f"Expected diffcell tag on line 1 (A). ranges={ranges}"
    assert has_line2, f"Expected diffcell tag on line 2 (B). ranges={ranges}"

    # Cleanup tk
    try:
        app.root.destroy()
    except Exception:
        pass

    print("GUI_SELF_TEST_CURSOR_HIGHLIGHT_OK")


if __name__ == "__main__":
    main()
