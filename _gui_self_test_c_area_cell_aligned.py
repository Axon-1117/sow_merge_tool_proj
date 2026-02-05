"""GUI self-test: C区 cell-aligned view populates diff cells.

Run:
  .venv\\Scripts\\python.exe _gui_self_test_c_area_cell_aligned.py

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
    a_rows = [["h1", "h2", "h3"], ["a", "b", "X"]]
    b_rows = [["h1", "h2", "h3"], ["a", "b", "Y"]]

    td1 = tempfile.mkdtemp(prefix="sow_merge_gui_test_c_area_a_")
    td2 = tempfile.mkdtemp(prefix="sow_merge_gui_test_c_area_b_")
    fa = os.path.join(td1, "same.xlsx")
    fb = os.path.join(td2, "same.xlsx")
    _make_xlsx(fa, a_rows)
    _make_xlsx(fb, b_rows)

    import sys
    sys.path.insert(0, r"D:\Tools\sow_merge_tool")
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

    view.only_diff_var.set(0)
    view.refresh(row_only=None, rescan=True)

    # Enable C区单元格对齐 rendering for this test
    view._enable_c_cell = True

    # Select row 2
    view.left.mark_set("insert", "2.0")
    view.right.mark_set("insert", "2.0")
    view._update_cursor_lines()

    # C区-单元格对齐 now shows values only in only-diff mode; ensure it includes X/Y
    txt = view.cell_cmp_text.get("1.0", "end")
    assert "X" in txt and "Y" in txt, f"Expected X/Y in C区 text, got:\n{txt}"

    try:
        app.root.destroy()
    except Exception:
        pass

    print("GUI_SELF_TEST_C_AREA_CELL_ALIGNED_OK")


if __name__ == "__main__":
    main()
