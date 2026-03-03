import os
import sys
import argparse
import re
import difflib
import tempfile
import subprocess
import traceback
from datetime import datetime
import time
import stat
import shutil

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import threading

from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.formula import ArrayFormula
# Note: formulas will be treated as cached values only (data_only), with fallback when cache is missing.
from openpyxl.utils import get_column_letter


APP_NAME = "sow_merge_tool"
APP_VERSION = "2026-03-02.perf1"
APP_BUILD_TAG = "new10-savefidelity"

# Debug logging (writes to %TEMP%\sow_merge_tool_debug.log)
_DEBUG_LOG_PATH = os.path.join(tempfile.gettempdir(), f"{APP_NAME}_debug.log")
_DEBUG_ENABLED = False

# Save performance: fast mode writes directly to target (faster, less safe than atomic replace)
_FAST_SAVE_ENABLED = True
# Save correctness: keep workbook fidelity (styles/formulas/metadata).
# values-only fast save can make unrelated sheets look modified in SVN diff.
_FAST_SAVE_VALUES_ONLY = False
# Open performance: skip background preloads and global scans (loads on demand)
_FAST_OPEN_ENABLED = True
# Global mode: compare and save cached values only (ignore formulas)
_USE_CACHED_VALUES_ONLY = True
# When cached values are missing for formulas, try to recalc via Excel (if available)
_AUTO_RECALC_MISSING_CACHE = False
_AUTO_RECALC_FORMULAS_ALWAYS = False
_CACHE_CHECK_MAX_CELLS = 3000
# Render performance: limit initial rows rendered (user can load full)
_FAST_RENDER_ROW_LIMIT = 800
_FAST_RENDER_BATCH = 500
_LARGE_SHEET_ROW_THRESHOLD = 1000
_LARGE_SHEET_INITIAL_ROWS = 200
_LARGE_SHEET_BLOCK_ROWS = 1000
_LARGE_SHEET_DIRECT_PAIR_THRESHOLD = 5000
_TABMARK_QUICK_TAIL_ROWS = 2000

# Settings (persist UI prefs)
_SETTINGS_PATH = os.path.join(os.environ.get("LOCALAPPDATA", tempfile.gettempdir()), APP_NAME, "settings.json")


def _dlog(msg: str):
    if not _DEBUG_ENABLED:
        return
    try:
        ts = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        with open(_DEBUG_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass


def _val_to_str(v):
    """Render a cell value as single-line text for the Text widget.

    IMPORTANT: We must keep each Excel row rendered as exactly ONE line in tk.Text.
    So we sanitize embedded newlines/tabs that would otherwise break line alignment
    and cause diff highlights to drift.
    """
    if v is None:
        return ""
    s = str(v)
    # Normalize line breaks and tabs to keep one-row-per-line invariant
    s = s.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    s = s.replace("\t", "    ")
    return s


def _effective_bounds(ws):
    """Return (max_row, max_col) based on actual non-empty cells.

    Some workbooks have an inaccurate ws.max_row/ws.max_column (e.g. only first
    N rows reported). We scan ws._cells (when available) to derive a safer bound.
    """
    max_r = ws.max_row or 1
    max_c = ws.max_column or 1
    try:
        cells = getattr(ws, "_cells", None)
        if cells:
            last_r = 1
            last_c = 1
            found = False
            for cell in cells.values():
                v = cell.value
                if v not in (None, ""):
                    found = True
                    if cell.row > last_r:
                        last_r = cell.row
                    if cell.column > last_c:
                        last_c = cell.column
            if found:
                max_r = max(max_r, last_r)
                max_c = max(max_c, last_c)
    except Exception:
        pass
    return max(1, max_r), max(1, max_c)


def _save_values_only_from_wb(src_wb, target_path: str):
    """Fast save: values only, no styles. Drops formatting."""
    def _trim_bounds_ws(ws):
        # Find last non-empty row/col (scan ws._cells then fallback)
        max_r = ws.max_row or 1
        max_c = ws.max_column or 1
        last_r = 1
        last_c = 1
        found = False
        try:
            cells = getattr(ws, "_cells", None)
            if cells:
                for cell in cells.values():
                    v = cell.value
                    if v not in (None, ""):
                        found = True
                        if cell.row > last_r:
                            last_r = cell.row
                        if cell.column > last_c:
                            last_c = cell.column
        except Exception:
            pass
        if not found:
            for r in range(max_r, max(1, max_r - 5000), -1):
                row = next(ws.iter_rows(min_row=r, max_row=r, min_col=1, max_col=max_c, values_only=True), ())
                if any(v not in (None, "") for v in row):
                    found = True
                    last_r = r
                    for ci in range(len(row), 0, -1):
                        v = row[ci - 1]
                        if v not in (None, ""):
                            last_c = ci
                            break
                    break
        if not found:
            return max_r, max_c
        use_r = min(max_r, last_r + 50)
        use_c = min(max_c, last_c + 50)
        return max(1, use_r), max(1, use_c)

    dst = Workbook(write_only=True)
    # Remove default sheet
    try:
        if dst.sheetnames:
            dst.remove(dst[dst.sheetnames[0]])
    except Exception:
        pass
    for name in src_wb.sheetnames:
        ws_src = src_wb[name]
        ws_dst = dst.create_sheet(title=name)
        max_row, max_col = _trim_bounds_ws(ws_src)
        if max_row <= 0 or max_col <= 0:
            continue
        for r in range(1, max_row + 1):
            row_vals = []
            for c in range(1, max_col + 1):
                row_vals.append(ws_src.cell(row=r, column=c).value)
            ws_dst.append(row_vals)
    dst.save(target_path)


def _cell_display_and_equal(ws_a_val, ws_b_val, ws_a_edit, ws_b_edit, r: int, c: int):
    va_val = ws_a_val.cell(row=r, column=c).value
    vb_val = ws_b_val.cell(row=r, column=c).value

    if _USE_CACHED_VALUES_ONLY:
        # If cache missing but edit has a literal value, use it for display/compare.
        if ws_a_edit is not None and ws_b_edit is not None:
            try:
                if va_val is None:
                    va_edit = ws_a_edit.cell(row=r, column=c).value
                    if va_edit is not None and not _formula_text(va_edit):
                        va_val = va_edit
                if vb_val is None:
                    vb_edit = ws_b_edit.cell(row=r, column=c).value
                    if vb_edit is not None and not _formula_text(vb_edit):
                        vb_val = vb_edit
            except Exception:
                pass
        # If cache missing on one side but both formulas are the same, treat as equal and display the available value.
        if ws_a_edit is not None and ws_b_edit is not None and ((va_val is None) != (vb_val is None)):
            va_edit = ws_a_edit.cell(row=r, column=c).value
            vb_edit = ws_b_edit.cell(row=r, column=c).value
            fa = _formula_text(va_edit)
            fb = _formula_text(vb_edit)
            if fa and fb and fa == fb:
                v = va_val if va_val is not None else vb_val
                return v, v, True
        # Compare with numeric/string normalization to avoid false diffs
        eq = (_merge_cmp_value(va_val) == _merge_cmp_value(vb_val))
        return va_val, vb_val, eq

    eq = (_merge_cmp_value(va_val) == _merge_cmp_value(vb_val))
    return va_val, vb_val, eq


def _cell_display_and_equal_by_row(ws_a_val, ws_b_val, ws_a_edit, ws_b_edit, ra: int | None, rb: int | None, c: int):
    va_val = ws_a_val.cell(row=ra, column=c).value if ra is not None else None
    vb_val = ws_b_val.cell(row=rb, column=c).value if rb is not None else None

    if _USE_CACHED_VALUES_ONLY:
        if ws_a_edit is not None and ws_b_edit is not None:
            try:
                if va_val is None and ra is not None:
                    va_edit = ws_a_edit.cell(row=ra, column=c).value
                    if va_edit is not None and not _formula_text(va_edit):
                        va_val = va_edit
                if vb_val is None and rb is not None:
                    vb_edit = ws_b_edit.cell(row=rb, column=c).value
                    if vb_edit is not None and not _formula_text(vb_edit):
                        vb_val = vb_edit
            except Exception:
                pass
            if (va_val is None) != (vb_val is None):
                try:
                    va_edit = ws_a_edit.cell(row=ra, column=c).value if ra is not None else None
                    vb_edit = ws_b_edit.cell(row=rb, column=c).value if rb is not None else None
                    fa = _formula_text(va_edit)
                    fb = _formula_text(vb_edit)
                    if fa and fb and fa == fb:
                        v = va_val if va_val is not None else vb_val
                        return v, v, True
                except Exception:
                    pass
        eq = (_merge_cmp_value(va_val) == _merge_cmp_value(vb_val))
        return va_val, vb_val, eq

    eq = (_merge_cmp_value(va_val) == _merge_cmp_value(vb_val))
    return va_val, vb_val, eq


def _formula_text(v):
    if isinstance(v, str) and v.startswith("="):
        return v
    if isinstance(v, ArrayFormula):
        return getattr(v, "text", None)
    t = getattr(v, "text", None)
    if isinstance(t, str) and t.startswith("="):
        return t
    return None


def _merge_cmp_value(v):
    """Normalize values for merge conflict comparison to match UI display."""
    try:
        if v is None:
            return ""
        s = _val_to_str(v)
        if isinstance(s, str):
            # Normalize line endings and trim trailing whitespace to avoid false conflicts
            s = s.replace("\r\n", "\n").rstrip()
            # Normalize numeric strings
            try:
                num = float(s)
                if num.is_integer():
                    return str(int(num))
                return str(num)
            except Exception:
                return s
        return s
    except Exception:
        return v


def _scan_formula_cache(path: str):
    """Return (has_formula, missing_cache) based on a sample scan."""
    try:
        wb_val = load_workbook(path, data_only=True, read_only=True)
        wb_edit = load_workbook(path, data_only=False, read_only=True)
    except Exception as e:
        _dlog(f"cache check open failed: {e}")
        return False, False

    checked = 0
    has_formula = False
    missing_cache = False
    try:
        for sheet in wb_edit.sheetnames:
            ws_e = wb_edit[sheet]
            ws_v = wb_val[sheet]
            for row in ws_e.iter_rows(values_only=False):
                for cell in row:
                    if checked >= _CACHE_CHECK_MAX_CELLS:
                        return has_formula, missing_cache
                    f = _formula_text(cell.value)
                    if not f:
                        continue
                    has_formula = True
                    checked += 1
                    try:
                        v = ws_v.cell(row=cell.row, column=cell.column).value
                    except Exception:
                        v = None
                    if v is None:
                        missing_cache = True
                        return has_formula, missing_cache
    finally:
        try:
            wb_val.close()
            wb_edit.close()
        except Exception:
            pass
    return has_formula, missing_cache


def _recalc_with_excel(path: str) -> str | None:
    """Use Excel COM to recalc formulas and update cached values in a temp copy."""
    try:
        base = os.path.basename(path)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        tmp = os.path.join(tempfile.gettempdir(), f"{APP_NAME}_recalc_{os.getpid()}_{ts}_{base}")
        shutil.copy2(path, tmp)
    except Exception as e:
        _dlog(f"recalc copy failed: {e}")
        return None

    try:
        ps = (
            "$ErrorActionPreference='Stop';"
            "$p='" + tmp.replace("'", "''") + "';"
            "$xl=New-Object -ComObject Excel.Application;"
            "$xl.Visible=$false;"
            "$xl.DisplayAlerts=$false;"
            "$xl.AskToUpdateLinks=$false;"
            "$xl.EnableEvents=$false;"
            "$wb=$xl.Workbooks.Open($p,$false,$false);"
            "try{$xl.Calculation=-4105}catch{};"
            "try{$xl.CalculateFullRebuild()}catch{};"
            "try{$wb.RefreshAll();$xl.CalculateFullRebuild()}catch{};"
            "$wb.Save();"
            "$wb.Close($true);"
            "$xl.Quit();"
        )
        no_window = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps],
            capture_output=True,
            text=True,
            timeout=120,
            creationflags=no_window,
        )
        if r.returncode != 0:
            _dlog(f"excel recalc ps failed: {r.stderr.strip()}")
            return None
        return tmp
    except Exception as e:
        _dlog(f"excel recalc failed: {e}")
        return None


def _prepare_val_path(path: str) -> str:
    if not _USE_CACHED_VALUES_ONLY or not _AUTO_RECALC_MISSING_CACHE:
        return path
    try:
        has_formula, missing_cache = _scan_formula_cache(path)
        if has_formula and (_AUTO_RECALC_FORMULAS_ALWAYS or missing_cache):
            _dlog(f"formula cache recalc: has_formula={has_formula} missing_cache={missing_cache} path={path}")
            tmp = _recalc_with_excel(path)
            if tmp:
                _dlog(f"recalc cache OK: {tmp}")
                return tmp
    except Exception as e:
        _dlog(f"prepare val path failed: {e}")
    return path


def _recalc_and_prepare_val_path(path: str) -> str | None:
    """Force Excel recalc to refresh cached values and return temp path."""
    try:
        tmp = _recalc_with_excel(path)
        return tmp
    except Exception:
        return None


def _launch_deferred_copy(src: str, dst: str, retries: int = 60, delay_ms: int = 500):
    """Launch a background copy that retries for a while (to avoid lock issues)."""
    try:
        ps = (
            f"$src='{src}';$dst='{dst}';"
            f"for($i=0;$i -lt {retries};$i++){{"
            "try{Copy-Item -LiteralPath $src -Destination $dst -Force;"
            "Remove-Item -LiteralPath $src -Force;exit 0}catch{Start-Sleep -Milliseconds "
            f"{delay_ms}}};exit 1"
        )
        creationflags = 0
        try:
            creationflags = subprocess.CREATE_NO_WINDOW
        except Exception:
            creationflags = 0
        subprocess.Popen(["powershell", "-NoProfile", "-Command", ps], creationflags=creationflags)
    except Exception as e:
        _dlog(f"deferred copy launch failed: {e}")


def _find_tortoise_merge_exe():
    candidates = [
        os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "TortoiseSVN", "bin", "TortoiseMerge.exe"),
        os.path.join(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"), "TortoiseSVN", "bin", "TortoiseMerge.exe"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return "TortoiseMerge.exe"


def _find_tortoise_proc_exe():
    candidates = [
        os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "TortoiseSVN", "bin", "TortoiseProc.exe"),
        os.path.join(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"), "TortoiseSVN", "bin", "TortoiseProc.exe"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return "TortoiseProc.exe"




def _try_export_svn_revision_from_merge_temp(path: str) -> str:
    """If path looks like *.merge-left.r#### or *.merge-right.r####, export that revision from WC.

    Returns replacement path if export succeeded; otherwise returns original path.
    """
    try:
        if not path:
            return path
        p = os.path.abspath(path)
        m = re.match(r"^(?P<base>.+)\.merge-(left|right)\.r(?P<rev>\d+)$", p, flags=re.IGNORECASE)
        if not m:
            return path
        base_path = m.group("base")
        rev = m.group("rev")
        if not os.path.exists(base_path):
            # Try same dir + original base filename
            base_path = os.path.join(os.path.dirname(p), os.path.basename(m.group("base")))
        if not os.path.exists(base_path):
            _dlog(f"svn export skip: base not found for {path}")
            return path

        proc_exe = _find_tortoise_proc_exe()
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = os.path.basename(base_path)
        save_path = os.path.join(tempfile.gettempdir(), f"{APP_NAME}_svncat_r{rev}_{ts}_{base_name}")
        if not save_path.lower().endswith(".xlsx"):
            save_path += ".xlsx"

        try:
            _dlog(f"svn export: {base_path} r{rev} -> {save_path}")
        except Exception:
            pass

        # TortoiseProc may show UI; run and wait briefly for file to appear.
        try:
            subprocess.Popen([
                proc_exe,
                "/command:cat",
                f"/path:{base_path}",
                f"/revision:{rev}",
                f"/savepath:{save_path}",
                "/closeonend:1",
            ])
        except Exception as e:
            _dlog(f"svn export failed launch: {e}")
            return path

        # Wait for file to be created (best-effort)
        for _ in range(50):
            try:
                if os.path.exists(save_path) and os.path.getsize(save_path) > 0:
                    return save_path
            except Exception:
                pass
            time.sleep(0.1)

        _dlog(f"svn export timeout: {save_path}")
        return path
    except Exception as e:
        _dlog(f"svn export error: {e}")
        return path



def _find_handle_exe():
    candidates = [
        os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "handle.exe"),
        r"C:\Sysinternals\handle.exe",
        r"C:\Tools\Sysinternals\handle.exe",
        r"D:\Tools\Sysinternals\handle.exe",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return shutil.which("handle.exe")


def _log_lock_holders(path: str) -> bool:
    """Return True if Excel is detected holding the file."""
    excel_found = False
    try:
        handle_exe = _find_handle_exe()
        if not handle_exe:
            _dlog(f"lock holders: handle.exe not found for {path}")
            return False
        no_window = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        r = subprocess.run(
            [handle_exe, "-accepteula", path],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=no_window,
        )
        out = (r.stdout or "") + (r.stderr or "")
        if out.strip():
            for line in out.splitlines():
                if path.lower() in line.lower():
                    _dlog(f"lock holders: {line.strip()}")
                    if "excel.exe" in line.lower():
                        excel_found = True
            return
        _dlog(f"lock holders: no output for {path}")
    except Exception as e:
        _dlog(f"lock holders: failed {e}")
    return excel_found


def _try_svn_resolve(path: str) -> bool:
    """Attempt to mark conflict as resolved in SVN."""
    try:
        svn_exe = shutil.which("svn")
        if svn_exe:
            subprocess.run(
                [svn_exe, "resolve", "--accept", "working", path],
                check=False,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return True
    except Exception:
        pass
    # Fallback to TortoiseProc (may show UI)
    try:
        proc_exe = _find_tortoise_proc_exe()
        subprocess.Popen([proc_exe, "/command:resolve", f"/path:{path}", "/closeonend:1"])
        return True
    except Exception:
        return False


def open_tortoise_merge(left_txt: str, right_txt: str, title: str):
    exe = _find_tortoise_merge_exe()
    args = [exe, "/base", left_txt, "/mine", right_txt, "/title", title]
    subprocess.Popen(args)


def _show_conflict_popup(conflicts):
    try:
        root = tk.Tk()
        root.withdraw()
        win = tk.Toplevel(root)
        win.title("发现冲突")
        win.resizable(False, False)
        win.geometry("+{}+{}".format(root.winfo_screenwidth() // 2 - 220, root.winfo_screenheight() // 2 - 180))
        msg = "与其他同学冲突，请联系确认后再修改保存！！！"
        lbl = tk.Label(win, text=msg, fg="red", font=("Microsoft YaHei", 12, "bold"), padx=16, pady=10)
        lbl.pack()

        detail_lines = []
        for sheet, r, c, _vm, _vt in conflicts[:3]:
            col = get_column_letter(c)
            detail_lines.append(f"{sheet}!{col}{r}")
        if len(conflicts) > 3:
            detail_lines.append("...")
        detail_text = "\n".join(detail_lines) if detail_lines else "（无）"
        txt = tk.Text(win, height=12, width=60)
        txt.insert("1.0", detail_text)
        txt.configure(state="disabled")
        txt.pack(padx=12, pady=(0, 10))

        tk.Button(win, text="确定", command=win.destroy).pack(pady=(0, 10))
        win.grab_set()
        win.wait_window()
        root.destroy()
    except Exception:
        pass


def excel_to_text(path: str, out_path: str, thick_sep_char: str = "="):
    val_path = _prepare_val_path(path)
    wb = load_workbook(val_path, data_only=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(f"{APP_NAME} text export\n")
        f.write(f"Source: {path}\n")
        f.write(f"Time: {datetime.now().isoformat(sep=' ', timespec='seconds')}\n\n")

        for idx, name in enumerate(wb.sheetnames):
            ws = wb[name]
            max_row = ws.max_row or 1
            max_col = ws.max_column or 1

            if idx != 0:
                f.write("\n" + (thick_sep_char * 120) + "\n")

            title = f"SHEET: {name}"
            pad = max(0, 120 - len(title) - 2)
            left = thick_sep_char * (pad // 2)
            right = thick_sep_char * (pad - (pad // 2))
            f.write(f"{left} {title} {right}\n")

            cols = [ws.cell(row=1, column=c).coordinate[:-1] for c in range(1, max_col + 1)]
            f.write("ROW\t" + "\t".join(cols) + "\n")

            for r in range(1, max_row + 1):
                vals = []
                for c in range(1, max_col + 1):
                    vals.append(_val_to_str(ws.cell(row=r, column=c).value))
                f.write(str(r) + "\t" + "\t".join(vals) + "\n")


def pick_two_files_same_name():
    root = tk.Tk()
    root.withdraw()

    a = filedialog.askopenfilename(title="Select first .xlsx file", filetypes=[("Excel Workbook", "*.xlsx")])
    if not a:
        return None, None
    b = filedialog.askopenfilename(title="Select second .xlsx file (same filename)", filetypes=[("Excel Workbook", "*.xlsx")])
    if not b:
        return None, None

    if os.path.basename(a).lower() != os.path.basename(b).lower():
        messagebox.showerror(
            "Filename mismatch",
            f"The two files must have the same filename.\n\nA: {os.path.basename(a)}\nB: {os.path.basename(b)}",
        )
        return None, None

    return a, b


def _detect_svn_conflict_files(target_path: str):
    folder = os.path.dirname(target_path)
    base_name = os.path.basename(target_path)
    # SVN conflict artifacts:
    # - file.merge-left.r#### / file.merge-right.r#### (newer SVN)
    # - file.r<rev> (older SVN), possibly .mine
    merge_left = []
    merge_right = []
    for name in os.listdir(folder):
        if name.startswith(base_name + ".merge-left.r"):
            suffix = name[len(base_name) + len(".merge-left.r"):]
            if suffix.isdigit():
                merge_left.append((int(suffix), os.path.join(folder, name)))
        elif name.startswith(base_name + ".merge-right.r"):
            suffix = name[len(base_name) + len(".merge-right.r"):]
            if suffix.isdigit():
                merge_right.append((int(suffix), os.path.join(folder, name)))
        elif name == base_name + ".merge-left":
            merge_left.append((0, os.path.join(folder, name)))
        elif name == base_name + ".merge-right":
            merge_right.append((0, os.path.join(folder, name)))
    if merge_left and merge_right:
        merge_left.sort(key=lambda x: x[0])
        merge_right.sort(key=lambda x: x[0])
        base_path = merge_left[-1][1]
        theirs_path = merge_right[-1][1]
        mine_path = target_path
        merged_path = target_path
        return base_path, mine_path, theirs_path, merged_path

    # Older SVN conflict artifacts: file.r<rev> (numeric), possibly .mine
    r_files = []
    for name in os.listdir(folder):
        if not name.startswith(base_name + ".r"):
            continue
        suffix = name[len(base_name) + 2:]
        if suffix.isdigit():
            r_files.append((int(suffix), os.path.join(folder, name)))
    if len(r_files) >= 2:
        r_files.sort(key=lambda x: x[0])
        base_path = r_files[0][1]
        theirs_path = r_files[-1][1]
        mine_path = target_path
        merged_path = target_path
        return base_path, mine_path, theirs_path, merged_path
    # Fallback for rOLD/rNEW naming
    r_old = os.path.join(folder, base_name + ".rOLD")
    r_new = os.path.join(folder, base_name + ".rNEW")
    if os.path.exists(r_old) and os.path.exists(r_new):
        return r_old, target_path, r_new, target_path
    return None


def _has_svn_conflict_artifacts(target_path: str) -> bool:
    try:
        folder = os.path.dirname(target_path)
        base_name = os.path.basename(target_path)
        mine = os.path.join(folder, base_name + ".mine")
        if os.path.exists(mine):
            return True
        for name in os.listdir(folder):
            if name.startswith(base_name + ".merge-left") or name.startswith(base_name + ".merge-right"):
                return True
        r_old = os.path.join(folder, base_name + ".rOLD")
        r_new = os.path.join(folder, base_name + ".rNEW")
        if os.path.exists(r_old) or os.path.exists(r_new):
            return True
        for name in os.listdir(folder):
            if name.startswith(base_name + ".r"):
                suffix = name[len(base_name) + 2:]
                if suffix.isdigit():
                    return True
    except Exception:
        pass
    return False


def _find_conflict_in_dir(folder: str):
    try:
        # If there is exactly one conflicted file in folder, return it.
        base_names = set()
        for name in os.listdir(folder):
            if ".merge-left" in name:
                base = name.split(".merge-left")[0]
                base_names.add(base)
                continue
            if ".merge-right" in name:
                base = name.split(".merge-right")[0]
                base_names.add(base)
                continue
            if ".r" in name:
                base = name.split(".r")[0]
                base_names.add(base)
        candidates = []
        for base in base_names:
            target = os.path.join(folder, base)
            if os.path.exists(target) and _has_svn_conflict_artifacts(target):
                candidates.append(target)
        if len(candidates) == 1:
            return candidates[0]
    except Exception:
        pass
    return None


def _auto_pick_conflict_file():
    # Best-effort: try current working directory
    try:
        cwd = os.getcwd()
        p = _find_conflict_in_dir(cwd)
        if p:
            return p
        # Walk up to find SVN working copy root (.svn)
        cur = cwd
        wc_root = None
        while True:
            if os.path.isdir(os.path.join(cur, ".svn")):
                wc_root = cur
                break
            parent = os.path.dirname(cur)
            if parent == cur:
                break
            cur = parent
        if wc_root:
            # If exactly one conflicted file exists in the working copy, auto-pick it
            candidates = []
            for root, _dirs, files in os.walk(wc_root):
                base_names = set()
                for name in files:
                    if ".r" in name:
                        base = name.split(".r")[0]
                        base_names.add(base)
                for base in base_names:
                    target = os.path.join(root, base)
                    if os.path.exists(target) and _has_svn_conflict_artifacts(target):
                        candidates.append(target)
                        if len(candidates) > 1:
                            return None
            if len(candidates) == 1:
                return candidates[0]
    except Exception:
        pass
    return None


def pick_files_or_conflict():
    root = tk.Tk()
    root.withdraw()

    auto = _auto_pick_conflict_file()
    if auto:
        conflict = _detect_svn_conflict_files(auto)
        if conflict:
            return ("merge",) + conflict + (True,)

    a = filedialog.askopenfilename(title="Select .xlsx file", filetypes=[("Excel Workbook", "*.xlsx")])
    if not a:
        return None

    conflict = _detect_svn_conflict_files(a)
    if conflict:
        return ("merge",) + conflict + (True,)

    b = filedialog.askopenfilename(title="Select second .xlsx file (same filename)", filetypes=[("Excel Workbook", "*.xlsx")])
    if not b:
        return None

    if os.path.basename(a).lower() != os.path.basename(b).lower():
        messagebox.showerror(
            "Filename mismatch",
            f"The two files must have the same filename.\n\nA: {os.path.basename(a)}\nB: {os.path.basename(b)}",
        )
        return None

    return ("diff", a, b)


def _atomic_save_wb(wb, target_path: str):
    """Safely overwrite a workbook."""
    folder = os.path.dirname(target_path)
    if folder:
        os.makedirs(folder, exist_ok=True)
    base = os.path.basename(target_path)
    tmp_path = os.path.join(folder, f"~{base}.{os.getpid()}.tmp")
    if _FAST_SAVE_VALUES_ONLY and _USE_CACHED_VALUES_ONLY:
        _save_values_only_from_wb(wb, tmp_path)
    else:
        wb.save(tmp_path)
    os.replace(tmp_path, target_path)


def _ensure_xlsx_copy(path: str) -> str:
    """If path is not .xlsx, copy to a temp .xlsx and return new path."""
    if not path:
        return path
    if os.path.splitext(path)[1].lower() == ".xlsx":
        return _ensure_stable_copy(path)
    try:
        base = os.path.basename(path)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        tmp = os.path.join(tempfile.gettempdir(), f"{APP_NAME}_svn_{base}_{ts}.xlsx")
        shutil.copy2(path, tmp)
        return tmp
    except Exception:
        return path


def _ensure_stable_copy(path: str) -> str:
    """If path looks like a temp/svn artifact, copy to a stable temp file."""
    if not path:
        return path
    try:
        temp_root = os.path.abspath(tempfile.gettempdir()).lower()
        p_abs = os.path.abspath(path)
        p_low = p_abs.lower()
        base = os.path.basename(path)
        looks_temp = p_low.startswith(temp_root)
        looks_svn = ".svn" in p_low or ".r" in base or "revbase" in base.lower() or "rev" in base.lower()
        if looks_temp or looks_svn:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            tmp = os.path.join(tempfile.gettempdir(), f"{APP_NAME}_stable_{ts}_{base}")
            if not tmp.lower().endswith(".xlsx"):
                tmp += ".xlsx"
            shutil.copy2(path, tmp)
            return tmp
    except Exception:
        pass
    return path


def _is_temp_base_path(path: str) -> bool:
    if not path:
        return False
    p = os.path.abspath(path).lower()
    base = os.path.basename(path).lower()
    if p.startswith(os.path.abspath(tempfile.gettempdir()).lower()):
        return True
    if "revbase" in base or ".svn" in p or base.endswith(".tmp.xlsx") or ".r" in base:
        return True
    return False


def _merge_three_way(base_path: str, mine_path: str, theirs_path: str, merged_path: str, save_merged: bool = True):
    """3-way merge: apply theirs onto mine when no conflict.

    Conflict: mine and theirs both changed a cell differently vs base.
    Returns (conflicts, merged_preview_path, conflict_cells_by_sheet).
    """
    # XLSM not supported in current scope
    for p in (base_path, mine_path, theirs_path, merged_path):
        if p and os.path.splitext(p)[1].lower() == ".xlsm":
            raise ValueError("当前版本暂不支持 .xlsm 文件合并")

    base_path = _ensure_xlsx_copy(base_path)
    mine_path = _ensure_xlsx_copy(mine_path)
    theirs_path = _ensure_xlsx_copy(theirs_path)

    base_val_path = _prepare_val_path(base_path)
    mine_val_path = _prepare_val_path(mine_path)
    theirs_val_path = _prepare_val_path(theirs_path)

    wb_base_val = load_workbook(base_val_path, data_only=True)
    wb_mine_val = load_workbook(mine_val_path, data_only=True)
    wb_theirs_val = load_workbook(theirs_val_path, data_only=True)
    wb_mine = load_workbook(mine_path, data_only=False)
    wb_base_edit = load_workbook(base_path, data_only=False)
    wb_theirs_edit = load_workbook(theirs_path, data_only=False)

    # Start merged as mine (copy in-memory by reusing workbook; then save to merged_path)
    wb_merged = wb_mine

    set_base = set(wb_base_val.sheetnames)
    set_mine = set(wb_mine_val.sheetnames)
    set_theirs = set(wb_theirs_val.sheetnames)

    # If a sheet exists only in theirs, copy values into merged.
    only_theirs = sorted(set_theirs - set_mine)
    for name in only_theirs:
        ws_t = wb_theirs_val[name]
        ws_m = wb_merged.create_sheet(title=name)
        max_row = ws_t.max_row or 1
        max_col = ws_t.max_column or 1
        for r in range(1, max_row + 1):
            row = next(ws_t.iter_rows(min_row=r, max_row=r, min_col=1, max_col=max_col, values_only=True), ())
            if len(row) < max_col:
                row = tuple(row) + (None,) * (max_col - len(row))
            for c, v in enumerate(row, start=1):
                ws_m.cell(row=r, column=c).value = v

    conflicts = []
    conflict_cells_by_sheet = {}

    common = sorted(set_mine & set_theirs)
    for name in common:
        ws_b = wb_base_val[name] if name in set_base else None
        ws_m_val = wb_mine_val[name]
        ws_t = wb_theirs_val[name]
        ws_m_edit = wb_mine[name]
        ws_b_edit = wb_base_edit[name] if name in wb_base_edit.sheetnames else None
        ws_t_edit = wb_theirs_edit[name] if name in wb_theirs_edit.sheetnames else None

        max_row = max(ws_m_val.max_row or 1, ws_t.max_row or 1, (ws_b.max_row or 1) if ws_b else 1)
        max_col = max(ws_m_val.max_column or 1, ws_t.max_column or 1, (ws_b.max_column or 1) if ws_b else 1)

        it_m = ws_m_val.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col, values_only=True)
        it_t = ws_t.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col, values_only=True)
        if ws_b:
            it_b = ws_b.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col, values_only=True)
        else:
            it_b = ((None,) * max_col for _ in range(max_row))

        for r, (row_m, row_t, row_b) in enumerate(zip(it_m, it_t, it_b), start=1):
            if len(row_m) < max_col:
                row_m = tuple(row_m) + (None,) * (max_col - len(row_m))
            if len(row_t) < max_col:
                row_t = tuple(row_t) + (None,) * (max_col - len(row_t))
            if len(row_b) < max_col:
                row_b = tuple(row_b) + (None,) * (max_col - len(row_b))

            for c, (vm, vt, vb) in enumerate(zip(row_m, row_t, row_b), start=1):
                # Normalize formula objects to text for comparison
                vm_cmp = vm
                vt_cmp = vt
                vb_cmp = vb
                # If cached value missing but edit cell has a literal (non-formula) value, use it
                try:
                    if vm_cmp is None:
                        vme = ws_m_edit.cell(row=r, column=c).value
                        if vme is not None and not _formula_text(vme):
                            vm_cmp = vme
                    if vt_cmp is None:
                        vte = ws_t_edit.cell(row=r, column=c).value if ws_t_edit is not None else None
                        if vte is not None and not _formula_text(vte):
                            vt_cmp = vte
                    if vb_cmp is None and ws_b_edit is not None:
                        vbe = ws_b_edit.cell(row=r, column=c).value
                        if vbe is not None and not _formula_text(vbe):
                            vb_cmp = vbe
                except Exception:
                    pass
                # If base cache is missing but base/theirs formulas are identical, treat base as theirs
                if vb_cmp is None and vt_cmp is not None and ws_b_edit is not None and ws_t_edit is not None:
                    try:
                        fb = _formula_text(ws_b_edit.cell(row=r, column=c).value)
                        ft = _formula_text(ws_t_edit.cell(row=r, column=c).value)
                        if fb and ft and fb == ft:
                            vb_cmp = vt_cmp
                    except Exception:
                        pass

                vm_key = _merge_cmp_value(vm_cmp)
                vt_key = _merge_cmp_value(vt_cmp)
                vb_key = _merge_cmp_value(vb_cmp)

                mine_changed = (vm_key != vb_key)
                theirs_changed = (vt_key != vb_key)
                if mine_changed and theirs_changed:
                    if vm_key != vt_key:
                        conflicts.append((name, r, c, vm_cmp, vt_cmp))
                        conflict_cells_by_sheet.setdefault(name, {}).setdefault(r, set()).add(c)
                    else:
                        # same change; keep as is
                        continue
                elif (not mine_changed) and theirs_changed:
                    # safe to apply theirs
                    ws_m_edit.cell(row=r, column=c).value = vt

    # Always save a preview for UI if needed
    if conflicts or (not save_merged):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        preview = os.path.join(tempfile.gettempdir(), f"{APP_NAME}_merged_preview_{os.getpid()}_{ts}.xlsx")
        _atomic_save_wb(wb_merged, preview)
        return conflicts, preview, conflict_cells_by_sheet

    # No conflicts: save directly to merged path
    _atomic_save_wb(wb_merged, merged_path)
    return [], None, {}


class SheetView:
    """TortoiseMerge-like side-by-side full-sheet viewer.

    Performance notes (optimized for responsiveness):
    - Avoids O(N) tag_remove across the whole document on every click.
    - Avoids per-cell ws.cell access loops during normal interactions.
    - Keeps per-row cached text and per-row diff columns; row merge refreshes only that row.
    """

    def __init__(self, parent, app, sheet_name: str):
        self.parent = parent
        self.app = app
        self.sheet = sheet_name
        # Support lazy tab containers: if parent is already a tab frame, reuse it.
        if isinstance(parent, ttk.Frame) and not parent.winfo_children():
            self.frame = parent
        else:
            self.frame = ttk.Frame(parent)

        self.max_row = 1
        self.max_col = 1
        self._bounds_checked = False

        # Cached row text and diff cols
        self.row_text_a: dict[int, str] = {}
        self.row_text_b: dict[int, str] = {}
        self.diff_cols_by_row: dict[int, set[int]] = {}
        self._display_diff_row_count: int = 0
        self._sample_scan_started = False
        # Row alignment (pair-wise) caches
        self.row_pairs: list[tuple[int | None, int | None]] = []
        self.pair_text_a: dict[int, str] = {}
        self.pair_text_b: dict[int, str] = {}
        self.pair_diff_cols: dict[int, set[int]] = {}
        self.row_a_to_pair_idx: dict[int, int] = {}
        self.row_b_to_pair_idx: dict[int, int] = {}

        # Render state
        # display_rows stores pair indices (into row_pairs)
        self.display_rows: list[int] = []
        self._full_display_rows: list[int] = []
        self._render_limit: int = _FAST_RENDER_ROW_LIMIT
        self.row_to_line: dict[int, int] = {}
        self._pending_yview: float | None = None
        self._render_cache = {}
        self._data_version = 0
        self.selected_excel_row: int | None = None
        self.selected_excel_row_a: int | None = None
        self.selected_excel_row_b: int | None = None
        self.selected_pair_idx: int | None = None
        self._last_selected_line: int | None = None
        self._is_large_sheet = False
        self._prefer_only_diff_when_ready = False
        self._diff_partial = False
        self._align_rows_enabled = True
        # Set to True once initial diff data has been computed (background or manual).
        # Prevents refresh(rescan=False) from triggering a full rescan on empty initial state.
        self._data_ready = False

        # Rows that were modified via overwrite in this session.
        # In "只看差异" mode, we keep these rows visible even if diffs are resolved.
        self.touched_rows: set[int] = set()

        # Snapshot mode: build the diff row list once, then keep the row list stable.
        # Overwrites only update per-row highlight (to show "已处理") and keep the row visible.
        self.snapshot_only_diff = True

        # Toolbar
        bar = ttk.Frame(self.frame)
        bar.pack(fill="x", padx=8, pady=(8, 6))

        ttk.Label(bar, text=f"Sheet: {sheet_name}", font=("Segoe UI", 11, "bold")).pack(side="left")
        self.info = ttk.Label(bar, text="", foreground="#444")
        self.info.pack(side="left", padx=(10, 0))

        # Diff block navigation (fixed position on the right; does not shift with label lengths)
        self.next_diff_btn = tk.Button(bar, text="下一处差异", padx=10, pady=2, command=self._goto_next_diff_block)
        self.prev_diff_btn = tk.Button(bar, text="上一处差异", padx=10, pady=2, command=self._goto_prev_diff_block)
        # Pack on right so it stays at a stable location across sheets
        self.next_diff_btn.pack(side="right", padx=(6, 0))
        self.prev_diff_btn.pack(side="right", padx=(6, 0))
        self._diff_blocks_cache = []

        # Some environments fail to toggle BooleanVar reliably; use IntVar with explicit on/off values.
        self.only_diff_var = tk.IntVar(value=int(getattr(self.app, "only_diff_default", 0)))

        self.only_diff_cb = tk.Checkbutton(
            bar,
            text="只看差异内容",
            variable=self.only_diff_var,
            onvalue=1,
            offvalue=0,
            command=self._toggle_only_diff,
            padx=6,
        )
        # Put on the right for a stable position
        self.only_diff_cb.pack(side="right", padx=(6, 0))
        if getattr(self.app, "merge_conflict_mode", False):
            try:
                self.only_diff_var.set(1)
                self.only_diff_cb.select()
                self.only_diff_cb.configure(state="disabled")
            except Exception:
                pass

        # Apply initial visual state from persisted setting
        try:
            if self.only_diff_var.get():
                self.only_diff_cb.select()
            else:
                self.only_diff_cb.deselect()
        except Exception:
            pass
        self._last_only_diff_value = int(self.only_diff_var.get())

        # Debug: provide a force-toggle button to prove the filtering path works even if UI toggling fails.
        if _DEBUG_ENABLED:
            tk.Button(
                bar,
                text="强制切换",
                command=lambda: (self.only_diff_var.set(0 if self.only_diff_var.get() else 1), self._toggle_only_diff()),
                padx=6,
                pady=1,
            ).pack(side="right", padx=(6, 0))

        # Debug: log click + resulting value
        def _log_cb_click(_evt=None):
            _dlog(f"CHECKBOX_CLICK sheet={self.sheet} var={self.only_diff_var.get()}")
            try:
                self.frame.after_idle(lambda: _dlog(f"CHECKBOX_AFTER_IDLE sheet={self.sheet} var={self.only_diff_var.get()}"))
            except Exception:
                pass

        try:
            self.only_diff_cb.bind("<ButtonRelease-1>", _log_cb_click)
        except Exception:
            pass

        # Context merge buttons (always visible)
        # No UI state changes on selection; logic will no-op if there is no diff.
        self.use_left_btn = tk.Button(
            bar,
            text="使用左侧(A)",
            bg="#eaf2ff",
            padx=10,
            pady=2,
            command=lambda: self._copy_selected_row("A2B"),
        )
        self.use_right_btn = tk.Button(
            bar,
            text="使用右侧(B)",
            bg="#ffecec",
            padx=10,
            pady=2,
            command=lambda: self._copy_selected_row("B2A"),
        )
        self.undo_btn = tk.Button(
            bar,
            text="回退",
            bg="#f2f2f2",
            padx=8,
            pady=2,
            command=self._undo_last_action,
        )
        if getattr(self.app, "merge_conflict_mode", False):
            try:
                self.use_left_btn.configure(text="保留我的(A)")
                self.use_right_btn.configure(text="采用对方(B)")
            except Exception:
                pass
        # Keep at top-right (avoid misclick)
        self.use_right_btn.pack(side="right", padx=(6, 0))
        self.use_left_btn.pack(side="right")
        self.undo_btn.pack(side="right", padx=(6, 0))

        ttk.Button(bar, text="刷新本Sheet", command=lambda: self.refresh(row_only=None, rescan=True)).pack(side="right", padx=(6, 0))
        self._full_render = False
        self._load_all_btn = ttk.Button(bar, text="加载全部", command=self._load_all_rows)
        if _FAST_OPEN_ENABLED:
            self._load_all_btn.pack(side="right", padx=(6, 0))

        # Path bar (requested red-box area): show full paths above the diff panes
        path_bar = ttk.Frame(self.frame)
        path_bar.pack(fill="x", padx=8, pady=(0, 4))

        self._path_font = ("Segoe UI", 9)
        path_bar.grid_columnconfigure(0, weight=1)
        path_bar.grid_columnconfigure(1, weight=1)

        self._path_font = ("Segoe UI", 9, "bold")
        label_a = self.app.file_a
        label_b = self.app.file_b
        if getattr(self.app, "merge_conflict_mode", False):
            label_a = f"临时合并预览(merged预览): {self.app.file_a}"
            label_b = f"对方(theirs): {self.app.file_b}"
        tk.Label(
            path_bar,
            text=label_a,
            font=self._path_font,
            bg="#eaf2ff",
            anchor="center",
            padx=6,
            pady=2,
        ).grid(row=0, column=0, sticky="ew")
        tk.Label(
            path_bar,
            text=label_b,
            font=self._path_font,
            bg="#ffecec",
            anchor="center",
            padx=6,
            pady=2,
        ).grid(row=0, column=1, sticky="ew")

        # Extra vertical scrollbar (left side) for convenience; controls both panes.
        # NOTE: must be packed BEFORE the paned window so it remains visible.
        self.vsb_left = ttk.Scrollbar(self.frame, orient="vertical", command=self._yview_both)
        self.vsb_left.pack(side="left", fill="y")

        # Panes
        paned = ttk.PanedWindow(self.frame, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self._main_paned = paned

        left_wrap = ttk.Frame(paned)
        right_wrap = ttk.Frame(paned)
        paned.add(left_wrap, weight=1)
        paned.add(right_wrap, weight=1)

        def _keep_panes_equal(_evt=None):
            # Keep A/B content panes at 50:50 to avoid visual width mismatch.
            try:
                total = self._main_paned.winfo_width()
                if total and total > 2:
                    self._main_paned.sashpos(0, total // 2)
            except Exception:
                pass

        self._keep_panes_equal = _keep_panes_equal
        self._main_paned.bind("<Configure>", self._keep_panes_equal)
        self._main_paned.bind("<ButtonRelease-1>", self._keep_panes_equal)
        self.frame.after(0, self._keep_panes_equal)

        ttk.Label(left_wrap, text="A(左)", background="#eaf2ff").pack(fill="x")
        ttk.Label(right_wrap, text="B(右)", background="#ffecec").pack(fill="x")

        # Font size tuned closer to TortoiseMerge (+~20%)
        self.editor_font = ("Consolas", 11)
        self.left = tk.Text(left_wrap, wrap="none", undo=False, font=self.editor_font)
        self.right = tk.Text(right_wrap, wrap="none", undo=False, font=self.editor_font)

        # Scrollbars
        # Per-pane vertical scrollbars (user requested visible scrollbars on both A and B)
        self.vsb_a = ttk.Scrollbar(left_wrap, orient="vertical", command=self._yview_both)
        self.vsb_b = ttk.Scrollbar(right_wrap, orient="vertical", command=self._yview_both)
        self.vsb_a.pack(side="right", fill="y")
        self.vsb_b.pack(side="right", fill="y")

        # Shared vertical scrollbar on far right (keep)
        self.vsb = ttk.Scrollbar(self.frame, orient="vertical", command=self._yview_both)
        self.vsb.pack(side="right", fill="y")

        # Horizontal scroll sync: keep A/B panes aligned when scrolling horizontally.
        self._xsyncing = False

        def _xscroll_left(first, last):
            # called when left xview changes
            if self._xsyncing:
                self.hsb_left.set(first, last)
                return
            self._xsyncing = True
            try:
                self.hsb_left.set(first, last)
                self.right.xview_moveto(first)
                # ensure right scrollbar matches
                rf, rl = self.right.xview()
                self.hsb_right.set(rf, rl)
            finally:
                self._xsyncing = False

        def _xscroll_right(first, last):
            if self._xsyncing:
                self.hsb_right.set(first, last)
                return
            self._xsyncing = True
            try:
                self.hsb_right.set(first, last)
                self.left.xview_moveto(first)
                lf, ll = self.left.xview()
                self.hsb_left.set(lf, ll)
            finally:
                self._xsyncing = False

        def _xview_left(*args):
            # scrollbar drag/click on left
            self._xsyncing = True
            try:
                self.left.xview(*args)
                first, last = self.left.xview()
                self.right.xview_moveto(first)
                self.hsb_left.set(first, last)
                rf, rl = self.right.xview()
                self.hsb_right.set(rf, rl)
            finally:
                self._xsyncing = False

        def _xview_right(*args):
            self._xsyncing = True
            try:
                self.right.xview(*args)
                first, last = self.right.xview()
                self.left.xview_moveto(first)
                self.hsb_right.set(first, last)
                lf, ll = self.left.xview()
                self.hsb_left.set(lf, ll)
            finally:
                self._xsyncing = False

        self.hsb_left = ttk.Scrollbar(left_wrap, orient="horizontal", command=_xview_left)
        self.hsb_right = ttk.Scrollbar(right_wrap, orient="horizontal", command=_xview_right)
        self.left.configure(xscrollcommand=_xscroll_left)
        self.right.configure(xscrollcommand=_xscroll_right)

        self.left.configure(yscrollcommand=self._yscroll_left)
        self.right.configure(yscrollcommand=self._yscroll_right)
        self.vsb.configure(command=self._yview_both)
        self.vsb_left.configure(command=self._yview_both)
        self.vsb_a.configure(command=self._yview_both)
        self.vsb_b.configure(command=self._yview_both)

        self.left.pack(fill="both", expand=True)
        self.hsb_left.pack(fill="x")

        # Save action row: keep a fixed height on both sides so horizontal
        # scrollbars stay aligned even when only one side has a button.
        save_row_height = 34

        # Save A button (bottom-right of A pane)
        save_a_row = ttk.Frame(left_wrap, height=save_row_height)
        save_a_row.pack(fill="x", pady=(2, 0))
        save_a_row.pack_propagate(False)
        if getattr(self.app, "merge_mode", False):
            tk.Button(save_a_row, text="保存Merged并退出", bg="#eaf2ff", padx=14, pady=4,
                      command=self.app.save_merged_and_exit).pack(side="right")
        else:
            if not _is_temp_base_path(getattr(self.app, "file_a", "")):
                tk.Button(save_a_row, text="保存A", bg="#eaf2ff", padx=14, pady=4,
                          command=self.app.save_a_inplace).pack(side="right")

        self.right.pack(fill="both", expand=True)
        self.hsb_right.pack(fill="x")

        # Save B button (bottom-right of B pane)
        save_b_row = ttk.Frame(right_wrap, height=save_row_height)
        save_b_row.pack(fill="x", pady=(2, 0))
        save_b_row.pack_propagate(False)
        if not getattr(self.app, "merge_mode", False):
            tk.Button(save_b_row, text="保存B", bg="#ffecec", padx=14, pady=4,
                      command=self.app.save_b_inplace).pack(side="right")

        # Tags (order matters: diffcell should be applied after diffrow)
        # Closer to TortoiseMerge vibe: left diff block = orange, right diff block = yellow
        self.left.tag_configure("diffrow", background="#F6C16B")
        self.right.tag_configure("diffrow", background="#FFF176")

        # Cell-level highlight (red) for exact diffs
        self.left.tag_configure("diffcell", background="#FF2D2D")
        self.right.tag_configure("diffcell", background="#FF2D2D")

        # Alignment padding: grey slot for rows that exist only on the other side.
        # tag_raise ensures paddingrow background overrides diffrow on the empty slot.
        self.left.tag_configure("paddingrow", background="#A0A0A0")
        self.right.tag_configure("paddingrow", background="#A0A0A0")
        self.left.tag_raise("paddingrow")
        self.right.tag_raise("paddingrow")

        # selection should not overwrite diff colors
        self.left.tag_configure("selrow", underline=1, font=("Consolas", 11, "bold"))
        self.right.tag_configure("selrow", underline=1, font=("Consolas", 11, "bold"))

        # Bindings
        self._syncing = False
        for w in (self.left, self.right):
            w.bind("<MouseWheel>", self._on_mousewheel)
            w.bind("<Button-4>", self._on_mousewheel)
            w.bind("<Button-5>", self._on_mousewheel)
            w.bind("<KeyRelease>", lambda e: self._update_cursor_lines())
            w.bind("<ButtonRelease-1>", lambda e: self._update_cursor_lines())
            if getattr(self.app, "merge_conflict_mode", False):
                #快捷键：下一处/上一处冲突
                w.bind("<F4>", lambda e: (self._goto_next_diff_block(), "break"))
                w.bind("<Shift-F4>", lambda e: (self._goto_prev_diff_block(), "break"))
                w.bind("<Control-n>", lambda e: (self._goto_next_diff_block(), "break"))
                w.bind("<Control-p>", lambda e: (self._goto_prev_diff_block(), "break"))

        # Click handling (selection + arrow action)
        self.left.bind("<Button-1>", lambda e: self._on_click_with_arrow(self.left, e, "A2B"))
        self.right.bind("<Button-1>", lambda e: self._on_click_with_arrow(self.right, e, "B2A"))
        # Hover arrows for row merge
        self._hover_line_left = None
        self._hover_line_right = None
        self._left_cursor_default = self.left.cget("cursor")
        self._right_cursor_default = self.right.cget("cursor")
        self.left.bind("<Motion>", lambda e: self._on_hover(self.left, e, "A2B"))
        self.right.bind("<Motion>", lambda e: self._on_hover(self.right, e, "B2A"))
        self.left.bind("<Leave>", lambda e: self._clear_hover(self.left))
        self.right.bind("<Leave>", lambda e: self._clear_hover(self.right))

        # Double-click merge (single cell)
        self.left.bind("<Double-Button-1>", lambda e: self._copy_cell("A2B", e))
        self.right.bind("<Double-Button-1>", lambda e: self._copy_cell("B2A", e))

        # Sync bar (append right-side data to the end of left-side)
        sync_bar = ttk.Frame(self.frame)
        sync_bar.pack(fill="x", padx=8, pady=(0, 6))
        self.sync_all_btn = tk.Button(
            sync_bar,
            text="同步所有",
            bg="#ffecec",
            padx=10,
            pady=2,
            command=self._append_right_all_to_left_end,
        )
        self.sync_one_btn = tk.Button(
            sync_bar,
            text="同步单行",
            bg="#eaf2ff",
            padx=10,
            pady=2,
            command=self._append_right_row_to_left_end,
        )
        self.sync_all_btn.pack(side="right", padx=(6, 0))
        self.sync_one_btn.pack(side="right")

        # C区: compact cursor compare block + cell-aligned view
        self.c_area = ttk.Notebook(self.frame)
        self.c_area.pack(fill="x", padx=8, pady=(0, 4))

        # ---- C1: two-line row text (A on top, B bottom) ----
        c_text_frame = ttk.Frame(self.c_area)
        self.c_area.add(c_text_frame, text="C区-行对比")

        self.cursor_cmp = tk.Text(c_text_frame, height=2, wrap="none", font=self.editor_font, bd=1, relief="solid")
        # Make base colors stronger (user feedback: previous too light)
        self.cursor_cmp.tag_configure("a", background="#CFE3FF")
        self.cursor_cmp.tag_configure("b", background="#FFD1D1")
        # Diff cell highlight (match main panes)
        self.cursor_cmp.tag_configure("diffcell", background="#FF2D2D")
        self.cursor_cmp.pack(side="top", fill="x", expand=True)

        # Horizontal scrollbar for C区行对比
        self.cursor_hsb = ttk.Scrollbar(c_text_frame, orient="horizontal", command=self.cursor_cmp.xview)
        self.cursor_cmp.configure(xscrollcommand=self.cursor_hsb.set)
        self.cursor_hsb.pack(side="top", fill="x")

        # ---- C2: cell-aligned view (optional; can be hidden if not useful/performance) ----
        self._enable_c_cell = False  # user feedback: not useful; keep hidden by default
        c_cell_frame = ttk.Frame(self.c_area)
        self.c_area.add(c_cell_frame, text="C区-单元格对齐")
        if not self._enable_c_cell:
            try:
                self.c_area.tab(c_cell_frame, state="hidden")
            except Exception:
                pass

        top_row = ttk.Frame(c_cell_frame)
        top_row.pack(fill="x", pady=(2, 2))
        self.c_only_diff_cells = tk.IntVar(value=1)
        tk.Checkbutton(
            top_row,
            text="只显示差异单元格",
            variable=self.c_only_diff_cells,
            onvalue=1,
            offvalue=0,
            command=lambda: self._update_cursor_lines(),
        ).pack(side="left")

        self.cell_cmp_text = tk.Text(c_cell_frame, height=6, wrap="none", font=self.editor_font, bd=1, relief="solid")
        self.cell_cmp_text.tag_configure("a", background="#CFE3FF")
        self.cell_cmp_text.tag_configure("b", background="#FFD1D1")
        self.cell_cmp_text.tag_configure("diffcell", background="#FF2D2D")

        self.cell_cmp_hsb = ttk.Scrollbar(c_cell_frame, orient="horizontal", command=self.cell_cmp_text.xview)
        self.cell_cmp_text.configure(xscrollcommand=self.cell_cmp_hsb.set)

        self.cell_cmp_text.pack(side="top", fill="x", expand=True)
        self.cell_cmp_hsb.pack(side="top", fill="x")

        # initial render should respect the persisted only-diff setting
        # Defer heavy initial refresh; SowMergeApp will lazy-load the active sheet.
        # (Still create the UI widgets now.)
        # self.refresh(row_only=None, rescan=True)
        # self._update_cursor_lines()

    # ---------- Scrolling sync ----------
    def _yscroll_all(self, first, last):
        for sb in (self.vsb, self.vsb_left, self.vsb_a, self.vsb_b):
            try:
                sb.set(first, last)
            except Exception:
                pass

    def _yscroll_left(self, first, last):
        if self._syncing:
            return
        self._syncing = True
        self.right.yview_moveto(first)
        self._yscroll_all(first, last)
        self._syncing = False
        self._maybe_load_more_rows(last)

    def _yscroll_right(self, first, last):
        if self._syncing:
            return
        self._syncing = True
        self.left.yview_moveto(first)
        self._yscroll_all(first, last)
        self._syncing = False
        self._maybe_load_more_rows(last)

    def _yview_both(self, *args):
        self._syncing = True
        self.left.yview(*args)
        self.right.yview(*args)
        try:
            first, last = self.left.yview()
            self._yscroll_all(first, last)
        except Exception:
            pass
        self._syncing = False
        try:
            _first, last = self.left.yview()
            self._maybe_load_more_rows(last)
        except Exception:
            pass

    def _on_mousewheel(self, event):
        if getattr(event, "num", None) == 4:
            delta = 120
        elif getattr(event, "num", None) == 5:
            delta = -120
        else:
            delta = event.delta
        steps = int(-1 * (delta / 120))
        self._yview_both("scroll", steps, "units")
        return "break"

    # ---------- Selection + toolbar buttons ----------
    def _widget_line(self, w: tk.Text):
        try:
            idx = w.index("@%d,%d" % (w.winfo_pointerx() - w.winfo_rootx(), w.winfo_pointery() - w.winfo_rooty()))
        except Exception:
            idx = w.index("insert")
        return int(str(idx).split(".")[0])

    def _pair_idx_for_line(self, line: int) -> int | None:
        if not (1 <= line <= len(self.display_rows)):
            return None
        return self.display_rows[line - 1]

    def _pair_for_line(self, line: int):
        idx = self._pair_idx_for_line(line)
        if idx is None or idx >= len(self.row_pairs):
            return None
        return self.row_pairs[idx]

    @staticmethod
    def _row_for_side(pair, side: str) -> int | None:
        if not pair:
            return None
        return pair[0] if side == "A" else pair[1]

    def _select_from_widget(self, w: tk.Text, event):
        # Set insert mark to the clicked position so follow-up actions can use it.
        try:
            idx = w.index(f"@{event.x},{event.y}")
            w.mark_set("insert", idx)
        except Exception:
            idx = None

        line = self._widget_line(w)

        # Keep both panes aligned for the cursor compare block:
        # when user clicks either side, set BOTH cursors to the same line.
        other = self.right if w is self.left else self.left
        try:
            other.mark_set("insert", f"{line}.0")
        except Exception:
            pass

        self._highlight_selected_line(line)
        pair = self._pair_for_line(line)
        self.selected_pair_idx = self._pair_idx_for_line(line)
        self.selected_excel_row_a = self._row_for_side(pair, "A")
        self.selected_excel_row_b = self._row_for_side(pair, "B")
        self.selected_excel_row = self.selected_excel_row_a or self.selected_excel_row_b
        # No button state updates (performance): buttons are always visible and logic no-ops when no diff.

        self._update_cursor_lines()
        self._update_diff_nav_state()

    def _on_click_with_arrow(self, w: tk.Text, event, direction: str):
        # select row first
        self._select_from_widget(w, event)
        try:
            idx = w.index(f"@{event.x},{event.y}")
            line = int(idx.split(".")[0])
            col = int(idx.split(".")[1])
        except Exception:
            return
        if not (1 <= line <= len(self.display_rows)):
            return
        pair = self._pair_for_line(line)
        r = self._row_for_side(pair, "A" if w is self.left else "B")
        pair_idx = self._pair_idx_for_line(line)
        cols = self.pair_diff_cols.get(pair_idx, set()) if pair_idx is not None else set()
        if not cols:
            return
        if r is None:
            return
        rownum_len = len(str(r))
        if col <= rownum_len:
            self._copy_selected_row(direction)

    def _on_hover(self, w: tk.Text, event, direction: str):
        try:
            idx = w.index(f"@{event.x},{event.y}")
            line = int(idx.split(".")[0])
            col = int(idx.split(".")[1])
        except Exception:
            self._clear_hover(w)
            return
        if not (1 <= line <= len(self.display_rows)):
            self._clear_hover(w)
            return
        pair = self._pair_for_line(line)
        r = self._row_for_side(pair, "A" if w is self.left else "B")
        pair_idx = self._pair_idx_for_line(line)
        cols = self.pair_diff_cols.get(pair_idx, set()) if pair_idx is not None else set()
        if not cols:
            self._clear_hover(w)
            return
        if r is None:
            self._clear_hover(w)
            return
        rownum_len = len(str(r))
        if col > rownum_len:
            self._clear_hover(w)
            return
        self._show_hover_arrow(w, line, r, direction)

    def _clear_hover(self, w: tk.Text):
        line = self._hover_line_left if w is self.left else self._hover_line_right
        if line is None:
            return
        self._restore_rownum(w, line)
        try:
            if w is self.left:
                w.configure(cursor=self._left_cursor_default)
            else:
                w.configure(cursor=self._right_cursor_default)
        except Exception:
            pass
        if w is self.left:
            self._hover_line_left = None
        else:
            self._hover_line_right = None

    def _show_hover_arrow(self, w: tk.Text, line: int, r: int, direction: str):
        if w is self.left:
            if self._hover_line_left == line:
                return
        else:
            if self._hover_line_right == line:
                return
        # restore previous
        self._clear_hover(w)
        self._replace_rownum_with_arrow(w, line, r, direction)
        try:
            w.configure(cursor="hand2")
        except Exception:
            pass
        if w is self.left:
            self._hover_line_left = line
        else:
            self._hover_line_right = line

    def _replace_rownum_with_arrow(self, w: tk.Text, line: int, r: int, direction: str):
        arrow = "→" if direction == "A2B" else "←"
        rownum = str(r)
        new_label = arrow + (" " * max(0, len(rownum) - 1))
        start = f"{line}.0"
        end = f"{line}.{len(rownum)}"
        try:
            w.delete(start, end)
            w.insert(start, new_label)
        except Exception:
            pass

    def _restore_rownum(self, w: tk.Text, line: int):
        if not (1 <= line <= len(self.display_rows)):
            return
        pair = self._pair_for_line(line)
        r = self._row_for_side(pair, "A" if w is self.left else "B")
        if r is None:
            return
        rownum = str(r)
        start = f"{line}.0"
        end = f"{line}.{len(rownum)}"
        try:
            w.delete(start, end)
            w.insert(start, rownum)
        except Exception:
            pass

    def _highlight_selected_line(self, line: int):
        # Remove highlight only from the previously selected line (O(1))
        if self._last_selected_line is not None:
            prev = self._last_selected_line
            for t in (self.left, self.right):
                t.tag_remove("selrow", f"{prev}.0", f"{prev}.end")
        for t in (self.left, self.right):
            t.tag_add("selrow", f"{line}.0", f"{line}.end")
        self._last_selected_line = line

    def _update_cursor_lines(self):
        """Update per-sheet compact A/B compare block (2 lines).

        - Line 1 is A, line 2 is B
        - Highlight differing cells with background color matching the main panes
        """
        try:
            la = int(self.left.index("insert").split(".")[0])
            lb = int(self.right.index("insert").split(".")[0])

            a_text = self.left.get(f"{la}.0", f"{la}.end") if la >= 1 else ""
            b_text = self.right.get(f"{lb}.0", f"{lb}.end") if lb >= 1 else ""

            # Determine selected pair (based on line in the view)
            pair_idx = self._pair_idx_for_line(la)
            pair = self.row_pairs[pair_idx] if pair_idx is not None and pair_idx < len(self.row_pairs) else None
            diff_cols = self.pair_diff_cols.get(pair_idx, set()) if pair_idx is not None else set()

            # Force a strict 2-line rendering: line1=A, line2=B
            self.cursor_cmp.configure(state="normal")
            self.cursor_cmp.delete("1.0", "end")
            self.cursor_cmp.insert("1.0", f"{a_text}\n{b_text}")

            # Clear & apply base tags
            self.cursor_cmp.tag_remove("a", "1.0", "end")
            self.cursor_cmp.tag_remove("b", "1.0", "end")
            self.cursor_cmp.tag_remove("diffcell", "1.0", "end")
            self.cursor_cmp.tag_add("a", "1.0", "1.end")
            self.cursor_cmp.tag_add("b", "2.0", "2.end")

            # Cell-level diff highlight
            if diff_cols:
                spans_a = self._spans_for_line(a_text)
                spans_b = self._spans_for_line(b_text)
                for c in diff_cols:
                    if c in spans_a:
                        s, e = spans_a[c]
                        self.cursor_cmp.tag_add("diffcell", f"1.{s}", f"1.{e}")
                    if c in spans_b:
                        s, e = spans_b[c]
                        self.cursor_cmp.tag_add("diffcell", f"2.{s}", f"2.{e}")

            self.cursor_cmp.configure(state="disabled")

            # ---- Update C区单元格对齐（可选） ----
            if getattr(self, "_enable_c_cell", False) and hasattr(self, "cell_cmp_text"):
                try:
                    self.cell_cmp_text.configure(state="normal")
                    self.cell_cmp_text.delete("1.0", "end")
                    self.cell_cmp_text.tag_remove("a", "1.0", "end")
                    self.cell_cmp_text.tag_remove("b", "1.0", "end")
                    self.cell_cmp_text.tag_remove("diffcell", "1.0", "end")

                    if pair is not None:
                        ws_a_val = self.app.ws_a_val(self.sheet)
                        ws_b_val = self.app.ws_b_val(self.sheet)
                        ra = self._row_for_side(pair, "A")
                        rb = self._row_for_side(pair, "B")

                        show_only_diff = bool(self.c_only_diff_cells.get())
                        cols_to_show = sorted(diff_cols) if show_only_diff else list(range(1, self.max_col + 1))

                        if show_only_diff:
                            parts_a = []
                            parts_b = []
                            for c in cols_to_show:
                                va = ws_a_val.cell(row=ra, column=c).value if ra is not None else None
                                vb = ws_b_val.cell(row=rb, column=c).value if rb is not None else None
                                parts_a.append(_val_to_str(va))
                                parts_b.append(_val_to_str(vb))

                            a_line = "\t".join(parts_a)
                            b_line = "\t".join(parts_b)
                            self.cell_cmp_text.insert("end", a_line + "\n" + b_line + "\n")

                            self.cell_cmp_text.tag_add("a", "1.0", "1.end")
                            self.cell_cmp_text.tag_add("b", "2.0", "2.end")

                            if a_line:
                                spans = self._spans_for_line("0\t" + a_line)
                                for idx in range(1, len(parts_a) + 1):
                                    if idx in spans:
                                        s, e = spans[idx]
                                        self.cell_cmp_text.tag_add("diffcell", f"1.{s}", f"1.{e}")
                            if b_line:
                                spans = self._spans_for_line("0\t" + b_line)
                                for idx in range(1, len(parts_b) + 1):
                                    if idx in spans:
                                        s, e = spans[idx]
                                        self.cell_cmp_text.tag_add("diffcell", f"2.{s}", f"2.{e}")
                        else:
                            line_no = 1
                            for c in cols_to_show:
                                va = ws_a_val.cell(row=ra, column=c).value if ra is not None else None
                                vb = ws_b_val.cell(row=rb, column=c).value if rb is not None else None
                                a_s = _val_to_str(va)
                                b_s = _val_to_str(vb)

                                self.cell_cmp_text.insert("end", a_s + "\n")
                                self.cell_cmp_text.insert("end", b_s + "\n")

                                self.cell_cmp_text.tag_add("a", f"{line_no}.0", f"{line_no}.end")
                                self.cell_cmp_text.tag_add("b", f"{line_no+1}.0", f"{line_no+1}.end")

                                if va != vb:
                                    self.cell_cmp_text.tag_add("diffcell", f"{line_no}.0", f"{line_no}.end")
                                    self.cell_cmp_text.tag_add("diffcell", f"{line_no+1}.0", f"{line_no+1}.end")

                                line_no += 2

                    self.cell_cmp_text.configure(state="disabled")
                except Exception:
                    try:
                        self.cell_cmp_text.configure(state="disabled")
                    except Exception:
                        pass
        except Exception:
            pass

    def _update_merge_buttons_for_row(self, excel_row: int):
        # Buttons are always visible; no UI updates needed.
        return

    # ---------- Diff block navigation ----------
    def _compute_diff_blocks(self):
        """Return list of (start_line, end_line) diff blocks in current view."""
        blocks = []
        start = None
        for line_idx, pair_idx in enumerate(self.display_rows, start=1):
            has = bool(self.pair_diff_cols.get(pair_idx, set()))
            if has and start is None:
                start = line_idx
            elif (not has) and start is not None:
                blocks.append((start, line_idx - 1))
                start = None
        if start is not None:
            blocks.append((start, len(self.display_rows)))
        self._diff_blocks_cache = blocks
        return blocks

    def _current_line(self) -> int:
        try:
            return int(self.left.index("insert").split(".")[0])
        except Exception:
            return 1

    def _update_diff_nav_state(self):
        blocks = self._compute_diff_blocks()
        if not blocks:
            self.prev_diff_btn.configure(state="disabled")
            self.next_diff_btn.configure(state="disabled")
            return

        cur = self._current_line()
        has_prev = any(b[0] < cur for b in blocks)
        has_next = any(b[0] > cur for b in blocks)
        self.prev_diff_btn.configure(state=("normal" if has_prev else "disabled"))
        self.next_diff_btn.configure(state=("normal" if has_next else "disabled"))

    def _goto_block_start(self, start_line: int):
        # Scroll so the line is visible
        try:
            for w in (self.left, self.right):
                w.mark_set("insert", f"{start_line}.0")
                w.see(f"{start_line}.0")
            self._highlight_selected_line(start_line)
            pair = self._pair_for_line(start_line)
            self.selected_pair_idx = self._pair_idx_for_line(start_line)
            self.selected_excel_row_a = self._row_for_side(pair, "A")
            self.selected_excel_row_b = self._row_for_side(pair, "B")
            self.selected_excel_row = self.selected_excel_row_a or self.selected_excel_row_b
            self._update_cursor_lines()
        except Exception:
            pass
        self._update_diff_nav_state()

    def _goto_next_diff_block(self):
        blocks = self._compute_diff_blocks()
        cur = self._current_line()
        for start, _end in blocks:
            if start > cur:
                self._goto_block_start(start)
                return
        self._update_diff_nav_state()

    def _goto_prev_diff_block(self):
        blocks = self._compute_diff_blocks()
        cur = self._current_line()
        prev = None
        for start, _end in blocks:
            if start < cur:
                prev = start
            else:
                break
        if prev is not None:
            self._goto_block_start(prev)
        self._update_diff_nav_state()

    # ---------- Diff calculation helpers ----------
    def _get_row_values(self, ws, r: int):
        # Fast row read using iter_rows(values_only=True)
        try:
            row = next(ws.iter_rows(min_row=r, max_row=r, min_col=1, max_col=self.max_col, values_only=True))
        except StopIteration:
            row = ()
        if row is None:
            row = ()
        # Ensure length == max_col
        if len(row) < self.max_col:
            row = tuple(row) + (None,) * (self.max_col - len(row))
        return row

    def _show_loading(self):
        """Show a loading placeholder while background diff computation is in progress."""
        try:
            for w in (self.left, self.right):
                w.configure(state="normal")
                w.delete("1.0", "end")
                w.insert("1.0", "计算中...\n")
            self.info.configure(text="正在后台计算差异，请稍候...")
        except Exception:
            pass

    @staticmethod
    def _row_label(r: int | None) -> str:
        return str(r) if r is not None else ""

    def _build_line_from_row_label(self, label: str, row_vals) -> str:
        return label + "\t" + "\t".join(_val_to_str(v) for v in row_vals)

    def _build_row_and_diff_pair(self, ws_a_val, ws_b_val, ws_a_edit, ws_b_edit, ra: int | None, rb: int | None):
        parts_a = []
        parts_b = []
        cols = set()
        for c in range(1, self.max_col + 1):
            da, db, eq = _cell_display_and_equal_by_row(ws_a_val, ws_b_val, ws_a_edit, ws_b_edit, ra, rb, c)
            parts_a.append(_val_to_str(da))
            parts_b.append(_val_to_str(db))
            if not eq:
                cols.add(c)
        line_a = self._row_label(ra) + "\t" + "\t".join(parts_a)
        line_b = self._row_label(rb) + "\t" + "\t".join(parts_b)
        return line_a, line_b, cols

    def _build_row_pairs(self, ws_a_val, ws_b_val):
        # Align rows between A and B to avoid cascading diffs on insert/delete.
        max_row_a = ws_a_val.max_row or 1
        max_row_b = ws_b_val.max_row or 1
        max_row = max(max_row_a, max_row_b)
        if max_row <= 0:
            return []

        def _row_sig_list(ws, max_row_local: int):
            # Read all rows in one pass (much faster than per-row iter_rows calls)
            try:
                all_rows = list(ws.iter_rows(
                    min_row=1, max_row=max_row_local,
                    min_col=1, max_col=self.max_col,
                    values_only=True,
                ))
            except Exception:
                all_rows = []
            sigs = []
            for row in all_rows:
                if row is None:
                    row = ()
                sigs.append("\x1f".join(_merge_cmp_value(v) for v in row))
            return sigs

        sig_a = _row_sig_list(ws_a_val, max_row_a)
        sig_b = _row_sig_list(ws_b_val, max_row_b)

        sm = difflib.SequenceMatcher(a=sig_a, b=sig_b, autojunk=False)
        pairs: list[tuple[int | None, int | None]] = []
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                for i, j in zip(range(i1, i2), range(j1, j2)):
                    pairs.append((i + 1, j + 1))
            elif tag == "replace":
                len_a = i2 - i1
                len_b = j2 - j1
                common = min(len_a, len_b)
                for k in range(common):
                    pairs.append((i1 + k + 1, j1 + k + 1))
                for k in range(common, len_a):
                    pairs.append((i1 + k + 1, None))
                for k in range(common, len_b):
                    pairs.append((None, j1 + k + 1))
            elif tag == "delete":
                for i in range(i1, i2):
                    pairs.append((i + 1, None))
            elif tag == "insert":
                for j in range(j1, j2):
                    pairs.append((None, j + 1))
        return pairs

    @staticmethod
    def _build_row_pairs_direct(max_row_a: int, max_row_b: int):
        """Direct row pairing (1:1 by row number), used for very large sheets."""
        max_row = max(max_row_a, max_row_b)
        pairs: list[tuple[int | None, int | None]] = []
        for r in range(1, max_row + 1):
            ra = r if r <= max_row_a else None
            rb = r if r <= max_row_b else None
            pairs.append((ra, rb))
        return pairs

    def _precompute_large_diff_by_blocks(self, ws_a_val, ws_b_val, ws_a_edit, ws_b_edit, max_row_a: int, max_row_b: int):
        """Large-sheet only-diff precompute using tail-first block scan."""
        max_row = max(max_row_a, max_row_b)
        block = _LARGE_SHEET_BLOCK_ROWS
        for block_end in range(max_row, 0, -block):
            block_start = max(1, block_end - block + 1)
            block_len = block_end - block_start + 1

            rows_a = {}
            rows_b = {}
            if block_start <= max_row_a:
                for idx, row in enumerate(
                    ws_a_val.iter_rows(
                        min_row=block_start,
                        max_row=min(block_end, max_row_a),
                        min_col=1,
                        max_col=self.max_col,
                        values_only=True,
                    ),
                    start=block_start,
                ):
                    rows_a[idx] = row or ()
            if block_start <= max_row_b:
                for idx, row in enumerate(
                    ws_b_val.iter_rows(
                        min_row=block_start,
                        max_row=min(block_end, max_row_b),
                        min_col=1,
                        max_col=self.max_col,
                        values_only=True,
                    ),
                    start=block_start,
                ):
                    rows_b[idx] = row or ()

            sig_a = []
            sig_b = []
            for r in range(block_start, block_end + 1):
                row_a = rows_a.get(r, ())
                row_b = rows_b.get(r, ())
                if len(row_a) < self.max_col:
                    row_a = tuple(row_a) + (None,) * (self.max_col - len(row_a))
                if len(row_b) < self.max_col:
                    row_b = tuple(row_b) + (None,) * (self.max_col - len(row_b))
                sig_a.append(tuple(_merge_cmp_value(v) for v in row_a))
                sig_b.append(tuple(_merge_cmp_value(v) for v in row_b))

            if sig_a == sig_b:
                continue

            # Tail-first within changed block (newer rows first).
            for off in range(block_len - 1, -1, -1):
                if sig_a[off] == sig_b[off]:
                    continue
                r = block_start + off
                pair_idx = self.row_a_to_pair_idx.get(r)
                if pair_idx is None:
                    pair_idx = self.row_b_to_pair_idx.get(r)
                if pair_idx is None:
                    continue
                ra, rb = self.row_pairs[pair_idx]
                line_a, line_b, cols = self._build_row_and_diff_pair(ws_a_val, ws_b_val, ws_a_edit, ws_b_edit, ra, rb)
                self.pair_diff_cols[pair_idx] = cols
                self.pair_text_a[pair_idx] = line_a
                self.pair_text_b[pair_idx] = line_b

    def _build_row_and_diff(self, ws_a_val, ws_b_val, ws_a_edit, ws_b_edit, r: int):
        parts_a = []
        parts_b = []
        cols = set()
        for c in range(1, self.max_col + 1):
            da, db, eq = _cell_display_and_equal(ws_a_val, ws_b_val, ws_a_edit, ws_b_edit, r, c)
            parts_a.append(_val_to_str(da))
            parts_b.append(_val_to_str(db))
            if not eq:
                cols.add(c)
        line_a = str(r) + "\t" + "\t".join(parts_a)
        line_b = str(r) + "\t" + "\t".join(parts_b)
        return line_a, line_b, cols

    def _compute_diff_cols_from_rows(self, row_a, row_b):
        cols = set()
        # row tuples are 0-indexed; cols are 1-indexed
        for i, (va, vb) in enumerate(zip(row_a, row_b), start=1):
            if va != vb:
                cols.add(i)
        return cols

    def _build_line_from_row(self, r: int, row_vals) -> str:
        return str(r) + "\t" + "\t".join(_val_to_str(v) for v in row_vals)

    @staticmethod
    def _spans_for_line(line: str):
        # returns {colIndex: (start,end)} where colIndex is 1..N
        fields = line.split("\t")
        spans = {}
        pos = 0
        # rownum field
        pos += len(fields[0])
        if len(fields) == 1:
            return spans
        pos += 1
        for c in range(1, len(fields)):
            start = pos
            pos += len(fields[c])
            end = pos
            spans[c] = (start, end)
            pos += 1
        return spans

    # ---------- Only-diff toggle ----------
    def _toggle_only_diff(self):
        # Snapshot mode confirmed by user: diff rows list is generated once when opening (or manual refresh).
        # Toggling "只看差异" only switches display, without recomputing the diff map.
        try:
            _dlog(f"TOGGLE only_diff={bool(self.only_diff_var.get())} raw={self.only_diff_var.get()} sheet={self.sheet}")
        except Exception:
            pass

        cur = int(self.only_diff_var.get())
        self._last_only_diff_value = cur

        # User-requested behavior: whenever only-diff state changes,
        # run the same refresh path as "刷新本Sheet" for deterministic UI update.
        self.refresh(row_only=None, rescan=True)
        self._update_cursor_lines()
        self._update_diff_nav_state()

        # Persist setting (debounced: write 1 s after last toggle to avoid per-keypress I/O)
        try:
            self.app.only_diff_default = int(self.only_diff_var.get())
            if hasattr(self, "_settings_save_id"):
                try:
                    self.frame.after_cancel(self._settings_save_id)
                except Exception:
                    pass
            self._settings_save_id = self.frame.after(1000, self._flush_settings)
        except Exception as e:
            _dlog(f"settings debounce failed: {e}")

    def _flush_settings(self):
        """Debounced settings write: called 1 s after the last only-diff toggle."""
        try:
            os.makedirs(os.path.dirname(_SETTINGS_PATH), exist_ok=True)
            with open(_SETTINGS_PATH, "w", encoding="utf-8") as f:
                json.dump({"only_diff": int(self.only_diff_var.get())}, f, ensure_ascii=False)
        except Exception as e:
            _dlog(f"settings save failed: {e}")

    # ---------- Merge operations ----------
    def _copy_cell(self, direction: str, event):
        try:
            src = self.left if direction == "A2B" else self.right
            idx = src.index(f"@{event.x},{event.y}")
            src.mark_set("insert", idx)
            line = int(idx.split(".")[0])
            col_char = int(idx.split(".")[1])

            if not (1 <= line <= len(self.display_rows)):
                return
            pair = self._pair_for_line(line)
            ra = self._row_for_side(pair, "A")
            rb = self._row_for_side(pair, "B")
            if direction == "A2B":
                if ra is None or rb is None:
                    return
                src_r = ra
                dst_r = rb
            else:
                if ra is None or rb is None:
                    return
                src_r = rb
                dst_r = ra

            line_text = src.get(f"{line}.0", f"{line}.end")
            before = line_text[:col_char]
            tab_count = before.count("\t")
            # Clicking row header/arrow area should apply whole-row overwrite,
            # not a single-cell copy of the first column.
            if tab_count <= 0:
                self._copy_selected_row(direction)
                return
            c = max(1, tab_count)  # 1..N
            if c > self.max_col:
                c = self.max_col

            # Merge conflict mode:
            # - "A2B" means keep mine, just mark resolved.
            # - "B2A" means apply theirs to mine, then mark resolved.
            if getattr(self.app, "merge_conflict_mode", False):
                if direction == "A2B":
                    self.app.user_touched_conflicts = True
                    self._resolve_conflict_cell(dst_r, c)
                    return

            if direction == "A2B":
                old_edit = self.app.ws_b_edit(self.sheet).cell(row=dst_r, column=c).value
                old_val = self.app.ws_b_val(self.sheet).cell(row=dst_r, column=c).value
                v_edit = self.app.ws_a_edit(self.sheet).cell(row=src_r, column=c).value
                v_val = self.app.ws_a_val(self.sheet).cell(row=src_r, column=c).value
                # Cached-value mode: always write the cached value
                self.app.ws_b_edit(self.sheet).cell(row=dst_r, column=c).value = v_val if _USE_CACHED_VALUES_ONLY else v_edit
                self.app.ws_b_val(self.sheet).cell(row=dst_r, column=c).value = v_val
                self.app.modified_b = True
                self.app.modified_sheets_b.add(self.sheet)
                self.app.push_undo({"sheet": self.sheet, "target": "B", "cells": [(dst_r, c, old_edit, old_val)]})
            else:
                old_edit = self.app.ws_a_edit(self.sheet).cell(row=dst_r, column=c).value
                old_val = self.app.ws_a_val(self.sheet).cell(row=dst_r, column=c).value
                v_edit = self.app.ws_b_edit(self.sheet).cell(row=src_r, column=c).value
                v_val = self.app.ws_b_val(self.sheet).cell(row=src_r, column=c).value
                self.app.ws_a_edit(self.sheet).cell(row=dst_r, column=c).value = v_val if _USE_CACHED_VALUES_ONLY else v_edit
                self.app.ws_a_val(self.sheet).cell(row=dst_r, column=c).value = v_val
                self.app.modified_a = True
                self.app.modified_sheets_a.add(self.sheet)
                self.app.push_undo({"sheet": self.sheet, "target": "A", "cells": [(dst_r, c, old_edit, old_val)]})
                # In conflict mode, B2A applies theirs; mark conflict resolved.
                if getattr(self.app, "merge_conflict_mode", False):
                    self.app.user_touched_conflicts = True
                    self._resolve_conflict_cell(dst_r, c)
                    return

            # Mark as touched: keep row visible in "只看差异" even if diffs are resolved.
            pair = self._pair_for_line(line)
            touched_r = self._row_for_side(pair, "A") or self._row_for_side(pair, "B")
            if touched_r is not None:
                self.touched_rows.add(touched_r)
            self._invalidate_render_cache()

            # Recompute this row immediately, then do a full rescan in row-align mode
            # so pair alignment and full-sheet diff state stay up-to-date.
            if bool(self.only_diff_var.get()) and self.snapshot_only_diff:
                self._recalc_row_diff_and_update(dst_r)
            self.refresh(row_only=dst_r, rescan=False)
            if self._align_rows_enabled:
                self.refresh(row_only=None, rescan=True)
            self._update_cursor_lines()
        except Exception as e:
            messagebox.showerror("Error", f"覆盖单元格失败：\n{e}")

    def _copy_selected_row(self, direction: str):
        t0 = datetime.now()
        try:
            resolved_only = False
            # use last selected excel row (set on click); fallback to cursor line
            pair_idx = self.selected_pair_idx
            if pair_idx is None:
                widget = self.left
                try:
                    focus = self.root.focus_get()
                    if focus == self.right:
                        widget = self.right
                except Exception:
                    pass
                try:
                    line = int((widget.index("insert").split(".")[0]))
                except Exception:
                    line = 1
                if not (1 <= line <= len(self.display_rows)):
                    return
                pair_idx = self.display_rows[line - 1]
            pair = self.row_pairs[pair_idx] if pair_idx is not None and pair_idx < len(self.row_pairs) else None
            ra = self._row_for_side(pair, "A")
            rb = self._row_for_side(pair, "B")
            if direction == "A2B":
                if ra is None or rb is None:
                    return
                src_r = ra
                dst_r = rb
            else:
                if ra is None or rb is None:
                    return
                src_r = rb
                dst_r = ra
            ws_a_val = self.app.ws_a_val(self.sheet)
            ws_b_val = self.app.ws_b_val(self.sheet)
            ws_a_edit = self.app.ws_a_edit(self.sheet)
            ws_b_edit = self.app.ws_b_edit(self.sheet)

            # Row action should overwrite the full row range (not only diff cols),
            # so "使用左侧/使用右侧" behaves as full-row adopt.
            full_max_col = max(
                self.max_col,
                ws_a_val.max_column or 1,
                ws_b_val.max_column or 1,
                ws_a_edit.max_column or 1,
                ws_b_edit.max_column or 1,
            )
            cols = set(range(1, full_max_col + 1))

            # Merge conflict mode:
            # - "A2B" means keep mine, just mark resolved.
            # - "B2A" means apply theirs to mine, then mark resolved.
            if getattr(self.app, "merge_conflict_mode", False):
                rows = self.app.merge_conflict_cells_by_sheet.get(self.sheet) if getattr(self.app, "merge_conflict_cells_by_sheet", None) else None
                conflict_row = ra or rb
                if rows and conflict_row in rows:
                    cols = set(rows.get(conflict_row, set())) if direction == "A2B" else cols
                if direction == "A2B":
                    self.app.user_touched_conflicts = True
                    self._resolve_conflict_row(conflict_row, cols)
                    resolved_only = True

            if not cols:
                return

            if direction == "A2B":
                if not resolved_only:
                    undo_cells = []
                    for c in cols:
                        old_edit = ws_b_edit.cell(row=dst_r, column=c).value
                        old_val = ws_b_val.cell(row=dst_r, column=c).value
                        v_edit = ws_a_edit.cell(row=src_r, column=c).value
                        v_val = ws_a_val.cell(row=src_r, column=c).value
                        ws_b_edit.cell(row=dst_r, column=c).value = v_val if _USE_CACHED_VALUES_ONLY else v_edit
                        ws_b_val.cell(row=dst_r, column=c).value = v_val
                        undo_cells.append((dst_r, c, old_edit, old_val))
                    self.app.modified_b = True
                    self.app.modified_sheets_b.add(self.sheet)
                    if undo_cells:
                        self.app.push_undo({"sheet": self.sheet, "target": "B", "cells": undo_cells})
            else:
                undo_cells = []
                for c in cols:
                    old_edit = ws_a_edit.cell(row=dst_r, column=c).value
                    old_val = ws_a_val.cell(row=dst_r, column=c).value
                    v_edit = ws_b_edit.cell(row=src_r, column=c).value
                    v_val = ws_b_val.cell(row=src_r, column=c).value
                    ws_a_edit.cell(row=dst_r, column=c).value = v_val if _USE_CACHED_VALUES_ONLY else v_edit
                    ws_a_val.cell(row=dst_r, column=c).value = v_val
                    undo_cells.append((dst_r, c, old_edit, old_val))
                self.app.modified_a = True
                self.app.modified_sheets_a.add(self.sheet)
                if undo_cells:
                    self.app.push_undo({"sheet": self.sheet, "target": "A", "cells": undo_cells})
                # In conflict mode, B2A applies theirs; mark conflict resolved.
                if getattr(self.app, "merge_conflict_mode", False):
                    self.app.user_touched_conflicts = True
                    self._resolve_conflict_row(conflict_row, cols)
                    resolved_only = True

            # Mark as touched: keep row visible in "只看差异" even if diffs are resolved.
            touched_r = ra or rb
            if touched_r is not None:
                self.touched_rows.add(touched_r)
            self._invalidate_render_cache()

            # Recompute this row immediately, then do a full rescan in row-align mode
            # so pair alignment and full-sheet diff state stay up-to-date.
            if bool(self.only_diff_var.get()) and self.snapshot_only_diff:
                self._recalc_row_diff_and_update(dst_r)
            self.refresh(row_only=dst_r, rescan=False)
            if self._align_rows_enabled:
                self.refresh(row_only=None, rescan=True)
            self._update_cursor_lines()
        except Exception as e:
            messagebox.showerror("Error", f"覆盖整行失败：\n{e}")
        finally:
            try:
                dt = (datetime.now() - t0).total_seconds() * 1000.0
                _dlog(f"OVERWRITE_ROW {self.sheet} dir={direction} ms={dt:.1f}")
            except Exception:
                pass

    def _undo_last_action(self):
        try:
            action = self.app.pop_undo()
            if not action:
                return
            sheet = action.get("sheet")
            target = action.get("target")
            if target == "A_APPEND":
                start_row = action.get("start_row")
                count = action.get("count")
                if not start_row or not count:
                    return
                ws_edit = self.app.ws_a_edit(sheet)
                ws_val = self.app.ws_a_val(sheet)
                ws_edit.delete_rows(start_row, count)
                ws_val.delete_rows(start_row, count)
                self.app.modified_a = True
                self.app.modified_sheets_a.add(sheet)
                if sheet == self.sheet:
                    self._invalidate_render_cache()
                    self.refresh(row_only=None, rescan=True)
                    self._update_cursor_lines()
                return
            cells = action.get("cells", [])
            if not cells:
                return
            if target == "A":
                ws_edit = self.app.ws_a_edit(sheet)
                ws_val = self.app.ws_a_val(sheet)
                self.app.modified_a = True
                self.app.modified_sheets_a.add(sheet)
            else:
                ws_edit = self.app.ws_b_edit(sheet)
                ws_val = self.app.ws_b_val(sheet)
                self.app.modified_b = True
                self.app.modified_sheets_b.add(sheet)
            rows = set()
            for r, c, old_edit, old_val in cells:
                ws_edit.cell(row=r, column=c).value = old_edit
                ws_val.cell(row=r, column=c).value = old_val
                rows.add(r)
            # refresh current sheet if applicable
            if sheet == self.sheet:
                for r in rows:
                    self.touched_rows.add(r)
                    if bool(self.only_diff_var.get()) and self.snapshot_only_diff:
                        self._recalc_row_diff_and_update(r)
                    self.refresh(row_only=r, rescan=False)
                if self._align_rows_enabled:
                    self.refresh(row_only=None, rescan=True)
                self._update_cursor_lines()
        except Exception:
            pass

    def _get_current_right_excel_row(self) -> int | None:
        try:
            line = int(self.right.index("insert").split(".")[0])
        except Exception:
            line = 1
        if not (1 <= line <= len(self.display_rows)):
            return None
        pair = self._pair_for_line(line)
        return self._row_for_side(pair, "B")

    def _append_rows_from_right_to_left_end(self, rows: list[int]):
        if not rows:
            return
        ws_a_val = self.app.ws_a_val(self.sheet)
        ws_b_val = self.app.ws_b_val(self.sheet)
        ws_a_edit = self.app.ws_a_edit(self.sheet)
        ws_b_edit = self.app.ws_b_edit(self.sheet)

        a_r_val, _a_c_val = _effective_bounds(ws_a_val)
        a_r_edit, _a_c_edit = _effective_bounds(ws_a_edit)
        start_row = max(a_r_val, a_r_edit) + 1

        _b_r_val, b_c_val = _effective_bounds(ws_b_val)
        _b_r_edit, b_c_edit = _effective_bounds(ws_b_edit)
        max_col = max(1, b_c_val, b_c_edit)

        for i, r in enumerate(rows):
            target_r = start_row + i
            for c in range(1, max_col + 1):
                v_val = ws_b_val.cell(row=r, column=c).value
                v_edit = ws_b_edit.cell(row=r, column=c).value
                ws_a_val.cell(row=target_r, column=c).value = v_val
                ws_a_edit.cell(row=target_r, column=c).value = v_val if _USE_CACHED_VALUES_ONLY else v_edit

        self.app.modified_a = True
        self.app.modified_sheets_a.add(self.sheet)
        self.app.push_undo({"sheet": self.sheet, "target": "A_APPEND", "start_row": start_row, "count": len(rows)})
        self._invalidate_render_cache()
        self.refresh(row_only=None, rescan=True)
        self._update_cursor_lines()

    def _append_right_row_to_left_end(self):
        try:
            r = self._get_current_right_excel_row()
            if r is None:
                return
            self._append_rows_from_right_to_left_end([r])
        except Exception as e:
            messagebox.showerror("Error", f"同步单行失败：\n{e}")

    def _append_right_all_to_left_end(self):
        try:
            ws_b_val = self.app.ws_b_val(self.sheet)
            b_r, _b_c = _effective_bounds(ws_b_val)
            if b_r <= 0:
                return
            rows = list(range(1, b_r + 1))

            if len(rows) > 200:
                self.app._with_progress(
                    "同步中",
                    f"正在同步右侧到左侧（{len(rows)}行）...",
                    lambda: self._append_rows_from_right_to_left_end(rows),
                )
            else:
                self._append_rows_from_right_to_left_end(rows)
        except Exception as e:
            messagebox.showerror("Error", f"同步所有失败：\n{e}")

    def _resolve_conflict_cell(self, r: int, c: int):
        try:
            if self.app.resolve_conflict_cell(self.sheet, r, c):
                # update view based on updated conflict map
                self.refresh(row_only=None, rescan=False)
                self._update_cursor_lines()
        except Exception:
            pass

    def _resolve_conflict_row(self, r: int, cols):
        try:
            if self.app.resolve_conflict_row(self.sheet, r, cols):
                self.refresh(row_only=None, rescan=False)
                self._update_cursor_lines()
        except Exception:
            pass

    def _refresh_row_text_only(self, r: int):
        """Update the rendered row text for row r without recomputing diff_cols_by_row."""
        try:
            ws_a = self.app.ws_a_val(self.sheet)
            ws_b = self.app.ws_b_val(self.sheet)
            ws_a_edit = self.app.ws_a_edit(self.sheet)
            ws_b_edit = self.app.ws_b_edit(self.sheet)

            self.max_col = max(ws_a.max_column or 1, ws_b.max_column or 1)

            pair_idx = self.row_a_to_pair_idx.get(r)
            if pair_idx is None:
                pair_idx = self.row_b_to_pair_idx.get(r)
            if pair_idx is None:
                return
            ra, rb = self.row_pairs[pair_idx]
            line_a, line_b, _cols = self._build_row_and_diff_pair(ws_a, ws_b, ws_a_edit, ws_b_edit, ra, rb)
            self.pair_text_a[pair_idx] = line_a
            self.pair_text_b[pair_idx] = line_b

            line = self.row_to_line.get(pair_idx)
            if line is None:
                return

            self.left.delete(f"{line}.0", f"{line}.end")
            self.right.delete(f"{line}.0", f"{line}.end")
            self.left.insert(f"{line}.0", self.pair_text_a[pair_idx])
            self.right.insert(f"{line}.0", self.pair_text_b[pair_idx])
        except Exception:
            pass

    def _recalc_row_diff_and_update(self, r: int):
        """Recompute diff for row r and update its highlight, without changing the row list (snapshot mode)."""
        try:
            ws_a = self.app.ws_a_val(self.sheet)
            ws_b = self.app.ws_b_val(self.sheet)
            ws_a_edit = self.app.ws_a_edit(self.sheet)
            ws_b_edit = self.app.ws_b_edit(self.sheet)

            self.max_col = max(ws_a.max_column or 1, ws_b.max_column or 1)
            pair_idx = self.row_a_to_pair_idx.get(r)
            if pair_idx is None:
                pair_idx = self.row_b_to_pair_idx.get(r)
            if pair_idx is None:
                return
            ra, rb = self.row_pairs[pair_idx]
            line_a, line_b, cols = self._build_row_and_diff_pair(ws_a, ws_b, ws_a_edit, ws_b_edit, ra, rb)
            self.pair_diff_cols[pair_idx] = cols
            self.pair_text_a[pair_idx] = line_a
            self.pair_text_b[pair_idx] = line_b

            line = self.row_to_line.get(pair_idx)
            if line is None:
                # if not visible and touched, rebuild snapshot to include it
                if bool(self.only_diff_var.get()) and (r in self.touched_rows):
                    self.refresh(row_only=None, rescan=False)
                return

            # update text
            self.left.delete(f"{line}.0", f"{line}.end")
            self.right.delete(f"{line}.0", f"{line}.end")
            self.left.insert(f"{line}.0", self.pair_text_a[pair_idx])
            self.right.insert(f"{line}.0", self.pair_text_b[pair_idx])

            # update tags for this line
            for w in (self.left, self.right):
                w.tag_remove("diffrow", f"{line}.0", f"{line}.end")
                w.tag_remove("diffcell", f"{line}.0", f"{line}.end")

            if cols:
                self.left.tag_add("diffrow", f"{line}.0", f"{line}.end")
                self.right.tag_add("diffrow", f"{line}.0", f"{line}.end")

                spans_a = self._spans_for_line(self.pair_text_a[pair_idx])
                spans_b = self._spans_for_line(self.pair_text_b[pair_idx])
                for c in cols:
                    if c in spans_a:
                        s, e = spans_a[c]
                        self.left.tag_add("diffcell", f"{line}.{s}", f"{line}.{e}")
                    if c in spans_b:
                        s, e = spans_b[c]
                        self.right.tag_add("diffcell", f"{line}.{s}", f"{line}.{e}")
        except Exception:
            pass

    def _invalidate_render_cache(self):
        self._data_version += 1
        self._render_cache.clear()

    # ---------- Rendering ----------
    def _load_all_rows(self):
        self._full_render = True
        self.refresh(row_only=None, rescan=False)

    def _append_rows(self, new_rows: list[int]):
        if not new_rows:
            return
        ws_a = self.app.ws_a_val(self.sheet)
        ws_b = self.app.ws_b_val(self.sheet)
        try:
            wb_a_edit = getattr(self.app, "_wb_a_edit", None)
            ws_a_edit = wb_a_edit[self.sheet] if wb_a_edit is not None else None
        except Exception:
            ws_a_edit = None
        try:
            wb_b_edit = getattr(self.app, "_wb_b_edit", None)
            ws_b_edit = wb_b_edit[self.sheet] if wb_b_edit is not None else None
        except Exception:
            ws_b_edit = None

        start_line = len(self.display_rows) + 1
        # Preserve current scroll position to avoid jumps
        try:
            first, _last = self.left.yview()
        except Exception:
            first = None

        for idx, pair_idx in enumerate(new_rows, start=0):
            if pair_idx not in self.pair_text_a or pair_idx not in self.pair_text_b:
                ra, rb = self.row_pairs[pair_idx]
                line_a, line_b, cols = self._build_row_and_diff_pair(ws_a, ws_b, ws_a_edit, ws_b_edit, ra, rb)
                self.pair_diff_cols[pair_idx] = cols
                self.pair_text_a[pair_idx] = line_a
                self.pair_text_b[pair_idx] = line_b
            else:
                cols = self.pair_diff_cols.get(pair_idx, set())
                line_a = self.pair_text_a.get(pair_idx, "")
                line_b = self.pair_text_b.get(pair_idx, "")

            line_no = start_line + idx
            self.left.insert("end", line_a + "\n")
            self.right.insert("end", line_b + "\n")

            if cols:
                self._display_diff_row_count += 1
                self.left.tag_add("diffrow", f"{line_no}.0", f"{line_no}.end")
                self.right.tag_add("diffrow", f"{line_no}.0", f"{line_no}.end")
                spans_a = self._spans_for_line(line_a)
                spans_b = self._spans_for_line(line_b)
                for c in cols:
                    if c in spans_a:
                        s, e = spans_a[c]
                        self.left.tag_add("diffcell", f"{line_no}.{s}", f"{line_no}.{e}")
                    if c in spans_b:
                        s, e = spans_b[c]
                        self.right.tag_add("diffcell", f"{line_no}.{s}", f"{line_no}.{e}")

        self.display_rows.extend(new_rows)
        for i, pair_idx in enumerate(new_rows, start=start_line):
            self.row_to_line[pair_idx] = i

        mode = "只看差异" if self.only_diff_var.get() else "全量"
        total_rows = len(self.row_pairs) if self.row_pairs else self.max_row
        self.info.configure(text=f"{mode} | RowsShown: {len(self.display_rows)} / {total_rows}   Cols: {self.max_col}   DiffRows: {self._display_diff_row_count}")

        if first is not None:
            try:
                self.left.yview_moveto(first)
                self.right.yview_moveto(first)
            except Exception:
                pass

        self._invalidate_render_cache()

    def _maybe_load_more_rows(self, last_fraction: float):
        if not _FAST_OPEN_ENABLED:
            return
        try:
            last_fraction = float(last_fraction)
        except Exception:
            return
        if self._full_render:
            return
        if bool(self.only_diff_var.get()):
            return
        if getattr(self.app, "merge_conflict_mode", False):
            return
        # Only for full-list mode (not only-diff or conflict-only)
        if not self._full_display_rows:
            return
        if last_fraction < 0.98:
            return
        if len(self.display_rows) >= len(self._full_display_rows):
            return
        old_limit = len(self.display_rows)
        new_limit = min(len(self._full_display_rows), self._render_limit + _FAST_RENDER_BATCH)
        self._render_limit = new_limit
        new_rows = self._full_display_rows[old_limit:new_limit]
        self._append_rows(new_rows)

    def refresh(self, row_only: int | None, rescan: bool):
        _dlog(f"REFRESH sheet={self.sheet} row_only={row_only} rescan={rescan} only_diff={bool(self.only_diff_var.get())} raw={self.only_diff_var.get()}")
        if rescan and (not self._full_render):
            self._render_limit = _FAST_RENDER_ROW_LIMIT
        conflict_cells_by_row = None
        if getattr(self.app, "merge_conflict_mode", False):
            rows_map = getattr(self.app, "merge_conflict_cells_by_sheet", None)
            conflict_cells_by_row = rows_map.get(self.sheet) if rows_map else None
        ws_a = self.app.ws_a_val(self.sheet)
        ws_b = self.app.ws_b_val(self.sheet)
        # Non-blocking edit sheets: use loaded edit workbook if already available.
        # Do not trigger expensive load_workbook() during pure view refresh/toggle.
        try:
            wb_a_edit = getattr(self.app, "_wb_a_edit", None)
            ws_a_edit = wb_a_edit[self.sheet] if wb_a_edit is not None else None
        except Exception:
            ws_a_edit = None
        try:
            wb_b_edit = getattr(self.app, "_wb_b_edit", None)
            ws_b_edit = wb_b_edit[self.sheet] if wb_b_edit is not None else None
        except Exception:
            ws_b_edit = None

        if rescan or (not self._bounds_checked):
            a_r, a_c = _effective_bounds(ws_a)
            b_r, b_c = _effective_bounds(ws_b)
            self.max_row = max(a_r, b_r)
            self.max_col = max(a_c, b_c)
            self._bounds_checked = True
            self._is_large_sheet = self.max_row >= _LARGE_SHEET_ROW_THRESHOLD
            if self._is_large_sheet:
                self._prefer_only_diff_when_ready = True

        # Full rescan diff map + cache row text if requested
        # Use _data_ready flag instead of checking pair_diff_cols emptiness:
        # pair_diff_cols can legitimately be empty (no diffs found) while still being valid data.
        if rescan or not self._data_ready:
            if not rescan:
                # Data not yet ready (background computation still running).
                # Skip this call; _apply_sheet_cache will call refresh() when done.
                return
            self.pair_diff_cols = {}
            self.pair_text_a = {}
            self.pair_text_b = {}
            self.row_a_to_pair_idx = {}
            self.row_b_to_pair_idx = {}
            self._diff_partial = False

            if conflict_cells_by_row is not None:
                # Conflict-only fast path: avoid full-sheet diff scan.
                self._align_rows_enabled = False
                conflict_rows = sorted(conflict_cells_by_row.keys())
                self.row_pairs = [(r, r) for r in conflict_rows]
                for idx, (ra, rb) in enumerate(self.row_pairs):
                    if ra is not None:
                        self.row_a_to_pair_idx[ra] = idx
                    if rb is not None:
                        self.row_b_to_pair_idx[rb] = idx
                    line_a, line_b, _cols = self._build_row_and_diff_pair(ws_a, ws_b, ws_a_edit, ws_b_edit, ra, rb)
                    cols = set(conflict_cells_by_row.get(ra, set())) if ra is not None else set()
                    self.pair_diff_cols[idx] = cols
                    self.pair_text_a[idx] = line_a
                    self.pair_text_b[idx] = line_b
            else:
                max_row_a = ws_a.max_row or 1
                max_row_b = ws_b.max_row or 1

                # Very large sheets: skip expensive row-alignment on open.
                if self.max_row >= _LARGE_SHEET_DIRECT_PAIR_THRESHOLD:
                    self._align_rows_enabled = False
                    self.row_pairs = self._build_row_pairs_direct(max_row_a, max_row_b)
                else:
                    self._align_rows_enabled = (not getattr(self.app, "merge_conflict_mode", False))
                    if self._align_rows_enabled:
                        self.row_pairs = self._build_row_pairs(ws_a, ws_b)
                    else:
                        self.row_pairs = self._build_row_pairs_direct(max_row_a, max_row_b)

                for idx, (ra, rb) in enumerate(self.row_pairs):
                    if ra is not None:
                        self.row_a_to_pair_idx[ra] = idx
                    if rb is not None:
                        self.row_b_to_pair_idx[rb] = idx

                # Large-sheet strategy:
                # - full mode: lazy row compute (first 200 visible rows only)
                # - only-diff mode: block scan from tail to head (1000 rows/block)
                if self._is_large_sheet and bool(self.only_diff_var.get()):
                    self._precompute_large_diff_by_blocks(ws_a, ws_b, ws_a_edit, ws_b_edit, max_row_a, max_row_b)
                elif not self._is_large_sheet:
                    for idx, (ra, rb) in enumerate(self.row_pairs):
                        line_a, line_b, cols = self._build_row_and_diff_pair(ws_a, ws_b, ws_a_edit, ws_b_edit, ra, rb)
                        self.pair_diff_cols[idx] = cols
                        self.pair_text_a[idx] = line_a
                        self.pair_text_b[idx] = line_b

            self._data_ready = True

        # Build display rows list (pair indices)
        if conflict_cells_by_row is not None:
            # Always show conflict rows only
            rows = []
            for r in sorted(conflict_cells_by_row.keys()):
                idx = self.row_a_to_pair_idx.get(r)
                if idx is not None:
                    rows.append(idx)
            self._full_display_rows = rows
        elif bool(self.only_diff_var.get()):
            if (not self.snapshot_only_diff) or rescan or (row_only is None) or (not self.display_rows):
                # build snapshot: diff rows + touched rows
                rows = [idx for idx, cols in self.pair_diff_cols.items() if cols]
                rows_set = set(rows)
                for r in self.touched_rows:
                    idx = self.row_a_to_pair_idx.get(r)
                    if idx is not None:
                        rows_set.add(idx)
                self._full_display_rows = sorted(rows_set)
            else:
                # snapshot mode: keep existing row list stable
                pass
        else:
            self._full_display_rows = list(range(0, len(self.row_pairs)))

        # Fast render: limit initial rows unless user opted to load all
        if self._full_render or (not _FAST_OPEN_ENABLED):
            self.display_rows = list(self._full_display_rows)
        else:
            # reset render limit if full list shrank
            self._render_limit = min(self._render_limit, len(self._full_display_rows)) if self._full_display_rows else _FAST_RENDER_ROW_LIMIT
            if len(self._full_display_rows) > _FAST_RENDER_ROW_LIMIT and self._render_limit < _FAST_RENDER_ROW_LIMIT:
                self._render_limit = _FAST_RENDER_ROW_LIMIT
            if self._is_large_sheet and rescan:
                self._render_limit = min(_LARGE_SHEET_INITIAL_ROWS, len(self._full_display_rows)) if self._full_display_rows else _LARGE_SHEET_INITIAL_ROWS
            self.display_rows = self._full_display_rows[:self._render_limit]
        _dlog(f"  build display_rows: {len(self.display_rows)} / {self.max_row} (only_diff={bool(self.only_diff_var.get())} raw={self.only_diff_var.get()})")

        # Ensure pair text/diff exists for currently displayed rows (lazy fill)
        if self.display_rows:
            missing = [idx for idx in self.display_rows if idx not in self.pair_text_a or idx not in self.pair_text_b]
            if missing:
                for idx in missing:
                    ra, rb = self.row_pairs[idx]
                    line_a, line_b, cols = self._build_row_and_diff_pair(ws_a, ws_b, ws_a_edit, ws_b_edit, ra, rb)
                    self.pair_diff_cols[idx] = cols
                    self.pair_text_a[idx] = line_a
                    self.pair_text_b[idx] = line_b

        self.row_to_line = {r: i + 1 for i, r in enumerate(self.display_rows)}

        # Partial refresh: update a single excel row if it is visible
        if row_only is not None:
            r = row_only
            pair_idx = self.row_a_to_pair_idx.get(r)
            if pair_idx is None:
                pair_idx = self.row_b_to_pair_idx.get(r)
            if pair_idx is None:
                return
            ra, rb = self.row_pairs[pair_idx]
            # recompute diff cols + cache text for that pair only
            line_a, line_b, cols = self._build_row_and_diff_pair(ws_a, ws_b, ws_a_edit, ws_b_edit, ra, rb)
            self.pair_text_a[pair_idx] = line_a
            self.pair_text_b[pair_idx] = line_b
            if conflict_cells_by_row is not None and ra is not None:
                self.pair_diff_cols[pair_idx] = set(conflict_cells_by_row.get(ra, set()))
            else:
                self.pair_diff_cols[pair_idx] = cols

            # If only-diff enabled, row might need to be added/removed
            if bool(self.only_diff_var.get()):
                visible = pair_idx in self.row_to_line
                has = bool(self.pair_diff_cols[pair_idx])

                # If diffs are resolved but this row was touched, keep it visible as a record.
                keep = (r in self.touched_rows)

                if self.snapshot_only_diff:
                    # Snapshot mode: never auto-remove rows from the list.
                    # If a touched row is not visible (was not in initial snapshot), allow adding it.
                    if (not visible) and keep:
                        self.refresh(row_only=None, rescan=False)
                        return
                else:
                    if visible and (not has) and (not keep):
                        # remove the line
                        line = self.row_to_line[pair_idx]
                        self.left.delete(f"{line}.0", f"{line + 1}.0")
                        self.right.delete(f"{line}.0", f"{line + 1}.0")
                        # rebuild
                        self.refresh(row_only=None, rescan=False)
                        return

                    if (not visible) and (has or keep):
                        # add row: simplest is full rebuild (diff list is small)
                        self.refresh(row_only=None, rescan=False)
                        return

            line = self.row_to_line.get(pair_idx)
            if line is None:
                # not visible
                return

            line_a = self.pair_text_a.get(pair_idx, "")
            line_b = self.pair_text_b.get(pair_idx, "")

            # update text
            self.left.delete(f"{line}.0", f"{line}.end")
            self.right.delete(f"{line}.0", f"{line}.end")
            self.left.insert(f"{line}.0", line_a)
            self.right.insert(f"{line}.0", line_b)

            # clear tags on this line then apply diff highlight (unless touched row resolved)
            for w in (self.left, self.right):
                w.tag_remove("diffrow", f"{line}.0", f"{line}.end")
                w.tag_remove("diffcell", f"{line}.0", f"{line}.end")

            cols = self.pair_diff_cols.get(pair_idx, set())
            # If this row was touched and has no diffs anymore, keep it visible but don't show diff highlight.
            show_diff = bool(cols)
            if show_diff:
                self.left.tag_add("diffrow", f"{line}.0", f"{line}.end")
                self.right.tag_add("diffrow", f"{line}.0", f"{line}.end")

                spans_a = self._spans_for_line(line_a)
                spans_b = self._spans_for_line(line_b)
                for c in cols:
                    if c in spans_a:
                        s, e = spans_a[c]
                        self.left.tag_add("diffcell", f"{line}.{s}", f"{line}.{e}")
                    if c in spans_b:
                        s, e = spans_b[c]
                        self.right.tag_add("diffcell", f"{line}.{s}", f"{line}.{e}")

            # keep fast; do not rebuild sheet nav here
            try:
                self._display_diff_row_count = sum(1 for idx in self.display_rows if self.pair_diff_cols.get(idx))
                mode = "只看差异" if self.only_diff_var.get() else "全量"
                total_rows = len(self.row_pairs) if self.row_pairs else self.max_row
                self.info.configure(text=f"{mode} | RowsShown: {len(self.display_rows)} / {total_rows}   Cols: {self.max_col}   DiffRows: {self._display_diff_row_count}")
            except Exception:
                pass
            return

        # Full render (use cache when possible)
        mode_key = "diff" if (conflict_cells_by_row is not None or bool(self.only_diff_var.get())) else "full"
        head = tuple(self.display_rows[:5])
        tail = tuple(self.display_rows[-5:]) if len(self.display_rows) > 5 else tuple(self.display_rows)
        cache_key = (mode_key, self._render_limit, len(self.display_rows), head, tail, self._data_version)
        if row_only is None and (not rescan):
            cached = self._render_cache.get(cache_key)
            if cached is not None:
                text_a, text_b, tag_rows, tag_cells, diff_row_count = cached
                self.left.delete("1.0", "end")
                self.right.delete("1.0", "end")
                self.left.insert("1.0", text_a)
                self.right.insert("1.0", text_b)
                # clear tags
                self.left.tag_remove("diffrow", "1.0", "end")
                self.right.tag_remove("diffrow", "1.0", "end")
                self.left.tag_remove("diffcell", "1.0", "end")
                self.right.tag_remove("diffcell", "1.0", "end")
                self.left.tag_remove("paddingrow", "1.0", "end")
                self.right.tag_remove("paddingrow", "1.0", "end")
                # apply cached tags in bulk (one Tcl call per tag per widget)
                if tag_rows:
                    cached_diffrow_args = []
                    for line_idx in tag_rows:
                        cached_diffrow_args.extend([f"{line_idx}.0", f"{line_idx}.end"])
                    self.left.tag_add("diffrow", *cached_diffrow_args)
                    self.right.tag_add("diffrow", *cached_diffrow_args)
                if tag_cells:
                    cached_cell_left = []
                    cached_cell_right = []
                    for line_idx, spans_a, spans_b in tag_cells:
                        for s, e in spans_a:
                            cached_cell_left.extend([f"{line_idx}.{s}", f"{line_idx}.{e}"])
                        for s, e in spans_b:
                            cached_cell_right.extend([f"{line_idx}.{s}", f"{line_idx}.{e}"])
                    if cached_cell_left:
                        self.left.tag_add("diffcell", *cached_cell_left)
                    if cached_cell_right:
                        self.right.tag_add("diffcell", *cached_cell_right)

                # paddingrow: grey slot for one-sided pairs (computed from row_pairs, not cached)
                _padding_left = []
                _padding_right = []
                for _i, _pidx in enumerate(self.display_rows):
                    if _pidx < len(self.row_pairs):
                        _ra, _rb = self.row_pairs[_pidx]
                        _ln = _i + 1
                        if _ra is None:
                            _padding_left.extend([f"{_ln}.0", f"{_ln}.end"])
                        elif _rb is None:
                            _padding_right.extend([f"{_ln}.0", f"{_ln}.end"])
                if _padding_left:
                    self.left.tag_add("paddingrow", *_padding_left)
                if _padding_right:
                    self.right.tag_add("paddingrow", *_padding_right)

                mode = "只看差异" if self.only_diff_var.get() else "全量"
                total_rows = len(self.row_pairs) if self.row_pairs else self.max_row
                self.info.configure(text=f"{mode} | RowsShown: {len(self.display_rows)} / {total_rows}   Cols: {self.max_col}   DiffRows: {diff_row_count}")
                self._display_diff_row_count = diff_row_count
                self.app.set_sheet_has_diff(self.sheet, diff_row_count > 0, confirmed=True)
                self.app.refresh_sheet_nav()
                self._update_diff_nav_state()
                return

        # Full render
        self.left.delete("1.0", "end")
        self.right.delete("1.0", "end")
        self.left.tag_remove("diffrow", "1.0", "end")
        self.right.tag_remove("diffrow", "1.0", "end")
        self.left.tag_remove("diffcell", "1.0", "end")
        self.right.tag_remove("diffcell", "1.0", "end")
        self.left.tag_remove("paddingrow", "1.0", "end")
        self.right.tag_remove("paddingrow", "1.0", "end")

        # Build full text in memory and insert once (faster)
        lines_a = []
        lines_b = []
        for pair_idx in self.display_rows:
            lines_a.append(self.pair_text_a.get(pair_idx, ""))
            lines_b.append(self.pair_text_b.get(pair_idx, ""))
        self.left.insert("1.0", "\n".join(lines_a) + ("\n" if lines_a else ""))
        self.right.insert("1.0", "\n".join(lines_b) + ("\n" if lines_b else ""))

        # On some environments/large documents, forcing an idle layout pass improves tag correctness.
        try:
            self.left.update_idletasks()
            self.right.update_idletasks()
        except Exception:
            pass

        # Restore scroll position if we just appended more rows
        if self._pending_yview is not None:
            try:
                self.left.yview_moveto(self._pending_yview)
                self.right.yview_moveto(self._pending_yview)
            except Exception:
                pass
            self._pending_yview = None

        diff_row_count = 0
        tag_rows = []
        tag_cells = []
        # Collect all tag ranges first; apply in bulk (one Tcl call per tag instead of N).
        # tag_add(tagName, index1, *args) accepts multiple index pairs in a single call.
        diffrow_args = []
        diffcell_args_left = []
        diffcell_args_right = []
        for line_idx, pair_idx in enumerate(self.display_rows, start=1):
            cols = self.pair_diff_cols.get(pair_idx, set())
            if cols:
                diff_row_count += 1
                diffrow_args.extend([f"{line_idx}.0", f"{line_idx}.end"])
                tag_rows.append(line_idx)

                line_a = lines_a[line_idx - 1] if (line_idx - 1) < len(lines_a) else ""
                line_b = lines_b[line_idx - 1] if (line_idx - 1) < len(lines_b) else ""

                spans_a = self._spans_for_line(line_a)
                spans_b = self._spans_for_line(line_b)
                spans_a_ranges = []
                spans_b_ranges = []
                for c in cols:
                    if c in spans_a:
                        s, e = spans_a[c]
                        diffcell_args_left.extend([f"{line_idx}.{s}", f"{line_idx}.{e}"])
                        spans_a_ranges.append((s, e))
                    if c in spans_b:
                        s, e = spans_b[c]
                        diffcell_args_right.extend([f"{line_idx}.{s}", f"{line_idx}.{e}"])
                        spans_b_ranges.append((s, e))
                if spans_a_ranges or spans_b_ranges:
                    tag_cells.append((line_idx, spans_a_ranges, spans_b_ranges))

        # Apply all diffrow tags in one call per widget
        if diffrow_args:
            self.left.tag_add("diffrow", *diffrow_args)
            self.right.tag_add("diffrow", *diffrow_args)
        # Apply all diffcell tags in one call per widget
        if diffcell_args_left:
            self.left.tag_add("diffcell", *diffcell_args_left)
        if diffcell_args_right:
            self.right.tag_add("diffcell", *diffcell_args_right)
        # Apply paddingrow (grey) to empty slots of one-sided pairs
        _padding_left = []
        _padding_right = []
        for _i, _pidx in enumerate(self.display_rows):
            if _pidx < len(self.row_pairs):
                _ra, _rb = self.row_pairs[_pidx]
                _ln = _i + 1
                if _ra is None:
                    _padding_left.extend([f"{_ln}.0", f"{_ln}.end"])
                elif _rb is None:
                    _padding_right.extend([f"{_ln}.0", f"{_ln}.end"])
        if _padding_left:
            self.left.tag_add("paddingrow", *_padding_left)
        if _padding_right:
            self.right.tag_add("paddingrow", *_padding_right)

        mode = "只看差异" if self.only_diff_var.get() else "全量"
        total_rows = len(self.row_pairs) if self.row_pairs else self.max_row
        self.info.configure(text=f"{mode} | RowsShown: {len(self.display_rows)} / {total_rows}   Cols: {self.max_col}   DiffRows: {diff_row_count}")
        self._display_diff_row_count = diff_row_count

        self.app.set_sheet_has_diff(self.sheet, diff_row_count > 0, confirmed=True)
        self.app.refresh_sheet_nav()
        self._update_diff_nav_state()

        # Cache rendered result for fast toggle
        if row_only is None:
            text_a = "\n".join(lines_a) + ("\n" if lines_a else "")
            text_b = "\n".join(lines_b) + ("\n" if lines_b else "")
            self._render_cache[cache_key] = (text_a, text_b, tag_rows, tag_cells, diff_row_count)


class SowMergeApp:
    def __init__(self, file_a: str, file_b: str, merge_mode: bool = False, merged_path: str | None = None,
                 merge_conflict_cells_by_sheet: dict | None = None, merge_conflict_mode: bool = False):
        self.file_a = file_a
        self.file_b = file_b
        self.merge_mode = merge_mode
        self.merged_path = merged_path
        self.merge_conflict_cells_by_sheet = merge_conflict_cells_by_sheet or {}
        self.merge_conflict_mode = merge_conflict_mode
        self.user_touched_conflicts = False
        self.undo_stack = []
        # reset debug log each run
        try:
            with open(_DEBUG_LOG_PATH, "w", encoding="utf-8") as f:
                f.write(f"{APP_NAME} {APP_VERSION}\n")
                f.write(f"A={self.file_a}\nB={self.file_b}\n")
        except Exception:
            pass
        # load settings
        self.settings = {}
        self.only_diff_default = 0
        try:
            os.makedirs(os.path.dirname(_SETTINGS_PATH), exist_ok=True)
            if os.path.exists(_SETTINGS_PATH):
                with open(_SETTINGS_PATH, "r", encoding="utf-8") as f:
                    self.settings = json.load(f) or {}
            self.only_diff_default = int(self.settings.get("only_diff", 0))
        except Exception as e:
            _dlog(f"settings load failed: {e}")

        _dlog(f"SowMergeApp init only_diff_default={self.only_diff_default}")

        # Fast open: load value workbooks first; defer editable workbooks until first modification/save.
        self._wb_a_edit = None
        self._wb_b_edit = None
        self._edit_loaded_event = threading.Event()
        self._edit_loading_started = False

        t0 = datetime.now()
        self._file_a_val_path = _prepare_val_path(file_a)
        self._wb_a_val = load_workbook(self._file_a_val_path, data_only=True)
        _dlog(f"load wb_a_val: {(datetime.now()-t0).total_seconds():.3f}s")
        t0 = datetime.now()
        self._file_b_val_path = _prepare_val_path(file_b)
        self._wb_b_val = load_workbook(self._file_b_val_path, data_only=True)
        _dlog(f"load wb_b_val: {(datetime.now()-t0).total_seconds():.3f}s")

        # Preload editable workbooks in background to make the first overwrite fast.
        if not _FAST_OPEN_ENABLED:
            def _preload_edit():
                try:
                    _dlog("preload edit workbooks (background) start")
                    t1 = datetime.now()
                    a_edit = load_workbook(self.file_a, data_only=False)
                    _dlog(f"preload wb_a_edit: {(datetime.now()-t1).total_seconds():.3f}s")
                    t2 = datetime.now()
                    b_edit = load_workbook(self.file_b, data_only=False)
                    _dlog(f"preload wb_b_edit: {(datetime.now()-t2).total_seconds():.3f}s")
                    self._wb_a_edit = a_edit
                    self._wb_b_edit = b_edit
                except Exception as e:
                    _dlog(f"preload edit failed: {e}")
                finally:
                    self._edit_loaded_event.set()
                    _dlog("preload edit workbooks (background) done")

            self._edit_loading_started = True
            threading.Thread(target=_preload_edit, daemon=True).start()

        # Determine sheets from value workbooks (available immediately)
        set_a = set(self._wb_a_val.sheetnames)
        set_b = set(self._wb_b_val.sheetnames)
        self.common_sheets = sorted(set_a & set_b)
        if self.merge_conflict_mode and self.merge_conflict_cells_by_sheet:
            # Only keep sheets that actually have conflicts
            conflict_sheets = sorted(self.merge_conflict_cells_by_sheet.keys())
            self.common_sheets = [s for s in conflict_sheets if s in self.common_sheets]
        self.only_a = sorted(set_a - set_b)
        self.only_b = sorted(set_b - set_a)

        self.modified_a = False
        self.modified_b = False
        self.modified_sheets_a = set()
        self.modified_sheets_b = set()

        # sheet diff state: 0=none, 1=maybe (sampled), 2=confirmed
        self.sheet_diff_state = {s: 0 for s in self.common_sheets}

        self.root = tk.Tk()
        self._window_title_suffix = f"{APP_NAME} {APP_VERSION} [{APP_BUILD_TAG}]"
        self.root.title(self._window_title_suffix)
        ttk.Style().theme_use("clam")
        if self.merge_mode:
            self.root.title(f"{self._window_title_suffix} (SVN Merge)")
        else:
            self.root.title(f"{self._window_title_suffix} (TortoiseMerge-like)")
        self.root.geometry("1450x860")

        self._build_ui()

    def _ensure_edit_loaded(self):
        if self._wb_a_edit is not None and self._wb_b_edit is not None:
            return

        # If background preload is running, wait briefly.
        if getattr(self, "_edit_loading_started", False):
            _dlog("waiting for background edit preload")
            self._edit_loaded_event.wait(timeout=10)
            if self._wb_a_edit is not None and self._wb_b_edit is not None:
                return

        _dlog("loading edit workbooks (fallback)")
        t0 = datetime.now()
        self._wb_a_edit = load_workbook(self.file_a, data_only=False)
        _dlog(f"load wb_a_edit: {(datetime.now()-t0).total_seconds():.3f}s")
        t0 = datetime.now()
        self._wb_b_edit = load_workbook(self.file_b, data_only=False)
        _dlog(f"load wb_b_edit: {(datetime.now()-t0).total_seconds():.3f}s")

    def ws_a_edit(self, sheet: str):
        self._ensure_edit_loaded()
        return self._wb_a_edit[sheet]

    def ws_b_edit(self, sheet: str):
        self._ensure_edit_loaded()
        return self._wb_b_edit[sheet]

    def ws_a_val(self, sheet: str):
        return self._wb_a_val[sheet]

    def ws_b_val(self, sheet: str):
        return self._wb_b_val[sheet]

    def set_sheet_has_diff(self, sheet: str, has: bool, confirmed: bool = True):
        # Keep API: mark sheet diff state
        if sheet not in self.sheet_diff_state:
            return
        if has:
            self.sheet_diff_state[sheet] = 2 if confirmed else max(self.sheet_diff_state[sheet], 1)
        else:
            # only downgrade when confirmed
            if confirmed:
                self.sheet_diff_state[sheet] = 0

    def _build_ui(self):
        top = ttk.Frame(self.root)
        top.pack(fill="x", padx=10, pady=8)

        # Keep top area minimal (summary + buttons). Paths are shown inside each Sheet (requested).
        ttk.Label(top, text="左侧(A):").grid(row=0, column=0, sticky="w")
        ttk.Label(top, text=os.path.basename(self.file_a)).grid(row=0, column=1, sticky="w")
        ttk.Label(top, text="右侧(B):").grid(row=1, column=0, sticky="w")
        ttk.Label(top, text=os.path.basename(self.file_b)).grid(row=1, column=1, sticky="w")

        summary = f"同名Sheet: {len(self.common_sheets)}   仅A: {len(self.only_a)}   仅B: {len(self.only_b)}"
        ttk.Label(top, text=summary).grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Label(top, text=f"Build: {APP_BUILD_TAG}", foreground="#666").grid(row=0, column=3, sticky="ne", padx=(16, 0))

        ttk.Button(top, text="重算并刷新", command=self.recalc_and_refresh).grid(row=0, column=2, rowspan=2, sticky="ne", padx=(10, 0))

        ttk.Separator(self.root, orient="horizontal").pack(fill="x", padx=10, pady=(0, 6))

        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(8, 6))

        # Bottom bar: sheet nav (only)
        self.bottom = ttk.Frame(self.root)
        self.bottom.pack(fill="x", padx=10, pady=(0, 10))

        self.nav = ttk.Frame(self.bottom)
        self.nav.pack(side="left", fill="x", expand=True)
        ttk.Label(self.nav, text="Sheets:").pack(side="left")
        self.nav_canvas = tk.Canvas(self.nav, height=28, highlightthickness=0)
        self.nav_canvas.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.nav_scroll = ttk.Scrollbar(self.nav, orient="horizontal", command=self.nav_canvas.xview)
        self.nav_scroll.pack(side="bottom", fill="x")
        self.nav_canvas.configure(xscrollcommand=self.nav_scroll.set)
        self.nav_inner = ttk.Frame(self.nav_canvas)
        self.nav_canvas.create_window((0, 0), window=self.nav_inner, anchor="nw")
        self.nav_inner.bind("<Configure>", lambda e: self.nav_canvas.configure(scrollregion=self.nav_canvas.bbox("all")))

        if self.only_a:
            self._add_missing_tab("仅在A存在", self.only_a)
        if self.only_b:
            self._add_missing_tab("仅在B存在", self.only_b)

        # Tabs are created up-front, but heavy SheetView is created lazily on first activation.
        self.sheet_views = {}
        self._sheet_loaded = {}
        self._sheet_containers = {}
        for s in self.common_sheets:
            container = ttk.Frame(self.nb)
            self._sheet_containers[s] = container
            self.nb.add(container, text=s)
            self.sheet_views[s] = None
            self._sheet_loaded[s] = False

        # Background compute queue for sheet diffs
        self._compute_lock = threading.Lock()
        self._compute_queue = []  # list of sheet names
        self._compute_inflight = set()
        self._ui_task_lock = threading.Lock()
        self._ui_tasks = []

        def _enqueue_sheet(sheet: str, front: bool = False):
            with self._compute_lock:
                if sheet in self._compute_inflight:
                    return
                if sheet in self._compute_queue:
                    # move to front if requested
                    if front:
                        self._compute_queue.remove(sheet)
                        self._compute_queue.insert(0, sheet)
                    return
                if front:
                    self._compute_queue.insert(0, sheet)
                else:
                    self._compute_queue.append(sheet)

        def _queue_ui_task(fn):
            with self._ui_task_lock:
                self._ui_tasks.append(fn)

        def _drain_ui_tasks():
            tasks = []
            try:
                with self._ui_task_lock:
                    if self._ui_tasks:
                        tasks = self._ui_tasks
                        self._ui_tasks = []
            except Exception:
                tasks = []
            for fn in tasks:
                try:
                    fn()
                except Exception as e:
                    _dlog(f"ui task failed: {e}")
            try:
                self.root.after(50, _drain_ui_tasks)
            except Exception:
                pass

        def _compute_trim_bounds(ws):
            # Find last non-empty row/col; then +50 buffer.
            # Prefer ws._cells (only stored non-empty cells) to avoid missing data
            # when max_row/max_col are inflated by styles.
            max_r = ws.max_row or 1
            max_c = ws.max_column or 1
            last_r = 1
            last_c = 1
            found = False

            try:
                cells = getattr(ws, "_cells", None)
                if cells:
                    for cell in cells.values():
                        v = cell.value
                        if v not in (None, ""):
                            found = True
                            if cell.row > last_r:
                                last_r = cell.row
                            if cell.column > last_c:
                                last_c = cell.column
            except Exception:
                pass

            if not found:
                # Fallback: scan backwards up to 5000 rows.
                # If still not found, do NOT trim (avoid cutting real data).
                for r in range(max_r, max(1, max_r - 5000), -1):
                    row = next(ws.iter_rows(min_row=r, max_row=r, min_col=1, max_col=max_c, values_only=True), ())
                    if any(v not in (None, "") for v in row):
                        found = True
                        last_r = r
                        # determine last non-empty col from that row
                        for ci in range(len(row), 0, -1):
                            v = row[ci - 1]
                            if v not in (None, ""):
                                last_c = ci
                                break
                        break
                if not found:
                    return max_r, max_c

            use_r = min(max_r, last_r + 50)
            use_c = min(max_c, last_c + 50)
            return max(1, use_r), max(1, use_c)

        def _compute_row_pairs_bg(ws_a, ws_b, max_row_a: int, max_row_b: int, max_col: int):
            """Compute row alignment pairs using difflib.SequenceMatcher (background-safe)."""
            if max(max_row_a, max_row_b) >= _LARGE_SHEET_DIRECT_PAIR_THRESHOLD:
                max_row = max(max_row_a, max_row_b)
                pairs = []
                for r in range(1, max_row + 1):
                    ra = r if r <= max_row_a else None
                    rb = r if r <= max_row_b else None
                    pairs.append((ra, rb))
                return pairs

            def _bulk_sig_list(ws, max_row_local: int):
                try:
                    all_rows = list(ws.iter_rows(
                        min_row=1, max_row=max_row_local,
                        min_col=1, max_col=max_col,
                        values_only=True,
                    ))
                except Exception:
                    all_rows = []
                return ["\x1f".join(_merge_cmp_value(v) for v in (row or ())) for row in all_rows]

            sig_a = _bulk_sig_list(ws_a, max_row_a)
            sig_b = _bulk_sig_list(ws_b, max_row_b)
            sm = difflib.SequenceMatcher(a=sig_a, b=sig_b, autojunk=False)
            pairs: list[tuple[int | None, int | None]] = []
            for tag, i1, i2, j1, j2 in sm.get_opcodes():
                if tag == "equal":
                    for i, j in zip(range(i1, i2), range(j1, j2)):
                        pairs.append((i + 1, j + 1))
                elif tag == "replace":
                    len_a = i2 - i1
                    len_b = j2 - j1
                    common = min(len_a, len_b)
                    for k in range(common):
                        pairs.append((i1 + k + 1, j1 + k + 1))
                    for k in range(common, len_a):
                        pairs.append((i1 + k + 1, None))
                    for k in range(common, len_b):
                        pairs.append((None, j1 + k + 1))
                elif tag == "delete":
                    for i in range(i1, i2):
                        pairs.append((i + 1, None))
                elif tag == "insert":
                    for j in range(j1, j2):
                        pairs.append((None, j + 1))
            return pairs

        def _has_diff_by_blocks_bg(ws_a, ws_b, max_row_a: int, max_row_b: int, max_col: int):
            max_row = max(max_row_a, max_row_b)
            block = _LARGE_SHEET_BLOCK_ROWS
            for block_end in range(max_row, 0, -block):
                block_start = max(1, block_end - block + 1)
                rows_a = {}
                rows_b = {}
                if block_start <= max_row_a:
                    for idx, row in enumerate(
                        ws_a.iter_rows(
                            min_row=block_start,
                            max_row=min(block_end, max_row_a),
                            min_col=1,
                            max_col=max_col,
                            values_only=True,
                        ),
                        start=block_start,
                    ):
                        rows_a[idx] = row or ()
                if block_start <= max_row_b:
                    for idx, row in enumerate(
                        ws_b.iter_rows(
                            min_row=block_start,
                            max_row=min(block_end, max_row_b),
                            min_col=1,
                            max_col=max_col,
                            values_only=True,
                        ),
                        start=block_start,
                    ):
                        rows_b[idx] = row or ()
                for r in range(block_end, block_start - 1, -1):
                    row_a = rows_a.get(r, ())
                    row_b = rows_b.get(r, ())
                    if len(row_a) < max_col:
                        row_a = tuple(row_a) + (None,) * (max_col - len(row_a))
                    if len(row_b) < max_col:
                        row_b = tuple(row_b) + (None,) * (max_col - len(row_b))
                    sig_a = tuple(_merge_cmp_value(v) for v in row_a)
                    sig_b = tuple(_merge_cmp_value(v) for v in row_b)
                    if sig_a != sig_b:
                        return True
            return False

        def _compute_sheet_cache(wb_a_val, wb_b_val, wb_a_edit, wb_b_edit, sheet: str):
            ws_a = wb_a_val[sheet]
            ws_b = wb_b_val[sheet]
            max_r_a, max_c_a = _compute_trim_bounds(ws_a)
            max_r_b, max_c_b = _compute_trim_bounds(ws_b)
            max_row = max(max_r_a, max_r_b)
            max_col = max(max_c_a, max_c_b)

            # Compute row-aligned pairs (same algorithm as SheetView._build_row_pairs)
            row_pairs = _compute_row_pairs_bg(ws_a, ws_b, max_r_a, max_r_b, max_col)

            pair_diff_cols: dict[int, set] = {}
            pair_text_a: dict[int, str] = {}
            pair_text_b: dict[int, str] = {}
            row_a_to_pair_idx: dict[int, int] = {}
            row_b_to_pair_idx: dict[int, int] = {}

            for idx, (ra, rb) in enumerate(row_pairs):
                if ra is not None:
                    row_a_to_pair_idx[ra] = idx
                if rb is not None:
                    row_b_to_pair_idx[rb] = idx

            # Large-sheet fast open: avoid full cell-by-cell precompute.
            if max_row >= _LARGE_SHEET_ROW_THRESHOLD:
                has_diff = _has_diff_by_blocks_bg(ws_a, ws_b, max_r_a, max_r_b, max_col)
            else:
                ws_a_e = wb_a_edit[sheet]
                ws_b_e = wb_b_edit[sheet]
                for idx, (ra, rb) in enumerate(row_pairs):
                    cols = set()
                    parts_a = []
                    parts_b = []
                    for c in range(1, max_col + 1):
                        da, db, eq = _cell_display_and_equal_by_row(ws_a, ws_b, ws_a_e, ws_b_e, ra, rb, c)
                        parts_a.append(_val_to_str(da))
                        parts_b.append(_val_to_str(db))
                        if not eq:
                            cols.add(c)
                    label_a = str(ra) if ra is not None else ""
                    label_b = str(rb) if rb is not None else ""
                    pair_text_a[idx] = label_a + "\t" + "\t".join(parts_a)
                    pair_text_b[idx] = label_b + "\t" + "\t".join(parts_b)
                    pair_diff_cols[idx] = cols
                has_diff = any(bool(v) for v in pair_diff_cols.values())
            return {
                "sheet": sheet,
                "max_row": max_row,
                "max_col": max_col,
                "row_pairs": row_pairs,
                "pair_diff_cols": pair_diff_cols,
                "pair_text_a": pair_text_a,
                "pair_text_b": pair_text_b,
                "row_a_to_pair_idx": row_a_to_pair_idx,
                "row_b_to_pair_idx": row_b_to_pair_idx,
                "has_diff": has_diff,
            }

        def _apply_sheet_cache(cache: dict):
            sheet = cache["sheet"]
            # update confirmed state first; this also works when the tab view
            # is not created yet (lazy sheet loading).
            self.set_sheet_has_diff(sheet, cache.get("has_diff", False), confirmed=True)
            view = self.sheet_views.get(sheet)
            if view is None:
                self.refresh_sheet_nav()
                return
            # Skip if the user has made edits in this view; background data (from read-only copies)
            # would be stale relative to the user's in-memory changes.
            if getattr(view, "_data_ready", False) and view.touched_rows:
                self.refresh_sheet_nav()
                return
            view.max_row = cache["max_row"]
            view.max_col = cache["max_col"]
            view._is_large_sheet = view.max_row >= _LARGE_SHEET_ROW_THRESHOLD
            view._bounds_checked = True

            # Apply row-aligned pair data (computed in background with row alignment)
            view.row_pairs = cache["row_pairs"]
            view.pair_diff_cols = cache["pair_diff_cols"]
            view.pair_text_a = cache["pair_text_a"]
            view.pair_text_b = cache["pair_text_b"]
            view.row_a_to_pair_idx = cache["row_a_to_pair_idx"]
            view.row_b_to_pair_idx = cache["row_b_to_pair_idx"]
            view._align_rows_enabled = True
            view._diff_partial = False
            # Mark data as ready so refresh(rescan=False) uses it without rescanning
            view._data_ready = True
            view._invalidate_render_cache()

            if view._prefer_only_diff_when_ready:
                # For large sheets, keep full-mode initial render (first 200 rows) for responsiveness.
                if view.max_row >= _LARGE_SHEET_ROW_THRESHOLD:
                    view.only_diff_var.set(0)
                    view._full_render = False
                    view._render_limit = min(_LARGE_SHEET_INITIAL_ROWS, view.max_row)
                elif cache.get("has_diff", False):
                    view.only_diff_var.set(1)
                else:
                    view.only_diff_var.set(0)
                    view._full_render = False
                    view._render_limit = min(_LARGE_SHEET_INITIAL_ROWS, view.max_row)
                view._prefer_only_diff_when_ready = False
            view.refresh(row_only=None, rescan=False)
            # Reset cursor to line 1 after full re-render so _update_cursor_lines
            # shows row 1 instead of a stale/out-of-range position.
            try:
                view.left.mark_set("insert", "1.0")
                view.right.mark_set("insert", "1.0")
            except Exception:
                pass
            view._update_cursor_lines()
            self.refresh_sheet_nav()

        def _compute_worker():
            wb_a_ro = None
            wb_b_ro = None
            wb_a_e = None
            wb_b_e = None
            try:
                # Use separate read-only workbooks to avoid threading issues
                wb_a_ro = load_workbook(self._file_a_val_path, data_only=True, read_only=True)
                wb_b_ro = load_workbook(self._file_b_val_path, data_only=True, read_only=True)
                wb_a_e = load_workbook(self.file_a, data_only=False, read_only=True)
                wb_b_e = load_workbook(self.file_b, data_only=False, read_only=True)
            except Exception as e:
                _dlog(f"bg compute open read-only failed: {e}")
                return
            if wb_a_ro is None or wb_b_ro is None or wb_a_e is None or wb_b_e is None:
                _dlog("bg compute read-only workbooks not available; skip background compute")
                return
            while True:
                with self._compute_lock:
                    if not self._compute_queue:
                        break
                    sheet = self._compute_queue.pop(0)
                    self._compute_inflight.add(sheet)
                try:
                    _dlog(f"bg compute sheet: {sheet}")
                    cache = _compute_sheet_cache(wb_a_ro, wb_b_ro, wb_a_e, wb_b_e, sheet)
                    # Never call tkinter APIs from background threads.
                    _queue_ui_task(lambda c=cache: _apply_sheet_cache(c))
                except Exception as e:
                    _dlog(f"bg compute failed {sheet}: {e}")
                finally:
                    with self._compute_lock:
                        self._compute_inflight.discard(sheet)
            try:
                wb_a_ro.close()
                wb_b_ro.close()
                wb_a_e.close()
                wb_b_e.close()
            except Exception:
                pass

        def _kick_worker():
            # start a worker if not running
            with self._compute_lock:
                running = bool(self._compute_inflight)
            if running:
                return
            threading.Thread(target=_compute_worker, daemon=True).start()
        self._kick_worker = _kick_worker

        # Lazy-create SheetView UI immediately; compute diff in background.
        def _on_tab_changed(_evt=None):
            try:
                tab_id = self.nb.select()
                tab_text = self.nb.tab(tab_id, "text")
                self.selected_sheet = tab_text
                self.refresh_sheet_nav()
                if tab_text in self._sheet_containers and not self._sheet_loaded.get(tab_text, False):
                    _dlog(f"lazy create SheetView (ui only): {tab_text}")
                    view = SheetView(self._sheet_containers[tab_text], self, tab_text)
                    self.sheet_views[tab_text] = view
                    self._sheet_loaded[tab_text] = True
                    # Show loading placeholder immediately (non-blocking).
                    # The background worker will compute diffs and call _apply_sheet_cache
                    # which sets _data_ready=True and calls refresh(rescan=False).
                    view._show_loading()
                if tab_text in self._sheet_containers:
                    # Skip background recompute if data is already ready (no edits pending).
                    # Reopening workbooks on every tab switch is the main perf regression.
                    _view = self.sheet_views.get(tab_text)
                    if not (_view and getattr(_view, "_data_ready", False)):
                        _enqueue_sheet(tab_text, front=True)
                        _kick_worker()
                        # Fallback: if background path stalls, force sync refresh on UI thread.
                        # This guarantees that switching tabs eventually shows compare results.
                        def _force_refresh_if_still_loading(sheet_name=tab_text):
                            try:
                                cur_id = self.nb.select()
                                cur_sheet = self.nb.tab(cur_id, "text")
                                if cur_sheet != sheet_name:
                                    return
                                v = self.sheet_views.get(sheet_name)
                                if not v:
                                    return
                                if getattr(v, "_data_ready", False):
                                    return
                                v.refresh(row_only=None, rescan=True)
                                v._update_cursor_lines()
                            except Exception as e:
                                _dlog(f"force refresh fallback failed {sheet_name}: {e}")
                        try:
                            self.root.after(700, _force_refresh_if_still_loading)
                        except Exception:
                            pass
            except Exception as e:
                _dlog(f"tab changed handler failed: {e}")

        try:
            self.nb.bind("<<NotebookTabChanged>>", _on_tab_changed)
        except Exception:
            pass

        # Main-thread UI task pump (for background compute/sample updates).
        try:
            self.root.after(50, _drain_ui_tasks)
        except Exception:
            pass

        # Load the initially selected tab immediately so first-open state is ready.
        _on_tab_changed()

        self.refresh_sheet_nav()

        # Background fast pre-mark for sheet tabs:
        # exact by cached values (tail-first block scan), no random sampling.
        def _apply_fast_mark_result(sheet: str, has: bool):
            self.set_sheet_has_diff(sheet, has, confirmed=True)
            view = self.sheet_views.get(sheet)
            if has and view and view._prefer_only_diff_when_ready and view.only_diff_var.get() == 0:
                view.only_diff_var.set(1)
                view.refresh(row_only=None, rescan=False)
                view._update_cursor_lines()
            self.refresh_sheet_nav()

        def _sheet_has_diff_fast_tail(ws_a, ws_b, max_row: int, max_col: int, min_row: int = 1):
            none_sig = tuple("" for _ in range(max_col))
            block = _LARGE_SHEET_BLOCK_ROWS
            max_row_a = ws_a.max_row or 1
            max_row_b = ws_b.max_row or 1

            for block_end in range(max_row, 0, -block):
                block_start = max(1, block_end - block + 1)
                if block_end < min_row:
                    break
                if block_start < min_row:
                    block_start = min_row
                end_a = min(block_end, max_row_a)
                end_b = min(block_end, max_row_b)

                rows_a = []
                rows_b = []
                if block_start <= end_a:
                    rows_a = list(ws_a.iter_rows(
                        min_row=block_start,
                        max_row=end_a,
                        min_col=1,
                        max_col=max_col,
                        values_only=True,
                    ))
                if block_start <= end_b:
                    rows_b = list(ws_b.iter_rows(
                        min_row=block_start,
                        max_row=end_b,
                        min_col=1,
                        max_col=max_col,
                        values_only=True,
                    ))

                sig_a = [tuple(_merge_cmp_value(v) for v in (row or ())) for row in rows_a]
                sig_b = [tuple(_merge_cmp_value(v) for v in (row or ())) for row in rows_b]

                for r in range(block_end, block_start - 1, -1):
                    if r <= max_row_a:
                        ia = r - block_start
                        sa = sig_a[ia] if 0 <= ia < len(sig_a) else none_sig
                    else:
                        sa = none_sig
                    if r <= max_row_b:
                        ib = r - block_start
                        sb = sig_b[ib] if 0 <= ib < len(sig_b) else none_sig
                    else:
                        sb = none_sig
                    if sa != sb:
                        return True
            return False

        def _sheet_has_diff_quick_tail(ws_a, ws_b, max_row: int, max_col: int):
            # Phase-1 quick check: scan only the tail window.
            # True means "confirmed diff"; False means "unknown yet".
            quick_rows = min(max_row, _TABMARK_QUICK_TAIL_ROWS)
            if quick_rows <= 0:
                return False
            start = max(1, max_row - quick_rows + 1)
            return _sheet_has_diff_fast_tail(ws_a, ws_b, max_row, max_col, min_row=start)

        def _scan_sheet_has_diff_fast():
            wb_a_ro = None
            wb_b_ro = None
            try:
                wb_a_ro = load_workbook(self._file_a_val_path, data_only=True, read_only=True)
                wb_b_ro = load_workbook(self._file_b_val_path, data_only=True, read_only=True)
                ordered = list(self.common_sheets)
                if ordered:
                    # Prefer currently selected sheet first, then newer tabs first.
                    cur = getattr(self, "selected_sheet", None)
                    if cur in ordered:
                        ordered.remove(cur)
                        ordered = [cur] + list(reversed(ordered))
                    else:
                        ordered = list(reversed(ordered))

                unknown_sheets = []

                # Phase-1: quick tail scan to surface diff tabs early.
                for s in ordered:
                    ws_a = wb_a_ro[s]
                    ws_b = wb_b_ro[s]
                    max_row = max(ws_a.max_row or 1, ws_b.max_row or 1)
                    max_col = max(ws_a.max_column or 1, ws_b.max_column or 1)
                    has_quick = _sheet_has_diff_quick_tail(ws_a, ws_b, max_row, max_col)
                    if has_quick:
                        _queue_ui_task(lambda s=s: _apply_fast_mark_result(s, True))
                    else:
                        unknown_sheets.append((s, ws_a, ws_b, max_row, max_col))

                # Phase-2: full exact scan for unresolved sheets.
                for s, ws_a, ws_b, max_row, max_col in unknown_sheets:
                    has = _sheet_has_diff_fast_tail(ws_a, ws_b, max_row, max_col)
                    _queue_ui_task(lambda s=s, has=has: _apply_fast_mark_result(s, has))
            except Exception as e:
                _dlog(f"fast diff mark scan failed: {e}")
            finally:
                try:
                    if wb_a_ro:
                        wb_a_ro.close()
                    if wb_b_ro:
                        wb_b_ro.close()
                except Exception:
                    pass

        try:
            threading.Thread(target=_scan_sheet_has_diff_fast, daemon=True).start()
        except Exception:
            pass

        # Enqueue all sheets for background confirmation (slow compute)
        try:
            for s in self.common_sheets:
                _enqueue_sheet(s, front=False)
            _kick_worker()
        except Exception as e:
            _dlog(f"enqueue all sheets failed: {e}")

    def push_undo(self, action: dict):
        try:
            self.undo_stack.append(action)
            if len(self.undo_stack) > 3:
                self.undo_stack.pop(0)
        except Exception:
            pass

    def pop_undo(self) -> dict | None:
        try:
            if not self.undo_stack:
                return None
            return self.undo_stack.pop()
        except Exception:
            return None

    def _add_missing_tab(self, title: str, items):
        frame = ttk.Frame(self.nb)
        self.nb.add(frame, text=title)
        ttk.Label(frame, text=title, font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=8, pady=(8, 4))
        txt = tk.Text(frame, wrap="none", height=10)
        txt.pack(fill="both", expand=True, padx=8, pady=8)
        txt.insert("1.0", "\n".join(items))
        txt.configure(state="disabled")

    def _select_tab(self, tab_text: str):
        for tab_id in self.nb.tabs():
            if self.nb.tab(tab_id, "text") == tab_text:
                self.nb.select(tab_id)
                return

    def refresh_sheet_nav(self):
        for child in list(self.nav_inner.winfo_children()):
            child.destroy()

        try:
            from tkinter import font as tkfont
            if not hasattr(self, "_nav_font"):
                self._nav_font = tkfont.nametofont("TkDefaultFont")
                self._nav_font_bold = self._nav_font.copy()
                self._nav_font_bold.configure(weight="bold")
        except Exception:
            self._nav_font = None
            self._nav_font_bold = None

        def add_btn(label: str, tab_text: str, kind: str, state: int = 0):
            if kind in ("onlyA", "onlyB"):
                bg = "#FFE5E5"
            else:
                # 0=none, 1=maybe (pale), 2=confirmed (bright)
                if state >= 2:
                    bg = "#FFD400"  # bright yellow
                elif state == 1:
                    bg = "#FFF3B0"  # pale yellow
                else:
                    bg = "#F2F2F2"
            is_selected = (tab_text == getattr(self, "selected_sheet", None))
            if is_selected:
                bg = "#D9D9D9"
            b = tk.Button(self.nav_inner, text=label,
                          relief="sunken" if is_selected else "groove",
                          bd=2 if is_selected else 1,
                          padx=8, pady=2, bg=bg,
                          command=lambda: self._select_tab(tab_text))
            try:
                if is_selected and self._nav_font_bold:
                    b.configure(font=self._nav_font_bold)
                elif self._nav_font:
                    b.configure(font=self._nav_font)
            except Exception:
                pass
            b.pack(side="left", padx=4)

        if self.only_a:
            add_btn(f"仅A({len(self.only_a)})", "仅在A存在", "onlyA")
        if self.only_b:
            add_btn(f"仅B({len(self.only_b)})", "仅在B存在", "onlyB")

        for s in self.common_sheets:
            add_btn(s, s, "common", state=int(self.sheet_diff_state.get(s, 0)))

        self.nav_canvas.update_idletasks()
        self.nav_canvas.configure(scrollregion=self.nav_canvas.bbox("all"))

    def open_textdiff(self):
        try:
            temp_root = os.path.join(os.environ.get("LOCALAPPDATA", tempfile.gettempdir()), "Temp", "TortoiseXlsTemp")
            os.makedirs(temp_root, exist_ok=True)
        except Exception:
            temp_root = tempfile.gettempdir()

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        left_txt = os.path.join(temp_root, f"{APP_NAME}_left_{ts}.txt")
        right_txt = os.path.join(temp_root, f"{APP_NAME}_right_{ts}.txt")
        excel_to_text(self.file_a, left_txt, thick_sep_char="=")
        excel_to_text(self.file_b, right_txt, thick_sep_char="=")
        open_tortoise_merge(left_txt, right_txt, title=f"{APP_NAME}: {os.path.basename(self.file_a)}")

    def recalc_and_refresh(self):
        # Manual: force Excel recalc to refresh cached values, then reload view.
        def _do_recalc():
            new_a = _recalc_and_prepare_val_path(self.file_a)
            new_b = _recalc_and_prepare_val_path(self.file_b)
            if new_a:
                self._file_a_val_path = new_a
                self._wb_a_val = load_workbook(new_a, data_only=True)
            if new_b:
                self._file_b_val_path = new_b
                self._wb_b_val = load_workbook(new_b, data_only=True)

            # Refresh current sheet immediately
            try:
                tab_id = self.nb.select()
                tab_text = self.nb.tab(tab_id, "text")
                view = self.sheet_views.get(tab_text)
                if view:
                    view.refresh(row_only=None, rescan=True)
            except Exception:
                pass
            # Recompute diff states in background
            try:
                for s in self.common_sheets:
                    self.set_sheet_has_diff(s, False, confirmed=False)
                self._compute_queue = [s for s in self.common_sheets if s not in self._compute_inflight]
                self._kick_worker()
            except Exception:
                pass

        try:
            self._with_progress("重算中", "正在重算并刷新，请稍候...", _do_recalc)
        except Exception as e:
            messagebox.showerror("重算失败", f"重算失败：\n{e}")

    def _with_progress(self, title: str, message: str, fn):
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)
        dlg.geometry("+{}+{}".format(self.root.winfo_rootx() + 200, self.root.winfo_rooty() + 150))
        ttk.Label(dlg, text=message, padding=10).pack(fill="x")
        pb = ttk.Progressbar(dlg, mode="indeterminate")
        pb.pack(fill="x", padx=10, pady=(0, 10))
        pb.start(10)
        self.root.update_idletasks()
        try:
            fn()
        finally:
            try:
                pb.stop()
                dlg.destroy()
            except Exception:
                pass

    def _atomic_save(self, wb, target_path: str):
        """Safely overwrite a workbook.

        Writes to a temp file in the same directory, then os.replace.
        This avoids corrupting the target if the process is interrupted.
        """
        if _FAST_SAVE_ENABLED:
            # Fast path: write directly (faster, but not atomic)
            try:
                if _FAST_SAVE_VALUES_ONLY and _USE_CACHED_VALUES_ONLY:
                    _save_values_only_from_wb(wb, target_path)
                else:
                    wb.save(target_path)
                return
            except PermissionError:
                # fallback to safe path below
                pass
        folder = os.path.dirname(target_path)
        base = os.path.basename(target_path)
        tmp_path = os.path.join(folder, f"~{base}.{os.getpid()}.tmp")
        if _FAST_SAVE_VALUES_ONLY and _USE_CACHED_VALUES_ONLY:
            _save_values_only_from_wb(wb, tmp_path)
        else:
            wb.save(tmp_path)
        try:
            os.replace(tmp_path, target_path)
            return
        except PermissionError:
            # Try clearing readonly flag then retry a few times (file may be locked briefly)
            try:
                if os.path.exists(target_path):
                    os.chmod(target_path, stat.S_IWRITE | stat.S_IREAD)
            except Exception:
                pass
            # If we can delete the target, try that once (replace may fail on readonly)
            try:
                if os.path.exists(target_path):
                    os.remove(target_path)
            except Exception:
                pass
            for _ in range(8):
                try:
                    os.replace(tmp_path, target_path)
                    return
                except PermissionError:
                    time.sleep(0.5)
            # If replace keeps failing, try overwrite-in-place (requires write but not delete)
            try:
                with open(tmp_path, "rb") as src, open(target_path, "wb") as dst:
                    shutil.copyfileobj(src, dst, length=1024 * 1024)
                return
            except Exception:
                raise
        except Exception:
            # Last-resort fallback to shutil.move
            try:
                shutil.move(tmp_path, target_path)
                return
            except Exception:
                raise
        finally:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass

    def _atomic_save_with_retry(self, wb, target_path: str, retries: int = 6, delay_sec: float = 0.5):
        """Retry save when target is temporarily locked."""
        last_err = None
        for _ in range(max(1, retries)):
            try:
                self._atomic_save(wb, target_path)
                return
            except Exception as e:
                if getattr(e, "winerror", None) in (5, 32) or isinstance(e, PermissionError):
                    last_err = e
                    time.sleep(delay_sec)
                    continue
                raise
        if last_err:
            raise last_err

    def _alt_save_path(self, path: str, which: str):
        folder = os.path.dirname(path)
        base, ext = os.path.splitext(os.path.basename(path))
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        return os.path.join(folder, f"{base}_{which}_saved_{ts}{ext or '.xlsx'}")

    def _try_alt_save(self, wb, path: str, which: str) -> bool:
        alt = self._alt_save_path(path, which)
        try:
            self._atomic_save(wb, alt)
            messagebox.showinfo("另存为成功", f"无法覆盖原文件，已另存为：\n{alt}")
            return True
        except Exception as e:
            messagebox.showerror("另存为失败", f"另存为失败：\n{e}")
            return False

    def _path_diagnostics(self, path: str) -> str:
        try:
            folder = os.path.dirname(path)
            exists = os.path.exists(path)
            readonly = False
            if exists:
                try:
                    readonly = not os.access(path, os.W_OK)
                except Exception:
                    readonly = False
            dir_writable = False
            try:
                test_file = os.path.join(folder, f"~perm_test_{os.getpid()}.tmp")
                with open(test_file, "w", encoding="utf-8") as f:
                    f.write("x")
                os.remove(test_file)
                dir_writable = True
            except Exception:
                dir_writable = False
            return f"exists={exists}, readonly={readonly}, dir_writable={dir_writable}"
        except Exception:
            return "diagnostics_failed"

    def _confirm_overwrite(self, which: str, path: str) -> bool:
        if which == "A":
            modified = self.modified_a
        else:
            modified = self.modified_b

        if not modified:
            return messagebox.askyesno("提示", f"{which} 没有检测到改动。仍然要覆盖保存原文件吗？\n\n{path}")

        return messagebox.askyesno(
            "确认保存",
            f"将直接覆盖保存 {which} 文件（原路径、原文件名）：\n\n{path}\n\n建议确保该 Excel 未在 WPS/Excel 中打开。继续吗？",
        )

    def save_b_inplace(self):
        self._ensure_edit_loaded()
        path = self.file_b
        if not self._confirm_overwrite("B", path):
            return
        try:
            self._with_progress("保存中", f"正在保存：\n{path}", lambda: self._atomic_save(self._wb_b_edit, path))
            self.modified_b = False
            messagebox.showinfo("Saved", f"已保存并覆盖：\n{path}")
        except Exception as e:
            # If the file is locked or denied, offer save-as fallback
            if getattr(e, "winerror", None) in (5, 32) or isinstance(e, PermissionError):
                diag = self._path_diagnostics(path)
                if messagebox.askyesno("保存失败", f"保存 B 失败（可能文件被占用/无权限）：\n{e}\n\n诊断：{diag}\n\n是否另存为？"):
                    if self._try_alt_save(self._wb_b_edit, path, "B"):
                        self.modified_b = False
                        return
            messagebox.showerror("保存失败", f"保存 B 失败：\n{e}")

    def save_a_inplace(self):
        self._ensure_edit_loaded()
        path = self.file_a
        if not self._confirm_overwrite("A", path):
            return
        try:
            self._with_progress("保存中", f"正在保存：\n{path}", lambda: self._atomic_save(self._wb_a_edit, path))
            self.modified_a = False
            messagebox.showinfo("Saved", f"已保存并覆盖：\n{path}")
        except Exception as e:
            if getattr(e, "winerror", None) in (5, 32) or isinstance(e, PermissionError):
                diag = self._path_diagnostics(path)
                if messagebox.askyesno("保存失败", f"保存 A 失败（可能文件被占用/无权限）：\n{e}\n\n诊断：{diag}\n\n是否另存为？"):
                    if self._try_alt_save(self._wb_a_edit, path, "A"):
                        self.modified_a = False
                        return
            messagebox.showerror("保存失败", f"保存 A 失败：\n{e}")

    def save_merged_and_exit(self, auto: bool = False):
        if not self.merged_path:
            return
        self._ensure_edit_loaded()
        if not auto:
            if not messagebox.askyesno("确认保存", f"将保存合并结果到：\n\n{self.merged_path}\n\n继续吗？"):
                return
        try:
            # Small delay to allow SVN/Tortoise to release file locks
            try:
                time.sleep(1.2)
            except Exception:
                pass
            self._with_progress("保存中", f"正在保存合并结果：\n{self.merged_path}",
                                lambda: self._atomic_save_with_retry(self._wb_a_edit, self.merged_path))
            self.modified_a = False
            try:
                messagebox.showinfo("Saved", f"已保存合并结果：\n{self.merged_path}")
            except Exception:
                pass
        except Exception as e:
            if getattr(e, "winerror", None) in (5, 32) or isinstance(e, PermissionError):
                excel_locked = _log_lock_holders(self.merged_path)
                if excel_locked:
                    try:
                        messagebox.showwarning("文件被占用", "检测到 Excel 正在占用目标文件。\n请关闭 Excel 后再保存。")
                    except Exception:
                        pass
                # In conflict UI, target file might still be locked by SVN/Tortoise.
                # Save to a temp file and schedule a deferred replace.
                try:
                    folder = os.path.dirname(self.merged_path)
                    base = os.path.basename(self.merged_path)
                    tmp_path = os.path.join(folder, f"~{base}.deferred.{os.getpid()}.tmp")
                    if _FAST_SAVE_VALUES_ONLY and _USE_CACHED_VALUES_ONLY:
                        _save_values_only_from_wb(self._wb_a_edit, tmp_path)
                    else:
                        self._wb_a_edit.save(tmp_path)
                    _launch_deferred_copy(tmp_path, self.merged_path)
                    messagebox.showinfo("保存中", f"目标文件被占用，已写入临时文件并将在关闭后自动覆盖：\n{self.merged_path}")
                    try:
                        self.root.destroy()
                    except Exception:
                        pass
                    sys.exit(0)
                except Exception:
                    diag = self._path_diagnostics(self.merged_path)
                    if messagebox.askyesno("保存失败", f"保存合并结果失败（可能文件被占用/无权限）：\n{e}\n\n诊断：{diag}\n\n是否另存为？"):
                        if self._try_alt_save(self._wb_a_edit, self.merged_path, "MERGED"):
                            try:
                                self.root.destroy()
                            except Exception:
                                pass
                            sys.exit(0)
                        return
            messagebox.showerror("保存失败", f"保存合并结果失败：\n{e}")
            return
        # Try auto-resolve in SVN if conflict artifacts exist
        try:
            if _has_svn_conflict_artifacts(self.merged_path):
                _try_svn_resolve(self.merged_path)
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass
        sys.exit(0)

    def resolve_conflict_cell(self, sheet: str, r: int, c: int) -> bool:
        rows = self.merge_conflict_cells_by_sheet.get(sheet)
        if not rows:
            return False
        cols = rows.get(r)
        if not cols or c not in cols:
            return False
        cols.discard(c)
        if not cols:
            rows.pop(r, None)
        if not rows:
            self.merge_conflict_cells_by_sheet.pop(sheet, None)
        self._auto_save_if_no_conflicts()
        return True

    def resolve_conflict_row(self, sheet: str, r: int, cols) -> bool:
        rows = self.merge_conflict_cells_by_sheet.get(sheet)
        if not rows or r not in rows:
            return False
        for c in list(cols):
            rows[r].discard(c)
        if not rows[r]:
            rows.pop(r, None)
        if not rows:
            self.merge_conflict_cells_by_sheet.pop(sheet, None)
        self._auto_save_if_no_conflicts()
        return True

    def _auto_save_if_no_conflicts(self):
        if not self.merge_conflict_cells_by_sheet:
            # If user has manually touched conflicts, require explicit save.
            if getattr(self, "user_touched_conflicts", False):
                return
            # all conflicts resolved
            self.save_merged_and_exit(auto=True)

    def run(self):
        self.root.mainloop()


def main():
    try:
        # Log raw args early for troubleshooting (even if argparse fails)
        try:
            _dlog(f"argv: {' '.join(sys.argv[1:])}")
        except Exception:
            pass

        def _parse_slash_args(argv):
            out = {}
            for a in argv:
                la = a.lower()
                if la.startswith("/path:"):
                    out["path"] = a[6:]
                elif la.startswith("/path2:"):
                    out["path2"] = a[7:]
                elif la.startswith("/base:"):
                    out["base"] = a[6:]
                elif la.startswith("/mine:"):
                    out["mine"] = a[6:]
                elif la.startswith("/theirs:"):
                    out["theirs"] = a[8:]
                elif la.startswith("/merged:"):
                    out["merged"] = a[8:]
            return out

        slash_args = _parse_slash_args(sys.argv[1:])

        parser = argparse.ArgumentParser(add_help=True)
        parser.add_argument("file_a", nargs="?")
        parser.add_argument("file_b", nargs="?")
        # SVN/TortoiseSVN style args
        parser.add_argument("--base")
        parser.add_argument("--mine")
        parser.add_argument("--theirs")
        parser.add_argument("--merged")
        parser.add_argument("--title")
        parser.add_argument("--textdiff", action="store_true", help="Only generate text files and open TortoiseMerge")
        args, unknown = parser.parse_known_args()
        if unknown:
            try:
                _dlog(f"unknown args: {unknown}")
            except Exception:
                pass

        # Map /path:/path2:/base: style args (TortoiseProc)
        if not args.base and "base" in slash_args:
            args.base = slash_args.get("base")
        if not args.mine and "mine" in slash_args:
            args.mine = slash_args.get("mine")
        if not args.theirs and "theirs" in slash_args:
            args.theirs = slash_args.get("theirs")
        if not args.merged and "merged" in slash_args:
            args.merged = slash_args.get("merged")
        if not args.file_a and "path" in slash_args:
            args.file_a = slash_args.get("path")
        if not args.file_b and "path2" in slash_args:
            args.file_b = slash_args.get("path2")

        # Map SVN-style args to our 2-pane viewer (diff mode)
        if args.base and args.mine and not args.theirs:
            a, b = args.base, args.mine
        elif args.file_a and args.file_b:
            a, b = args.file_a, args.file_b
        elif args.file_a and (not args.file_b) and (not args.base) and (not args.mine) and (not args.theirs):
            # Single file provided (e.g., from Explorer/TortoiseSVN). If it's a conflicted file, auto-merge it.
            conflict = _detect_svn_conflict_files(args.file_a)
            if conflict:
                args.base, args.mine, args.theirs, args.merged = conflict
                args.force_ui = True
            else:
                a, b = args.file_a, None
        else:
            sel = pick_files_or_conflict()
            if not sel:
                return
            if sel[0] == "merge":
                _mode, base_p, mine_p, theirs_p, merged_p, force_ui = sel
                args.base = _ensure_xlsx_copy(base_p)
                args.mine = _ensure_xlsx_copy(mine_p)
                args.theirs = _ensure_xlsx_copy(theirs_p)
                args.merged = merged_p
                args.force_ui = bool(force_ui)
            else:
                _mode, a, b = sel

        if args.file_a and (args.file_b is None) and (not args.base) and b is None:
            # Need second file for diff mode
            root = tk.Tk()
            root.withdraw()
            b = filedialog.askopenfilename(title="Select second .xlsx file (same filename)", filetypes=[("Excel Workbook", "*.xlsx")])
            if not b:
                return
            if os.path.basename(args.file_a).lower() != os.path.basename(b).lower():
                messagebox.showerror(
                    "Filename mismatch",
                    f"The two files must have the same filename.\n\nA: {os.path.basename(args.file_a)}\nB: {os.path.basename(b)}",
                )
                return
            a = args.file_a

        # Normalize SVN merge temp files (merge-left/right.r####) by exporting true revision.
        if args.base:
            args.base = _try_export_svn_revision_from_merge_temp(args.base)
        if args.mine:
            args.mine = _try_export_svn_revision_from_merge_temp(args.mine)
        if args.theirs:
            args.theirs = _try_export_svn_revision_from_merge_temp(args.theirs)
        if args.file_a:
            args.file_a = _try_export_svn_revision_from_merge_temp(args.file_a)
        if args.file_b:
            args.file_b = _try_export_svn_revision_from_merge_temp(args.file_b)

        # Merge mode: apply theirs onto mine, detect conflicts, save merged, then exit.
        if args.base and args.mine and args.theirs and args.merged:
            conflicts = []
            preview_path = None
            conflict_map = {}
            force_ui = bool(getattr(args, "force_ui", False))
            if unknown and any(str(u).lower() in (":", "working", "base") for u in unknown):
                force_ui = True
            if not force_ui and _has_svn_conflict_artifacts(args.merged):
                # SVN conflict state detected: force UI review even if auto-merge finds no conflicts.
                force_ui = True
            try:
                _dlog(f"merge args: base={args.base} mine={args.mine} theirs={args.theirs} merged={args.merged}")
                _dlog(f"merge force_ui={force_ui} unknown={unknown}")
            except Exception:
                pass
            try:
                # Ensure .r#### / .mine files are converted to temp .xlsx for openpyxl
                args.base = _ensure_xlsx_copy(args.base)
                args.mine = _ensure_xlsx_copy(args.mine)
                args.theirs = _ensure_xlsx_copy(args.theirs)
                try:
                    _dlog("merge start: calling _merge_three_way")
                except Exception:
                    pass
                conflicts, preview_path, conflict_map = _merge_three_way(
                    args.base, args.mine, args.theirs, args.merged, save_merged=(not force_ui)
                )
                try:
                    _dlog(f"merge result: conflicts={len(conflicts)} preview={preview_path} conflict_sheets={len(conflict_map)}")
                except Exception:
                    pass
            except Exception as e:
                try:
                    _dlog(f"merge exception: {e}")
                except Exception:
                    pass
                try:
                    messagebox.showerror("Merge failed", f"合并失败：\n{e}")
                except Exception:
                    print(f"Merge failed: {e}", file=sys.stderr)
                sys.exit(1)

            if conflicts:
                # Open merge UI in conflict-only mode (mine vs theirs), with merged preview as A.
                _show_conflict_popup(conflicts)

                app = SowMergeApp(
                    preview_path,
                    args.theirs,
                    merge_mode=True,
                    merged_path=args.merged,
                    merge_conflict_cells_by_sheet=conflict_map,
                    merge_conflict_mode=True,
                )
                app.run()
                sys.exit(1)

            # No conflicts
            if force_ui and (not preview_path):
                try:
                    preview_path = _ensure_xlsx_copy(args.mine)
                    _dlog(f"force_ui: created preview from mine: {preview_path}")
                except Exception:
                    pass
            if force_ui and preview_path:
                try:
                    messagebox.showinfo(
                        "Merge ok",
                        f"未检测到冲突，但将进入确认界面（可检查后保存）：\n{args.merged}",
                    )
                except Exception:
                    pass
                app = SowMergeApp(
                    preview_path,
                    args.theirs,
                    merge_mode=True,
                    merged_path=args.merged,
                    merge_conflict_cells_by_sheet={},
                    merge_conflict_mode=False,
                )
                try:
                    _dlog("open UI: force_ui no-conflict")
                except Exception:
                    pass
                app.run()
                sys.exit(0)
            try:
                messagebox.showinfo("Merge ok", f"合并完成，已保存：\n{args.merged}")
            except Exception:
                print(f"Merge ok: {args.merged}", file=sys.stderr)
            sys.exit(0)

        if args.textdiff:
            try:
                temp_root = os.path.join(os.environ.get("LOCALAPPDATA", tempfile.gettempdir()), "Temp", "TortoiseXlsTemp")
                os.makedirs(temp_root, exist_ok=True)
            except Exception:
                temp_root = tempfile.gettempdir()

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            left_txt = os.path.join(temp_root, f"{APP_NAME}_left_{ts}.txt")
            right_txt = os.path.join(temp_root, f"{APP_NAME}_right_{ts}.txt")
            excel_to_text(a, left_txt, thick_sep_char="=")
            excel_to_text(b, right_txt, thick_sep_char="=")
            open_tortoise_merge(left_txt, right_txt, title=f"{APP_NAME}: {os.path.basename(a)}")
            return

        app = SowMergeApp(a, b)
        app.run()

    except Exception:
        err = traceback.format_exc()
        try:
            messagebox.showerror("Error", err)
        except Exception:
            print(err, file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
