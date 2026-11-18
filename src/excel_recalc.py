# -*- coding: utf-8 -*-
"""Excel COM recalculation — refresh cached formula values in xlsx files."""

import os
import re
import sys
import time
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

XL_CALCULATION_AUTOMATIC = -4105

_SS_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_CELL_REF_RE = re.compile(r"([A-Z]+)([0-9]+)")


def _col_letters_to_index(letters):
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - 64)
    return idx


def _sheet_xml_targets(zf):
    """Map sheet name -> sheet xml path inside the xlsx archive."""
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    targets = {rel.get("Id"): rel.get("Target")
               for rel in rels.iter(f"{_REL_NS}Relationship")}
    out = {}
    for sheet in wb.iter(f"{_SS_NS}sheet"):
        target = targets.get(sheet.get(f"{_R_NS}id"))
        if not target:
            continue
        target = target.lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        out[sheet.get("name")] = target
    return out


def _configure_excel(excel):
    """Put a private Excel instance into a predictable automation state."""
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    excel.ScreenUpdating = False
    try:
        excel.Calculation = XL_CALCULATION_AUTOMATIC
    except Exception:
        pass
    try:
        excel.CalculateBeforeSave = True
    except Exception:
        pass


def _calculate_workbook(excel, workbook):
    """Force Excel to rebuild formula caches before saving."""
    try:
        workbook.RefreshAll()
    except Exception:
        pass
    try:
        workbook.UpdateLink()
    except Exception:
        pass
    try:
        excel.CalculateUntilAsyncQueriesDone()
    except Exception:
        pass
    try:
        excel.CalculateFullRebuild()
    except Exception:
        try:
            excel.CalculateFull()
        except Exception:
            pass
    try:
        workbook.Calculate()
    except Exception:
        pass
    try:
        excel.CalculateUntilAsyncQueriesDone()
    except Exception:
        pass


def com_available():
    """本机是否可用 Excel COM（win32com）。写回路径判定的单一来源。"""
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def write_cells_via_com(target_path, cells, timeout=120):
    """COM 直写静态值到目标 xlsx：UpdateLinks=0 打开 → 按行序写值 → Save → 关闭。

    cells: {sheet_name: [(row, col, value), ...]}，value 必须为静态值。
    公式字符串（str 以 '=' 开头）→ 抛 ValueError（公式写入会触发未受控求值）。
    win32com 不可用 → 抛 RuntimeError，由调用方决定降级。
    不调用 CalculateFull：仅依赖 Save 时的局部重算（打开时链接数据已加载，
    XLOOKUP 局部求值可用），外部引用缓存/公式/shared formula 原样保留。
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise RuntimeError("win32com not available, cannot write cells via COM") from e

    # 公式值提前拒绝（不触发任何求值，先于文件检查）
    for sheet_name, cell_list in cells.items():
        for row, col, value in cell_list:
            if isinstance(value, str) and value.startswith("="):
                raise ValueError(
                    f"COM 直写不支持公式写入，请改为静态值: {sheet_name} "
                    f"r{row}c{col}"
                )

    fullpath = str(Path(target_path).resolve())
    if not os.path.isfile(fullpath):
        raise FileNotFoundError(f"File not found: {fullpath}")

    excel = None
    workbook = None
    start = time.time()

    try:
        pythoncom.CoInitialize()
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
        except Exception as e:
            # 同 recalc_xlsx：禁止 Dispatch() 附着用户实例
            raise RuntimeError(
                f"Failed to start a private Excel instance: {e}"
            ) from e
        _configure_excel(excel)
        # 禁止 Save 时全表重算：CalculateBeforeSave=True 会让 Excel 在 Save
        # 前重算整个工作簿，把「空段/名称缺失」等存量数据问题成批暴露为
        # 新的 #N/A 缓存（实测 Effect #事件效果 2 → 235）。COM 直写只依赖
        # Save 对受影响依赖链的局部重算（打开时链接数据已加载，局部求值
        # 可用），存量 #N/A 反而会被顺带修复。
        try:
            excel.CalculateBeforeSave = False
        except Exception:
            pass

        workbook = excel.Workbooks.Open(fullpath, UpdateLinks=0)
        if workbook is None:
            workbook = excel.ActiveWorkbook
        if workbook is None:
            raise RuntimeError(f"Excel failed to open workbook: {fullpath}")

        for sheet_name, cell_list in cells.items():
            try:
                ws = workbook.Worksheets(sheet_name)
            except Exception as e:
                raise RuntimeError(f"Sheet not found: {sheet_name}") from e
            # 行序写入，避免同列依赖链错乱
            for row, col, value in sorted(cell_list, key=lambda c: c[0]):
                ws.Cells(row, col).Value = value

        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    elapsed = time.time() - start
    print(f"[com-write] {Path(target_path).name} done in {elapsed:.1f}s", file=sys.stderr)
    return target_path


def apply_write_plan_via_com(target_path, plan, timeout=120):
    """COM 应用 resolve 写计划：UpdateLinks=0 打开 → 删尾行 → 插行 → 写值 → Save。

    plan: {sheet_name: {
        "created": bool,                  # theirs-only sheet，需新建
        "headers": {(row, col): value},   # created 时的新表头（3 行）
        "inserts": [(out_row, {col: value})],  # 新增行（out_row 为最终行号）
        "updates": [(row, col, value)],   # 既有行修改单元格
        "deletes": [row, ...],            # 中间删除的 mine 行号（降序执行）
        "delete_from": int | None,        # 尾部删除起始行（含）
    }}
    公式值**受控写入**（resolve 场景决策行公式必须保留，proposal 验证
    「公式保持公式」）：str 以 '=' 开头 → .Formula；ArrayFormula 对象
    （有 .text）→ .FormulaArray（保留 CSE 数组语义）。UpdateLinks=0 打开
    时链接数据已加载，Excel 写公式后局部求值可用，不会成批写 #N/A。
    静态值 → .Value。
    执行顺序：中间删行（降序）→ 尾部删除 → 行插入（按最终行号升序，上方
    插入会整体下推下方内容，先插上方行号才稳定；Excel 自动调整公式引用）
    → 写值（按行序）。
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise RuntimeError("win32com not available, cannot apply write plan via COM") from e

    fullpath = str(Path(target_path).resolve())
    if not os.path.isfile(fullpath):
        raise FileNotFoundError(f"File not found: {fullpath}")

    excel = None
    workbook = None
    start = time.time()

    try:
        pythoncom.CoInitialize()
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
        except Exception as e:
            # 同 recalc_xlsx：禁止 Dispatch() 附着用户实例
            raise RuntimeError(
                f"Failed to start a private Excel instance: {e}"
            ) from e
        _configure_excel(excel)
        # 同 write_cells_via_com：禁止 Save 全表重算，只依赖局部重算
        try:
            excel.CalculateBeforeSave = False
        except Exception:
            pass

        workbook = excel.Workbooks.Open(fullpath, UpdateLinks=0)
        if workbook is None:
            workbook = excel.ActiveWorkbook
        if workbook is None:
            raise RuntimeError(f"Excel failed to open workbook: {fullpath}")

        for sn, sp in plan.items():
            try:
                ws = workbook.Worksheets(sn)
            except Exception:
                ws = None
            if ws is None:
                if not sp.get("created"):
                    raise RuntimeError(f"Sheet not found: {sn}")
                last_ws = workbook.Worksheets(workbook.Worksheets.Count)
                ws = workbook.Worksheets.Add(After=last_ws)
                ws.Name = sn

            # 1) 中间删行（降序，先删行号大的不影响小行号）
            for del_row in sorted(sp.get("deletes", []), reverse=True):
                ws.Rows(del_row).Delete()

            # 2) 尾部删除（自上而下，先删行号才稳定）
            delete_from = sp.get("delete_from")
            if delete_from:
                try:
                    last_row = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
                except Exception:
                    last_row = delete_from
                if last_row >= delete_from:
                    ws.Rows(f"{delete_from}:{last_row}").Delete()

            # 3) 新行插入（按最终行号升序：上方插入整体下推下方内容，
            #    先插行号小的，后插的不会影响已就位的上方行）
            for out_row in sorted(r for r, _ in sp.get("inserts", [])):
                ws.Rows(out_row).Insert()

            # 3) 写值（表头 + 更新 + 新行），按行序避免依赖链错乱。
            #    公式受控写入：普通公式 → .Formula，数组公式（ArrayFormula
            #    对象）→ .FormulaArray 保留 CSE 语义
            writes = {}
            for (hrow, hcol), hval in (sp.get("headers") or {}).items():
                writes.setdefault(hrow, {})[hcol] = hval
            for row, col, value in sp.get("updates", []):
                writes.setdefault(row, {})[col] = value
            for out_row, cells in sp.get("inserts", []):
                writes.setdefault(out_row, {}).update(cells)
            for row in sorted(writes):
                for col in sorted(writes[row]):
                    value = writes[row][col]
                    if hasattr(value, "text"):  # ArrayFormula → 数组公式
                        ws.Cells(row, col).FormulaArray = value.text
                    elif isinstance(value, str) and value.startswith("="):
                        ws.Cells(row, col).Formula = value
                    else:
                        ws.Cells(row, col).Value = value

        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    elapsed = time.time() - start
    print(f"[com-write] {Path(target_path).name} done in {elapsed:.1f}s", file=sys.stderr)
    return target_path


def count_formula_error_caches(filepath, data_start_row=4):
    """按 # sheet 统计缓存值为错误值（#N/A、#REF! 等）的公式单元格数。

    返回 {sheet_name: count}（仅含 >0 的 sheet）。XML 层解析（同
    count_missing_formula_caches），错误值特征：cell t="e"，或 <v> 文本以
    '#' 开头（#N/A / #REF! / #VALUE! 等）。只统计数据行的公式单元格。
    """
    fullpath = str(Path(filepath).resolve())
    if not os.path.isfile(fullpath):
        raise FileNotFoundError(f"File not found: {fullpath}")

    counts = {}
    with zipfile.ZipFile(fullpath) as zf:
        for sheet_name, xml_path in _sheet_xml_targets(zf).items():
            if not sheet_name.startswith("#"):
                continue
            n = 0
            with zf.open(xml_path) as stream:
                for _, cell in ET.iterparse(stream, events=("end",)):
                    if cell.tag != f"{_SS_NS}c":
                        continue
                    ref = _CELL_REF_RE.fullmatch(cell.get("r") or "")
                    if ref and int(ref.group(2)) >= data_start_row:
                        if cell.find(f"{_SS_NS}f") is not None:
                            v = cell.find(f"{_SS_NS}v")
                            if v is not None and (
                                cell.get("t") == "e"
                                or (v.text or "").startswith("#")
                            ):
                                n += 1
                    cell.clear()
            if n:
                counts[sheet_name] = n
    return counts


def verify_formula_cache_guard(filepath, baseline, data_start_row=4):
    """写回后错误值基线校验：与写回前 baseline 对比。

    baseline: 写回前 count_formula_error_caches() 的结果（{sheet: count}）。
    返回 {ok, baseline, after, added_sheets, total}：
      ok=False 当且仅当新增了错误值；基线已存在的错误值视为表格数据问题，
      不算写回失败（由策划修正数据）。
    """
    after = count_formula_error_caches(filepath, data_start_row=data_start_row)
    added_sheets = {
        sheet: count - baseline.get(sheet, 0)
        for sheet, count in after.items()
        if count > baseline.get(sheet, 0)
    }
    return {
        "ok": not added_sheets,
        "baseline": dict(baseline or {}),
        "after": after,
        "added_sheets": added_sheets,
        "total": sum(added_sheets.values()),
    }


def _has_external_links(filepath):
    """zip 层检测文件是否含外部链接部件（xl/externalLinks/）。"""
    try:
        with zipfile.ZipFile(filepath) as zf:
            return any(
                n.startswith("xl/externalLinks/") or n.startswith("externalLinks/")
                for n in zf.namelist()
            )
    except Exception:
        return False


def count_missing_formula_caches(filepath, important_columns=(1,), data_start_row=4):
    """Count formula cells whose cached values are missing.

    By default this checks only data-row column A because RowName formulas are
    identity keys during merge and diff. Other formulas may legitimately return
    blank, but a blank RowName formula cache breaks key matching.

    This parses the sheet XML directly instead of openpyxl data_only reads:
    Excel caches an empty-string formula result as ``t="str"`` + ``<v/>``,
    which openpyxl reports as ``None`` — indistinguishable from a truly
    missing cache.  A formula cell counts as missing only when it has no
    usable cached value: no ``<v>`` element, or an empty ``<v>`` without the
    ``t="str"`` marker that Excel puts on empty-string results (openpyxl
    writes cache-less formulas as ``<f>..</f><v></v>`` with no cell type).
    """
    fullpath = str(Path(filepath).resolve())
    if not os.path.isfile(fullpath):
        raise FileNotFoundError(f"File not found: {fullpath}")

    columns = set(important_columns or ())
    missing = 0
    with zipfile.ZipFile(fullpath) as zf:
        for sheet_name, xml_path in _sheet_xml_targets(zf).items():
            if not sheet_name.startswith("#"):
                continue
            with zf.open(xml_path) as stream:
                for _, cell in ET.iterparse(stream, events=("end",)):
                    if cell.tag != f"{_SS_NS}c":
                        continue
                    ref = _CELL_REF_RE.fullmatch(cell.get("r") or "")
                    if ref:
                        col_idx = _col_letters_to_index(ref.group(1))
                        row_idx = int(ref.group(2))
                        col_checked = col_idx in columns if columns else True
                        if (col_checked and row_idx >= data_start_row
                                and cell.find(f"{_SS_NS}f") is not None):
                            v = cell.find(f"{_SS_NS}v")
                            cache_present = (
                                v is not None
                                and ((v.text or "") != "" or cell.get("t") == "str")
                            )
                            if not cache_present:
                                missing += 1
                    cell.clear()
    return missing


def ensure_formula_cache(filepath, timeout=120, raise_on_error=False,
                         important_columns=(1,), data_start_row=4,
                         retries=2):
    """Recalculate a workbook and verify important formula caches exist.

    Excel occasionally saves before its background/multi-threaded calculation
    has committed every cache, leaving a handful of formula cells without
    cached values.  That is transient, so recalc again instead of failing the
    whole writeback on the first pass.

    只用于无外部链接的表：文件含外部链接时拒绝全表重算（CalculateFull 会
    重算所有外部引用并成批写回 #N/A 缓存），写回请走 COM 直写。
    """
    fullpath = str(Path(filepath).resolve())
    if _has_external_links(fullpath):
        msg = (
            "文件含外部链接，全表重算会重算外部引用并成批写回 #N/A 缓存；"
            "写回请走 COM 直写（write_cells_via_com），不要手动全表重算"
        )
        if raise_on_error:
            raise RuntimeError(msg)
        print(f"[recalc] {Path(fullpath).name}: {msg}", file=sys.stderr)
        return {"ok": False, "error": msg}

    attempts = max(1, int(retries) + 1)
    missing = None
    for attempt in range(attempts):
        recalc_xlsx(filepath, timeout=timeout, raise_on_error=True)
        missing = count_missing_formula_caches(
            filepath,
            important_columns=important_columns,
            data_start_row=data_start_row,
        )
        if not missing:
            break
        if attempt + 1 < attempts:
            print(
                f"[recalc] {missing} important formula cache(s) missing after "
                f"pass {attempt + 1}, recalculating again",
                file=sys.stderr,
            )
    result = {
        "ok": missing == 0,
        "missing_important_caches": missing,
    }
    if missing and raise_on_error:
        raise RuntimeError(
            f"Excel recalculated the workbook, but {missing} important formula "
            "cache(s) are still missing"
        )
    return result


def recalc_xlsx(filepath, timeout=60, raise_on_error=False):
    """Open xlsx in Excel via COM, recalculate all formulas with
    external links resolved, save and close.

    Uses CalculateFull to force complete recalculation of all open
    workbooks, ensuring cross-sheet and cross-file XLOOKUP references
    resolve correctly.  Per-sheet ws.Calculate() is NOT sufficient —
    it can leave cross-sheet references using stale cached values.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        if raise_on_error:
            raise RuntimeError("win32com not available, cannot refresh formula cache") from e
        print("[recalc] win32com not available, skipping recalc", file=sys.stderr)
        return filepath

    fullpath = str(Path(filepath).resolve())
    if not os.path.isfile(fullpath):
        raise FileNotFoundError(f"File not found: {fullpath}")

    excel = None
    workbook = None
    opened_links = []
    start = time.time()

    try:
        pythoncom.CoInitialize()
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
        except Exception as e:
            # Do NOT fall back to Dispatch() here — it may attach to the
            # user's running Excel instance, which we would then hide,
            # reconfigure and Quit(), destroying unsaved work.
            if raise_on_error:
                raise
            print(f"[recalc] Failed to start a private Excel instance: {e}", file=sys.stderr)
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            return filepath
        _configure_excel(excel)

        workbook = excel.Workbooks.Open(fullpath, UpdateLinks=3)
        if workbook is None:
            workbook = excel.ActiveWorkbook
        if workbook is None:
            if raise_on_error:
                raise RuntimeError(f"Excel failed to open workbook: {fullpath}")
            print("[recalc] Failed to open workbook", file=sys.stderr)
            return filepath

        # Open all linked workbooks so XLOOKUP references resolve
        try:
            link_sources = list(workbook.LinkSources(1))  # 1 = xlExcelLinks
        except Exception:
            link_sources = []
        for link_path in link_sources:
            resolved = _resolve_link_path(link_path, os.path.dirname(fullpath))
            if resolved:
                try:
                    lw = excel.Workbooks.Open(resolved, UpdateLinks=0, ReadOnly=True)
                    opened_links.append(lw)
                except Exception:
                    pass

        _calculate_workbook(excel, workbook)

        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None

        # Close linked workbooks without saving
        for lw in opened_links:
            try:
                lw.Close(SaveChanges=False)
            except Exception:
                pass
        opened_links = []

    except Exception as e:
        if raise_on_error:
            raise
        print(f"[recalc] Excel recalc failed: {e}", file=sys.stderr)
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        for lw in opened_links:
            try:
                lw.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    elapsed = time.time() - start
    print(f"[recalc] {Path(filepath).name} done in {elapsed:.1f}s", file=sys.stderr)
    return filepath


def _resolve_link_path(link_path, base_dir):
    """Resolve an external link path to an existing absolute path."""
    if not link_path:
        return None
    if os.path.isfile(link_path):
        return link_path
    # Try relative to base dir
    alt = os.path.join(base_dir, os.path.basename(link_path))
    if os.path.isfile(alt):
        return alt
    # Try relative to Raw dir
    raw = os.path.dirname(base_dir)
    alt2 = os.path.join(raw, os.path.basename(link_path))
    if os.path.isfile(alt2):
        return alt2
    return None


def remove_with_retry(path, attempts=5, delay=0.5):
    """Delete a file, retrying while Excel/AV still holds a handle on it.

    EXCEL.EXE often lingers briefly after Quit() and keeps the just-recalc'd
    copy locked; a single os.remove() then fails with PermissionError and the
    temp file leaks into the working directory. Retrying rides out the lock.
    Returns True when the file is gone.
    """
    for i in range(attempts):
        try:
            os.remove(path)
            return True
        except FileNotFoundError:
            return True
        except PermissionError:
            if i < attempts - 1:
                time.sleep(delay)
        except Exception as e:
            print(f"[recalc] cannot remove temp file {path}: {e}", file=sys.stderr)
            return False
    print(f"[recalc] cannot remove temp file {path}: still locked after "
          f"{attempts} attempts", file=sys.stderr)
    return False


def schedule_deletion(path, attempts=60, delay=5):
    """Best-effort deferred delete for a temp file that is still locked.

    remove_with_retry() gives up when a lingering EXCEL.EXE (or AV scan)
    keeps the file open past the retry window; without a follow-up the
    recalc copy leaks into the working-copy directory. This spawns a
    detached PowerShell that retries Remove-Item until the lock clears
    (or attempts run out), so the copy disappears shortly after Excel
    exits. Windows-only; no-op elsewhere.
    """
    if os.name != "nt":
        return
    import subprocess
    ps_path = str(Path(path)).replace("'", "''")
    command = (
        f"for ($i = 0; $i -lt {int(attempts)}; $i++) {{"
        f" if (-not (Test-Path -LiteralPath '{ps_path}')) {{ exit 0 }};"
        f" Remove-Item -LiteralPath '{ps_path}' -Force -ErrorAction SilentlyContinue;"
        f" Start-Sleep -Seconds {int(delay)}"
        f" }}"
    )
    try:
        # CREATE_NO_WINDOW: no console window and survives our process exit.
        # (DETACHED_PROCESS would be cleaner but powershell.exe fails to
        # start without any console attached.)
        subprocess.Popen(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
            creationflags=subprocess.CREATE_NO_WINDOW | subprocess.CREATE_NEW_PROCESS_GROUP,
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            close_fds=True,
        )
    except Exception as e:
        print(f"[recalc] cannot schedule deletion of {path}: {e}", file=sys.stderr)


def _file_ctime(path):
    """Creation time of a file (tests monkeypatch this to fake file age)."""
    return path.stat().st_ctime


def cleanup_stale_recalc_copies(directory, min_age_seconds=600):
    """Remove leftover `.*_recalc*.xlsx` temp copies from previous runs.

    Only this tool creates files matching that pattern (leading dot, _recalc
    infix — with or without the uuid suffix). Files younger than
    min_age_seconds or still locked are skipped — a concurrent diff may be
    using them.

    Age is measured by CREATION time: copy2() preserves the source file's
    mtime, so a fresh copy of a long-unmodified workbook would look stale
    and could be swept from under a concurrent run.
    """
    try:
        entries = list(Path(directory).glob(".*_recalc*.xlsx"))
    except Exception:
        return
    now = time.time()
    for entry in entries:
        try:
            if now - _file_ctime(entry) < min_age_seconds:
                continue
            os.remove(str(entry))
        except Exception:
            pass


def recalc_to_temp(filepath, temp_dir, raise_on_error=False):
    """Copy file to a temp name in the SAME directory, recalc the copy,
    return copy path. Original file is NOT modified.
    We stay in the source directory so external link references remain valid."""
    import uuid

    src = Path(filepath)
    # Sweep leftovers from previous crashed/failed runs before adding a new one.
    cleanup_stale_recalc_copies(str(src.parent))
    # Keep the copy beside the source so relative external links still resolve.
    # Add a suffix to avoid collisions between consecutive diff operations.
    dst = src.parent / f".{src.stem}_recalc_{uuid.uuid4().hex[:8]}{src.suffix}"
    shutil.copy2(str(src), str(dst))
    try:
        recalc_xlsx(str(dst), raise_on_error=raise_on_error)
    except Exception:
        if raise_on_error:
            raise
        pass
    return str(dst)


def recalc_batch(filepaths, timeout=120):
    """Recalculate multiple xlsx files in ONE Excel session.
    Opens all files + linked workbooks, recalculates, saves, closes.
    Much faster than calling recalc_xlsx() per file — saves COM startup cost.

    Returns list of successfully recalc'd file paths.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print("[recalc] win32com not available, skipping batch recalc", file=sys.stderr)
        return filepaths

    valid = []
    for fp in filepaths:
        full = str(Path(fp).resolve())
        if os.path.isfile(full):
            valid.append(full)

    if not valid:
        return valid

    excel = None
    workbooks = []
    opened_links = []
    start = time.time()

    try:
        pythoncom.CoInitialize()
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
        except Exception as e:
            # No Dispatch() fallback — attaching to the user's running Excel
            # instance would let us hide/reconfigure/Quit() it (data loss).
            print(f"[recalc] Failed to start a private Excel instance: {e}", file=sys.stderr)
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            return valid
        _configure_excel(excel)

        # Collect all unique link sources across all files
        all_link_sources = set()
        for fp in valid:
            wb = excel.Workbooks.Open(fp, UpdateLinks=0)
            workbooks.append(wb)
            try:
                for link in (wb.LinkSources(1) or []):
                    all_link_sources.add(link)
            except Exception:
                pass

        # Open linked workbooks
        for link_path in all_link_sources:
            resolved = _resolve_link_path(link_path, os.path.dirname(valid[0]))
            if resolved:
                try:
                    lw = excel.Workbooks.Open(resolved, UpdateLinks=0, ReadOnly=True)
                    opened_links.append(lw)
                except Exception:
                    pass

        for wb in workbooks:
            _calculate_workbook(excel, wb)

        # Save and close all target workbooks
        for wb in workbooks:
            try:
                wb.Save()
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        workbooks = []

        # Close linked workbooks
        for lw in opened_links:
            try:
                lw.Close(SaveChanges=False)
            except Exception:
                pass
        opened_links = []

    except Exception as e:
        print(f"[recalc] batch recalc failed: {e}", file=sys.stderr)
    finally:
        for wb in workbooks:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        for lw in opened_links:
            try:
                lw.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    elapsed = time.time() - start
    print(f"[recalc] batch {len(valid)} files done in {elapsed:.1f}s", file=sys.stderr)
    return valid
