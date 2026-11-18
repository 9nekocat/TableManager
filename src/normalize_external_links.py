# -*- coding: utf-8 -*-
"""外部链接 Target 规范化 + 轮询重算工具（最小实施子集）。

背景：externalLink rels 中的 Target 若为 `file:///A:\\...` 反斜杠 file URL，
Excel 的链接更新机制无法解析（交互式弹「更新值」框、COM 自动化静默失败），
导致 CalculateFullRebuild 后外部引用公式缓存被重算成 #N/A。
规范化 Target（file:///A:\\ → A:\\ 无前缀）后链接可自动更新，
且链接加载为异步（0~2 秒），需轮询等待就绪再重算。

用法：
  python normalize_external_links.py <file.xlsx>            # 规范化 + 轮询重算 + 保存 + 校验
  python normalize_external_links.py <file.xlsx> --no-recalc  # 仅规范化 rels Target（不重算）

退出码：0 = 成功（#N/A 未异常增加）；2 = #N/A 异常增加或轮询超时。
"""
import argparse
import os
import re
import shutil
import sys
import time
import zipfile

# ---------- zip 层 rels 规范化 ----------

def normalize_targets(filepath):
    """把 externalLinkN.xml.rels 的 Target 规范化为 A:\\ 无前缀（或相对路径）。

    仅改写 rels 字符串，其余 zip 条目字节不变。幂等。
    返回变更的 rels 文件数。
    """
    tmp = filepath + '.tmp'
    changed = 0
    zin = zipfile.ZipFile(filepath)
    try:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename.endswith('.rels') and '_rels/externalLink' in item.filename:
                    text = zin.read(item.filename).decode('utf-8')
                    new_text = text
                    new_text = new_text.replace('file:///A:\\', 'A:\\')
                    new_text = new_text.replace('file:///D:\\catsintheshell\\', 'A:\\')
                    new_text = re.sub(r'Target="/DataTable/Raw/', 'Target="', new_text)
                    if new_text != text:
                        changed += 1
                        text = new_text
                    zout.writestr(item, text.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item.filename))
    finally:
        zin.close()
    if changed:
        os.replace(tmp, filepath)
    else:
        os.remove(tmp)
    return changed

# ---------- #N/A / #REF 统计 ----------

def count_errors(filepath):
    """统计全表缓存中的 #N/A 与 #REF 错误值数量（逐 sheet）。"""
    import openpyxl
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    by_sheet = {}
    total = 0
    try:
        for ws in wb.worksheets:
            n = 0
            for row in ws.iter_rows(min_row=4, values_only=True):
                for v in row:
                    if v is not None and ('#N/A' in str(v) or '#REF' in str(v)):
                        n += 1
                        break
            by_sheet[ws.title] = n
            total += n
    finally:
        wb.close()
    return total, by_sheet

# ---------- 轮询重算 ----------

def _build_probe_formula(zf):
    """构造探测公式：引用第一个外部链接（[1]）的第一个 sheet 的 A1。"""
    try:
        rels = zf.read('xl/externalLinks/_rels/externalLink1.xml.rels').decode('utf-8')
        # 取第一个 Target 的文件名（basename）用于报错信息（实际用 [1] 编号）
        m = re.search(r'Target="([^"]+)"', rels)
        sheet_xml = zf.read('xl/externalLinks/externalLink1.xml').decode('utf-8')
        sheets = re.findall(r'<sheetName val="([^"]+)"', sheet_xml)
        if not sheets:
            return None
        return "='[1]#%s'!A1" % sheets[0]
    except KeyError:
        return None


def recalc_with_poll(filepath, timeout=24, interval=2):
    """COM UpdateLinks=3 打开 → 轮询重算直到外部链接就绪 → 保存 → 关闭。

    返回 (ready, tries, probe_value)。ready=False 表示轮询超时（链接未就绪）。
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print('[normalize] 需要 pywin32（Excel COM 自动化）')
        return False, 0, None

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    probe_value = None
    ready = False
    tries = 0
    try:
        wb = excel.Workbooks.Open(filepath, UpdateLinks=3)
        # 新建临时 sheet 放探测公式（不污染原表）
        probe_ws = wb.Worksheets.Add()
        formula = None
        with zipfile.ZipFile(filepath) as zf:
            formula = _build_probe_formula(zf)
        if not formula:
            ready = True  # 无外部链接，直接就绪
            tries = 0
        else:
            probe_ws.Range('A1').Formula = formula
            max_tries = max(1, int(timeout / interval))
            for i in range(max_tries):
                time.sleep(interval)
                try:
                    excel.CalculateFullRebuild()
                except Exception:
                    pass
                v = probe_ws.Range('A1').Value
                probe_value = v
                tries = i + 1
                # -2146826265 = #REF! ; -2146826246 = #N/A （链接未加载时的表现）
                if v not in (-2146826265, -2146826246):
                    ready = True
                    break
        try:
            probe_ws.Delete()
        except Exception:
            pass
        wb.Save()
        wb.Close(False)
    finally:
        excel.Quit()
    return ready, tries, probe_value


# ---------- 主流程 ----------

def main():
    parser = argparse.ArgumentParser(description='外部链接 Target 规范化 + 轮询重算（修复 #N/A）')
    parser.add_argument('file', help='目标 xlsx 路径')
    parser.add_argument('--no-recalc', action='store_true', help='仅规范化 rels Target，不重算')
    parser.add_argument('--timeout', type=int, default=24, help='轮询超时秒数（默认 24）')
    parser.add_argument('--no-backup', action='store_true', help='不生成 .bak 备份')
    args = parser.parse_args()

    path = os.path.abspath(args.file)
    if not os.path.isfile(path):
        print('[normalize] 文件不存在: %s' % path)
        return 2

    if not args.no_backup:
        bak = path + '.bak'
        shutil.copy(path, bak)
        print('[normalize] 备份: %s' % bak)

    before, before_by = count_errors(path)
    print('[normalize] 处理前 #N/A/#REF: %d %s' % (before, before_by))

    changed = normalize_targets(path)
    print('[normalize] 规范化 rels Target: %d 个文件变更' % changed)
    if changed == 0:
        print('[normalize] 无 file:/// 变体，无需规范化')

    if not args.no_recalc:
        print('[normalize] 轮询重算（等待外部链接异步加载）...')
        ready, tries, probe = recalc_with_poll(path, timeout=args.timeout)
        if ready:
            print('[normalize] 链接就绪（第 %d 次轮询，探测值=%s）' % (tries, probe))
        else:
            print('[normalize] !!! 轮询超时，链接未就绪（探测值=%s）。' % probe)
            print('[normalize]     请用交互式 Excel 打开 → 弹窗点「更新」→ 保存后重试。')
            return 2

    after, after_by = count_errors(path)
    delta = after - before
    print('[normalize] 处理后 #N/A/#REF: %d %s (变化 %+d)' % (after, after_by, delta))
    if delta > 0:
        print('[normalize] !!! #N/A/#REF 增加 %d 处：多为表格数据问题（名称列空段/名称缺失），' % delta)
        print('[normalize]     请检查新增 sheet 的对应行（可用 TableManager diff 定位）。')
        return 2
    print('[normalize] OK：缓存错误值未增加。')
    return 0


if __name__ == '__main__':
    sys.exit(main())
