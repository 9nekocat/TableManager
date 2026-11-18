# -*- coding: utf-8 -*-
"""
Effect.xlsx 重算脚本 — 修复 #事件效果/#Main 重算产生的 #N/A

背景（2026-08-12 调研结论）：
  Effect.xlsx 的 Params（H 列）与 ConditionIDs（F 列）是 TEXTJOIN+XLOOKUP 公式，
  XLOOKUP 的查找值来自名称列（事件名称/目标名称/效果名称1~6）。
  原公式保护写法为 IF(X="无",0,XLOOKUP(...))——只覆盖"无"，不覆盖"空单元格"。
  Excel 中空单元格参与公式时视为 0，XLOOKUP(0, C:C, B:B) 在查找数组中找不到 0
  → 返回 #N/A → TEXTJOIN 整体 #N/A。
  由于 calcPr calcMode=manual + calcCompleted=0（缓存未完成全算），打开文件时
  显示的是旧缓存（部分有效），一旦全量重算（F9 / CalculateFull）即批量暴露 #N/A。

修复：
  1. 公式读回写（Excel COM）：把 IF(X="无",0, → IF(OR(X="无",X=""),0,（覆盖空单元格）
  2. CalculateFull 全量重算（私有 Excel 实例，不更新外部链接，用内嵌缓存）
  3. 保存（Excel 原生保存维护公式缓存/外部链接）
  4. 校验：统计残留 #N/A（应只剩数据缺失类，如名称在目标表不存在）

用法：
  python effect_recalc.py                # 重算 A:/DataTable/Raw/Effect.xlsx
  python effect_recalc.py --file <path>  # 重算指定文件
  python effect_recalc.py --dry-run      # 只诊断（备份+统计+报告），不改文件
  python effect_recalc.py --no-backup    # 跳过备份（不推荐）

输出：控制台摘要 + 报告文件（默认 <文件目录>/../effect_recalc_report.txt，Raw 外）。
"""
import argparse
import datetime
import os
import re
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import pythoncom
    import win32com.client
except ImportError:
    print("[错误] 需要 pywin32：pip install pywin32")
    sys.exit(1)

import excel_recalc  # Tools/TableManager/src/excel_recalc.py（同目录）

DEFAULT_FILE = r"A:/DataTable/Raw/Effect.xlsx"
BACKUP_DIR = r"A:/DataTable/Raw/../effect_recalc_backups"  # Raw 目录之外
SHEETS = ["#Main", "#事件效果", "#BOSS"]
FORMULA_COLS = ["F", "H"]  # ConditionIDs / Params

# 空值保护模板（两处替换共用）：
#   IF(X="无",0, → IF(OR(X="无",X=""),IF(X="无",0,""),   （"无"→0 保持原始语义，空→"" 被 TEXTJOIN 忽略，不产生 ,0,0,0）
#   未保护的 XLOOKUP(X, → IF(OR(X="无",X=""),IF(X="无",0,""),XLOOKUP(X,
OR_PATTERN = re.compile(r'IF\(([A-Z]+\d+)="无",0,')

# 兼容 COM 层（XLOOKUP）与 XML 层（_xlfn.XLOOKUP）的函数名前缀
XLOOKUP_CELL_PATTERN = re.compile(r'(?:_xlfn\.)?XLOOKUP\(([A-Z]+\d+),')

PROTECT_PREFIX = 'IF(OR(%s="无",%s=""),IF(%s="无",0,""),'


def _is_protected(prefix_text, ref):
    return PROTECT_PREFIX % (ref, ref, ref) in prefix_text


def protect_formula(f):
    """为公式加空值保护（幂等，可重复执行）：
    1. IF(X="无",0, → IF(OR(X="无",X=""),IF(X="无",0,""),（"无"→0 保持原始语义；空→"" 被 TEXTJOIN 忽略）
    2. 未被保护的直接 XLOOKUP(单元格, → 同样包裹（自动补右括号，兼容 _xlfn. 前缀）
    """
    f2 = OR_PATTERN.sub(
        lambda m: PROTECT_PREFIX % (m.group(1), m.group(1), m.group(1)), f)
    out = []
    pos = 0
    i = 0
    while i < len(f2):
        m = XLOOKUP_CELL_PATTERN.search(f2, i)
        if not m:
            break
        ref = m.group(1)
        # 检查该 XLOOKUP 前 80 字符内是否已有同单元格的 OR 保护（避免重复包裹）
        prefix = f2[max(0, m.start() - 80):m.start()]
        if _is_protected(prefix, ref):
            i = m.end()
            continue
        # 找到 XLOOKUP( 的匹配右括号（从函数名后开始扫描）
        depth = 0
        j = m.end()
        while j < len(f2):
            ch = f2[j]
            if ch == "(":
                depth += 1
            elif ch == ")":
                if depth == 0:
                    break
                depth -= 1
            j += 1
        out.append(f2[pos:m.start()])
        out.append(PROTECT_PREFIX % (ref, ref, ref))
        out.append(f2[m.start():j + 1])
        out.append(")")
        pos = j + 1
        i = j + 1
    out.append(f2[pos:])
    return ''.join(out)


def is_open_in_excel(path):
    """检查文件是否正被 Excel 打开（存在 ~$ 锁文件即视为占用）"""
    d = os.path.dirname(path)
    name = os.path.basename(path)
    lock = os.path.join(d, "~$" + name[1:] if name.startswith("~") else "~$" + name)
    return os.path.exists(lock)


def backup(path):
    """OS 级字节复制备份到 Raw 目录之外"""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(BACKUP_DIR, "Effect_%s.xlsx" % ts)
    shutil.copy2(path, dst)
    return dst


def _fix_formulas_in_xml(xml_text):
    """在 sheet XML 文本中处理全部公式：加空值保护（只改含 XLOOKUP 的公式）。
    返回 (新文本, 修改处数)。"""
    count = [0]

    def repl(m):
        f = m.group(1)
        if 'XLOOKUP' not in f:
            return m.group(0)
        f2 = protect_formula(f)
        if f2 != f:
            count[0] += 1
        # XML 转义：& < >（公式里出现的 & 在 XML 中是 &amp;，保护插入内容不含这些字符）
        return '<f>' + f2 + '</f>'

    # 排除自闭合共享公式 <f t="shared" si="N"/>（其内容在从公式处不重复存储）
    new = re.sub(r'<f(?![^>]*/>)[^>]*>(.*?)</f>', repl, xml_text, flags=re.S)
    return new, count[0]


def recalc(path):
    """修复重算（zip 层 + COM）：
    1. zip 解包 → 仅修改 sheet XML 中的公式文本（空值保护），非公式单元格零接触
    2. 重打包 → COM 打开（UpdateLinks=0）→ CalculateFull
    3. 内存 #N/A 验证（超过阈值不保存，防止写坏）→ Excel 原生保存 → 校验
    """
    import zipfile
    import shutil as _shutil

    path = os.path.abspath(path)
    tmp = os.path.join(os.path.dirname(path), '.effect_recalc_tmp')
    if os.path.exists(tmp):
        _shutil.rmtree(tmp)
    os.makedirs(tmp)
    with zipfile.ZipFile(path) as zf:
        zf.extractall(tmp)

    # 处理全部 sheet XML
    total_changed = 0
    sheets_dir = os.path.join(tmp, 'xl/worksheets')
    for name in sorted(os.listdir(sheets_dir)):
        if not name.endswith('.xml'):
            continue
        p = os.path.join(sheets_dir, name)
        with open(p, encoding='utf-8') as f:
            data = f.read()
        new_data, n = _fix_formulas_in_xml(data)
        if new_data != data:
            with open(p, 'w', encoding='utf-8') as f:
                f.write(new_data)
        total_changed += n

    # 重打包（zip 条目路径必须用正斜杠）
    out = path
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root_d, dirs, files in os.walk(tmp):
            for fn in files:
                full = os.path.join(root_d, fn)
                arc = os.path.relpath(full, tmp).replace('\\', '/')
                zf.write(full, arc)
    _shutil.rmtree(tmp)
    print("  zip 层公式保护修改: %d 处" % total_changed)

    # COM 打开 → CalculateFull → 内存 #N/A 验证 → 保存
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        try:
            excel.Calculation = -4135  # xlCalculationAutomatic
        except Exception:
            pass
        wb = excel.Workbooks.Open(path, UpdateLinks=0)
        excel.CalculateFull()
        # 内存 #N/A 验证（防止把异常结果写盘）
        mem_na = 0
        for sn in SHEETS:
            ws = wb.Worksheets(sn)
            vals = list(ws.UsedRange.Value)
            mem_na += sum(1 for row in vals for v in row
                          if isinstance(v, int) and v == -2146826246)
        print("  内存 #N/A: %d" % mem_na)
        if mem_na > 50:
            wb.Close(SaveChanges=False)
            return {"ok": False,
                    "error": "重算后内存 #N/A=%d（超过 50），未保存，请检查数据" % mem_na,
                    "or_protections": total_changed}
        wb.Save()
        wb.Close(SaveChanges=False)
        wb = None
        print("  已保存（内存 #N/A=%d）" % mem_na)
        return {"ok": True, "or_protections": total_changed}
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
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


def extract_error_details(path, data_start_row=4):
    """提取文件全部错误缓存明细（sheet、单元格坐标、错误值）。通用：任意 xlsx"""
    import xml.etree.ElementTree as ET
    details = []
    with zipfile.ZipFile(path) as zf:
        for sheet_name, xml_path in excel_recalc._sheet_xml_targets(zf).items():
            if not sheet_name.startswith("#"):
                continue
            ns = excel_recalc._SS_NS
            with zf.open(xml_path) as stream:
                for _, cell in ET.iterparse(stream, events=("end",)):
                    if cell.tag != ns + "c":
                        cell.clear()
                        continue
                    ref = cell.get("r") or ""
                    m = re.match(r"^([A-Z]+)(\d+)$", ref)
                    if m and int(m.group(2)) >= data_start_row:
                        v = cell.find(ns + "v")
                        if v is not None and (cell.get("t") == "e" or (v.text or "").startswith("#")):
                            if cell.find(ns + "f") is not None:
                                err = (v.text or "").strip() or "#N/A"
                                details.append({
                                    "sheet": sheet_name,
                                    "cell": ref,
                                    "error": err[:30],
                                })
                    cell.clear()
    return details


def report_residual(path):
    """校验并列出残留 #N/A 明细（sheet、行、错误值），写入报告文件"""
    counts = excel_recalc.count_formula_error_caches(path)
    lines = []
    lines.append("重算后错误缓存统计: %s" % (counts if counts else "无（零 #N/A）"))
    if counts:
        details = extract_error_details(path)
        lines.append("残留 #N/A 明细（%d 个）：" % len(details))
        for d in details[:200]:
            lines.append("  %s!%s = %s" % (d["sheet"], d["cell"], d["error"]))
        lines.append("")
        lines.append("注意：残留 #N/A 均为「数据缺失」类——XLOOKUP 查找值在目标表中不存在：")
        lines.append("  - #Main H 列：J 列值（Buff 名）在 Localization/Buff.xlsx D 列不存在（如 '寒冷名刀'）")
        lines.append("  - #事件效果 H 列：K~R 列值（效果名称）在 #Main C 列不存在（如 '添加1层[冻结计数]'）")
        lines.append("修复方式：改名称列与目标表一致，或补目标表数据；不能用重算消除。")
    report_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(path))),
        "effect_recalc_report.txt",
    )
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print("\n".join(lines))
    print("\n报告已写入: %s" % report_path)
    return counts


def main():
    ap = argparse.ArgumentParser(description="Effect.xlsx 重算脚本（修复 #N/A）")
    ap.add_argument("--file", default=DEFAULT_FILE)
    ap.add_argument("--dry-run", action="store_true", help="只诊断不改文件")
    ap.add_argument("--no-backup", action="store_true", help="跳过备份")
    args = ap.parse_args()

    path = args.file
    if not os.path.isfile(path):
        print("[错误] 文件不存在: %s" % path)
        sys.exit(1)

    print("=== Effect.xlsx 重算脚本 ===")
    print("文件: %s" % path)

    if is_open_in_excel(path):
        print("[警告] 检测到 %s 正被 Excel 打开（存在 ~$ 锁文件）。" % os.path.basename(path))
        print("       请先关闭 Excel 中的该文件再运行本脚本，否则可能保存失败或只读打开。")
        if not args.dry_run:
            sys.exit(2)

    before = excel_recalc.count_formula_error_caches(path)
    print("重算前错误缓存: %s" % (before if before else "无"))

    if not args.dry_run:
        if not args.no_backup:
            bk = backup(path)
            print("已备份: %s" % bk)
        print("开始 COM 重算（读回写 + OR 保护 + CalculateFull + 保存）...")
        result = recalc(path)
        if not result or not result.get("ok"):
            print("[错误] 重算失败")
            sys.exit(1)
        print("  OR 保护合计: %d 处" % result.get("or_protections", 0))
    else:
        print("--dry-run：跳过修改，仅统计与报告。")

    after = excel_recalc.count_formula_error_caches(path)
    print("重算后错误缓存: %s" % (after if after else "无"))
    report_residual(path)


if __name__ == "__main__":
    main()
