# -*- coding: utf-8 -*-
"""
Effect.xlsx 公式静态诊断器（只读，不修改任何文件）

原理：
  #事件效果 / #Main 的 Params 列是 TEXTJOIN+XLOOKUP 公式：
    - I 列(事件名称) → Event.xlsx #Main C列(事件描述) → B列(事件ID)
    - J 列(目标名称) → Target.xlsx #Main D列(目标名称) → B列(目标ID)
    - K~R 列(效果名称) → 本文件 #Main C列(效果描述，含 =I&J 拼接公式) → B列(效果ID)
  #Main Params 列另有模式：XLOOKUP 查自身 C:C/B:B、查 [1]映射 / [2]Buff D列 / [5]Target D列
  公式计算出 #N/A = XLOOKUP 的查找值在查找数组中不存在（数据缺失 / 名称不一致）。

  本脚本用 openpyxl 只读解析：
    1. 计算本文件 #Main C 列（=I&J 拼接公式）与 B 列（=Bxx+1 链式公式）的真实值
    2. 读取外部表（Event/Target/Buff/ShopFilterCondition/AttributeTable）磁盘值
    3. 对每个 XLOOKUP 检查查找值是否命中
  输出：全部 #N/A 单元格清单 + 根因（哪个名称、在哪个表、缺失）。

用法：
  python effect_diag.py                 # 诊断 Effect.xlsx（默认 A:/DataTable/Raw/Effect.xlsx）
  python effect_diag.py <path.xlsx>     # 诊断指定文件副本
"""
import io
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import OrderedDict

try:
    import openpyxl
except ImportError:
    print("需要 openpyxl：pip install openpyxl")
    sys.exit(1)

RAW_DIR = r"A:/DataTable/Raw"
XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def norm_text(v):
    """单元格值 → 用于匹配的文本（Excel 中数字 9000001 与文本 '9000001' 匹配行为不同，这里统一 str）"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def load_sheet_values(path, sheet_name, data_only=True):
    """读取某表某 sheet 的全部单元格值（dict: (row, col) -> value）"""
    wb = openpyxl.load_workbook(path, data_only=data_only)
    if sheet_name not in wb.sheetnames:
        return {}
    ws = wb[sheet_name]
    out = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                out[(cell.row, cell.column)] = cell.value
    return out


class Table:
    """一张外部表的列值集合：col_index -> {值文本: 行号}，以及 col -> [行值]"""

    def __init__(self, path, sheet, col_map):
        self.path = path
        self.sheet = sheet
        self.col_map = col_map  # {col_idx: (lookup_col_idx, return_col_idx)}
        self.cells = load_sheet_values(path, sheet, data_only=True)
        self.formula_cells = load_sheet_values(path, sheet, data_only=False)
        # 列值索引：col -> dict(文本 -> 行号集合)
        self.col_index = {}
        for (r, c), v in self.cells.items():
            if c not in self.col_index:
                self.col_index[c] = OrderedDict()
            self.col_index[c].setdefault(norm_text(v), set()).add(r)

    def lookup(self, value, lookup_col, return_col):
        """模拟 XLOOKUP(value, lookup_col 整列, return_col 整列)：精确匹配第一个命中"""
        key = norm_text(value)
        idx = self.col_index.get(lookup_col)
        if idx is None:
            return None, "查找列 %s 不存在" % lookup_col
        rows = idx.get(key)
        if not rows:
            return None, "查找值 %r 在 %s!%s 列不存在" % (value, self.sheet, openpyxl.utils.get_column_letter(lookup_col))
        r = min(rows)
        return self.cells.get((r, return_col)), None


def col_letter_to_idx(letter):
    return openpyxl.utils.column_index_from_string(letter)


def parse_xlookup_args(fragment):
    """从公式片段中解析 XLOOKUP(lookup, array, return_array) —— 只处理整列引用"""
    m = re.match(r"_xlfn\.XLOOKUP\((.+)\)", fragment.strip())
    if not m:
        return None
    args = split_top_level(m.group(1))
    if len(args) < 3:
        return None
    return args[0], args[1], args[2]


def split_top_level(s):
    """按逗号分割，忽略引号内与括号内的逗号"""
    parts = []
    depth = 0
    cur = []
    in_str = False
    for ch in s:
        if ch == '"':
            in_str = not in_str
            cur.append(ch)
        elif ch == '(' and not in_str:
            depth += 1
            cur.append(ch)
        elif ch == ')' and not in_str:
            depth -= 1
            cur.append(ch)
        elif ch == ',' and depth == 0 and not in_str:
            parts.append("".join(cur).strip())
            cur = []
        else:
            cur.append(ch)
    parts.append("".join(cur).strip())
    return parts


class RefResolver:
    """解析公式中的引用：'[7]#Main'!$C:$C → 外部表；'#Main'!C:C → 本文件 #Main；I4 → 本 sheet 单元格"""

    def __init__(self, effect_path, raw_dir):
        self.effect_path = effect_path
        self.raw_dir = raw_dir
        # 外部链接编号 → 文件路径（从 effect 文件的 externalLinks rels 解析）
        self.ext_map = self._load_external_links()
        # 本文件各 sheet 的公式/值
        self.wb = openpyxl.load_workbook(effect_path, data_only=False)
        self.wb_v = openpyxl.load_workbook(effect_path, data_only=True)
        self.sheet_cells = {}  # sheet -> {(r,c): value}
        self.sheet_formulas = {}
        for sn in self.wb.sheetnames:
            ws = self.wb[sn]
            self.sheet_formulas[sn] = {}
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        self.sheet_formulas[sn][(cell.row, cell.column)] = cell.value
            ws_v = self.wb_v[sn]
            self.sheet_cells[sn] = {}
            for row in ws_v.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        self.sheet_cells[sn][(cell.row, cell.column)] = cell.value
        # 外部表缓存（懒加载）
        self.ext_tables = {}

    def _load_external_links(self):
        """解析 Effect.xlsx 内 externalLinks 的 rels：编号 -> 文件路径"""
        mapping = {}
        try:
            zf = zipfile.ZipFile(self.effect_path)
        except Exception:
            return mapping
        REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
        rels = None
        for name in zf.namelist():
            if name == "xl/_rels/workbook.xml.rels":
                rels = zf.read(name)
                break
        if not rels:
            return mapping
        root = ET.fromstring(rels)
        for rel in root.findall(REL_NS + "Relationship"):
            if rel.get("Type", "").endswith("externalLink"):
                target = rel.get("Target", "")
                m = re.search(r"externalLink(\d+)\.xml", target)
                if m:
                    idx = int(m.group(1))
                    rels_path = "xl/externalLinks/_rels/externalLink%d.xml.rels" % idx
                    try:
                        rels_xml = zf.read(rels_path)
                    except KeyError:
                        continue
                    rroot = ET.fromstring(rels_xml)
                    for rrel in rroot.findall(REL_NS + "Relationship"):
                        if rrel.get("Type", "").endswith("externalLinkPath"):
                            tgt = rrel.get("Target", "")
                            if tgt.startswith("file:///"):
                                tgt = tgt[len("file:///"):]
                            mapping[idx] = tgt.replace("\\", "/")
                            break
        return mapping

    def get_ext_table(self, link_idx):
        """外部链接编号 → Table（懒加载）"""
        if link_idx in self.ext_tables:
            return self.ext_tables[link_idx]
        path = self.ext_map.get(link_idx)
        table = None
        if path:
            full = os.path.join(self.raw_dir, path) if not os.path.isabs(path) else path
            # 相对路径基于 Raw 目录解析
            if not os.path.exists(full):
                # 尝试从 effect 文件所在目录解析
                full = os.path.join(os.path.dirname(self.effect_path), path)
            if os.path.exists(full):
                try:
                    wb = openpyxl.load_workbook(full, data_only=True)
                    sheets = {sn.lower(): sn for sn in wb.sheetnames}
                    table = {"path": full, "wb": wb, "sheets": sheets}
                except Exception as e:
                    table = {"error": str(e)}
            else:
                table = {"error": "外部文件不存在: %s" % full}
        else:
            table = {"error": "未知外部链接编号 [%d]" % link_idx}
        self.ext_tables[link_idx] = table
        return table

    def resolve_ref(self, ref, sheet_name):
        """解析引用 → ('value', v) 或 ('column', sheet, col) 或 ('error', msg)
        支持：'[7]#Main'!$C:$C / '[7]#Main'!$C:$C / '#Main'!C:C / I4 / B42"""
        ref = ref.strip()
        m = re.match(r"^'\[(\d+)\]([^']+)'!\$?([A-Z]+):\$?([A-Z]+)$", ref)
        if m:
            link_idx = int(m.group(1))
            sheet = m.group(2)
            col_l = col_letter_to_idx(m.group(3))
            table = self.get_ext_table(link_idx)
            if "error" in table:
                return ("error", table["error"])
            real_sheet = table["sheets"].get(sheet.lower())
            if real_sheet is None:
                return ("error", "外部表 %s 无 sheet %s" % (table["path"], sheet))
            ws = table["wb"][real_sheet]
            col_vals = []
            for row in ws.iter_rows(min_col=col_l, max_col=col_l):
                cell = row[0]
                if cell.value is not None:
                    col_vals.append((cell.row, cell.value))
            return ("column", table["path"] + "!" + sheet, col_l, col_vals)
        m = re.match(r"^'#([^']+)'!\$?([A-Z]+):\$?([A-Z]+)$", ref)
        if m:
            sheet = "#" + m.group(1)
            col_l = col_letter_to_idx(m.group(2))
            return ("local_column", sheet, col_l)
        m = re.match(r"^'?([A-Z]+):([A-Z]+)'?$", ref)
        if m:
            col_l = col_letter_to_idx(m.group(1))
            return ("local_column", sheet_name, col_l)
        m = re.match(r"^'?([A-Z]+)(\d+)'?$", ref)
        if m:
            col_l = col_letter_to_idx(m.group(1))
            r = int(m.group(2))
            cells = self.sheet_formulas.get(sheet_name, {})
            return ("value", cells.get((r, col_l)))
        return ("error", "无法解析引用 %r" % ref)

    def compute_local_column(self, sheet_name, col):
        """计算本文件某 sheet 某列的全部值（含公式结果，行序）：
        支持公式：=Bxx+1（ID 链）、=I&J[&K...]（描述拼接）、普通值"""
        key = (sheet_name, col)
        if key in self._local_col_cache:
            return self._local_col_cache[key]
        formulas = self.sheet_formulas[sheet_name]
        cells = self.sheet_cells[sheet_name]
        max_row = max([r for (r, c) in formulas.keys() if c == col] or [0])
        # 先按行升序处理
        result = OrderedDict()
        for r in range(1, max_row + 2):
            v = formulas.get((r, col))
            if v is None:
                if (r, col) in cells:
                    result[r] = cells[(r, col)]
                continue
            if isinstance(v, str) and v.startswith("="):
                result[r] = self._eval_formula(v, sheet_name, r, col)
            else:
                result[r] = v
        self._local_col_cache[key] = result
        return result

    def _eval_formula(self, formula, sheet_name, row, col):
        """求值本文件内的简单公式"""
        f = formula[1:].strip()
        # =Bxx+1
        m = re.match(r"^([A-Z]+)(\d+)\+1$", f)
        if m:
            c = col_letter_to_idx(m.group(1))
            r = int(m.group(2))
            base = self._get_cell_value(sheet_name, r, c)
            if isinstance(base, (int, float)):
                return base + 1
            return "#VALUE!%s" % norm_text(base)
        # =I&J&K...
        if "&" in f and not f.startswith("_xlfn"):
            parts = []
            for tok in re.split(r"&", f):
                tok = tok.strip()
                v = self._resolve_cell(tok, sheet_name)
                parts.append(norm_text(v) if v is not None else "")
            return "".join(parts)
        # 其它公式暂不模拟
        return None

    def _get_cell_value(self, sheet_name, r, c):
        """取本文件某单元格的值（公式则求值）"""
        formulas = self.sheet_formulas.get(sheet_name, {})
        v = formulas.get((r, c))
        if isinstance(v, str) and v.startswith("="):
            return self._eval_formula(v, sheet_name, r, c)
        if v is not None:
            return v
        return self.sheet_cells.get(sheet_name, {}).get((r, c))

    def _resolve_cell(self, tok, sheet_name):
        m = re.match(r"^([A-Z]+)(\d+)$", tok)
        if m:
            c = col_letter_to_idx(m.group(1))
            r = int(m.group(2))
            return self._get_cell_value(sheet_name, r, c)
        return None


# ---- 诊断主流程 ----

def diagnose(effect_path):
    resolver = RefResolver(effect_path, RAW_DIR)
    resolver._local_col_cache = {}
    # 预计算 #Main C 列与 B 列
    main_c = resolver.compute_local_column("#Main", 3)
    main_b = resolver.compute_local_column("#Main", 2)
    ev_c = resolver.compute_local_column("#事件效果", 3)
    ev_b = resolver.compute_local_column("#事件效果", 2)

    out = io.StringIO()

    def w(s=""):
        out.write(s + "\n")

    # 构建 #Main C 列查找索引
    main_c_index = OrderedDict()
    for r, v in main_c.items():
        if v is not None and not (isinstance(v, str) and v.startswith("#")):
            main_c_index.setdefault(norm_text(v), set()).add(r)

    def lookup_local(lookup_val, col_values, return_values, col_desc):
        """模拟 XLOOKUP：在 col_values(OrderedDict 行->值) 中精确匹配"""
        key = norm_text(lookup_val)
        for r, v in col_values.items():
            if v is not None and norm_text(v) == key:
                return return_values.get(r)
        return None

    issues = []

    # ============ #事件效果 H 列 ============
    ws = resolver.wb["#事件效果"]
    for r in range(4, ws.max_row + 1):
        h = ws.cell(row=r, column=8).value
        if not (isinstance(h, str) and h.startswith("=")):
            continue
        # 提取每个 IF(x="无",0,XLOOKUP(...)) 参数
        # 简化：找所有 XLOOKUP(...) 并配对前面的 I/J/K... 列引用
        xlookups = []
        for m in re.finditer(r"_xlfn\.XLOOKUP\(([^()]*(?:\([^()]*\))?[^()]*)\)", h):
            xlookups.append(m.group(1))
        row_issues = []
        for xl in xlookups:
            args = split_top_level(xl)
            if len(args) < 3:
                continue
            lookup_ref, arr_ref, ret_ref = args[0], args[1], args[2]
            # lookup_ref 通常是本行单元格（I4/J4/K4...）
            m = re.match(r"^([A-Z]+)(\d+)$", lookup_ref.strip())
            if not m:
                continue
            lc = col_letter_to_idx(m.group(1))
            lr = int(m.group(2))
            lookup_val = resolver._get_cell_value("#事件效果", lr, lc)
            if lookup_val is None or norm_text(lookup_val) == "":
                continue
            # IF(单元格="无",0,XLOOKUP(...)) 保护分支：查找值为"无"时不执行 XLOOKUP，输出 0
            if norm_text(lookup_val) == "无":
                continue
            # 判断 arr_ref 指向哪里
            arr = arr_ref.strip()
            if arr.startswith("'[") or arr.startswith("["):
                # 外部表
                m2 = re.match(r"^'\[(\d+)\]([^']+)'!\$?([A-Z]+):\$?([A-Z]+)$", arr)
                if m2:
                    link_idx = int(m2.group(1))
                    sheet = m2.group(2)
                    col_l = col_letter_to_idx(m2.group(3))
                    table = resolver.get_ext_table(link_idx)
                    if "error" in table:
                        row_issues.append("R%d %s: %s" % (r, openpyxl.utils.get_column_letter(lc), table["error"]))
                        continue
                    real_sheet = table["sheets"].get(sheet.lower())
                    ws_ext = table["wb"][real_sheet]
                    col_vals = []
                    ret_vals = {}
                    for row in ws_ext.iter_rows(min_col=col_l, max_col=col_l):
                        cell = row[0]
                        if cell.value is not None:
                            col_vals.append((cell.row, cell.value))
                    # return col: 公式里的第3参数
                    ret_letter = None
                    m3 = re.match(r"^'\[(\d+)\]([^']+)'!\$?([A-Z]+):\$?([A-Z]+)$", ret_ref.strip())
                    if m3:
                        ret_letter = m3.group(3)
                    if ret_letter:
                        ret_c = col_letter_to_idx(ret_letter)
                        for row in ws_ext.iter_rows(min_col=ret_c, max_col=ret_c):
                            cell = row[0]
                            if cell.value is not None:
                                ret_vals[cell.row] = cell.value
                    hit = None
                    for rr, vv in col_vals:
                        if norm_text(vv) == norm_text(lookup_val):
                            hit = rr
                            break
                    if hit is None:
                        row_issues.append("R%d %s列值 %r 在 %s!%s 列不存在(XLOOKUP→#N/A)" % (
                            r, openpyxl.utils.get_column_letter(lc), lookup_val, table["path"], openpyxl.utils.get_column_letter(col_l)))
            elif "'#Main'!" in arr or "#Main" in arr:
                hit = lookup_local(lookup_val, main_c, main_b, "#Main C")
                if hit is None:
                    row_issues.append("R%d %s列值 %r 在 Effect.xlsx #Main C列(描述) 不存在(XLOOKUP→#N/A)" % (
                        r, openpyxl.utils.get_column_letter(lc), lookup_val))
            else:
                row_issues.append("R%d 未识别的查找区域 %r" % (r, arr))
        if row_issues:
            issues.append((r, row_issues))

    w("===== #事件效果 H列(Params) #N/A 预测 =====")
    for r, ri in issues:
        w("行 %d:" % r)
        for x in ri:
            w("  - " + x)
    w("总计: %d 行" % len(issues))

    # ============ #Main H 列 ============
    ws = resolver.wb["#Main"]
    issues_main = []
    for r in range(4, ws.max_row + 1):
        h = ws.cell(row=r, column=8).value
        if not (isinstance(h, str) and h.startswith("=")):
            continue
        xlookups = re.findall(r"_xlfn\.XLOOKUP\(([^()]*(?:\([^()]*\))?[^()]*)\)", h)
        row_issues = []
        for xl in xlookups:
            args = split_top_level(xl)
            if len(args) < 3:
                continue
            lookup_ref, arr_ref, ret_ref = args[0], args[1], args[2]
            m = re.match(r"^([A-Z]+)(\d+)$", lookup_ref.strip())
            if not m:
                continue
            lc = col_letter_to_idx(m.group(1))
            lr = int(m.group(2))
            lookup_val = resolver._get_cell_value("#Main", lr, lc)
            if lookup_val is None or norm_text(lookup_val) == "":
                continue
            # IF(单元格="无",0,XLOOKUP(...)) 保护分支：查找值为"无"时不执行 XLOOKUP
            if norm_text(lookup_val) == "无":
                continue
            arr = arr_ref.strip()
            if arr.startswith("'[") or arr.startswith("["):
                m2 = re.match(r"^'\[(\d+)\]([^']+)'!\$?([A-Z]+):\$?([A-Z]+)$", arr)
                if m2:
                    link_idx = int(m2.group(1))
                    sheet = m2.group(2)
                    col_l = col_letter_to_idx(m2.group(3))
                    table = resolver.get_ext_table(link_idx)
                    if "error" in table:
                        row_issues.append("R%d %s: %s" % (r, openpyxl.utils.get_column_letter(lc), table["error"]))
                        continue
                    real_sheet = table["sheets"].get(sheet.lower())
                    if real_sheet is None:
                        row_issues.append("R%d %s: 外部表无 sheet %s" % (r, openpyxl.utils.get_column_letter(lc), sheet))
                        continue
                    ws_ext = table["wb"][real_sheet]
                    col_vals = []
                    for row in ws_ext.iter_rows(min_col=col_l, max_col=col_l):
                        cell = row[0]
                        if cell.value is not None:
                            col_vals.append((cell.row, cell.value))
                    hit = any(norm_text(vv) == norm_text(lookup_val) for _, vv in col_vals)
                    if not hit:
                        row_issues.append("R%d %s列值 %r 在 %s!%s 列不存在(XLOOKUP→#N/A)" % (
                            r, openpyxl.utils.get_column_letter(lc), lookup_val, table["path"], openpyxl.utils.get_column_letter(col_l)))
            elif re.match(r"^[A-Z]+:[A-Z]+$", arr) or "#Main" in arr:
                hit = lookup_local(lookup_val, main_c, main_b, "#Main C")
                if hit is None:
                    row_issues.append("R%d %s列值 %r 在 Effect.xlsx #Main C列(描述) 不存在(XLOOKUP→#N/A)" % (
                        r, openpyxl.utils.get_column_letter(lc), lookup_val))
        if row_issues:
            issues_main.append((r, row_issues))

    w("\n===== #Main H列(Params) #N/A 预测 =====")
    for r, ri in issues_main:
        w("行 %d:" % r)
        for x in ri:
            w("  - " + x)
    w("总计: %d 行" % len(issues_main))

    # ============ #事件效果 F 列 (ConditionIDs) ============
    issues_f = []
    for r in range(4, ws.max_row + 1):
        f = ws.cell(row=r, column=6).value
        if not (isinstance(f, str) and f.startswith("=")):
            continue
        m = re.search(r"_xlfn\.XLOOKUP\(([^)]+)\)", f)
        if not m:
            continue
        args = split_top_level(m.group(1))
        if len(args) < 3:
            continue
        lookup_ref = args[0].strip()
        mm = re.match(r"^([A-Z]+)(\d+)$", lookup_ref)
        if not mm:
            continue
        lc = col_letter_to_idx(mm.group(1))
        lr = int(mm.group(2))
        lookup_val = resolver._get_cell_value("#事件效果", lr, lc)
        if lookup_val is None or norm_text(lookup_val) == "":
            continue
        if norm_text(lookup_val) == "无":
            continue
        arr = args[1].strip()
        m2 = re.match(r"^'\[(\d+)\]([^']+)'!\$?([A-Z]+):\$?([A-Z]+)$", arr)
        if m2:
            link_idx = int(m2.group(1))
            sheet = m2.group(2)
            col_l = col_letter_to_idx(m2.group(3))
            table = resolver.get_ext_table(link_idx)
            if "error" in table:
                continue
            real_sheet = table["sheets"].get(sheet.lower())
            if real_sheet is None:
                continue
            ws_ext = table["wb"][real_sheet]
            col_vals = []
            for row in ws_ext.iter_rows(min_col=col_l, max_col=col_l):
                cell = row[0]
                if cell.value is not None:
                    col_vals.append((cell.row, cell.value))
            hit = any(norm_text(vv) == norm_text(lookup_val) for _, vv in col_vals)
            if not hit:
                issues_f.append((r, "%s列值 %r 在 %s!%s 列不存在" % (
                    openpyxl.utils.get_column_letter(lc), lookup_val, table["path"], openpyxl.utils.get_column_letter(col_l))))
    w("\n===== #事件效果 F列(ConditionIDs) #N/A 预测 =====")
    for r, x in issues_f:
        w("行 %d: %s" % (r, x))
    w("总计: %d 行" % len(issues_f))

    return out.getvalue()


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(RAW_DIR, "Effect.xlsx")
    text = diagnose(path)
    out_path = os.path.join(os.path.dirname(os.path.abspath(path)), "effect_diag_report.txt")
    with io.open(out_path, "w", encoding="utf-8") as f:
        f.write(text)
    print("诊断报告已写入: %s" % out_path)
    print("行数: %d" % len(text.splitlines()))
