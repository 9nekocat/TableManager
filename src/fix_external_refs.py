# -*- coding: utf-8 -*-
"""zip 层替换 Effect.xlsx 公式中的编号式外部引用 '[N]sheet'! → 文件名式 '[File.xlsx]sheet'!"""
import zipfile
import re
import shutil
import os
import sys
import xml.etree.ElementTree as ET

REL_NS = '{http://schemas.openxmlformats.org/package/2006/relationships}'


def replace_external_refs(src, out, absolute=False, strip_extlinks=False):
    """把公式中的编号式外部引用 '[N]sheet'! 替换为文件名式引用。

    absolute=True → 绝对路径引用（'A:/DataTable/Raw/[Event.xlsx]sheet'!）
    absolute=False → 相对文件名引用（'[Event.xlsx]sheet'!）

    strip_extlinks=True → 同时删除 externalLinks 目录、workbook.xml.rels 中的
    externalLink/calcChain 关系及 [Content_Types] 的 Override，让 Excel 打开时
    动态重建链接（Excel 2024 对新建链接保存为绝对路径引用，重算可靠）。
    """
    global REPLACE_COUNT
    REPLACE_COUNT = 0
    tmp = os.path.join(os.path.dirname(src), '.fix_tmp_effect')
    if os.path.exists(tmp):
        shutil.rmtree(tmp)
    os.makedirs(tmp)
    with zipfile.ZipFile(src) as zf:
        zf.extractall(tmp)

    # 编号 → 文件名映射
    mapping = {}
    root = ET.parse(os.path.join(tmp, 'xl/_rels/workbook.xml.rels')).getroot()
    for rel in root.findall(REL_NS + 'Relationship'):
        if rel.get('Type', '').endswith('externalLink'):
            tgt = rel.get('Target', '')
            m = re.search(r'externalLink(\d+)\.xml', tgt)
            if m:
                idx = int(m.group(1))
                rp = 'xl/externalLinks/_rels/externalLink%d.xml.rels' % idx
                if os.path.exists(os.path.join(tmp, rp)):
                    rroot = ET.parse(os.path.join(tmp, rp)).getroot()
                    for rrel in rroot.findall(REL_NS + 'Relationship'):
                        if rrel.get('Type', '').endswith('externalLinkPath'):
                            t = rrel.get('Target', '')
                            if t.startswith('file:///'):
                                t = t[len('file:///'):]
                            # 优先用相对路径（保留 Localization/ 子目录结构）
                            rel_t = t.replace('\\', '/')
                            fname = os.path.basename(rel_t)
                            mapping[idx] = (rel_t, fname)
                            break
    print('编号→文件名映射:', mapping)

    # 替换所有 sheet XML 中的编号引用
    count = 0
    sheets_dir = os.path.join(tmp, 'xl/worksheets')
    for name in sorted(os.listdir(sheets_dir)):
        if not name.endswith('.xml'):
            continue
        p = os.path.join(sheets_dir, name)
        with open(p, encoding='utf-8') as f:
            data = f.read()
        def repl(mo):
            global REPLACE_COUNT
            if mo.group(1) is not None:
                idx = int(mo.group(1))
                sheet = mo.group(2)
            else:
                idx = int(mo.group(3))
                sheet = mo.group(4)
            entry = mapping.get(idx)
            if not entry:
                return mo.group(0)
            REPLACE_COUNT += 1
            rel_t, fname = entry
            if absolute:
                # 绝对路径：基于源文件所在目录（externalLink 相对路径是相对源文件目录的）
                base = os.path.dirname(os.path.abspath(src)).replace('\\', '/')
                full = base + '/' + rel_t
                return "'%s\\[%s]%s'!" % (full, fname, sheet)
            return "'[%s]%s'!" % (fname, sheet)
        # 覆盖两种格式：'[7]#Main'! 与 [1]映射!（无引号）
        new = re.sub(r"'\[(\d+)\]([^'!]+?)'!|\[(\d+)\]([^'!]+?)!", repl, data)
        if new != data:
            with open(p, 'w', encoding='utf-8') as f:
                f.write(new)
    print('替换的引用数:', REPLACE_COUNT)

    # 重新打包（zip 条目路径必须用正斜杠，否则 Excel 无法加载）
    if strip_extlinks:
        # 删除 externalLinks 目录与 calcChain，并清理 rels/[Content_Types] 引用
        ext_dir = os.path.join(tmp, 'xl/externalLinks')
        if os.path.exists(ext_dir):
            shutil.rmtree(ext_dir)
        cc = os.path.join(tmp, 'xl/calcChain.xml')
        if os.path.exists(cc):
            os.remove(cc)
        # workbook.xml.rels：删除 externalLink / calcChain 关系
        rels_path = os.path.join(tmp, 'xl/_rels/workbook.xml.rels')
        if os.path.exists(rels_path):
            with open(rels_path, encoding='utf-8') as f:
                rels_data = f.read()
            rels_data = re.sub(
                r'<Relationship Id="rId\d+" Type="[^"]*(?:externalLink|calcChain)[^"]*"[^/]*/>',
                '', rels_data)
            # 删除可能的空 Relationships 自闭合残留（保留根元素）
            rels_data = re.sub(r'<Relationships[^>]*>\s*</Relationships>',
                               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
                               rels_data)
            with open(rels_path, 'w', encoding='utf-8') as f:
                f.write(rels_data)
        # [Content_Types].xml：删除 externalLink / calcChain Override
        ct_path = os.path.join(tmp, '[Content_Types].xml')
        if os.path.exists(ct_path):
            with open(ct_path, encoding='utf-8') as f:
                ct_data = f.read()
            ct_data = re.sub(
                r'<Override PartName="/xl/(?:externalLinks/[^"]+|calcChain\.xml)"[^/]*/>',
                '', ct_data)
            with open(ct_path, 'w', encoding='utf-8') as f:
                f.write(ct_data)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root_d, dirs, files in os.walk(tmp):
            for fn in files:
                full = os.path.join(root_d, fn)
                arc = os.path.relpath(full, tmp).replace('\\', '/')
                zf.write(full, arc)
    shutil.rmtree(tmp)
    print('打包完成:', out)


if __name__ == '__main__':
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else src + '.replaced.xlsx'
    strip = '--strip-extlinks' in sys.argv
    replace_external_refs(src, out, absolute=False, strip_extlinks=strip)
