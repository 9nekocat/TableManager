import shutil
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import openpyxl
import pytest

import merge_engine as merge_engine_module
from merge_engine import generate_merge_preview
from merge_workflow import (
    MergeWorkflowError,
    apply_merge_preview_transactional,
    create_merge_backup,
    recover_merge_preview_transactional,
)


def _write_book(path, value):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#Data"
    for row_idx, row in enumerate(
        [
            ("RowName", "Name", "Value"),
            ("Row", "Name", "Value"),
            ("string", "string", "int"),
        ],
        start=1,
    ):
        for col_idx, cell_value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=cell_value)
    ws.cell(row=4, column=1, value="Hero_001")
    ws.cell(row=4, column=2, value="Alice")
    ws.cell(row=4, column=3, value=value)
    wb.save(path)
    wb.close()


def _read_value(path):
    wb = openpyxl.load_workbook(path, data_only=False)
    try:
        return wb["#Data"].cell(row=4, column=3).value
    finally:
        wb.close()


def _raise_recalc(*_args, **_kwargs):
    raise RuntimeError("simulated recalc failure")


def _ok_recalc(*_args, **_kwargs):
    return {"ok": True, "missing_important_caches": 0}


def _write_external_link_rels(path, targets):
    """注入完整外链结构：externalLink1.xml 本体 + rels（缺一不可，否则
    Excel 加载时会把不完整的 rels 清理掉，COM 直写后外链丢失）。"""
    link_name = "xl/externalLinks/externalLink1.xml"
    rels_name = "xl/externalLinks/_rels/externalLink1.xml.rels"
    link = (
        '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<externalBook r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames>'
        '<sheetDataSet><sheetData/></sheetDataSet></externalBook></externalLink>'
    )
    rels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    for idx, target in enumerate(targets, start=1):
        rels.append(
            f'<Relationship Id="rId{idx}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" '
            f'Target="{target}" TargetMode="External"/>'
        )
    rels.append("</Relationships>")
    tmp = path.with_name(f".{path.name}.rels.tmp")
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename in (link_name, rels_name):
                continue
            zout.writestr(info, zin.read(info.filename))
        zout.writestr(link_name, link)
        zout.writestr(rels_name, "".join(rels))
    tmp.replace(path)


def _read_external_link_targets(path):
    targets = []
    with zipfile.ZipFile(path, "r") as z:
        for name in z.namelist():
            if name.startswith("xl/externalLinks/_rels/") and name.endswith(".rels"):
                text = z.read(name).decode("utf-8")
                for piece in text.split('Target="')[1:]:
                    targets.append(piece.split('"', 1)[0])
    return targets


def test_transactional_merge_restores_backup_when_com_write_fails(
    tmp_path, monkeypatch
):
    """COM 直写失败（如公式拒绝/COM 异常）时从事务备份回滚。"""
    import excel_recalc

    base = tmp_path / "base.xlsx"
    mine = tmp_path / "mine.xlsx"
    theirs = tmp_path / "theirs.xlsx"
    preview = tmp_path / "table_merge.xlsx"
    target = tmp_path / "table.xlsx"
    backup_root = tmp_path / "backups"

    _write_book(base, 100)
    _write_book(mine, 200)
    _write_book(theirs, 300)
    shutil.copy2(mine, target)
    original_bytes = target.read_bytes()
    generate_merge_preview(str(base), str(mine), str(theirs), str(preview))

    # 强制走 COM 直写路径并让写回抛错（模拟公式拒绝/COM 异常）
    def _fail_com_write(*_args, **_kwargs):
        raise RuntimeError("simulated COM write failure")

    import merge_workflow as merge_workflow_module

    monkeypatch.setattr(excel_recalc, "com_available", lambda: True)
    monkeypatch.setattr(
        merge_workflow_module, "com_available", lambda: True
    )
    monkeypatch.setattr(excel_recalc, "apply_write_plan_via_com", _fail_com_write)

    with pytest.raises(MergeWorkflowError) as exc:
        apply_merge_preview_transactional(
            preview,
            target,
            mine_path=target,
            theirs_path=theirs,
            backup_root=backup_root,
        )

    assert exc.value.restored is True
    assert exc.value.backup_path
    assert target.read_bytes() == original_bytes
    assert _read_value(target) == 200


def test_card_merge_keeps_excel_recalc_link_table_and_prefers_a_drive(tmp_path, monkeypatch):
    """Regression for Card: recalc link table must not be overwritten by stale cache.

    强制走 openpyxl 降级路径验证外链 zip 注入恢复 + A: 盘 normalize
    （COM 直写路径由真实 Excel 校验外链，手工注入的伪造外链结构会被
    Excel 清理——真实外链保留在回归任务中用真实文件副本验证）。
    """
    import excel_recalc
    import merge_workflow as merge_workflow_module

    monkeypatch.setattr(merge_engine_module, "RAW_DIR", tmp_path)
    monkeypatch.setattr(excel_recalc, "com_available", lambda: False)
    monkeypatch.setattr(merge_workflow_module, "com_available", lambda: False)

    base = tmp_path / "base.xlsx"
    mine = tmp_path / "mine.xlsx"
    theirs = tmp_path / "theirs.xlsx"
    preview = tmp_path / "table_merge.xlsx"
    target = tmp_path / "Localization" / "Card.xlsx"
    target.parent.mkdir()

    _write_book(base, 100)
    _write_book(mine, 200)
    _write_book(theirs, 300)
    generate_merge_preview(str(base), str(mine), str(theirs), str(preview))

    # 外链 rels 带非规范目标（Z: 盘 stale / D: 盘工作路径）；resolve 写回
    # 必须原样保留外链，随后 normalize 统一到 A: 盘共享路径
    _write_external_link_rels(
        mine,
        [
            "file:///Z:\\OldWorkspace\\DataTable\\Raw\\Localization\\Stale.xlsx",
        ],
    )

    result = apply_merge_preview_transactional(
        preview,
        target,
        mine_path=mine,
        theirs_path=theirs,
        backup_root=tmp_path / "backups",
    )

    assert result["ok"] is True
    assert result["external_link_targets"]["rewritten"] == 1
    assert _read_external_link_targets(target) == [
        "file:///A:\\DataTable\\Raw\\Localization\\Stale.xlsx",
    ]


def test_recover_merge_preview_can_restore_from_existing_backup(tmp_path):
    base = tmp_path / "base.xlsx"
    mine = tmp_path / "mine.xlsx"
    theirs = tmp_path / "theirs.xlsx"
    preview = tmp_path / "table_merge.xlsx"
    target = tmp_path / "table.xlsx"
    backup_root = tmp_path / "backups"

    _write_book(base, 100)
    _write_book(mine, 200)
    _write_book(theirs, 300)
    shutil.copy2(mine, target)
    clean_backup = create_merge_backup(target, backup_root=backup_root)
    generate_merge_preview(str(base), str(mine), str(theirs), str(preview))

    target.write_bytes(b"broken workbook")

    result = recover_merge_preview_transactional(
        preview,
        target,
        backup_path=clean_backup,
        theirs_path=theirs,
        backup_root=backup_root,
    )

    assert result["ok"] is True
    assert result["recovery_base"] == str(clean_backup)
    assert _read_value(target) == 300
