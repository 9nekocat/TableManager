# -*- coding: utf-8 -*-
"""Tests for formula cache detection used after merge resolve."""

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import openpyxl
import pytest
import zipfile

import excel_recalc
from excel_recalc import count_missing_formula_caches, ensure_formula_cache


def test_missing_rowname_formula_cache_is_detected():
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "#TestSheet"
        ws.append(["RowName", "Name", "Value"])
        ws.append(["RowName", "Name", "Value"])
        ws.append(["FString", "FString", "int32"])
        ws.append(["=B4", "Hero_001", 100])
        wb.save(path)
        wb.close()

        assert count_missing_formula_caches(path) == 1
    finally:
        os.unlink(path)


def test_non_key_formula_cache_is_ignored_by_default():
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "#TestSheet"
        ws.append(["RowName", "Name", "Value"])
        ws.append(["RowName", "Name", "Value"])
        ws.append(["FString", "FString", "int32"])
        ws.append(["Hero_001", "=A4", 100])
        wb.save(path)
        wb.close()

        assert count_missing_formula_caches(path) == 0
        assert count_missing_formula_caches(path, important_columns=(2,)) == 1
    finally:
        os.unlink(path)


def test_ensure_formula_cache_retries_transient_missing(monkeypatch, tmp_path):
    """A recalc pass that leaves caches missing is retried, not fatal."""
    target = tmp_path / "dummy.xlsx"
    target.write_bytes(b"x")
    recalc_calls = []
    missing_sequence = iter([3, 0])

    monkeypatch.setattr(
        excel_recalc, "recalc_xlsx",
        lambda path, **kw: recalc_calls.append(path) or path,
    )
    monkeypatch.setattr(
        excel_recalc, "count_missing_formula_caches",
        lambda path, **kw: next(missing_sequence),
    )

    result = ensure_formula_cache(str(target), raise_on_error=True)

    assert result == {"ok": True, "missing_important_caches": 0}
    assert len(recalc_calls) == 2


def test_ensure_formula_cache_raises_after_retries_exhausted(monkeypatch, tmp_path):
    target = tmp_path / "dummy.xlsx"
    target.write_bytes(b"x")
    recalc_calls = []

    monkeypatch.setattr(
        excel_recalc, "recalc_xlsx",
        lambda path, **kw: recalc_calls.append(path) or path,
    )
    monkeypatch.setattr(
        excel_recalc, "count_missing_formula_caches",
        lambda path, **kw: 3,
    )

    with pytest.raises(RuntimeError, match="3 important formula cache"):
        ensure_formula_cache(str(target), raise_on_error=True, retries=2)

    assert len(recalc_calls) == 3


def _inject_empty_string_cache(path, sheet_xml, cell_ref):
    """Rewrite one formula cell to ``t="str"`` with an empty ``<v/>`` cache,
    mimicking how Excel saves a formula that evaluates to an empty string."""
    import re
    import shutil
    import tempfile

    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == sheet_xml:
                text = data.decode("utf-8")
                pattern = re.compile(
                    r'(<c r="%s")[^>]*(><f[^>]*>[^<]*</f>)(?:<v\s*/>|<v>[^<]*</v>)?(</c>)' % cell_ref
                )
                text, n = pattern.subn(r'\1 t="str"\2<v/>\3', text)
                assert n == 1, f"cell {cell_ref} formula not found in {sheet_xml}"
                data = text.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)


def test_empty_string_formula_cache_is_not_missing(tmp_path):
    """Excel caches an empty-string result as <v/> — present, not missing."""
    path = str(tmp_path / "wb.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#TestSheet"
    ws.append(["RowName", "Name", "Value"])
    ws.append(["RowName", "Name", "Value"])
    ws.append(["FString", "FString", "int32"])
    ws.append(["=B4", "Hero_001", 100])
    wb.save(path)
    wb.close()

    assert count_missing_formula_caches(path) == 1

    _inject_empty_string_cache(path, "xl/worksheets/sheet1.xml", "A4")

    assert count_missing_formula_caches(path) == 0
