from pathlib import Path

from diff_standalone import _parse_diff_args


def test_parse_tortoisesvn_two_file_args(tmp_path):
    old_file = tmp_path / "Effect.xlsx.r1"
    new_file = tmp_path / "Effect.xlsx"
    old_file.write_bytes(b"old")
    new_file.write_bytes(b"new")

    old_path, new_path, display_path = _parse_diff_args([str(old_file), str(new_file)])

    assert old_path == str(old_file)
    assert new_path == str(new_file)
    assert display_path == new_file


def test_parse_svn_external_diff_args(tmp_path):
    old_file = tmp_path / "915b99b9.svn-base"
    new_file = tmp_path / "Effect.xlsx"
    old_file.write_bytes(b"old")
    new_file.write_bytes(b"new")
    label_old = "D:/work/DataTable/Raw/Effect.xlsx\t(revision 4728)"
    label_new = "D:/work/DataTable/Raw/Effect.xlsx\t(working copy)"

    old_path, new_path, display_path = _parse_diff_args([
        "-u",
        "-L",
        label_old,
        "-L",
        label_new,
        str(old_file),
        str(new_file),
    ])

    assert old_path == str(old_file)
    assert new_path == str(new_file)
    assert display_path == Path("D:/work/DataTable/Raw/Effect.xlsx")
