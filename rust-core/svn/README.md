# SVN Excel diff and three-way merge

These wrappers connect Subversion and TortoiseSVN to the Rust `table-core`
executable. Excel files use the table-aware diff/merge implementation; other
files are delegated to xlCompare.

Build the executable first:

```powershell
cargo build --release --manifest-path rust-core/Cargo.toml
```

## Subversion command line

Add these entries to `%APPDATA%\Subversion\config` (use absolute forward-slash
paths for this checkout):

```ini
[helpers]
diff-cmd = D:/Project/TableManager-refactor/rust-core/svn/svn-diff.cmd
diff-extensions =
merge-tool-cmd = D:/Project/TableManager-refactor/rust-core/svn/svn-merge.cmd

[miscellany]
diff-ignore-content-type = yes
```

## TortoiseSVN

Configure the `.xlsx`, `.xlsm`, `.xlsb`, and `.xls` entries under
**Settings > Diff Viewer > Advanced** and
**Settings > Merge Tool > Advanced**.

Diff:

```text
"D:\Project\TableManager-refactor\rust-core\svn\svn-diff.cmd" %base %mine
```

Merge:

```text
"D:\Project\TableManager-refactor\rust-core\svn\svn-merge.cmd" %merged %theirs %mine %base
```

## Behavior

- `svn-diff.cmd` writes a `Summary`, `Changes`, and `Issues` workbook and opens
  it for review.
- `svn-merge.cmd` writes the merged workbook plus `_MergeReview`,
  `_MergeConflicts`, `_MergeChanges`, and `_MergeIssues`.
- An unresolved merge conflict returns exit code `2`, so Subversion will not
  silently treat the merge as successful. The review workbook is still written
  and opened.
- The wrappers recognize filenames such as `.xlsx.r1234`, `.xlsx.mine`, and
  temporary SVN names.

Environment overrides:

```text
TABLE_CORE_EXE       Path to a table-core.exe build
TM_SVN_NO_OPEN=1     Generate files without opening Excel
TM_SVN_OUTPUT        Exact output path for diff (mainly for testing)
```
