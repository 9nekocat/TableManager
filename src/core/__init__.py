"""Core table model, diff, merge, and reporting APIs.

This package intentionally has no dependency on the GUI, SVN client, project
root discovery, or Excel COM automation.
"""

from .model import (
    Cell,
    Header,
    Issue,
    Row,
    Sheet,
    ValidationReport,
    WorkbookDocument,
)
from .diff import DiffResult, diff_documents, diff_files, format_text_diff, write_diff_excel
from .merge import (
    MergeChange,
    MergeConflictError,
    MergeResult,
    format_merge_text,
    merge_documents,
    merge_files,
    write_merge_excel,
)
from .operations import apply_operations
from .text import (
    TextColumn,
    TextRow,
    TextSheet,
    TextWorkbook,
    render_text_workbook,
    value_to_text,
    workbook_to_text,
)

__all__ = [
    "Cell",
    "Header",
    "Issue",
    "Row",
    "Sheet",
    "ValidationReport",
    "WorkbookDocument",
    "DiffResult",
    "diff_documents",
    "diff_files",
    "format_text_diff",
    "write_diff_excel",
    "MergeChange",
    "MergeConflictError",
    "MergeResult",
    "format_merge_text",
    "merge_documents",
    "merge_files",
    "write_merge_excel",
    "apply_operations",
    "TextColumn",
    "TextRow",
    "TextSheet",
    "TextWorkbook",
    "render_text_workbook",
    "value_to_text",
    "workbook_to_text",
]
