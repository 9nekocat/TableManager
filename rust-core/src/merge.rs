use std::collections::{BTreeMap, HashMap, HashSet};
use std::path::Path;

use serde::Serialize;

use crate::error::{CoreError, Result};
use crate::formula::{bind_formula, formula_signature};
use crate::model::{Cell, DATA_START_ROW, Header, Issue, Row, Sheet, Workbook, read_workbook};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum MergeStrategy {
    Fail,
    Ours,
    Theirs,
}

impl MergeStrategy {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Fail => "fail",
            Self::Ours => "ours",
            Self::Theirs => "theirs",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum ConflictDefault {
    Ours,
    Theirs,
    Empty,
}

impl ConflictDefault {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Ours => "ours",
            Self::Theirs => "theirs",
            Self::Empty => "empty",
        }
    }
}

fn deletion_choice_to_keep(
    strategy: MergeStrategy,
    conflict_default: ConflictDefault,
    missing_side: &str,
) -> bool {
    match strategy {
        MergeStrategy::Ours => missing_side != "ours",
        MergeStrategy::Theirs => missing_side != "theirs",
        MergeStrategy::Fail => match conflict_default {
            ConflictDefault::Ours => missing_side != "ours",
            ConflictDefault::Theirs | ConflictDefault::Empty => missing_side != "theirs",
        },
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Branch {
    Base,
    Ours,
    Theirs,
}

#[derive(Debug, Clone, Default, Serialize)]
pub struct MergeChange {
    pub sheet: String,
    pub kind: String,
    pub row: String,
    pub row_occurrence: usize,
    pub column: String,
    pub base: String,
    pub ours: String,
    pub theirs: String,
    pub result: String,
    pub conflict: bool,
    pub note: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct MergePreviewRow {
    pub sheet: String,
    pub row: String,
    pub occurrence: usize,
    pub kind: String,
    pub base: Vec<Cell>,
    pub ours: Vec<Cell>,
    pub theirs: Vec<Cell>,
    pub final_cells: Vec<Cell>,
    pub conflict_columns: Vec<usize>,
    pub default_decision: ConflictDefault,
    pub default_keep: bool,
}

#[derive(Debug, Clone, Serialize)]
pub struct MergeSummary {
    pub sheets: usize,
    pub rows: usize,
    pub columns: usize,
    pub changes: usize,
    pub conflicts: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct MergeResult {
    pub base: String,
    pub ours: String,
    pub theirs: String,
    pub strategy: MergeStrategy,
    pub changes: Vec<MergeChange>,
    pub issues: Vec<Issue>,
    #[serde(skip)]
    pub preview_rows: Vec<MergePreviewRow>,
    #[serde(skip)]
    pub workbook: Workbook,
}

impl MergeResult {
    pub fn conflicts(&self) -> impl Iterator<Item = &MergeChange> {
        self.changes.iter().filter(|change| change.conflict)
    }

    pub fn conflict_count(&self) -> usize {
        self.conflicts().count()
    }

    pub fn has_unresolved_conflicts(&self) -> bool {
        self.strategy == MergeStrategy::Fail && self.conflict_count() > 0
    }

    pub fn ok(&self) -> bool {
        !self.has_unresolved_conflicts()
    }

    pub fn summary(&self) -> MergeSummary {
        MergeSummary {
            sheets: self.workbook.sheets.len(),
            rows: self
                .workbook
                .sheets
                .iter()
                .map(|sheet| sheet.rows.len())
                .sum(),
            columns: self
                .workbook
                .sheets
                .iter()
                .map(|sheet| sheet.headers.len())
                .sum(),
            changes: self.changes.len(),
            conflicts: self.conflict_count(),
        }
    }

    pub fn require_resolved(&self) -> Result<()> {
        let conflicts = self.conflict_count();
        if conflicts > 0 {
            let sample = self
                .conflicts()
                .take(5)
                .map(|change| {
                    format!(
                        "{}!{}{}",
                        change.sheet,
                        change.row,
                        if change.column.is_empty() {
                            String::new()
                        } else {
                            format!(".{}", change.column)
                        }
                    )
                })
                .collect::<Vec<_>>()
                .join(", ");
            return Err(CoreError::Message(format!(
                "存在 {conflicts} 个未解决冲突: {sample}"
            )));
        }
        Ok(())
    }
}

#[derive(Debug, Clone)]
struct ColumnPlan {
    headers: Vec<Header>,
    base: Vec<Option<usize>>,
    ours: Vec<Option<usize>>,
    theirs: Vec<Option<usize>>,
    base_reverse: Vec<Option<usize>>,
    ours_reverse: Vec<Option<usize>>,
    theirs_reverse: Vec<Option<usize>>,
}

impl ColumnPlan {
    fn new(
        headers: Vec<Header>,
        base: Vec<Option<usize>>,
        ours: Vec<Option<usize>>,
        theirs: Vec<Option<usize>>,
    ) -> Self {
        let width = headers.len();
        Self {
            headers,
            base_reverse: reverse_map(&base, width),
            ours_reverse: reverse_map(&ours, width),
            theirs_reverse: reverse_map(&theirs, width),
            base,
            ours,
            theirs,
        }
    }

    fn map(&self, branch: Branch) -> &[Option<usize>] {
        match branch {
            Branch::Base => &self.base,
            Branch::Ours => &self.ours,
            Branch::Theirs => &self.theirs,
        }
    }

    fn reverse(&self, branch: Branch) -> &[Option<usize>] {
        match branch {
            Branch::Base => &self.base_reverse,
            Branch::Ours => &self.ours_reverse,
            Branch::Theirs => &self.theirs_reverse,
        }
    }
}

#[derive(Debug, Clone)]
struct PlannedCell {
    cell: Cell,
    branch: Branch,
    source_row: usize,
}

#[derive(Debug, Clone)]
struct PlannedRow {
    key: String,
    occurrence: usize,
    cells: Vec<Option<PlannedCell>>,
}

#[derive(Debug, Clone)]
struct OutputColumn {
    header: Header,
    base_index: Option<usize>,
    ours_index: Option<usize>,
    theirs_index: Option<usize>,
}

pub fn merge_files(
    base_path: impl AsRef<Path>,
    ours_path: impl AsRef<Path>,
    theirs_path: impl AsRef<Path>,
    strategy: MergeStrategy,
) -> Result<MergeResult> {
    merge_files_with_conflict_default(
        base_path,
        ours_path,
        theirs_path,
        strategy,
        ConflictDefault::Theirs,
    )
}

pub fn merge_files_with_conflict_default(
    base_path: impl AsRef<Path>,
    ours_path: impl AsRef<Path>,
    theirs_path: impl AsRef<Path>,
    strategy: MergeStrategy,
    conflict_default: ConflictDefault,
) -> Result<MergeResult> {
    let mut base = read_workbook(base_path)?;
    let mut ours = read_workbook(ours_path)?;
    let mut theirs = read_workbook(theirs_path)?;
    base.add_validation_issues();
    ours.add_validation_issues();
    theirs.add_validation_issues();
    let mut result =
        merge_workbooks_with_conflict_default(&base, &ours, &theirs, strategy, conflict_default);
    result.workbook.add_validation_issues();
    Ok(result)
}

pub fn merge_workbooks(
    base: &Workbook,
    ours: &Workbook,
    theirs: &Workbook,
    strategy: MergeStrategy,
) -> MergeResult {
    merge_workbooks_with_conflict_default(base, ours, theirs, strategy, ConflictDefault::Theirs)
}

pub fn merge_workbooks_with_conflict_default(
    base: &Workbook,
    ours: &Workbook,
    theirs: &Workbook,
    strategy: MergeStrategy,
    conflict_default: ConflictDefault,
) -> MergeResult {
    let mut merged = Workbook {
        path: "merged".to_string(),
        sheets: Vec::new(),
        issues: Vec::new(),
    };
    let mut changes = Vec::new();
    let mut sheet_names = Vec::new();
    for sheet in base
        .sheets
        .iter()
        .chain(ours.sheets.iter())
        .chain(theirs.sheets.iter())
    {
        if !sheet_names.contains(&sheet.name) {
            sheet_names.push(sheet.name.clone());
        }
    }
    for sheet_name in sheet_names {
        let base_sheet = base.sheet(&sheet_name);
        let ours_sheet = ours.sheet(&sheet_name);
        let theirs_sheet = theirs.sheet(&sheet_name);
        if let Some(sheet) = merge_sheet(
            &sheet_name,
            base_sheet,
            ours_sheet,
            theirs_sheet,
            strategy,
            conflict_default,
            &mut changes,
        ) {
            merged.sheets.push(sheet);
        }
    }
    let mut issues = base.issues.clone();
    issues.extend(ours.issues.clone());
    issues.extend(theirs.issues.clone());
    let mut result = MergeResult {
        base: base.path.clone(),
        ours: ours.path.clone(),
        theirs: theirs.path.clone(),
        strategy,
        changes,
        issues,
        preview_rows: Vec::new(),
        workbook: merged,
    };
    result.preview_rows = build_preview_rows(base, ours, theirs, &result, conflict_default);
    result
}

fn merge_sheet(
    name: &str,
    base: Option<&Sheet>,
    ours: Option<&Sheet>,
    theirs: Option<&Sheet>,
    strategy: MergeStrategy,
    conflict_default: ConflictDefault,
    changes: &mut Vec<MergeChange>,
) -> Option<Sheet> {
    match (base, ours, theirs) {
        (None, Some(ours), None) => {
            changes.push(MergeChange {
                sheet: name.to_string(),
                kind: "sheet_added".to_string(),
                note: "ours".to_string(),
                ..MergeChange::default()
            });
            Some(ours.clone())
        }
        (None, Some(ours), Some(theirs)) => Some(merge_existing_sheets(
            name,
            None,
            ours,
            theirs,
            strategy,
            conflict_default,
            changes,
        )),
        (None, None, Some(theirs)) => {
            changes.push(MergeChange {
                sheet: name.to_string(),
                kind: "sheet_added".to_string(),
                note: "theirs".to_string(),
                ..MergeChange::default()
            });
            Some(theirs.clone())
        }
        (None, None, None) => None,
        (Some(_base), None, None) => {
            changes.push(MergeChange {
                sheet: name.to_string(),
                kind: "sheet_deleted".to_string(),
                note: "both".to_string(),
                ..MergeChange::default()
            });
            None
        }
        (Some(base), Some(ours), Some(theirs)) => Some(merge_existing_sheets(
            name,
            Some(base),
            ours,
            theirs,
            strategy,
            conflict_default,
            changes,
        )),
        (Some(base), Some(present), None) => merge_sheet_delete(
            name,
            base,
            present,
            "theirs",
            strategy,
            conflict_default,
            changes,
        ),
        (Some(base), None, Some(present)) => merge_sheet_delete(
            name,
            base,
            present,
            "ours",
            strategy,
            conflict_default,
            changes,
        ),
    }
}

fn merge_sheet_delete(
    name: &str,
    base: &Sheet,
    present: &Sheet,
    missing_side: &str,
    strategy: MergeStrategy,
    conflict_default: ConflictDefault,
    changes: &mut Vec<MergeChange>,
) -> Option<Sheet> {
    if sheets_equal(base, present) {
        changes.push(MergeChange {
            sheet: name.to_string(),
            kind: "sheet_deleted".to_string(),
            note: missing_side.to_string(),
            ..MergeChange::default()
        });
        return None;
    }
    let keep = deletion_choice_to_keep(strategy, conflict_default, missing_side);
    changes.push(MergeChange {
        sheet: name.to_string(),
        kind: "sheet_delete_conflict".to_string(),
        conflict: true,
        result: if keep { "present" } else { "deleted" }.to_string(),
        note: format!("{missing_side} 删除，但另一侧修改了该 sheet"),
        ..MergeChange::default()
    });
    keep.then(|| present.clone())
}

fn merge_existing_sheets(
    name: &str,
    base: Option<&Sheet>,
    ours: &Sheet,
    theirs: &Sheet,
    strategy: MergeStrategy,
    conflict_default: ConflictDefault,
    changes: &mut Vec<MergeChange>,
) -> Sheet {
    let empty_base;
    let base = if let Some(base) = base {
        base
    } else {
        empty_base = Sheet {
            name: name.to_string(),
            headers: Vec::new(),
            rows: Vec::new(),
        };
        &empty_base
    };
    let (plan, _columns) = merge_columns(base, ours, theirs, strategy, conflict_default, changes);
    let rows = merge_rows(
        base,
        ours,
        theirs,
        &plan,
        strategy,
        conflict_default,
        changes,
    );
    Sheet {
        name: name.to_string(),
        headers: plan.headers,
        rows,
    }
}

fn reverse_map(map: &[Option<usize>], width: usize) -> Vec<Option<usize>> {
    let mut reverse = vec![None; width];
    for (source, target) in map.iter().enumerate() {
        if let Some(target) = *target
            && target < width
        {
            reverse[target] = Some(source);
        }
    }
    reverse
}

fn header_keys(headers: &[Header]) -> Vec<String> {
    let mut counts = HashMap::<String, usize>::new();
    headers
        .iter()
        .map(|header| {
            let base = header.base_key();
            let count = counts.entry(base.clone()).or_default();
            *count += 1;
            if *count == 1 {
                base
            } else {
                format!("{base}#{}", *count)
            }
        })
        .collect()
}

fn map_base_columns(
    base: &[Header],
    branch: &[Header],
) -> (Vec<Option<(usize, Header)>>, Vec<bool>) {
    let base_keys = header_keys(base);
    let branch_keys = header_keys(branch);
    let branch_pos: HashMap<&str, usize> = branch_keys
        .iter()
        .enumerate()
        .map(|(index, key)| (key.as_str(), index))
        .collect();
    let mut mapping = vec![None; base.len()];
    let mut used = vec![false; branch.len()];
    for (base_index, key) in base_keys.iter().enumerate() {
        if let Some(branch_index) = branch_pos.get(key.as_str()).copied() {
            mapping[base_index] = Some((branch_index, branch[branch_index].clone()));
            used[branch_index] = true;
        }
    }
    let unmatched_base: Vec<usize> = (0..base.len())
        .filter(|index| mapping[*index].is_none())
        .collect();
    let unmatched_branch: Vec<usize> = (0..branch.len()).filter(|index| !used[*index]).collect();
    for (base_index, branch_index) in unmatched_base.into_iter().zip(unmatched_branch.iter()) {
        mapping[base_index] = Some((*branch_index, branch[*branch_index].clone()));
        used[*branch_index] = true;
    }
    (mapping, used)
}

fn merge_columns(
    base: &Sheet,
    ours: &Sheet,
    theirs: &Sheet,
    strategy: MergeStrategy,
    conflict_default: ConflictDefault,
    changes: &mut Vec<MergeChange>,
) -> (ColumnPlan, Vec<OutputColumn>) {
    let (base_to_ours, _ours_used) = map_base_columns(&base.headers, &ours.headers);
    let (base_to_theirs, theirs_used) = map_base_columns(&base.headers, &theirs.headers);
    let mut specs: Vec<Option<OutputColumn>> = vec![None; base.headers.len()];
    for (base_index, base_header) in base.headers.iter().enumerate() {
        let ours_mapping = base_to_ours[base_index].clone();
        let theirs_mapping = base_to_theirs[base_index].clone();
        let ours_index = ours_mapping.as_ref().map(|item| item.0);
        let theirs_index = theirs_mapping.as_ref().map(|item| item.0);
        let ours_header = ours_mapping.as_ref().map(|item| item.1.clone());
        let theirs_header = theirs_mapping.as_ref().map(|item| item.1.clone());
        let ours_changed =
            ours_index.is_some_and(|index| column_changed(base, ours, base_index, index));
        let theirs_changed =
            theirs_index.is_some_and(|index| column_changed(base, theirs, base_index, index));

        if ours_index.is_none() && theirs_index.is_none() {
            changes.push(MergeChange {
                sheet: ours.name.clone(),
                kind: "column_deleted".to_string(),
                column: base_header.display_name(),
                base: base_header.display_name(),
                note: "both".to_string(),
                ..MergeChange::default()
            });
            continue;
        }
        if ours_index.is_none() || theirs_index.is_none() {
            let missing_side = if ours_index.is_none() {
                "ours"
            } else {
                "theirs"
            };
            let kept_changed = if missing_side == "ours" {
                theirs_changed
            } else {
                ours_changed
            };
            if !kept_changed {
                changes.push(MergeChange {
                    sheet: ours.name.clone(),
                    kind: "column_deleted".to_string(),
                    column: base_header.display_name(),
                    base: base_header.display_name(),
                    note: missing_side.to_string(),
                    ..MergeChange::default()
                });
                continue;
            }
            let keep = deletion_choice_to_keep(strategy, conflict_default, missing_side);
            let result_header = if missing_side == "ours" {
                theirs_header.clone().unwrap_or_else(|| base_header.clone())
            } else {
                ours_header.clone().unwrap_or_else(|| base_header.clone())
            };
            changes.push(MergeChange {
                sheet: ours.name.clone(),
                kind: "column_delete_conflict".to_string(),
                row: String::new(),
                row_occurrence: 0,
                column: base_header.display_name(),
                base: base_header.display_name(),
                ours: ours_header
                    .as_ref()
                    .map(Header::display_name)
                    .unwrap_or_default(),
                theirs: theirs_header
                    .as_ref()
                    .map(Header::display_name)
                    .unwrap_or_default(),
                result: if keep {
                    result_header.display_name()
                } else {
                    String::new()
                },
                conflict: true,
                note: format!("{missing_side} 删除，但另一侧修改了该列"),
            });
            if keep {
                specs[base_index] = Some(OutputColumn {
                    header: result_header,
                    base_index: Some(base_index),
                    ours_index,
                    theirs_index,
                });
            }
            continue;
        }

        let merged_header = merge_header(
            &ours.name,
            base_header,
            ours_header.as_ref().unwrap(),
            theirs_header.as_ref().unwrap(),
            strategy,
            conflict_default,
            changes,
        );
        specs[base_index] = Some(OutputColumn {
            header: merged_header,
            base_index: Some(base_index),
            ours_index,
            theirs_index,
        });
    }

    let ours_to_base: HashMap<usize, usize> = base_to_ours
        .iter()
        .enumerate()
        .filter_map(|(base, branch)| branch.as_ref().map(|item| (item.0, base)))
        .collect();
    let theirs_to_base: HashMap<usize, usize> = base_to_theirs
        .iter()
        .enumerate()
        .filter_map(|(base, branch)| branch.as_ref().map(|item| (item.0, base)))
        .collect();
    let mut output = Vec::<OutputColumn>::new();
    for (ours_index, header) in ours.headers.iter().enumerate() {
        if let Some(base_index) = ours_to_base.get(&ours_index).copied() {
            if let Some(spec) = specs[base_index].take() {
                output.push(spec);
            }
        } else {
            output.push(OutputColumn {
                header: header.clone(),
                base_index: None,
                ours_index: Some(ours_index),
                theirs_index: None,
            });
            changes.push(MergeChange {
                sheet: ours.name.clone(),
                kind: "column_added".to_string(),
                column: header.display_name(),
                ours: header.display_name(),
                result: header.display_name(),
                note: "ours".to_string(),
                ..MergeChange::default()
            });
        }
    }
    for spec in specs.into_iter().flatten() {
        output.push(spec);
    }
    for (theirs_index, header) in theirs.headers.iter().enumerate() {
        if theirs_used[theirs_index] || theirs_to_base.contains_key(&theirs_index) {
            continue;
        }
        let key = header.base_key();
        if let Some(existing_index) = output
            .iter()
            .position(|column| column.header.base_key() == key)
        {
            output[existing_index].theirs_index = Some(theirs_index);
            if output[existing_index].header.identity() != header.identity() {
                let chosen = choose_header(
                    strategy,
                    conflict_default,
                    &output[existing_index].header,
                    header,
                );
                changes.push(MergeChange {
                    sheet: ours.name.clone(),
                    kind: "column_header_conflict".to_string(),
                    row: String::new(),
                    row_occurrence: 0,
                    column: header.display_name(),
                    base: String::new(),
                    ours: output[existing_index].header.display_name(),
                    theirs: header.display_name(),
                    result: chosen.display_name(),
                    conflict: true,
                    note: "双方新增同名列但表头不同".to_string(),
                });
                output[existing_index].header = chosen;
            }
        } else {
            output.push(OutputColumn {
                header: header.clone(),
                base_index: None,
                ours_index: None,
                theirs_index: Some(theirs_index),
            });
            changes.push(MergeChange {
                sheet: ours.name.clone(),
                kind: "column_added".to_string(),
                column: header.display_name(),
                theirs: header.display_name(),
                result: header.display_name(),
                note: "theirs".to_string(),
                ..MergeChange::default()
            });
        }
    }

    let headers = output.iter().map(|column| column.header.clone()).collect();
    let mut base_map = vec![None; base.headers.len()];
    let mut ours_map = vec![None; ours.headers.len()];
    let mut theirs_map = vec![None; theirs.headers.len()];
    for (output_index, column) in output.iter().enumerate() {
        if let Some(index) = column.base_index {
            base_map[index] = Some(output_index);
        }
        if let Some(index) = column.ours_index {
            ours_map[index] = Some(output_index);
        }
        if let Some(index) = column.theirs_index {
            theirs_map[index] = Some(output_index);
        }
    }
    (
        ColumnPlan::new(headers, base_map, ours_map, theirs_map),
        output,
    )
}

fn choose_header(
    strategy: MergeStrategy,
    conflict_default: ConflictDefault,
    ours: &Header,
    theirs: &Header,
) -> Header {
    match strategy {
        MergeStrategy::Theirs => theirs.clone(),
        MergeStrategy::Ours => ours.clone(),
        MergeStrategy::Fail => match conflict_default {
            ConflictDefault::Ours => ours.clone(),
            ConflictDefault::Theirs | ConflictDefault::Empty => theirs.clone(),
        },
    }
}

fn merge_header(
    sheet: &str,
    base: &Header,
    ours: &Header,
    theirs: &Header,
    strategy: MergeStrategy,
    conflict_default: ConflictDefault,
    changes: &mut Vec<MergeChange>,
) -> Header {
    if ours.identity() == theirs.identity() {
        return ours.clone();
    }
    if ours.identity() == base.identity() {
        return theirs.clone();
    }
    if theirs.identity() == base.identity() {
        return ours.clone();
    }
    let chosen = choose_header(strategy, conflict_default, ours, theirs);
    changes.push(MergeChange {
        sheet: sheet.to_string(),
        kind: "column_header_conflict".to_string(),
        row: String::new(),
        row_occurrence: 0,
        column: base.display_name(),
        base: base.display_name(),
        ours: ours.display_name(),
        theirs: theirs.display_name(),
        result: chosen.display_name(),
        conflict: true,
        note: format!("表头双方均修改；按 {} 处理", strategy.as_str()),
    });
    chosen
}

fn column_changed(base: &Sheet, branch: &Sheet, base_index: usize, branch_index: usize) -> bool {
    if base.headers[base_index].identity() != branch.headers[branch_index].identity() {
        return true;
    }
    let base_rows = row_map(base);
    let branch_rows = row_map(branch);
    let identities: HashSet<_> = base_rows
        .keys()
        .chain(branch_rows.keys())
        .cloned()
        .collect();
    for identity in identities {
        let base_row = base_rows
            .get(&identity)
            .and_then(|index| base.rows.get(*index));
        let branch_row = branch_rows
            .get(&identity)
            .and_then(|index| branch.rows.get(*index));
        match (base_row, branch_row) {
            (Some(base_row), Some(branch_row)) => {
                let base_cell = base_row.cells.get(base_index).cloned().unwrap_or_default();
                let branch_cell = branch_row
                    .cells
                    .get(branch_index)
                    .cloned()
                    .unwrap_or_default();
                if !cell_equal(
                    &base_cell,
                    &branch_cell,
                    base_row.source_row.unwrap_or(DATA_START_ROW),
                    branch_row.source_row.unwrap_or(DATA_START_ROW),
                ) {
                    return true;
                }
            }
            _ => return true,
        }
    }
    false
}

fn row_map(sheet: &Sheet) -> HashMap<(String, usize), usize> {
    sheet
        .rows
        .iter()
        .enumerate()
        .map(|(index, row)| ((row.key.clone(), row.occurrence), index))
        .collect()
}

fn cell_equal(left: &Cell, right: &Cell, left_row: usize, right_row: usize) -> bool {
    if left.value != right.value && (!left.value.is_empty() || !right.value.is_empty()) {
        return false;
    }
    if left.formula == right.formula
        && left.formula.is_some()
        && left_row == right_row
        && left.value == right.value
    {
        return true;
    }
    if left.formula.is_some() || right.formula.is_some() {
        match (left.formula.as_deref(), right.formula.as_deref()) {
            (Some(left), Some(right)) => {
                formula_signature(left, left_row) == formula_signature(right, right_row)
            }
            _ => false,
        }
    } else {
        left.value == right.value
    }
}

fn sheets_equal(left: &Sheet, right: &Sheet) -> bool {
    if left
        .headers
        .iter()
        .map(Header::identity)
        .collect::<Vec<_>>()
        != right
            .headers
            .iter()
            .map(Header::identity)
            .collect::<Vec<_>>()
    {
        return false;
    }
    if left.rows.len() != right.rows.len() {
        return false;
    }
    let right_rows = row_map(right);
    for left_row in &left.rows {
        let Some(right_index) = right_rows.get(&(left_row.key.clone(), left_row.occurrence)) else {
            return false;
        };
        let right_row = &right.rows[*right_index];
        if left_row.cells.len() != right_row.cells.len() {
            return false;
        }
        for (index, left_cell) in left_row.cells.iter().enumerate() {
            if !cell_equal(
                left_cell,
                &right_row.cells[index],
                left_row.source_row.unwrap_or(DATA_START_ROW),
                right_row.source_row.unwrap_or(DATA_START_ROW),
            ) {
                return false;
            }
        }
    }
    true
}

fn row_changed(base_row: &Row, branch_row: &Row, plan: &ColumnPlan, branch_kind: Branch) -> bool {
    for output_index in 0..plan.headers.len() {
        let Some(base_index) = plan.base_reverse[output_index] else {
            continue;
        };
        let Some(branch_index) = plan.reverse(branch_kind)[output_index] else {
            continue;
        };
        let base_cell = base_row.cells.get(base_index).cloned().unwrap_or_default();
        let branch_cell = branch_row
            .cells
            .get(branch_index)
            .cloned()
            .unwrap_or_default();
        if !cell_equal(
            &base_cell,
            &branch_cell,
            base_row.source_row.unwrap_or(DATA_START_ROW),
            branch_row.source_row.unwrap_or(DATA_START_ROW),
        ) {
            return true;
        }
    }
    false
}

fn merge_rows(
    base: &Sheet,
    ours: &Sheet,
    theirs: &Sheet,
    plan: &ColumnPlan,
    strategy: MergeStrategy,
    conflict_default: ConflictDefault,
    changes: &mut Vec<MergeChange>,
) -> Vec<Row> {
    let base_rows = row_map(base);
    let ours_rows = row_map(ours);
    let theirs_rows = row_map(theirs);
    let mut identities = Vec::new();
    for row in ours
        .rows
        .iter()
        .chain(theirs.rows.iter())
        .chain(base.rows.iter())
    {
        let identity = (row.key.clone(), row.occurrence);
        if !identities.contains(&identity) {
            identities.push(identity);
        }
    }
    let mut planned: HashMap<(String, usize), PlannedRow> = HashMap::new();
    for identity in identities {
        let base_index = base_rows.get(&identity).copied();
        let ours_index = ours_rows.get(&identity).copied();
        let theirs_index = theirs_rows.get(&identity).copied();
        match (base_index, ours_index, theirs_index) {
            (None, Some(ours_index), None) => {
                planned.insert(
                    identity.clone(),
                    branch_row(ours, ours_index, plan, Branch::Ours),
                );
                changes.push(MergeChange {
                    sheet: ours.name.clone(),
                    kind: "row_added".to_string(),
                    row: identity.0,
                    note: "ours".to_string(),
                    ..MergeChange::default()
                });
            }
            (None, None, Some(theirs_index)) => {
                planned.insert(
                    identity.clone(),
                    branch_row(theirs, theirs_index, plan, Branch::Theirs),
                );
                changes.push(MergeChange {
                    sheet: ours.name.clone(),
                    kind: "row_added".to_string(),
                    row: identity.0,
                    note: "theirs".to_string(),
                    ..MergeChange::default()
                });
            }
            (None, Some(ours_index), Some(theirs_index)) => {
                let row = merge_existing_rows(
                    None,
                    Some(&ours.rows[ours_index]),
                    Some(&theirs.rows[theirs_index]),
                    ours,
                    plan,
                    strategy,
                    conflict_default,
                    &identity,
                    changes,
                );
                planned.insert(identity, row);
            }
            (Some(_), None, None) => {
                changes.push(MergeChange {
                    sheet: ours.name.clone(),
                    kind: "row_deleted".to_string(),
                    row: identity.0,
                    note: "both".to_string(),
                    ..MergeChange::default()
                });
            }
            (Some(base_index), Some(ours_index), Some(theirs_index)) => {
                let row = merge_existing_rows(
                    Some(&base.rows[base_index]),
                    Some(&ours.rows[ours_index]),
                    Some(&theirs.rows[theirs_index]),
                    ours,
                    plan,
                    strategy,
                    conflict_default,
                    &identity,
                    changes,
                );
                planned.insert(identity, row);
            }
            (Some(base_index), Some(ours_index), None) => {
                let base_row = &base.rows[base_index];
                let ours_row = &ours.rows[ours_index];
                if !row_changed(base_row, ours_row, plan, Branch::Ours) {
                    changes.push(MergeChange {
                        sheet: ours.name.clone(),
                        kind: "row_deleted".to_string(),
                        row: identity.0,
                        note: "theirs".to_string(),
                        ..MergeChange::default()
                    });
                } else {
                    let keep = deletion_choice_to_keep(strategy, conflict_default, "theirs");
                    changes.push(MergeChange {
                        sheet: ours.name.clone(),
                        kind: "row_delete_conflict".to_string(),
                        row: identity.0.clone(),
                        row_occurrence: identity.1,
                        result: if keep { "ours" } else { "deleted" }.to_string(),
                        conflict: true,
                        note: "仓库删除，但本地修改了该行".to_string(),
                        ..MergeChange::default()
                    });
                    if keep {
                        planned.insert(identity, branch_row(ours, ours_index, plan, Branch::Ours));
                    }
                }
            }
            (Some(base_index), None, Some(theirs_index)) => {
                let base_row = &base.rows[base_index];
                let theirs_row = &theirs.rows[theirs_index];
                if !row_changed(base_row, theirs_row, plan, Branch::Theirs) {
                    changes.push(MergeChange {
                        sheet: ours.name.clone(),
                        kind: "row_deleted".to_string(),
                        row: identity.0,
                        note: "ours".to_string(),
                        ..MergeChange::default()
                    });
                } else {
                    let keep = deletion_choice_to_keep(strategy, conflict_default, "ours");
                    changes.push(MergeChange {
                        sheet: ours.name.clone(),
                        kind: "row_delete_conflict".to_string(),
                        row: identity.0.clone(),
                        row_occurrence: identity.1,
                        result: if keep { "theirs" } else { "deleted" }.to_string(),
                        conflict: true,
                        note: "本地删除，但仓库修改了该行".to_string(),
                        ..MergeChange::default()
                    });
                    if keep {
                        planned.insert(
                            identity,
                            branch_row(theirs, theirs_index, plan, Branch::Theirs),
                        );
                    }
                }
            }
            (None, None, None) => {}
        }
    }

    let order = order_rows(ours, theirs, base, &planned);
    order
        .into_iter()
        .enumerate()
        .map(|(output_index, identity)| {
            let target_row = DATA_START_ROW + output_index;
            let planned_row = planned.remove(&identity).expect("ordered row exists");
            let mut cells = Vec::with_capacity(plan.headers.len());
            for cell in planned_row.cells {
                let mut cell = cell.unwrap_or_else(|| PlannedCell {
                    cell: Cell::default(),
                    branch: Branch::Base,
                    source_row: DATA_START_ROW,
                });
                if let Some(formula) = cell.cell.formula.as_mut() {
                    let bound = bind_formula(
                        formula,
                        cell.source_row,
                        target_row,
                        Some(plan.map(cell.branch)),
                    );
                    *formula = bound
                        .strip_prefix('=')
                        .unwrap_or(bound.as_str())
                        .to_string();
                }
                cells.push(cell.cell);
            }
            Row {
                key: planned_row.key,
                occurrence: planned_row.occurrence,
                cells,
                source_row: Some(target_row),
            }
        })
        .collect()
}

fn branch_row(sheet: &Sheet, row_index: usize, plan: &ColumnPlan, branch: Branch) -> PlannedRow {
    let row = &sheet.rows[row_index];
    let mut cells: Vec<Option<PlannedCell>> = vec![None; plan.headers.len()];
    for (source_index, cell) in row.cells.iter().enumerate() {
        if let Some(Some(output_index)) = plan.map(branch).get(source_index) {
            cells[*output_index] = Some(PlannedCell {
                cell: cell.clone(),
                branch,
                source_row: row.source_row.unwrap_or(DATA_START_ROW),
            });
        }
    }
    PlannedRow {
        key: row.key.clone(),
        occurrence: row.occurrence,
        cells,
    }
}

#[allow(clippy::too_many_arguments)]
fn merge_existing_rows(
    base: Option<&Row>,
    ours: Option<&Row>,
    theirs: Option<&Row>,
    sheet: &Sheet,
    plan: &ColumnPlan,
    strategy: MergeStrategy,
    conflict_default: ConflictDefault,
    identity: &(String, usize),
    changes: &mut Vec<MergeChange>,
) -> PlannedRow {
    let mut cells = Vec::with_capacity(plan.headers.len());
    for output_index in 0..plan.headers.len() {
        let base_index = plan.base_reverse[output_index];
        let ours_index = plan.reverse(Branch::Ours)[output_index];
        let theirs_index = plan.reverse(Branch::Theirs)[output_index];
        let base_cell = base
            .and_then(|row| base_index.and_then(|index| row.cells.get(index)))
            .cloned()
            .unwrap_or_default();
        let ours_cell = ours
            .and_then(|row| ours_index.and_then(|index| row.cells.get(index)))
            .cloned();
        let theirs_cell = theirs
            .and_then(|row| theirs_index.and_then(|index| row.cells.get(index)))
            .cloned();
        let planned = merge_cell(
            &sheet.name,
            &identity.0,
            identity.1,
            &plan.headers[output_index],
            &base_cell,
            ours_cell.as_ref(),
            theirs_cell.as_ref(),
            base.and_then(|row| row.source_row),
            ours.and_then(|row| row.source_row),
            theirs.and_then(|row| row.source_row),
            strategy,
            conflict_default,
            changes,
        );
        cells.push(planned);
    }
    PlannedRow {
        key: identity.0.clone(),
        occurrence: identity.1,
        cells,
    }
}

#[allow(clippy::too_many_arguments)]
fn merge_cell(
    sheet: &str,
    row_key: &str,
    row_occurrence: usize,
    header: &Header,
    base: &Cell,
    ours: Option<&Cell>,
    theirs: Option<&Cell>,
    base_row: Option<usize>,
    ours_row: Option<usize>,
    theirs_row: Option<usize>,
    strategy: MergeStrategy,
    conflict_default: ConflictDefault,
    changes: &mut Vec<MergeChange>,
) -> Option<PlannedCell> {
    let empty = Cell::default();
    let ours_cell = ours.unwrap_or(&empty);
    let theirs_cell = theirs.unwrap_or(&empty);
    let base_y = base_row.unwrap_or(DATA_START_ROW);
    let ours_y = ours_row.unwrap_or(DATA_START_ROW);
    let theirs_y = theirs_row.unwrap_or(DATA_START_ROW);
    let ours_equal_theirs = cell_equal(ours_cell, theirs_cell, ours_y, theirs_y);
    let ours_equal_base = cell_equal(ours_cell, base, ours_y, base_y);
    let theirs_equal_base = cell_equal(theirs_cell, base, theirs_y, base_y);
    if ours_equal_theirs {
        let (cell, branch, source_row) = if !ours_cell.is_empty() {
            (ours_cell, Branch::Ours, ours_y)
        } else if !theirs_cell.is_empty() {
            (theirs_cell, Branch::Theirs, theirs_y)
        } else {
            (&empty, Branch::Base, base_y)
        };
        if !ours_equal_base {
            changes.push(MergeChange {
                sheet: sheet.to_string(),
                kind: "cell_modified".to_string(),
                row: row_key.to_string(),
                row_occurrence,
                column: header.display_name(),
                base: display_cell(base),
                ours: display_cell(ours_cell),
                theirs: display_cell(theirs_cell),
                result: display_cell(cell),
                ..MergeChange::default()
            });
        }
        return (!cell.is_empty()).then(|| PlannedCell {
            cell: cell.clone(),
            branch,
            source_row,
        });
    }
    if ours_equal_base {
        if !theirs_equal_base {
            changes.push(MergeChange {
                sheet: sheet.to_string(),
                kind: "cell_modified".to_string(),
                row: row_key.to_string(),
                row_occurrence,
                column: header.display_name(),
                base: display_cell(base),
                ours: display_cell(ours_cell),
                theirs: display_cell(theirs_cell),
                result: display_cell(theirs_cell),
                note: "theirs".to_string(),
                ..MergeChange::default()
            });
        }
        return (!theirs_cell.is_empty()).then(|| PlannedCell {
            cell: theirs_cell.clone(),
            branch: Branch::Theirs,
            source_row: theirs_y,
        });
    }
    if theirs_equal_base {
        changes.push(MergeChange {
            sheet: sheet.to_string(),
            kind: "cell_modified".to_string(),
            row: row_key.to_string(),
            row_occurrence,
            column: header.display_name(),
            base: display_cell(base),
            ours: display_cell(ours_cell),
            theirs: display_cell(theirs_cell),
            result: display_cell(ours_cell),
            note: "ours".to_string(),
            ..MergeChange::default()
        });
        return (!ours_cell.is_empty()).then(|| PlannedCell {
            cell: ours_cell.clone(),
            branch: Branch::Ours,
            source_row: ours_y,
        });
    }
    let chosen = match strategy {
        MergeStrategy::Theirs => (theirs_cell, Branch::Theirs, theirs_y, "theirs"),
        MergeStrategy::Ours => (ours_cell, Branch::Ours, ours_y, "ours"),
        MergeStrategy::Fail => match conflict_default {
            ConflictDefault::Ours => (ours_cell, Branch::Ours, ours_y, "ours"),
            ConflictDefault::Theirs => (theirs_cell, Branch::Theirs, theirs_y, "theirs"),
            ConflictDefault::Empty => (&empty, Branch::Base, base_y, "empty"),
        },
    };
    changes.push(MergeChange {
        sheet: sheet.to_string(),
        kind: "cell_conflict".to_string(),
        row: row_key.to_string(),
        row_occurrence,
        column: header.display_name(),
        base: display_cell(base),
        ours: display_cell(ours_cell),
        theirs: display_cell(theirs_cell),
        result: display_cell(chosen.0),
        conflict: true,
        note: format!("按 {} 选择 {}", strategy.as_str(), chosen.3),
    });
    (!chosen.0.is_empty()).then(|| PlannedCell {
        cell: chosen.0.clone(),
        branch: chosen.1,
        source_row: chosen.2,
    })
}

fn display_cell(cell: &Cell) -> String {
    match (cell.formula.as_deref(), cell.value.is_empty()) {
        (Some(formula), false) => format!("={formula} -> {}", cell.value),
        (Some(formula), true) => format!("={formula}"),
        (None, _) => cell.value.clone(),
    }
}

fn order_rows(
    ours: &Sheet,
    theirs: &Sheet,
    base: &Sheet,
    planned: &HashMap<(String, usize), PlannedRow>,
) -> Vec<(String, usize)> {
    let mut order = Vec::new();
    for row in &ours.rows {
        let identity = (row.key.clone(), row.occurrence);
        if planned.contains_key(&identity) && !order.contains(&identity) {
            order.push(identity);
        }
    }
    let mut seen = order.iter().cloned().collect::<HashSet<_>>();
    let mut inserted_after: HashMap<(String, usize), usize> = HashMap::new();
    let theirs_ids: Vec<_> = theirs
        .rows
        .iter()
        .map(|row| (row.key.clone(), row.occurrence))
        .collect();
    for (index, identity) in theirs_ids.iter().enumerate() {
        if seen.contains(identity) || !planned.contains_key(identity) {
            continue;
        }
        if let Some(predecessor) = theirs_ids[..index]
            .iter()
            .rev()
            .find(|candidate| seen.contains(*candidate))
        {
            let base_position = order
                .iter()
                .position(|candidate| candidate == predecessor)
                .expect("predecessor exists");
            let offset = inserted_after.get(predecessor).copied().unwrap_or(0);
            let position = base_position + 1 + offset;
            order.insert(position, identity.clone());
            inserted_after.insert(predecessor.clone(), offset + 1);
        } else if let Some(successor) = theirs_ids[index + 1..]
            .iter()
            .find(|candidate| seen.contains(*candidate))
        {
            let position = order
                .iter()
                .position(|candidate| candidate == successor)
                .expect("successor exists");
            order.insert(position, identity.clone());
        } else {
            order.push(identity.clone());
        }
        seen.insert(identity.clone());
    }
    for row in &base.rows {
        let identity = (row.key.clone(), row.occurrence);
        if planned.contains_key(&identity) && !seen.contains(&identity) {
            order.push(identity.clone());
            seen.insert(identity);
        }
    }
    for identity in planned.keys() {
        if !seen.contains(identity) {
            order.push(identity.clone());
        }
    }
    order
}

fn build_preview_rows(
    base: &Workbook,
    ours: &Workbook,
    theirs: &Workbook,
    result: &MergeResult,
    conflict_default: ConflictDefault,
) -> Vec<MergePreviewRow> {
    #[derive(Default)]
    struct Group {
        kind: String,
        conflicts: Vec<MergeChange>,
    }

    let mut groups = BTreeMap::<(String, String, usize), Group>::new();
    for change in result.conflicts() {
        if !matches!(
            change.kind.as_str(),
            "cell_conflict" | "row_delete_conflict"
        ) {
            continue;
        }
        let key = (
            change.sheet.clone(),
            change.row.clone(),
            change.row_occurrence,
        );
        let group = groups.entry(key).or_default();
        if group.kind.is_empty() || change.kind == "row_delete_conflict" {
            group.kind = change.kind.clone();
        }
        group.conflicts.push(change.clone());
    }

    groups
        .into_iter()
        .map(|((sheet_name, row_key, occurrence), group)| {
            let target_headers = result
                .workbook
                .sheet(&sheet_name)
                .map(|sheet| sheet.headers.as_slice())
                .unwrap_or(&[]);
            let base_sheet = base.sheet(&sheet_name);
            let ours_sheet = ours.sheet(&sheet_name);
            let theirs_sheet = theirs.sheet(&sheet_name);
            let base_row = find_row(base_sheet, &row_key, occurrence);
            let ours_row = find_row(ours_sheet, &row_key, occurrence);
            let theirs_row = find_row(theirs_sheet, &row_key, occurrence);
            let merged_sheet = result.workbook.sheet(&sheet_name);
            let merged_row = find_row(merged_sheet, &row_key, occurrence);
            let base_cells = align_row_cells(base_sheet, base_row, target_headers);
            let ours_cells = align_row_cells(ours_sheet, ours_row, target_headers);
            let theirs_cells = align_row_cells(theirs_sheet, theirs_row, target_headers);
            let final_cells = merged_row.map(|row| row.cells.clone()).unwrap_or_else(|| {
                if ours_row.is_some() {
                    ours_cells.clone()
                } else {
                    theirs_cells.clone()
                }
            });
            let conflict_columns = group
                .conflicts
                .iter()
                .filter(|change| change.kind == "cell_conflict")
                .filter_map(|change| {
                    target_headers
                        .iter()
                        .position(|header| header.display_name() == change.column)
                })
                .collect::<Vec<_>>();
            let default_keep = merged_row.is_some();
            MergePreviewRow {
                sheet: sheet_name,
                row: row_key,
                occurrence,
                kind: group.kind,
                base: base_cells,
                ours: ours_cells,
                theirs: theirs_cells,
                final_cells,
                conflict_columns,
                default_decision: conflict_default,
                default_keep,
            }
        })
        .collect()
}

fn find_row<'a>(sheet: Option<&'a Sheet>, key: &str, occurrence: usize) -> Option<&'a Row> {
    sheet?
        .rows
        .iter()
        .find(|row| row.key == key && row.occurrence == occurrence)
}

fn align_row_cells(
    source_sheet: Option<&Sheet>,
    source_row: Option<&Row>,
    target_headers: &[Header],
) -> Vec<Cell> {
    let mut cells = vec![Cell::default(); target_headers.len()];
    let (Some(sheet), Some(row)) = (source_sheet, source_row) else {
        return cells;
    };
    let (mapping, _) = map_base_columns(&sheet.headers, target_headers);
    for (source_index, target_index) in mapping.into_iter().enumerate() {
        if let Some(target_index) = target_index {
            cells[target_index.0] = row.cells.get(source_index).cloned().unwrap_or_default();
        }
    }
    cells
}

pub fn format_merge_text(result: &MergeResult) -> String {
    let summary = result.summary();
    let mut lines = vec![
        format!("merge strategy: {}", result.strategy.as_str()),
        format!("changes: {}", summary.changes),
        format!("conflicts: {}", summary.conflicts),
    ];
    for change in &result.changes {
        lines.push(format!(
            "{} {} | {}!{}{} | base={:?} ours={:?} theirs={:?} result={:?}{}",
            if change.conflict { "!" } else { "~" },
            change.kind,
            change.sheet,
            change.row,
            if change.column.is_empty() {
                String::new()
            } else {
                format!(".{}", change.column)
            },
            change.base,
            change.ours,
            change.theirs,
            change.result,
            if change.note.is_empty() {
                String::new()
            } else {
                format!(" note={}", change.note)
            }
        ));
    }
    if !result.issues.is_empty() {
        lines.push(String::new());
        lines.push("issues:".to_string());
        for issue in &result.issues {
            lines.push(format!(
                "- [{}] {}: {}",
                issue.severity, issue.code, issue.message
            ));
        }
    }
    lines.join("\n")
}
