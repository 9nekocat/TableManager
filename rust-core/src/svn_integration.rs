use std::collections::HashSet;
use std::path::{Path, PathBuf};

use serde::Serialize;

use crate::error::{CoreError, Result};

#[derive(Debug, Clone, Serialize)]
pub struct SvnInstallResult {
    pub config_path: String,
    pub diff_cmd: String,
    pub merge_tool_cmd: String,
    pub changed: bool,
    pub dry_run: bool,
}

pub fn install_svn_integration(
    executable: impl AsRef<Path>,
    config_path: Option<&Path>,
    dry_run: bool,
) -> Result<SvnInstallResult> {
    let executable = executable.as_ref();
    let (diff_cmd, merge_tool_cmd) = find_svn_wrappers(executable)?;
    let config_path = config_path
        .map(Path::to_path_buf)
        .or_else(default_svn_config_path)
        .ok_or_else(|| {
            CoreError::Message(
                "无法确定 Subversion 用户配置路径；请通过 --config 指定 config 文件。".to_string(),
            )
        })?;
    let current = if config_path.exists() {
        std::fs::read_to_string(&config_path)?
    } else {
        String::new()
    };
    let updated = configure_svn_text(
        &current,
        &path_for_ini(&diff_cmd),
        &path_for_ini(&merge_tool_cmd),
    );
    let changed = updated != current;
    if changed && !dry_run {
        if let Some(parent) = config_path.parent() {
            std::fs::create_dir_all(parent)?;
        }
        if config_path.exists() {
            std::fs::copy(&config_path, config_path.with_extension("table-core.bak"))?;
        }
        std::fs::write(&config_path, updated)?;
    }
    Ok(SvnInstallResult {
        config_path: config_path.to_string_lossy().to_string(),
        diff_cmd: diff_cmd.to_string_lossy().to_string(),
        merge_tool_cmd: merge_tool_cmd.to_string_lossy().to_string(),
        changed,
        dry_run,
    })
}

pub fn find_svn_wrappers(executable: impl AsRef<Path>) -> Result<(PathBuf, PathBuf)> {
    if let Some(directory) = std::env::var_os("TABLE_CORE_SVN_DIR") {
        let directory = PathBuf::from(directory);
        if let Some(paths) = wrappers_in(&directory) {
            return Ok(paths);
        }
    }

    let executable = executable.as_ref();
    let executable_dir = executable
        .parent()
        .ok_or_else(|| CoreError::Message("无法确定 table-core.exe 所在目录。".to_string()))?;
    let mut directory = Some(executable_dir);
    for _ in 0..5 {
        let Some(current) = directory else {
            break;
        };
        if let Some(paths) = wrappers_in(&current.join("svn")) {
            return Ok(paths);
        }
        directory = current.parent();
    }
    Err(CoreError::Message(
        "未找到 svn\\svn-diff.cmd 和 svn\\svn-merge.cmd。请确认发布包完整，或设置 TABLE_CORE_SVN_DIR。"
            .to_string(),
    ))
}

fn wrappers_in(directory: &Path) -> Option<(PathBuf, PathBuf)> {
    let diff = directory.join("svn-diff.cmd");
    let merge = directory.join("svn-merge.cmd");
    (diff.is_file() && merge.is_file()).then_some((diff, merge))
}

fn default_svn_config_path() -> Option<PathBuf> {
    std::env::var_os("APPDATA")
        .map(PathBuf::from)
        .map(|path| path.join("Subversion").join("config"))
}

fn path_for_ini(path: &Path) -> String {
    path.to_string_lossy().replace('\\', "/")
}

fn configure_svn_text(text: &str, diff_cmd: &str, merge_tool_cmd: &str) -> String {
    let newline = if text.contains("\r\n") { "\r\n" } else { "\n" };
    let normalized = text.replace("\r\n", "\n");
    let mut lines = normalized
        .lines()
        .map(ToOwned::to_owned)
        .collect::<Vec<_>>();
    let mut section = String::new();
    let mut seen = HashSet::<&str>::new();
    let mut found_helpers = false;
    for line in &mut lines {
        let trimmed = line.trim();
        if trimmed.starts_with('[') && trimmed.ends_with(']') {
            section = trimmed[1..trimmed.len() - 1].trim().to_ascii_lowercase();
            if section == "helpers" {
                found_helpers = true;
            }
            continue;
        }
        if section != "helpers" {
            continue;
        }
        let Some((key, _)) = trimmed.split_once('=') else {
            continue;
        };
        let key = key.trim().to_ascii_lowercase();
        let replacement = match key.as_str() {
            "diff-cmd" => diff_cmd,
            "merge-tool-cmd" => merge_tool_cmd,
            "diff-extensions" => "",
            _ => continue,
        };
        if seen.insert(match key.as_str() {
            "diff-cmd" => "diff-cmd",
            "merge-tool-cmd" => "merge-tool-cmd",
            _ => "diff-extensions",
        }) {
            *line = format!("{key} = {replacement}");
        } else {
            *line = format!("# duplicate removed by table-core: {trimmed}");
        }
    }
    if !found_helpers {
        if !lines.is_empty() && lines.last().is_some_and(|line| !line.is_empty()) {
            lines.push(String::new());
        }
        lines.push("[helpers]".to_string());
    }
    for (key, value) in [
        ("diff-cmd", diff_cmd),
        ("diff-extensions", ""),
        ("merge-tool-cmd", merge_tool_cmd),
    ] {
        if !seen.contains(key) {
            lines.push(format!("{key} = {value}"));
        }
    }
    let mut result = lines.join(newline);
    if !result.ends_with(newline) {
        result.push_str(newline);
    }
    result
}

#[cfg(test)]
mod tests {
    use super::configure_svn_text;

    #[test]
    fn replaces_helper_commands_without_touching_other_sections() {
        let input = "[auth]\npassword = keep\n\n[helpers]\n# note\n diff-cmd = old.exe\ndiff-extensions = --x\nmerge-tool-cmd = old-merge.exe\ndiff-cmd = duplicate.exe\n";
        let output = configure_svn_text(
            input,
            "C:/Tools/table-core/svn/svn-diff.cmd",
            "C:/Tools/table-core/svn/svn-merge.cmd",
        );
        assert!(output.contains("password = keep"));
        assert!(output.contains("diff-cmd = C:/Tools/table-core/svn/svn-diff.cmd"));
        assert!(output.contains("diff-extensions = "));
        assert!(output.contains("merge-tool-cmd = C:/Tools/table-core/svn/svn-merge.cmd"));
        assert!(!output.contains("diff-cmd = old.exe"));
        assert_eq!(
            output
                .lines()
                .filter(|line| line.trim_start().starts_with("diff-cmd ="))
                .count(),
            1
        );
    }

    #[test]
    fn appends_helpers_section_when_missing() {
        let output = configure_svn_text("[auth]\npassword = keep\n", "C:/diff.cmd", "C:/merge.cmd");
        assert!(output.contains("[helpers]"));
        assert!(output.contains("diff-cmd = C:/diff.cmd"));
        assert!(output.contains("merge-tool-cmd = C:/merge.cmd"));
    }
}
