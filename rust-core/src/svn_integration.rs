use std::path::{Path, PathBuf};
use std::process::Command;

use crate::error::{CoreError, Result};

pub fn install_svn_integration(executable: impl AsRef<Path>, dry_run: bool) -> Result<()> {
    if !cfg!(windows) {
        return Err(CoreError::Message(
            "TortoiseSVN integration requires Windows.".to_string(),
        ));
    }

    let script = find_configure_script(executable.as_ref())?;
    let mut command = Command::new("powershell.exe");
    command
        .args(["-NoProfile", "-ExecutionPolicy", "Bypass", "-File"])
        .arg(script);
    if dry_run {
        command.arg("-DryRun");
    }
    let status = command.status()?;
    if !status.success() {
        return Err(CoreError::Message(format!(
            "TortoiseSVN configuration failed ({status})."
        )));
    }
    Ok(())
}

fn find_configure_script(executable: &Path) -> Result<PathBuf> {
    let override_directory = std::env::var_os("TABLE_CORE_SVN_DIR").map(PathBuf::from);
    let package_directories = executable
        .ancestors()
        .skip(1)
        .take(5)
        .map(|directory| directory.join("svn"));
    override_directory
        .into_iter()
        .chain(package_directories)
        .map(|directory| directory.join("configure-tortoisesvn.ps1"))
        .find(|script| script.is_file())
        .ok_or_else(|| {
            CoreError::Message(
                "未找到 svn\\configure-tortoisesvn.ps1。请确认发布包完整，或设置 TABLE_CORE_SVN_DIR。"
                    .to_string(),
            )
        })
}
