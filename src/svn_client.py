# -*- coding: utf-8 -*-
"""
SVN subprocess wrapper for DataTable Manager.
Handles svn status, update, and conflict file detection.
Auto-detects svn.exe from common install locations if not in PATH.
"""

import os
import locale
import re
import subprocess
import threading
from pathlib import Path

from manager_core import DATA_DIR, RAW_DIR

# Suppress console windows on Windows for all subprocess calls
_CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000) if os.name == "nt" else 0

# SVN on Windows may output in the system ANSI code page (GBK on Chinese
# Windows) when not forced to the C locale; _decode_svn tries UTF-8 first,
# then GBK-family encodings.

def _decode_svn(raw_bytes):
    """Decode SVN output bytes. UTF-8 strict first (LANG=C forces English
    output, which is valid UTF-8); fall back to GBK-family encodings for
    localized Windows SVN builds."""
    if not raw_bytes:
        return ""
    if os.name == "nt":
        for enc in ["utf-8", "gbk", "cp936", "gb2312"]:
            try:
                return raw_bytes.decode(enc)
            except (UnicodeDecodeError, LookupError):
                continue
        return raw_bytes.decode("utf-8", errors="replace")
    return raw_bytes.decode("utf-8", errors="replace")


def _svn_env():
    """Environment for svn subprocesses: force English (C locale) output so
    keyword parsing works regardless of the OS display language."""
    env = os.environ.copy()
    env["LANG"] = "C"
    env["LC_MESSAGES"] = "C"
    return env


# Common SVN install paths on Windows (fallback)
_SVN_SEARCH_PATHS = [
    r"C:\Program Files\TortoiseSVN\bin\svn.exe",
    r"C:\Program Files (x86)\TortoiseSVN\bin\svn.exe",
    r"C:\Program Files\SlikSvn\bin\svn.exe",
    r"C:\Program Files\VisualSVN\bin\svn.exe",
    r"C:\Program Files\Subversion\bin\svn.exe",
    r"C:\Program Files (x86)\Subversion\bin\svn.exe",
    r"C:\Program Files\CollabNet\Subversion Client\svn.exe",
    r"C:\Program Files\Apache-Subversion\bin\svn.exe",
]

_SVN_EXE = "svn"  # default, works if in PATH
_SVN_INSTALL_ATTEMPTED = False
_SVN_INSTALL_RESULT = None
_SVN_INSTALL_LOCK = threading.Lock()


def _find_svn():
    """Find svn.exe. Checks: 'where svn' + registry + common paths."""
    global _SVN_EXE

    # Method 1: 'where svn' (Windows, scans PATH)
    try:
        r = subprocess.run(
            ["where", "svn"],
            capture_output=True, timeout=5,
            creationflags=_CREATE_NO_WINDOW,
        )
        if r.returncode == 0 and r.stdout.strip():
            first = _decode_svn(r.stdout).strip().splitlines()[0].strip()
            if os.path.isfile(first):
                _SVN_EXE = first
                return True
    except Exception:
        pass

    # Method 2: Registry lookup (TortoiseSVN stores install dir)
    try:
        import winreg
        for root_key in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for subkey in (
                r"Software\TortoiseSVN",
                r"Software\WOW6432Node\TortoiseSVN",
                r"Software\VisualSVN",
            ):
                try:
                    with winreg.OpenKey(root_key, subkey) as key:
                        install_dir, _ = winreg.QueryValueEx(key, "Directory")
                        svn_exe = os.path.join(install_dir, "bin", "svn.exe")
                        if os.path.isfile(svn_exe):
                            _SVN_EXE = svn_exe
                            return True
                except (OSError, FileNotFoundError):
                    continue
    except Exception:
        pass

    # Method 3: Search common paths
    for p in _SVN_SEARCH_PATHS:
        if os.path.isfile(p):
            _SVN_EXE = p
            return True

    return False


def _find_env_svn_installer():
    """Return the bundled SVN installer path if the env package is present."""
    candidates = [
        DATA_DIR / "env",
        Path(__file__).resolve().parent / "env",
        Path.cwd() / "env",
    ]
    for env_dir in candidates:
        if not env_dir.exists():
            continue
        installers = sorted(env_dir.glob("Slik-Subversion-*.msi"))
        if installers:
            return installers[0]
    return None


def _install_svn_from_env(timeout=300):
    """Install bundled SVN command-line tools from env/*.msi.

    The MSI may still trigger Windows UAC. We use /passive so users can see
    that dependency setup is running, while keeping the installation automatic
    once permission is granted.
    """
    installer = _find_env_svn_installer()
    if not installer:
        return {
            "ok": False,
            "error": "未找到 env 目录下的 SVN 安装包",
            "installer": None,
        }
    if os.name != "nt":
        return {
            "ok": False,
            "error": "自动安装 SVN 仅支持 Windows",
            "installer": str(installer),
        }

    cmd = [
        "msiexec.exe",
        "/i",
        str(installer),
        "/passive",
        "/norestart",
    ]
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            timeout=timeout,
            creationflags=_CREATE_NO_WINDOW,
        )
    except subprocess.TimeoutExpired:
        return {
            "ok": False,
            "error": "SVN 安装超时",
            "installer": str(installer),
        }
    except Exception as e:
        return {
            "ok": False,
            "error": str(e),
            "installer": str(installer),
        }

    if result.returncode in (0, 3010):
        return {
            "ok": True,
            "restart_required": result.returncode == 3010,
            "installer": str(installer),
        }

    stderr = _decode_svn(result.stderr).strip()
    stdout = _decode_svn(result.stdout).strip()
    return {
        "ok": False,
        "returncode": result.returncode,
        "error": stderr or stdout or f"msiexec 返回 {result.returncode}",
        "installer": str(installer),
    }


_SVN_FOUND = None  # cached result: None=unchecked, True/False


def _missing_svn_message():
    installer = _find_env_svn_installer()
    local_hint = f"\n已检测到本地安装包: {installer}" if installer else "\n未找到 env\\Slik-Subversion-*.msi。"
    return (
        "未找到 SVN 命令行工具 svn.exe。TortoiseSVN 默认不包含 svn.exe。"
        f"{local_hint}\n"
        "请安装 env 目录下的 SlikSVN，或安装 VisualSVN/Apache Subversion，"
        "也可以重新运行 TortoiseSVN 安装器并启用 command line client tools。"
    )


def _ensure_svn(auto_install=True):
    """Ensure SVN is findable. Raises SvnError if not."""
    global _SVN_FOUND, _SVN_INSTALL_ATTEMPTED, _SVN_INSTALL_RESULT
    if _SVN_FOUND is None or _SVN_FOUND is False:
        _SVN_FOUND = _find_svn()
    if not _SVN_FOUND and auto_install and not _SVN_INSTALL_ATTEMPTED:
        with _SVN_INSTALL_LOCK:
            if _SVN_FOUND is None:
                _SVN_FOUND = _find_svn()
            if not _SVN_FOUND and not _SVN_INSTALL_ATTEMPTED:
                _SVN_INSTALL_ATTEMPTED = True
                _SVN_INSTALL_RESULT = _install_svn_from_env()
                if _SVN_INSTALL_RESULT.get("ok"):
                    _SVN_FOUND = _find_svn()
    if not _SVN_FOUND:
        install_error = ""
        if _SVN_INSTALL_RESULT:
            install_error = f"\n自动安装结果: {_SVN_INSTALL_RESULT.get('error') or _SVN_INSTALL_RESULT}"
        raise SvnError(_missing_svn_message() + install_error)


def ensure_svn_available(auto_install=True, raise_on_error=False):
    """Public environment check used by the GUI startup and diagnostics API."""
    global _SVN_FOUND
    try:
        _ensure_svn(auto_install=auto_install)
        return {
            "installed": True,
            "exe": _SVN_EXE,
            "installer": str(_find_env_svn_installer() or ""),
            "install_attempted": _SVN_INSTALL_ATTEMPTED,
            "install_result": _SVN_INSTALL_RESULT,
        }
    except SvnError as e:
        if raise_on_error:
            raise
        return {
            "installed": False,
            "exe": "",
            "installer": str(_find_env_svn_installer() or ""),
            "install_attempted": _SVN_INSTALL_ATTEMPTED,
            "install_result": _SVN_INSTALL_RESULT,
            "error": str(e),
        }


class SvnError(Exception):
    pass


class SvnClient:
    def __init__(self, working_dir=None):
        self.root = Path(working_dir) if working_dir else RAW_DIR.parent

    def _run(self, args, timeout=120):
        """Run svn command, return stdout. Raise SvnError on failure."""
        _ensure_svn()
        cmd = [_SVN_EXE] + args
        try:
            result = subprocess.run(
                cmd,
                cwd=str(self.root),
                capture_output=True,
                timeout=timeout,
                creationflags=_CREATE_NO_WINDOW,
                env=_svn_env(),
            )
        except FileNotFoundError:
            raise SvnError(f"SVN executable not found: {_SVN_EXE}")
        except subprocess.TimeoutExpired:
            raise SvnError(f"SVN command timed out: {' '.join(cmd)}")

        stdout = _decode_svn(result.stdout)
        stderr = _decode_svn(result.stderr)

        if result.returncode != 0:
            err = stderr.strip() or stdout.strip()
            raise SvnError(f"SVN error: {err}")

        return stdout.strip()

    def is_working_copy(self):
        """Check if the target directory is an SVN working copy."""
        _ensure_svn()
        try:
            self._run(["info", str(RAW_DIR)], timeout=10)
            return True
        except SvnError as e:
            # Check if error is about SVN itself or about the working copy
            msg = str(e)
            if "not a working copy" in msg.lower():
                return False
            if "svn: e" in msg.lower() and "working copy" in msg.lower():
                return False
            # Other errors (network, auth, etc.) — re-raise
            raise

    def info(self):
        """Return dict with url, revision, last_changed_author, etc."""
        out = self._run(["info", str(RAW_DIR)])
        info = {}
        for line in out.splitlines():
            if ":" in line:
                key, _, val = line.partition(": ")
                info[key.strip()] = val.strip()
        return info

    def status(self, paths=None):
        """
        Run svn status, return structured list.
        Each item: {path, status, conflicted}
        """
        target = []
        if paths:
            for p in paths:
                full = RAW_DIR / p if not Path(p).is_absolute() else Path(p)
                target.append(str(full))
        else:
            target = [str(RAW_DIR)]
        out = self._run(["status", "--"] + target)
        items = []
        for line in out.splitlines():
            if not line.strip():
                continue
            status_char = line[0] if len(line) > 0 else ""
            props_char = line[1] if len(line) > 1 else ""
            lock_char = line[2] if len(line) > 2 else ""
            rest = line[7:].strip() if len(line) > 7 else line.strip()

            item = {
                "path": rest,
                "status": status_char,
                "props": props_char,
                "locked": lock_char == "L",
                "conflicted": status_char == "C" or props_char == "C",
                "is_conflict": line.startswith("?") == False and status_char == "C",
            }
            items.append(item)
        return items

    def update(self):
        """Run svn update on DataTable/Raw. Return update summary."""
        out = self._run(["update", str(RAW_DIR)])
        lines = out.splitlines()
        summary = {
            "revision": "",
            "updated": 0,
            "merged": 0,
            "conflicted": 0,
            "deleted": 0,
            "added": 0,
            "raw": lines,
        }
        for line in lines:
            # Summary lines are not file changes — skip before status counting
            # ("Updated to revision 5." starts with 'U', "At revision 5." with 'A').
            if line.startswith("At revision") or line.startswith("Updated to revision"):
                summary["revision"] = line.split()[-1].rstrip(".")
                continue
            status = line[0] if line else ""
            if status == "U":
                summary["updated"] += 1
            elif status == "G":
                summary["merged"] += 1
            elif status == "C":
                summary["conflicted"] += 1
            elif status == "D":
                summary["deleted"] += 1
            elif status == "A":
                summary["added"] += 1
        return summary

    def resolved(self, filepath):
        """Run svn resolved on a file, accepting merged result."""
        full = RAW_DIR / filepath if not Path(filepath).is_absolute() else Path(filepath)
        self._run(["resolved", str(full)])

    def lock(self, filepath, message=None):
        """Run svn lock on a file. Returns dict with ok/error/locked_by."""
        full = RAW_DIR / filepath if not Path(filepath).is_absolute() else Path(filepath)
        args = ["lock", str(full)]
        if message:
            args += ["-m", message]
        try:
            out = self._run(args, timeout=30)
            token = ""
            for line in out.splitlines():
                if "locked by" in line.lower():
                    parts = line.split(":")
                    token = parts[-1].strip() if len(parts) > 1 else ""
            return {"ok": True, "token": token}
        except SvnError as e:
            msg = str(e)
            # Parse "already locked by <user>"
            locked_by = ""
            if "locked by" in msg.lower():
                parts = msg.split("locked by", 1)
                locked_by = parts[1].strip().rstrip(".") if len(parts) > 1 else "unknown"
            return {"ok": False, "error": msg, "locked_by": locked_by, "already_locked": True}

    def unlock(self, filepath):
        """Run svn unlock on a file. Returns dict with ok/error."""
        full = RAW_DIR / filepath if not Path(filepath).is_absolute() else Path(filepath)
        try:
            self._run(["unlock", str(full)], timeout=30)
            return {"ok": True}
        except SvnError as e:
            return {"ok": False, "error": str(e)}

    def log(self, limit=10):
        """Return recent svn log entries."""
        out = self._run(["log", str(RAW_DIR), "-l", str(limit)])
        entries = []
        current = None
        for line in out.splitlines():
            if line.startswith("---"):
                continue
            if line.startswith("r") and "|" in line:
                # r12345 | author | date | lines
                parts = line.split("|")
                revision = parts[0].strip()
                author = parts[1].strip() if len(parts) > 1 else ""
                date = parts[2].strip() if len(parts) > 2 else ""
                current = {"revision": revision, "author": author, "date": date, "message": ""}
                entries.append(current)
            elif current is not None and line.strip():
                current["message"] += line.strip() + "\n"
        for e in entries:
            e["message"] = e["message"].strip()
        return entries

    def log_file(self, filepath, limit=30):
        """Return SVN log entries for a specific file."""
        full = RAW_DIR / filepath if not Path(filepath).is_absolute() else Path(filepath)
        out = self._run(["log", str(full), "-l", str(limit)])
        entries = []
        current = None
        for line in out.splitlines():
            if line.startswith("---"):
                continue
            if line.startswith("r") and "|" in line:
                parts = line.split("|")
                revision = parts[0].strip()
                author = parts[1].strip() if len(parts) > 1 else ""
                date = parts[2].strip() if len(parts) > 2 else ""
                current = {"revision": revision, "author": author, "date": date, "message": ""}
                entries.append(current)
            elif current is not None and line.strip():
                current["message"] += line.strip() + "\n"
        for e in entries:
            e["message"] = e["message"].strip()
        return entries

    def cat_to_file(self, filepath, revision, dest_path):
        """Export a file at a specific revision to dest_path using svn cat."""
        full = RAW_DIR / filepath if not Path(filepath).is_absolute() else Path(filepath)
        _ensure_svn()
        rev_str = str(revision).lstrip("r")
        cmd = [_SVN_EXE, "cat", "-r", rev_str, str(full)]
        try:
            result = subprocess.run(
                cmd,
                cwd=str(self.root),
                capture_output=True,
                timeout=60,
                creationflags=_CREATE_NO_WINDOW,
                env=_svn_env(),
            )
        except FileNotFoundError:
            raise SvnError(f"SVN executable not found: {_SVN_EXE}")
        except subprocess.TimeoutExpired:
            raise SvnError(f"SVN cat timed out for {filepath}@r{rev_str}")

        if result.returncode != 0:
            err = _decode_svn(result.stderr).strip()
            raise SvnError(f"SVN cat error: {err}")

        with open(dest_path, "wb") as f:
            f.write(result.stdout)

    def commit(self, files, message):
        """Commit specific files with a message. Returns dict with ok/revision."""
        if not files or not message.strip():
            raise SvnError("commit requires files and a non-empty message")

        full_paths = []
        for fp in files:
            f = RAW_DIR / fp if not Path(fp).is_absolute() else Path(fp)
            full_paths.append(str(f))

        _ensure_svn()
        cmd = [_SVN_EXE, "commit", "-m", message] + full_paths
        try:
            result = subprocess.run(
                cmd,
                cwd=str(self.root),
                capture_output=True,
                timeout=120,
                creationflags=_CREATE_NO_WINDOW,
                env=_svn_env(),
            )
        except subprocess.TimeoutExpired:
            raise SvnError("SVN commit timed out")

        stdout = _decode_svn(result.stdout)
        stderr = _decode_svn(result.stderr)

        if result.returncode != 0:
            err = stderr.strip() or stdout.strip()
            raise SvnError(f"SVN commit error: {err}")

        revision = ""
        for line in stdout.splitlines():
            if "committed revision" in line.lower():
                revision = line.split()[-1].rstrip(".")
                break
        return {"ok": True, "revision": revision, "output": stdout.strip()}

    def check_updates(self):
        """Check if remote SVN has newer revision than local workspace.

        Returns: {local_rev, remote_rev, has_update, behind_by}
        On timeout or network error: returns has_update=False with error field.
        """
        try:
            local_info = self._run(["info", "--show-item", "revision", str(RAW_DIR)], timeout=10)
            local_rev = int(local_info.strip())
        except (SvnError, ValueError):
            return {"has_update": False, "error": "Cannot get local revision"}

        try:
            remote_info = self._run(
                ["info", "--show-item", "revision", "-r", "HEAD", str(RAW_DIR)],
                timeout=15,
            )
            remote_rev = int(remote_info.strip())
        except SvnError as e:
            msg = str(e)
            if "timed out" in msg.lower():
                return {"local_rev": local_rev, "has_update": False, "timeout": True}
            return {"local_rev": local_rev, "has_update": False, "error": msg}
        except ValueError:
            return {"local_rev": local_rev, "has_update": False, "error": "Cannot parse remote revision"}

        behind_by = max(0, remote_rev - local_rev)
        return {
            "local_rev": local_rev,
            "remote_rev": remote_rev,
            "has_update": behind_by > 0,
            "behind_by": behind_by,
        }

    def get_file_prev_rev(self, filepath):
        """Get the second-to-last revision that modified a file.
        Returns revision string (e.g., 'r4520') or None if only 1 revision.
        """
        entries = self.log_file(filepath, limit=2)
        if len(entries) >= 2:
            return entries[1]["revision"]
        return None

    def find_conflict_files(self):
        """
        Scan DataTable/Raw for SVN conflict artifacts (*.mine, *.rOLD, *.rNEW).
        Returns dict: {base_filepath: {mine, theirs, base}}
        """
        conflicts = {}
        raw = RAW_DIR
        rev_re = re.compile(r"^(?P<base>.+\.xlsx)\.r(?P<rev>\d+)$")
        for root, dirs, files in os.walk(raw):
            for fname in files:
                if fname.startswith("~$"):
                    continue
                full = Path(root) / fname
                rel = full.relative_to(raw)

                if fname.endswith(".mine"):
                    base_name = fname[:-5]  # remove .mine
                    base_rel = (rel.parent / base_name).as_posix() if str(rel.parent) != "." else base_name
                    conflicts.setdefault(base_rel, {})["mine"] = str(full)
                else:
                    m = rev_re.match(fname)
                    if m:
                        # filename.xlsx.rNNNNN — old/new base versions
                        base_name = m.group("base")
                        base_rel = (rel.parent / base_name).as_posix() if str(rel.parent) != "." else base_name
                        entry = conflicts.setdefault(base_rel, {})
                        entry.setdefault("rev_files", []).append((int(m.group("rev")), str(full)))

        # Sort rev files: lower rev = base (old), higher rev = theirs (new)
        for base_rel, entry in conflicts.items():
            rev_files = entry.get("rev_files", [])
            if rev_files:
                rev_files.sort(key=lambda x: x[0])
                if len(rev_files) >= 2:
                    entry["base"] = rev_files[0][1]
                    entry["theirs"] = rev_files[-1][1]
                elif len(rev_files) == 1:
                    entry["theirs"] = rev_files[0][1]
                del entry["rev_files"]

            # Fill missing: mine might not exist if local was same as base
            if "mine" not in entry:
                entry["mine"] = str(raw / base_rel)

        return conflicts
