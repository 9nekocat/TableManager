import svn_client


def _reset_svn_state(monkeypatch):
    monkeypatch.setattr(svn_client, "_SVN_EXE", "svn")
    monkeypatch.setattr(svn_client, "_SVN_FOUND", None)
    monkeypatch.setattr(svn_client, "_SVN_INSTALL_ATTEMPTED", False)
    monkeypatch.setattr(svn_client, "_SVN_INSTALL_RESULT", None)


def test_ensure_svn_installs_from_env_and_rechecks(monkeypatch):
    _reset_svn_state(monkeypatch)
    calls = {"find": 0, "install": 0}

    def fake_find():
        calls["find"] += 1
        if calls["find"] >= 2:
            svn_client._SVN_EXE = r"C:\Program Files\SlikSvn\bin\svn.exe"
            return True
        return False

    def fake_install():
        calls["install"] += 1
        return {"ok": True, "installer": r"D:\work\Tools\TableManager\env\Slik.msi"}

    monkeypatch.setattr(svn_client, "_find_svn", fake_find)
    monkeypatch.setattr(svn_client, "_install_svn_from_env", fake_install)

    status = svn_client.ensure_svn_available(auto_install=True)

    assert status["installed"] is True
    assert status["install_attempted"] is True
    assert status["exe"].endswith("svn.exe")
    assert calls == {"find": 2, "install": 1}


def test_ensure_svn_reports_failed_env_install(monkeypatch):
    _reset_svn_state(monkeypatch)

    monkeypatch.setattr(svn_client, "_find_svn", lambda: False)
    monkeypatch.setattr(
        svn_client,
        "_install_svn_from_env",
        lambda: {"ok": False, "error": "installer missing", "installer": None},
    )
    monkeypatch.setattr(svn_client, "_find_env_svn_installer", lambda: None)

    status = svn_client.ensure_svn_available(auto_install=True)

    assert status["installed"] is False
    assert status["install_attempted"] is True
    assert "installer missing" in status["error"]
