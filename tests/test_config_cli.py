from __future__ import annotations

import json
from pathlib import Path

from office_docs_mcp.common.config_manager import get_platform_config_path
from office_docs_mcp.main import build_root


def test_platform_config_path_resolution():
    home = Path("/home/testuser")
    # 1. Linux with XDG_CONFIG_HOME
    linux_xdg = get_platform_config_path(
        platform="linux",
        home=home,
        environ={"XDG_CONFIG_HOME": "/custom/config"},
    )
    assert linux_xdg == Path("/custom/config/office_docs_mcp/config.toml")

    # 2. Linux fallback (~/.config)
    linux_default = get_platform_config_path(platform="linux", home=home, environ={})
    assert linux_default == Path("/home/testuser/.config/office_docs_mcp/config.toml")

    # 3. macOS (Darwin)
    darwin = get_platform_config_path(platform="darwin", home=home)
    assert darwin == Path("/home/testuser/Library/Application Support/office_docs_mcp/config.toml")

    # 4. Windows (win32)
    win = get_platform_config_path(
        platform="win32",
        home=home,
        environ={"APPDATA": "C:/Users/testuser/AppData/Roaming"},
    )
    assert win == Path("C:/Users/testuser/AppData/Roaming/office_docs_mcp/config.toml")


def test_config_init_creates_file(tmp_path):
    cfg_file = tmp_path / "sub" / "custom_config.toml"
    root = build_root()
    code = root.execute(["config", "init", "--path", str(cfg_file)])
    assert code == 0
    assert cfg_file.exists()

    content = cfg_file.read_text(encoding="utf-8")
    assert "[app]" in content
    assert 'name = "office_docs_mcp"' in content
    assert "[logging]" in content
    assert 'level = "INFO"' in content


def test_config_init_platform_default(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))
    root = build_root()
    code = root.execute(["config", "init"])
    assert code == 0

    expected_file = tmp_path / "office_docs_mcp" / "config.toml"
    assert expected_file.exists()
    content = expected_file.read_text(encoding="utf-8")
    assert "office_docs_mcp" in content


def test_config_init_local(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    root = build_root()
    code = root.execute(["config", "init", "--local"])
    assert code == 0

    expected_file = tmp_path / "config.toml"
    assert expected_file.exists()


def test_config_init_existing_fails_without_force(tmp_path, capsys):
    cfg_file = tmp_path / "config.toml"
    cfg_file.write_text("existing = true\n", encoding="utf-8")

    root = build_root()
    code = root.execute(["config", "init", "--path", str(cfg_file)])
    assert code != 0
    captured = capsys.readouterr()
    assert "already exists" in captured.err.lower() or "already exists" in captured.out.lower()


def test_config_init_existing_succeeds_with_force(tmp_path):
    cfg_file = tmp_path / "config.toml"
    cfg_file.write_text("existing = true\n", encoding="utf-8")

    root = build_root()
    code = root.execute(["config", "init", "--path", str(cfg_file), "--force"])
    assert code == 0
    content = cfg_file.read_text(encoding="utf-8")
    assert "office_docs_mcp" in content


def test_config_path_default_platform(tmp_path, monkeypatch, capsys):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))
    root = build_root()
    code = root.execute(["config", "path"])
    assert code == 0
    captured = capsys.readouterr()
    expected = str((tmp_path / "office_docs_mcp" / "config.toml").resolve())
    assert expected in captured.out


def test_config_path_local(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    root = build_root()
    code = root.execute(["config", "path", "--local"])
    assert code == 0
    captured = capsys.readouterr()
    expected = str((tmp_path / "config.toml").resolve())
    assert expected in captured.out


def test_config_path_override(tmp_path, capsys):
    cfg_file = tmp_path / "my_config.toml"
    root = build_root()
    code = root.execute(["--config", str(cfg_file), "config", "path"])
    assert code == 0
    captured = capsys.readouterr()
    assert str(cfg_file.resolve()) in captured.out or str(cfg_file) in captured.out


def test_config_path_json(tmp_path, capsys):
    cfg_file = tmp_path / "my_config.toml"
    root = build_root()
    code = root.execute(["--config", str(cfg_file), "config", "path", "--json"])
    assert code == 0
    captured = capsys.readouterr()
    data = json.loads(captured.out)
    assert "path" in data
    assert data["exists"] is False


def test_config_show(tmp_path, capsys):
    root = build_root()
    code = root.execute(["config", "show"])
    assert code == 0
    captured = capsys.readouterr()
    data = json.loads(captured.out)
    assert data["app"]["name"] == "office_docs_mcp"


def test_config_set_platform_default(tmp_path, monkeypatch, capsys):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))
    root = build_root()

    code = root.execute(["config", "set", "app.name", "custom_mcp"])
    assert code == 0

    expected_file = tmp_path / "office_docs_mcp" / "config.toml"
    assert expected_file.exists()
    content = expected_file.read_text(encoding="utf-8")
    assert 'name = "custom_mcp"' in content


def test_config_set_and_get(tmp_path, capsys):
    cfg_file = tmp_path / "test_config.toml"
    root = build_root()

    # 1. Set key
    set_code = root.execute(["--config", str(cfg_file), "config", "set", "logging.level", "DEBUG"])
    assert set_code == 0
    assert cfg_file.exists()

    # 2. Get key
    get_code = root.execute(["--config", str(cfg_file), "config", "get", "logging.level"])
    assert get_code == 0
    captured = capsys.readouterr()
    assert "DEBUG" in captured.out

    # 3. Set another key with int and boolean
    root.execute(["--config", str(cfg_file), "config", "set", "app.workers", "4"])
    root.execute(["--config", str(cfg_file), "config", "set", "app.debug", "true"])

    get_workers = root.execute(["--config", str(cfg_file), "config", "get", "app.workers"])
    assert get_workers == 0
    captured = capsys.readouterr()
    assert "4" in captured.out

    get_debug = root.execute(["--config", str(cfg_file), "config", "get", "app.debug"])
    assert get_debug == 0
    captured = capsys.readouterr()
    assert "true" in captured.out


def test_config_get_missing_key(tmp_path, capsys):
    root = build_root()
    code = root.execute(["config", "get", "non.existent.key"])
    assert code != 0
    captured = capsys.readouterr()
    assert "not found" in captured.err.lower() or "not found" in captured.out.lower()
