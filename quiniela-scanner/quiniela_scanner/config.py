"""User settings, stored as config.json next to the program."""
from __future__ import annotations

import json
import os
import sys
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path


def app_dir() -> Path:
    """Folder that holds config.json, scans/ and the log (the exe folder when frozen)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


@dataclass
class Config:
    # --- camera -----------------------------------------------------------
    camera_index: int = 0
    camera_backend: str = "dshow"        # dshow | msmf | auto  (dshow is best for the C920s)
    frame_width: int = 1920
    frame_height: int = 1080
    autofocus: bool = True
    manual_focus: int = 30               # 0-255, used only when autofocus is False

    # --- slip -------------------------------------------------------------
    matches_per_slip: int = 9

    # --- Excel ------------------------------------------------------------
    backend: str = "auto"                # auto | com (live Excel, Windows) | file (.xlsx on disk)
    workbook: str = ""                   # "" = whichever workbook is active in Excel when you connect
    workbook_path: str = ""              # only for backend=file
    sheets: list = field(default_factory=lambda: ["ENTRY*"])
    name_header: str = "NOMBRE"          # header text that marks the name column of each ticket block
    uppercase_names: bool = True
    name_tag: str = ""                   # group tag such as "RRR" or "E" (remembered; the on/off switch always starts OFF)
    name_tag_position: str = "after"     # "before" or "after" the name
    tag_separator: str = " "             # put between tag and name
    save_after_write: bool = False       # Excel AutoSave normally handles this
    expected_jornada: str = ""           # "" = read it from the workbook name ("JORNADA 10 PC" -> 10)

    # --- name reading -----------------------------------------------------
    read_names_with_claude: bool = True  # needs ANTHROPIC_API_KEY; otherwise you type the name
    claude_model: str = "claude-opus-5-5"

    # --- empty boxes ------------------------------------------------------
    blank_marker: str = "F1"             # written in a pick box the player left empty ("" = leave the cell empty)
    blank_font_size: int = 9             # font size of that marker
    auto_enter_blanks: bool = True       # an empty box (F1) alone does not stop auto-enter ...
    auto_enter_max_blank: int = 2        # ... but more than this many empty boxes needs a human look

    # --- behaviour --------------------------------------------------------
    auto_enter: bool = True              # write clean scans without a click
    auto_enter_delay_s: float = 3.0      # review window before an automatic write (0 = instant)
    stable_frames: int = 7               # frames the slip must sit still before capture
    min_ink: float = 0.015               # mark detection thresholds (see slip_vision.Settings)
    multi_ratio: float = 0.45

    # ----------------------------------------------------------------------
    @classmethod
    def path(cls) -> Path:
        return app_dir() / "config.json"

    @classmethod
    def load(cls, path: Path | None = None) -> "Config":
        path = path or cls.path()
        cfg = cls()
        if path.exists():
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                return cfg
            known = {f.name for f in fields(cls)}
            for k, v in data.items():
                if k in known:
                    setattr(cfg, k, v)
        return cfg

    def save(self, path: Path | None = None) -> None:
        path = path or self.path()
        try:
            path.write_text(json.dumps(asdict(self), indent=2), encoding="utf-8")
        except OSError:
            pass


def load_dotenv() -> None:
    """Read ANTHROPIC_API_KEY=... from a .env file next to the program (if the variable is not already set)."""
    env = app_dir() / ".env"
    if not env.exists():
        return
    try:
        for line in env.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))
    except OSError:
        pass


def anthropic_key_present() -> bool:
    return bool(os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN"))
