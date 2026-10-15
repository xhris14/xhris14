"""User settings, stored as config.json next to the program."""
from __future__ import annotations

import codecs
import json
import locale
import math
import os
import shutil
import sys
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path

CONFIG_VERSION = 2
# Builds before version 2 saved every setting, so these three would pin the old (slow, hands-on) defaults forever.
RESET_ON_UPGRADE = ("stable_frames", "auto_enter_delay_s", "auto_enter")


# Hand-edited values that would break the app (0 matches, NaN timeouts, a negative camera ...) fall back to the default.
RANGES = {"matches_per_slip": (1, 30), "camera_index": (0, 20), "frame_width": (160, 8192), "frame_height": (120, 8192),
          "manual_focus": (0, 255), "stable_frames": (1, 120), "name_workers": (1, 64), "name_timeout_s": (0, 3600),
          "auto_enter_delay_s": (0, 600), "auto_enter_max_blank": (0, 30), "blank_font_size": (1, 72),
          "min_ink": (0, 1), "multi_ratio": (0, 1)}
CHOICES = {"backend": ("auto", "com", "file"), "camera_backend": ("dshow", "msmf", "auto"),
           "name_tag_position": ("before", "after")}


def read_text_any(path: Path) -> str:
    """Text of a hand-edited file whatever its encoding: UTF-8 (with or without BOM), UTF-16 (what PowerShell's
    '>' writes) or the Windows ANSI code page. Raises OSError only when the file cannot be read at all."""
    raw = Path(path).read_bytes()
    if raw.startswith(codecs.BOM_UTF8):
        return raw.decode("utf-8-sig", errors="replace")
    if raw.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        return raw.decode("utf-16", errors="replace")
    if b"\x00" in raw[:64]:                              # UTF-16 without a BOM
        return raw.decode("utf-16-be" if raw[:1] == b"\x00" else "utf-16-le", errors="replace")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        pass
    try:
        return raw.decode(locale.getpreferredencoding(False) or "cp1252")
    except (UnicodeDecodeError, LookupError):
        return raw.decode("cp1252", errors="replace")


def app_dir() -> Path:
    """Folder that holds config.json, scans/ and the log (the exe folder when frozen)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


@dataclass
class Config:
    config_version: int = CONFIG_VERSION

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
    claude_effort: str = "low"           # output_config effort sent with every name request
    name_workers: int = 6                # name reads in flight at once (they run in the background)
    name_timeout_s: float = 45.0         # give up on a name still not back after this long -> "?" + review
    unsure_marker: str = " ?"            # appended to a name Claude was not sure of ("" = no marker)

    # --- empty boxes ------------------------------------------------------
    blank_marker: str = "F1"             # written in a pick box the player left empty ("" = leave the cell empty)
    blank_font_size: int = 9             # font size of that marker
    auto_enter_blanks: bool = True       # an empty box (F1) alone does not stop auto-enter ...
    auto_enter_max_blank: int = 2        # ... but more than this many empty boxes needs a human look

    # --- behaviour --------------------------------------------------------
    auto_enter: bool = True              # write clean scans without a click
    auto_enter_delay_s: float = 0.0      # > 0: old hands-on mode - wait for the name, then count down (0 = instant)
    stable_frames: int = 4               # frames the slip must sit still before capture
    sounds: bool = True                  # short beeps: slip entered / needs you (Windows only)
    min_ink: float = 0.015               # mark detection thresholds (see slip_vision.Settings)
    multi_ratio: float = 0.45

    def __post_init__(self):
        self.migrated: list = []         # settings of an older config.json that were reset (not saved)
        self.notes: list = []            # problems with config.json / .env to show the user (not saved)
        self._session: dict = {}         # command-line overrides: name -> the value to keep in config.json

    def override(self, **values) -> None:
        """Use these values for this run only (command-line flags such as --xlsx must not end up in config.json)."""
        for k, v in values.items():
            self._session.setdefault(k, getattr(self, k))
            setattr(self, k, v)

    # ----------------------------------------------------------------------
    @classmethod
    def path(cls) -> Path:
        return app_dir() / "config.json"

    @classmethod
    def load(cls, path: Path | None = None) -> "Config":
        path = path or cls.path()
        cfg = cls()
        if not path.exists():
            return cfg
        try:
            data = json.loads(read_text_any(path))
            if not isinstance(data, dict):
                raise ValueError("not a JSON object")
        except (OSError, ValueError) as e:       # (UnicodeError is a ValueError)
            kept = "" if isinstance(e, OSError) else cls._keep_copy(path)
            cfg.notes.append(f"{path.name} could not be read ({e.__class__.__name__}: {e}) - using the default settings."
                             + (f" Your file was kept as {kept}." if kept else ""))
            return cfg
        version = data.get("config_version")
        if not isinstance(version, int) or isinstance(version, bool) or version < CONFIG_VERSION:
            cfg.migrated = [k for k in RESET_ON_UPGRADE if k in data]
            data = {k: v for k, v in data.items() if k not in RESET_ON_UPGRADE}
        default = cls()
        for f in fields(cls):
            if f.name == "config_version" or f.name not in data:
                continue
            v, d = data[f.name], getattr(default, f.name)
            if isinstance(d, bool):
                ok = isinstance(v, bool)
            elif isinstance(d, int):
                ok = isinstance(v, int) and not isinstance(v, bool)
            elif isinstance(d, float):
                ok = isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v)
                v = float(v) if ok else v
            elif isinstance(d, list):
                ok = isinstance(v, list) and bool(v) and all(isinstance(x, str) and x.strip() for x in v)
            else:
                ok = isinstance(v, type(d))
            if ok and f.name in RANGES:
                lo, hi = RANGES[f.name]
                ok = lo <= v <= hi
            if ok and f.name in CHOICES:
                ok = v in CHOICES[f.name]
            if ok:
                setattr(cfg, f.name, v)
            else:
                cfg.notes.append(f"{path.name}: {f.name}={v!r} is not valid - using {d!r}.")
        return cfg

    @staticmethod
    def _keep_copy(path: Path) -> str:
        """An unparseable config.json is about to be replaced by the defaults on close: keep what the user wrote."""
        try:
            if not path.read_bytes().strip():
                return ""
            for i in range(1, 50):
                bak = path.with_name(path.name + (".bak" if i == 1 else f".bak{i}"))
                if not bak.exists():
                    shutil.copy2(path, bak)
                    return bak.name
        except OSError:
            pass
        return ""

    def save(self, path: Path | None = None) -> None:
        """Write only what differs from the defaults (so new defaults reach everybody) plus config_version."""
        path = path or self.path()
        default = asdict(Config())
        data = {"config_version": CONFIG_VERSION}
        data.update({k: v for k, v in {**asdict(self), **self._session}.items()
                     if k != "config_version" and v != default[k]})
        try:
            tmp = path.with_name(path.name + ".tmp")
            tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
            os.replace(tmp, path)
        except OSError:
            pass


def load_dotenv() -> str:
    """Read ANTHROPIC_API_KEY=... from a .env file next to the program (if the variable is not already set).
    Returns "" or a sentence for the user when the file exists but could not be read."""
    env = app_dir() / ".env"
    if not env.exists():
        return ""
    try:
        text = read_text_any(env)
    except OSError as e:
        return f".env could not be read ({e.__class__.__name__}) - the API key was not loaded."
    for line in text.splitlines():
        line = line.strip().lstrip("\ufeff")
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            k = k.strip().removeprefix("export ").strip()
            if k and not os.environ.get(k):
                os.environ[k] = v.strip().strip('"').strip("'")
    return ""


def anthropic_key_present() -> bool:
    return bool(os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN"))
