"""Reading the handwritten name (and the printed jornada) off a slip.

Handwriting is the one thing classic OCR does badly, so the name goes to Claude
vision. If there is no API key / no internet the reader reports ``available =
False`` and the operator simply types the name (the UI always lets you correct it).

The reader runs in a background pool: ``read`` never raises, every failure comes back as a
``NameResult`` with ``error`` set.
"""
from __future__ import annotations

import base64
import json
import re
from dataclasses import dataclass

import cv2
import numpy as np

PROMPT = """Two crops from a Mexican football pool slip (quiniela).

Image 1 is the printed title strip (it normally reads "LIGA MX JORNADA <number>").
Image 2 is the bottom of the slip: the printed label "NOMBRE" followed by a name written by hand. \
Printed text such as "ULTIMO DIA ..." may be partly visible under it - ignore it.

Reply with ONLY a JSON object, no other text:
{"name": "<the handwritten name exactly as written, UPPERCASE, without the word NOMBRE or underscores>",
 "jornada": "<the number after JORNADA in the title, or null if unreadable>",
 "legible": <true if you are confident of the name, false if you had to guess>}

Keep punctuation and digits that are part of the name (for example "CLEMENTE O." or "1-2-3" or "PENJAMO 2")."""

MAX_IMAGE_W = 800           # a wider crop gains nothing for reading handwriting, only upload time
MAX_NAME_CHARS = 80


@dataclass
class NameResult:
    name: str = ""
    jornada: str | None = None
    legible: bool = False
    error: str | None = None
    systemic: bool = False     # the failure is about the service (key, credit, line, rate limit), not about this slip


def _jpeg_b64(img: np.ndarray, max_w: int = MAX_IMAGE_W) -> str:
    """Crop -> base64 JPEG at most max_w wide (about 6x smaller than PNG: that is upload time on a slow line)."""
    h, w = img.shape[:2]
    if w > max_w:
        img = cv2.resize(img, None, fx=max_w / w, fy=max_w / w, interpolation=cv2.INTER_AREA)
    ok, buf = cv2.imencode(".jpg", img, [cv2.IMWRITE_JPEG_QUALITY, 90])
    if not ok:
        raise ValueError("could not encode image")
    return base64.standard_b64encode(buf.tobytes()).decode("ascii")


def _truthy(v) -> bool:
    if isinstance(v, str):
        return v.strip().lower() in ("true", "yes", "si", "1")      # "false" must not count as True
    return bool(v)


def parse_reply(text: str) -> NameResult:
    """Pull the JSON object out of the model reply (tolerates code fences / chatter)."""
    m = re.search(r"\{.*\}", text, re.S)
    if not m:
        return NameResult(error="no JSON in reply")
    try:
        data = json.loads(m.group(0))
    except ValueError:
        return NameResult(error="bad JSON in reply")
    if not isinstance(data, dict):
        return NameResult(error="bad JSON in reply")
    name = str(data.get("name") or "").strip()
    name = re.sub(r"^\s*NOMBRE\s*[:_\-]*\s*", "", name, flags=re.I).strip(" _")[:MAX_NAME_CHARS]
    jr = data.get("jornada")
    jr = None if jr in (None, "", "null") else re.sub(r"\D", "", str(jr)) or None
    return NameResult(name=name, jornada=jr, legible=_truthy(data.get("legible", False)))


class ClaudeNameReader:
    # class-level defaults so a reader built without __init__ (test doubles) still has a complete policy
    effort = "low"
    max_tokens = 4000          # thinking tokens count against this: do NOT lower it for always-on thinking models

    def __init__(self, model: str = "claude-opus-5-5", timeout: float = 25.0, effort: str = "low",
                 max_retries: int = 1):
        import anthropic  # imported lazily so the app still runs without the package
        self._anthropic = anthropic
        self.model = model
        self.effort = effort
        # 5 s to connect (a dead line fails fast), `timeout` s of silence to read; one retry, so a slip is
        # given up on after ~1 minute at the very worst and never blocks a pool worker for good
        limit = getattr(anthropic, "Timeout", None)
        self.client = anthropic.Anthropic(timeout=limit(timeout, connect=5.0) if limit else timeout,
                                          max_retries=max_retries)

    def read(self, name_crop: np.ndarray, title_crop: np.ndarray | None = None) -> NameResult:
        try:
            return self._read(name_crop, title_crop)
        except Exception as e:                       # a pool worker must never die with an exception
            # (a picture that cannot be encoded is about this slip; any other surprise is the reader itself failing)
            return NameResult(error=f"name reading failed ({e.__class__.__name__})",
                              systemic=not isinstance(e, (ValueError, cv2.error)))

    def _read(self, name_crop: np.ndarray, title_crop: np.ndarray | None) -> NameResult:
        a = self._anthropic
        content = []
        for img in (title_crop, name_crop):
            if img is not None and img.size:
                content.append({"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                                            "data": _jpeg_b64(img)}})
        content.append({"type": "text", "text": PROMPT})
        try:
            resp = self.client.messages.create(
                model=self.model,
                max_tokens=self.max_tokens,
                output_config={"effort": self.effort},
                messages=[{"role": "user", "content": content}],
            )
        except a.AuthenticationError:                # (systemic: the same answer is coming for every slip)
            return NameResult(error="Anthropic API key rejected", systemic=True)
        except a.RateLimitError:
            return NameResult(error="Anthropic rate limit", systemic=True)
        except a.APITimeoutError:                    # before APIConnectionError, which it subclasses
            return NameResult(error="Claude took too long", systemic=True)
        except a.APIConnectionError:
            return NameResult(error="no internet", systemic=True)
        except a.APIStatusError as e:                # 400 = no credit left, 5xx = overloaded, ...
            return NameResult(error=f"Anthropic API error {e.status_code}", systemic=True)
        except a.APIError as e:
            return NameResult(error=f"Anthropic API error ({e.__class__.__name__})", systemic=True)
        if resp.stop_reason == "refusal":
            return NameResult(error="Claude declined to read this one")
        text = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
        if resp.stop_reason == "max_tokens" and not text.strip():
            return NameResult(error="Claude ran out of tokens")
        return parse_reply(text)


def make_reader(enabled: bool, model: str, effort: str = "low"):
    """Return a reader, or (None, reason) when handwriting reading is not available."""
    if not enabled:
        return None, "name reading is switched off - type the name"
    from .config import anthropic_key_present
    if not anthropic_key_present():
        return None, "no ANTHROPIC_API_KEY set - type the name"
    try:
        return ClaudeNameReader(model, effort=effort), "Claude reads the handwritten name"
    except Exception as e:  # missing package, bad environment
        return None, f"name reading unavailable ({e.__class__.__name__}) - type the name"
