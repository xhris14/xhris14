"""Reading the handwritten name (and the printed jornada) off a slip.

Handwriting is the one thing classic OCR does badly, so the name goes to Claude
vision. If there is no API key / no internet the reader reports ``available =
False`` and the operator simply types the name (the UI always lets you correct it).
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


@dataclass
class NameResult:
    name: str = ""
    jornada: str | None = None
    legible: bool = False
    error: str | None = None


def _png_b64(img: np.ndarray, max_w: int = 1000) -> str:
    h, w = img.shape[:2]
    if w > max_w:
        img = cv2.resize(img, None, fx=max_w / w, fy=max_w / w, interpolation=cv2.INTER_AREA)
    ok, buf = cv2.imencode(".png", img)
    if not ok:
        raise ValueError("could not encode image")
    return base64.standard_b64encode(buf.tobytes()).decode("ascii")


def parse_reply(text: str) -> NameResult:
    """Pull the JSON object out of the model reply (tolerates code fences / chatter)."""
    m = re.search(r"\{.*\}", text, re.S)
    if not m:
        return NameResult(error="no JSON in reply")
    try:
        data = json.loads(m.group(0))
    except ValueError:
        return NameResult(error="bad JSON in reply")
    name = str(data.get("name") or "").strip()
    name = re.sub(r"^\s*NOMBRE\s*[:_\-]*\s*", "", name, flags=re.I).strip(" _")
    jr = data.get("jornada")
    jr = None if jr in (None, "", "null") else re.sub(r"\D", "", str(jr)) or None
    return NameResult(name=name, jornada=jr, legible=bool(data.get("legible", False)))


class ClaudeNameReader:
    def __init__(self, model: str = "claude-opus-5-5", timeout: float = 25.0):
        import anthropic  # imported lazily so the app still runs without the package
        self._anthropic = anthropic
        self.model = model
        self.client = anthropic.Anthropic(timeout=timeout, max_retries=1)

    def read(self, name_crop: np.ndarray, title_crop: np.ndarray | None = None) -> NameResult:
        a = self._anthropic
        content = []
        if title_crop is not None and title_crop.size:
            content.append({"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                                        "data": _png_b64(title_crop)}})
        content.append({"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                                    "data": _png_b64(name_crop)}})
        content.append({"type": "text", "text": PROMPT})
        try:
            resp = self.client.messages.create(
                model=self.model,
                max_tokens=4000,                       # thinking tokens count against this
                output_config={"effort": "low"},
                messages=[{"role": "user", "content": content}],
            )
        except a.AuthenticationError:
            return NameResult(error="Anthropic API key rejected")
        except a.RateLimitError:
            return NameResult(error="Anthropic rate limit - type the name")
        except a.APIConnectionError:
            return NameResult(error="no internet - type the name")
        except a.APIStatusError as e:
            return NameResult(error=f"Anthropic API error {e.status_code}")
        if resp.stop_reason == "refusal":
            return NameResult(error="Claude declined to read this one")
        text = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
        return parse_reply(text)


def make_reader(enabled: bool, model: str):
    """Return a reader, or (None, reason) when handwriting reading is not available."""
    if not enabled:
        return None, "name reading is switched off - type the name"
    from .config import anthropic_key_present
    if not anthropic_key_present():
        return None, "no ANTHROPIC_API_KEY set - type the name"
    try:
        return ClaudeNameReader(model), "Claude reads the handwritten name"
    except Exception as e:  # missing package, bad environment
        return None, f"name reading unavailable ({e.__class__.__name__}) - type the name"
