import numpy as np

from quiniela_scanner.name_reader import ClaudeNameReader, parse_reply


def test_parse_clean_json():
    r = parse_reply('{"name": "CLEMENTE O.", "jornada": "12", "legible": true}')
    assert (r.name, r.jornada, r.legible, r.error) == ("CLEMENTE O.", "12", True, None)


def test_parse_tolerates_fences_label_and_junk():
    r = parse_reply('Sure!\n```json\n{"name": "NOMBRE: TONY 5000", "jornada": "JORNADA 12", "legible": false}\n```')
    assert r.name == "TONY 5000" and r.jornada == "12" and r.legible is False


def test_parse_null_jornada_and_garbage():
    assert parse_reply('{"name": "A", "jornada": null, "legible": true}').jornada is None
    assert parse_reply("I cannot read this").error


class _Block:
    def __init__(self, t, text=""):
        self.type, self.text = t, text


class _Resp:
    stop_reason = "end_turn"
    content = [_Block("thinking"), _Block("text", '{"name": "JUJU 2", "jornada": "12", "legible": true}')]


class _Messages:
    def __init__(self): self.kwargs = None
    def create(self, **kw):
        self.kwargs = kw
        return _Resp()


def test_reader_sends_both_crops_and_effort_low_without_thinking_param():
    rd = ClaudeNameReader.__new__(ClaudeNameReader)
    import anthropic
    rd._anthropic, rd.model = anthropic, "claude-opus-5-5"
    rd.client = type("C", (), {"messages": _Messages()})()
    res = rd.read(np.zeros((100, 400, 3), np.uint8), np.zeros((50, 400, 3), np.uint8))
    kw = rd.client.messages.kwargs
    assert res.name == "JUJU 2"
    assert kw["model"] == "claude-opus-5-5" and kw["output_config"] == {"effort": "low"}
    assert "thinking" not in kw and "temperature" not in kw
    assert [c["type"] for c in kw["messages"][0]["content"]] == ["image", "image", "text"]


def test_refusal_falls_back_to_manual_typing():
    class Refused:
        stop_reason, content = "refusal", []

    class M:
        def create(self, **kw): return Refused()

    import anthropic
    rd = ClaudeNameReader.__new__(ClaudeNameReader)
    rd._anthropic, rd.model, rd.client = anthropic, "claude-opus-5-5", type("C", (), {"messages": M()})()
    res = rd.read(np.zeros((50, 200, 3), np.uint8))
    assert res.error and not res.name
