# -*- coding: utf-8 -*-
"""T5-2：LLM 流式输出测试。

覆盖：
  - _stream_enabled 受 LLM_STREAM 控制。
  - _chat_stream 流式消费 chunk、实时回显、返回完整文本。
  - _chat 在 LLM_STREAM=1 时走流式、流式失败降级非流式。
  - json_mode 流式仍带 response_format。
  - 默认（未设 LLM_STREAM）不启用流式，行为不变。

全部 monkeypatch _client，不触真实网络。
"""
from src import llm_enhancer


# ---------------------------------------------------------------------------
#  _stream_enabled
# ---------------------------------------------------------------------------
def test_stream_disabled_by_default(monkeypatch):
    monkeypatch.delenv("LLM_STREAM", raising=False)
    assert llm_enhancer._stream_enabled() is False


def test_stream_enabled_via_env(monkeypatch):
    monkeypatch.setenv("LLM_STREAM", "1")
    assert llm_enhancer._stream_enabled() is True


def test_stream_disabled_via_env_zero(monkeypatch):
    monkeypatch.setenv("LLM_STREAM", "0")
    assert llm_enhancer._stream_enabled() is False


# ---------------------------------------------------------------------------
#  _chat_stream：流式消费 + 回显 + 拼接
# ---------------------------------------------------------------------------
class _FakeDelta:
    def __init__(self, content):
        self.content = content


class _FakeChunk:
    def __init__(self, content):
        choice = type("C", (), {"delta": _FakeDelta(content)})()
        self.choices = [choice]


class _FakeStreamResp:
    """可迭代的伪流式响应。"""

    def __init__(self, chunks):
        self._chunks = chunks

    def __iter__(self):
        return iter(self._chunks)


def _patch_stream_client(monkeypatch, chunks_factory):
    """打桩 _client，其 create(stream=True) 返回 chunks_factory()。"""
    calls = []

    class FakeCompletions:
        @staticmethod
        def create(**kw):
            calls.append(kw)
            return _FakeStreamResp(chunks_factory())

    fake_client = type("Cl", (), {})()
    fake_client.chat = type("Ch", (), {"completions": FakeCompletions})()
    monkeypatch.setattr(llm_enhancer, "_client", lambda: fake_client)
    return calls


def test_chat_stream_echoes_and_joins(monkeypatch, capsys):
    monkeypatch.setenv("LLM_STREAM", "1")
    chunks = ["综述", "正文", "第一段", "[1]。"]
    calls = _patch_stream_client(monkeypatch, lambda: [_FakeChunk(c) for c in chunks])

    text = llm_enhancer._chat_stream("sys", "usr")
    assert text == "综述正文第一段[1]。"
    assert calls[0]["stream"] is True
    out = capsys.readouterr().out
    # 增量逐块回显
    for c in chunks:
        assert c in out


def test_chat_stream_skips_none_delta(monkeypatch, capsys):
    """无 content 的 chunk（如首个 role chunk）被跳过，不报错。"""
    monkeypatch.setenv("LLM_STREAM", "1")
    chunks = [_FakeChunk(None), _FakeChunk("ok"), _FakeChunk(None)]
    _patch_stream_client(monkeypatch, lambda: chunks)

    text = llm_enhancer._chat_stream("sys", "usr")
    assert text == "ok"
    assert "ok" in capsys.readouterr().out


def test_chat_stream_json_mode_sets_response_format(monkeypatch):
    monkeypatch.setenv("LLM_STREAM", "1")
    calls = _patch_stream_client(monkeypatch, lambda: [_FakeChunk('{"a":1}')])
    llm_enhancer._chat_stream("sys", "usr", json_mode=True)
    assert calls[0]["response_format"] == {"type": "json_object"}
    assert calls[0]["stream"] is True


# ---------------------------------------------------------------------------
#  _chat：流式优先 + 降级
# ---------------------------------------------------------------------------
def test_chat_uses_stream_when_enabled(monkeypatch):
    """LLM_STREAM=1 时 _chat 走流式分支。"""
    monkeypatch.setenv("LLM_STREAM", "1")
    calls = _patch_stream_client(monkeypatch,
                                 lambda: [_FakeChunk("流"), _FakeChunk("式")])
    assert llm_enhancer._chat("sys", "usr") == "流式"
    assert calls[0]["stream"] is True
    assert len(calls) == 1  # 只调用一次（流式成功，不再非流式重试）


def test_chat_falls_back_to_non_stream_on_failure(monkeypatch, capsys):
    """流式抛错时降级非流式，仍返回结果。"""
    monkeypatch.setenv("LLM_STREAM", "1")

    class _NonStreamResp:
        def __init__(self, content):
            msg = type("M", (), {"content": content})()
            self.choices = [type("C", (), {"message": msg})()]

    calls = []

    class FakeCompletions:
        @staticmethod
        def create(**kw):
            calls.append(kw)
            if kw.get("stream"):
                raise RuntimeError("端点不支持流式")
            return _NonStreamResp("非流式结果")

    fake_client = type("Cl", (), {})()
    fake_client.chat = type("Ch", (), {"completions": FakeCompletions})()
    monkeypatch.setattr(llm_enhancer, "_client", lambda: fake_client)

    result = llm_enhancer._chat("sys", "usr")
    assert result == "非流式结果"
    out = capsys.readouterr().out
    assert "流式输出失败" in out
    assert "降级非流式" in out
    # 先流式（失败）再非流式（成功）
    assert calls[0]["stream"] is True
    assert "stream" not in calls[1]


def test_chat_no_stream_when_disabled(monkeypatch):
    """默认（未设 LLM_STREAM）不启用流式，走非流式分支。"""
    monkeypatch.delenv("LLM_STREAM", raising=False)

    class _NonStreamResp:
        def __init__(self, content):
            msg = type("M", (), {"content": content})()
            self.choices = [type("C", (), {"message": msg})()]

    calls = []

    class FakeCompletions:
        @staticmethod
        def create(**kw):
            calls.append(kw)
            return _NonStreamResp("普通结果")

    fake_client = type("Cl", (), {})()
    fake_client.chat = type("Ch", (), {"completions": FakeCompletions})()
    monkeypatch.setattr(llm_enhancer, "_client", lambda: fake_client)

    assert llm_enhancer._chat("sys", "usr") == "普通结果"
    assert len(calls) == 1
    assert "stream" not in calls[0]


# ---------------------------------------------------------------------------
#  端到端：_chat_json 在流式开启时仍正确解析 JSON
# ---------------------------------------------------------------------------
def test_chat_json_works_with_streaming(monkeypatch):
    """LLM_STREAM=1 时 _chat_json 走流式但仍能解析出 JSON。"""
    monkeypatch.setenv("LLM_STREAM", "1")
    # 分块到达（模拟 JSON 被切成多块）
    parts = ['{"paras"', ': ["段一"', ', "段二"]}']
    _patch_stream_client(monkeypatch, lambda: [_FakeChunk(p) for p in parts])

    data = llm_enhancer._chat_json("sys", "usr")
    assert data == {"paras": ["段一", "段二"]}
