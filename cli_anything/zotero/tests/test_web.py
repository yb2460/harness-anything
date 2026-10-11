"""Exercise native CLI dispatch through the real SDK and observed HTTP requests."""
from __future__ import annotations

import builtins
import json
from unittest import mock

import pytest
from click.testing import CliRunner

from cli_anything.zotero.utils import parallel_mcp
from cli_anything.zotero.zotero_cli import cli


def test_help_and_missing_extra_do_not_require_zotero():
    runner = CliRunner()
    assert runner.invoke(cli, ["web", "--help"]).exit_code == 0
    original_import = builtins.__import__

    def missing_mcp(name, *args, **kwargs):
        if name == "mcp":
            raise ImportError("Optional dependency not installed")
        return original_import(name, *args, **kwargs)

    with mock.patch("builtins.__import__", side_effect=missing_mcp):
        result = runner.invoke(cli, ["web", "search", "Research", "--query", "research source methods"])
    assert result.exit_code == 1
    assert 'python -m pip install -e ".[web]"' in result.output


@pytest.fixture
def mcp_http(monkeypatch):
    httpx = pytest.importorskip("httpx")
    pytest.importorskip("mcp")
    observed = []
    behavior = {"error": False, "status": 200, "missing_tool": False, "delay": False}
    original_client = httpx.AsyncClient

    async def respond(request):
        import anyio

        assert str(request.url) == parallel_mcp.MCP_URL
        assert request.headers["User-Agent"] == parallel_mcp.USER_AGENT
        assert "Authorization" not in request.headers
        payload = json.loads(request.content)
        observed.append(payload)
        if behavior["status"] != 200:
            return httpx.Response(behavior["status"], headers={"Retry-After": "60"})
        method = payload["method"]
        if method.startswith("notifications/"):
            return httpx.Response(202)
        if method == "initialize":
            result = {"protocolVersion": payload["params"]["protocolVersion"], "capabilities": {"tools": {}},
                      "serverInfo": {"name": "protocol-fixture", "version": "1.0"}}
        elif method == "tools/list":
            names = [] if behavior["missing_tool"] else ["web_search", "web_fetch"]
            result = {"tools": [{"name": name, "inputSchema": {"type": "object"}} for name in names]}
        elif method == "tools/call":
            if behavior["delay"]:
                await anyio.sleep(5)
            result = {"isError": behavior["error"], "content": [{"type": "text", "text": "fixture failure" if behavior["error"] else "Useful research excerpts"}],
                      "structuredContent": {"results": [{"url": "https://example.org/paper", "title": "A paper", "excerpts": ["Useful research excerpts"]}]}}
        else:
            raise AssertionError(f"Unexpected MCP method: {method}")
        return httpx.Response(200, json={"jsonrpc": "2.0", "id": payload["id"], "result": result})

    def client(**kwargs):
        return original_client(transport=httpx.MockTransport(respond), **kwargs)

    monkeypatch.setattr(httpx, "AsyncClient", client)
    # Neither library discovery nor session loading belongs in the web path.
    monkeypatch.setattr("cli_anything.zotero.zotero_cli.current_runtime", lambda *_: pytest.fail("Zotero discovery called"))
    monkeypatch.setattr("cli_anything.zotero.zotero_cli.current_session", lambda: pytest.fail("Saved session loaded"))
    monkeypatch.setenv("PARALLEL_API_KEY", "must-not-be-sent")
    monkeypatch.setenv("OPENAI_API_KEY", "must-not-be-sent")
    return observed, behavior


def test_search_native_dispatch_and_anonymous_requests(mcp_http):
    observed, _ = mcp_http
    result = CliRunner().invoke(cli, ["--json", "web", "search", "Find research methods", "--query", "literature review methods",
                                    "--query", "reproducible review workflow", "--session-id", "research-session"])
    assert result.exit_code == 0, result.output
    output = json.loads(result.output)
    assert output["structuredContent"]["results"][0]["url"] == "https://example.org/paper"
    assert output["content"][0]["text"] == "Useful research excerpts"
    assert {item["method"] for item in observed} >= {"initialize", "notifications/initialized", "tools/list", "tools/call"}
    call = next(item["params"] for item in observed if item["method"] == "tools/call")
    assert call == {"name": "web_search", "arguments": {"objective": "Find research methods", "search_queries": ["literature review methods", "reproducible review workflow"], "session_id": "research-session"}}


def test_fetch_native_dispatch_preserves_options(mcp_http):
    observed, _ = mcp_http
    result = CliRunner().invoke(cli, ["--json", "web", "fetch", "https://example.org/paper", "http://example.org/other", "--objective", "Research methods",
                                    "--query", "literature review methods", "--full-content", "--session-id", "research-session"])
    assert result.exit_code == 0, result.output
    assert json.loads(result.output)["structuredContent"]["results"][0]["excerpts"]
    call = next(item["params"] for item in observed if item["method"] == "tools/call")
    assert call == {"name": "web_fetch", "arguments": {"urls": ["https://example.org/paper", "http://example.org/other"], "objective": "Research methods",
                                                        "search_queries": ["literature review methods"], "full_content": True, "session_id": "research-session"}}


@pytest.mark.parametrize("behavior, expected", [("error", "fixture failure"), ("missing_tool", "does not advertise web_search"), ("delay", "timed out after 1 seconds")])
def test_protocol_failures_exit_nonzero(mcp_http, behavior, expected):
    _, state = mcp_http
    state[behavior] = True
    result = CliRunner().invoke(cli, ["web", "search", "Research", "--query", "literature review methods", "--timeout", "1"])
    assert result.exit_code == 1
    assert expected in result.output
    assert "Traceback" not in result.output


def test_http_failure_exits_nonzero(mcp_http):
    _, behavior = mcp_http
    behavior["status"] = 429
    result = CliRunner().invoke(cli, ["web", "search", "Research", "--query", "literature review methods"])
    assert result.exit_code == 1
    assert "429" in result.output


@pytest.mark.parametrize("args", [
    ["search", "Research"],
    ["search", " ", "--query", "valid query"],
    ["search", "Research", "--query", " "],
    ["search", "Research", "--query", "valid query", "--session-id", "a" * 101],
    ["fetch", "file:///local/paper"],
    ["fetch", "https://example.org", "--objective", "a" * 201],
    ["fetch", "https://example.org", "--timeout", "0"],
    ["fetch"] + ["https://example.org"] * 21,
])
def test_invalid_arguments_never_connect(args):
    with mock.patch.object(parallel_mcp, "call_tool", side_effect=AssertionError("Unexpected network call")):
        result = CliRunner().invoke(cli, ["web"] + args)
    assert result.exit_code in {1, 2}
