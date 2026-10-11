"""Anonymous, opt-in web research through Parallel's Streamable HTTP MCP."""
from __future__ import annotations

import asyncio
from datetime import timedelta
from typing import Any


MCP_URL = "https://search.parallel.ai/mcp"
USER_AGENT = "harness-anything/1.0 (+https://github.com/yb2460/harness-anything)"


def _error_message(exc: Exception) -> str:
    # SDK transport failures can be wrapped in task-group exceptions on Python 3.10+.
    nested = getattr(exc, "exceptions", None)
    if nested:
        return "; ".join(_error_message(child) for child in nested)
    return str(exc) or type(exc).__name__


def call_tool(name: str, arguments: dict[str, Any], *, timeout: int = 60) -> dict[str, Any]:
    """Return the MCP result envelope; do not read API keys or saved credentials."""
    try:
        import anyio
        import httpx
        from mcp import ClientSession
        from mcp.client.streamable_http import streamable_http_client
    except ImportError as exc:
        raise RuntimeError('Web research requires the optional extra: python -m pip install -e ".[web]"') from exc

    async def run() -> dict[str, Any]:
        # Bound the complete operation, including connection, discovery and cleanup.
        with anyio.fail_after(timeout):
            async with httpx.AsyncClient(headers={"User-Agent": USER_AGENT}, timeout=timeout) as client:
                async with streamable_http_client(MCP_URL, http_client=client) as (read, write, _):
                    async with ClientSession(read, write, read_timeout_seconds=timedelta(seconds=timeout)) as session:
                        await session.initialize()
                        tools = await session.list_tools()
                        if name not in {tool.name for tool in tools.tools}:
                            raise RuntimeError(f"Parallel Search MCP does not advertise {name}")
                        result = await session.call_tool(name, arguments)
                        if result.isError:
                            detail = "; ".join(block.text for block in result.content if block.type == "text")
                            raise RuntimeError(detail or f"{name} returned an MCP tool error")
                        return result.model_dump(mode="json", exclude_none=True)

    try:
        return asyncio.run(run())
    except TimeoutError as exc:
        raise RuntimeError(f"Parallel Search MCP timed out after {timeout} seconds") from exc
    except Exception as exc:
        raise RuntimeError(f"Parallel Search MCP request failed: {_error_message(exc)}") from exc
