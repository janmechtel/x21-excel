"""
WebSocket client for communicating with the X21 deno server.
"""

import asyncio
import json
import logging
from typing import Callable, Optional

import websockets

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# Suppress noisy websockets debug logging
logging.getLogger("websockets").setLevel(logging.WARNING)


def _extract_delta_text(msg: dict) -> str:
    """Extract text content from a stream:delta message."""
    # Try various payload structures
    payload = msg.get("payload", {})

    # Direct text field
    if "text" in payload:
        return payload["text"]

    # Anthropic-style delta
    delta = payload.get("delta", {})
    if "text" in delta:
        return delta["text"]

    # content_block_delta style
    if msg.get("type") == "content_block_delta":
        delta = msg.get("delta", {})
        if "text" in delta:
            return delta["text"]

    return ""


def _extract_tool_ids(msg: dict) -> list:
    """Extract tool IDs from a tool permission message."""
    tools = msg.get("data", {}).get("tools", [])
    if not tools:
        tools = msg.get("tools", [])
    if not tools:
        tools = msg.get("payload", {}).get("tools", [])
    if not tools:
        tools = msg.get("payload", {}).get("toolPermissions", [])

    tool_ids = []
    for t in tools:
        tool_id = t.get("id") or t.get("toolId")
        if tool_id:
            tool_ids.append(tool_id)
    return tool_ids


def _extract_tool_use_id_from_stream(msg: dict) -> str:
    """Extract tool use ID from stream:delta content_block_start events."""
    if msg.get("type") != "stream:delta":
        return ""
    payload = msg.get("payload", {})
    if payload.get("type") != "content_block_start":
        return ""
    content_block = payload.get("content_block", {})
    if content_block.get("type") != "tool_use":
        return ""
    return content_block.get("id") or ""


class AgentClient:
    """Simple WebSocket client to send prompts and receive responses from X21 agent."""

    def __init__(self, ws_url: str = "ws://localhost:8000/ws", timeout: float = 120.0):
        """Initialize the client with URL and timeout."""
        self.ws_url = ws_url
        self.timeout = timeout
        self.ws: Optional[websockets.WebSocketClientProtocol] = None

    async def connect(self):
        """Connect to the deno server WebSocket."""
        logger.info(f"Connecting to {self.ws_url}")
        self.ws = await websockets.connect(self.ws_url)
        logger.info("Connected successfully")

    async def close(self):
        """Close the WebSocket connection."""
        if self.ws:
            await self.ws.close()
            self.ws = None
            logger.info("Connection closed")

    async def run_prompt(
        self,
        prompt: str,
        workbook_name: str,
        workbook_path: str,
        active_tools: Optional[list] = None,
        attachments: Optional[list] = None,
        on_message: Optional[Callable[[dict], None]] = None,
    ) -> dict:
        """
        Send a prompt to the agent and wait for completion.

        Args:
            prompt: The user prompt to send
            workbook_name: Name of the Excel workbook
            workbook_path: Full path to the workbook
            active_tools: List of tool names to enable (default: all tools)
            attachments: List of file attachments with {name, type, size, base64}
            on_message: Optional callback for each received message

        Returns:
            Dictionary with keys:
                - success: bool indicating if execution succeeded
                - usage: dict with input_tokens, output_tokens, total_tokens
                  (if available)
                - model: str model name (if available)
        """
        # Reconnect for each prompt to ensure clean WebSocket state
        if self.ws:
            try:
                await self.ws.close()
            except Exception:
                pass
        self.ws = await websockets.connect(self.ws_url)
        logger.debug(f"Reconnected to {self.ws_url} for workbook: {workbook_name}")

        # Wait briefly and drain any welcome messages
        await asyncio.sleep(0.1)
        try:
            # Non-blocking drain of any pending messages (like welcome)
            while True:
                try:
                    raw = await asyncio.wait_for(self.ws.recv(), timeout=0.2)
                    msg = json.loads(raw)
                    logger.debug(f"Drained message: {msg.get('type', 'unknown')}")
                except asyncio.TimeoutError:
                    break
        except Exception:
            pass

        if active_tools is None:
            active_tools = [
                "read_values_batch",
                "write_values_batch",
                "read_format_batch",
                "write_format_batch",
                "drag_formula",
                "add_sheets",
                "remove_columns",
                "add_columns",
                "remove_rows",
                "add_rows",
                "vba_create",
                "vba_read",
                "vba_update",
            ]

        # Send stream:start message
        payload = {"prompt": prompt, "activeTools": active_tools}

        # Add attachments if provided
        if attachments:
            payload["documentsBase64"] = attachments
            logger.info(f"Including {len(attachments)} attachment(s)")

        start_message = {
            "type": "stream:start",
            "workbookName": workbook_name,
            "workbookPath": workbook_path,
            "payload": payload,
        }
        logger.info(f"Sending stream:start for workbook: {workbook_name}")
        logger.debug(f"Full message: {json.dumps(start_message, indent=2)}")
        await self.ws.send(json.dumps(start_message))

        # Listen for messages until stream:end or timeout
        # Accumulate token usage across all LLM calls in the conversation
        # (matching how the UI handles usage updates)
        streaming_started = False
        usage_info = {"input_tokens": 0, "output_tokens": 0, "total_tokens": 0}
        model_name = None
        tool_use_ids = []

        try:
            async with asyncio.timeout(self.timeout):
                while True:
                    raw_msg = await self.ws.recv()
                    msg = json.loads(raw_msg)
                    msg_type = msg.get("type", "")

                    # Handle streaming text - print inline as it arrives
                    if msg_type in ["stream:delta", "content_block_delta"]:
                        text = _extract_delta_text(msg)
                        if text:
                            if not streaming_started:
                                print("\n📝 Agent: ", end="", flush=True)
                                streaming_started = True
                            print(text, end="", flush=True)

                        # Capture tool use IDs from content_block_start events
                        tool_use_id = _extract_tool_use_id_from_stream(msg)
                        if tool_use_id and tool_use_id not in tool_use_ids:
                            tool_use_ids.append(tool_use_id)

                        # Accumulate usage from stream:delta events (like the UI does)
                        if msg_type == "stream:delta":
                            payload = msg.get("payload", {})
                            event_type = payload.get("type", "")

                            # Extract usage from message_start event
                            if event_type == "message_start":
                                message = payload.get("message", {})
                                if "model" in message:
                                    model_name = message["model"]
                                if "usage" in message:
                                    msg_usage = message["usage"]
                                    usage_info["input_tokens"] += msg_usage.get(
                                        "input_tokens", 0
                                    )
                                    usage_info["output_tokens"] += msg_usage.get(
                                        "output_tokens", 0
                                    )
                                    logger.debug(
                                        "Captured usage from message_start: input=%s, "
                                        "output=%s",
                                        msg_usage.get("input_tokens", 0),
                                        msg_usage.get("output_tokens", 0),
                                    )

                            # Extract usage from message_delta event (output token
                            # updates)
                            elif event_type == "message_delta":
                                if "usage" in payload:
                                    delta_usage = payload["usage"]
                                    output_delta = delta_usage.get("output_tokens", 0)
                                    usage_info["output_tokens"] += output_delta
                                    logger.debug(
                                        "Captured output delta: %s",
                                        output_delta,
                                    )
                    else:
                        # End streaming line if we were streaming
                        if streaming_started:
                            print("\n", flush=True)
                            streaming_started = False

                        logger.info(f"<< Received: {msg_type}")
                        logger.debug(
                            f"   Full message: {json.dumps(msg, indent=2)[:500]}"
                        )

                    if on_message:
                        on_message(msg)

                    # Auto-approve tool permissions - check multiple possible type names
                    if msg_type in [
                        "tool:permission:request",
                        "tool:permission",
                        "tools",
                    ]:
                        for tool_id in _extract_tool_ids(msg):
                            if tool_id not in tool_use_ids:
                                tool_use_ids.append(tool_id)
                        logger.info(
                            "Tool permission request detected, auto-approving..."
                        )
                        await self._approve_tools(msg, workbook_name)

                    # Stream ended - success
                    elif msg_type == "stream:end":
                        if streaming_started:
                            print("\n", flush=True)
                        logger.info("Stream ended successfully")

                        # Calculate total tokens from accumulated values
                        usage_info["total_tokens"] = (
                            usage_info["input_tokens"] + usage_info["output_tokens"]
                        )

                        logger.info(
                            "Total accumulated usage: input=%s, output=%s, total=%s",
                            usage_info["input_tokens"],
                            usage_info["output_tokens"],
                            usage_info["total_tokens"],
                        )

                        return {
                            "success": True,
                            "usage": usage_info,
                            "model": model_name,
                            "tool_use_ids": tool_use_ids,
                        }

                    # Error from server
                    elif msg_type == "error":
                        if streaming_started:
                            print("\n", flush=True)
                        error_msg = msg.get("message") or msg.get("data", {}).get(
                            "message", "Unknown error"
                        )
                        logger.error(f"Agent error: {error_msg}")
                        usage_info["total_tokens"] = (
                            usage_info["input_tokens"] + usage_info["output_tokens"]
                        )
                        return {
                            "success": False,
                            "usage": usage_info,
                            "model": model_name,
                            "tool_use_ids": tool_use_ids,
                        }

        except asyncio.TimeoutError:
            logger.error(f"Timeout after {self.timeout}s waiting for agent response")
            usage_info["total_tokens"] = (
                usage_info["input_tokens"] + usage_info["output_tokens"]
            )
            return {
                "success": False,
                "usage": usage_info,
                "model": model_name,
                "tool_use_ids": tool_use_ids,
            }
        except Exception as e:
            logger.exception(f"Error during agent communication: {e}")
            usage_info["total_tokens"] = (
                usage_info["input_tokens"] + usage_info["output_tokens"]
            )
            return {
                "success": False,
                "usage": usage_info,
                "model": model_name,
                "tool_use_ids": tool_use_ids,
            }

    async def _approve_tools(self, msg: dict, workbook_name: str):
        """Auto-approve all pending tool permissions."""
        # Try different message structures
        tools = msg.get("data", {}).get("tools", [])
        if not tools:
            tools = msg.get("tools", [])
        if not tools:
            tools = msg.get("payload", {}).get("tools", [])
        if not tools:
            # Handle toolPermissions structure:
            # { payload: { toolPermissions: [{ toolId, toolName }] } }
            tools = msg.get("payload", {}).get("toolPermissions", [])

        logger.debug(f"Tools in message: {tools}")

        # Extract tool IDs - check both "id" and "toolId" properties
        tool_ids = []
        for t in tools:
            tool_id = t.get("id") or t.get("toolId")
            if tool_id:
                tool_ids.append(tool_id)

        if tool_ids:
            # Server expects toolResponses array with { toolId, decision } objects
            approval_msg = {
                "type": "tool:permission:response",
                "workbookName": workbook_name,
                "toolResponses": [
                    {"toolId": tid, "decision": "approved"} for tid in tool_ids
                ],
            }
            logger.info(f">> Sending approval for tools: {tool_ids}")
            logger.debug(f"   Full message: {json.dumps(approval_msg, indent=2)}")
            await self.ws.send(json.dumps(approval_msg))
        else:
            logger.warning("No tool IDs found in message to approve!")
            logger.warning("Message structure: %s", json.dumps(msg, indent=2)[:1000])

    async def revert_tools(
        self,
        workbook_name: str,
        tool_use_ids: list,
        timeout: float = 30.0,
        delay_seconds: float = 2,
    ) -> bool:
        """Revert tool changes and wait for idle status."""
        if not tool_use_ids:
            return True

        if not self.ws:
            self.ws = await websockets.connect(self.ws_url)

        for tool_use_id in tool_use_ids:
            message = {
                "type": "tool:revert",
                "workbookName": workbook_name,
                "toolUseId": tool_use_id,
            }
            await self.ws.send(json.dumps(message))

            # Give the server a brief head start to apply the revert
            if delay_seconds:
                await asyncio.sleep(delay_seconds)

            try:
                async with asyncio.timeout(timeout):
                    while True:
                        raw_msg = await self.ws.recv()
                        msg = json.loads(raw_msg)
                        if msg.get("type") == "status:update":
                            status = msg.get("payload", {}).get("status")
                            if status == "idle":
                                break
                            if status == "error":
                                return False
            except asyncio.TimeoutError:
                logger.error(f"Timeout waiting for revert idle status after {timeout}s")
                return False

        return True


async def test_connection(ws_url: str = "ws://localhost:8000/ws") -> bool:
    """Test if the deno server is reachable."""
    try:
        async with websockets.connect(ws_url) as ws:
            await ws.close()
        return True
    except Exception:
        return False
