"""One-shot MCP tool call against the Release-built BcWordLayout server.

Usage:  python tools/e2e/call.py TOOL_NAME 'JSON_ARGS' [--images OUTDIR]

Prints the envelope JSON (first text block) to stdout. If --images OUTDIR is given, any image
content blocks (render_preview_pages) are saved there as page-N.png. Keeps stdin open until the
response arrives (stdio EOF race - see the pilot-harness notes). Build first: dotnet build -c Release.

Runs the server through scenarios.McpServer, so a server that dies fails the call at once with its
exit code and the tail of its stderr (exit 1), instead of after the 400 s timeout.
"""
import base64
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from scenarios import HOST, McpServer  # noqa: E402

CALL_TIMEOUT = 400  # seconds: the slowest legitimate call is a long preview_layout conversion


def main() -> int:
    if len(sys.argv) < 3:
        print("usage: call.py TOOL_NAME 'JSON_ARGS' [--images OUTDIR]", file=sys.stderr)
        return 2
    tool_name = sys.argv[1]
    args = json.loads(sys.argv[2])
    img_dir = sys.argv[4] if len(sys.argv) > 4 and sys.argv[3] == "--images" else None
    if not os.path.exists(HOST):
        print(f"server not built: {HOST}\nrun: dotnet build -c Release", file=sys.stderr)
        return 2

    try:
        server = McpServer()
    except (RuntimeError, OSError) as exc:  # ServerExited, TimeoutError, a refused handshake
        print(f"could not start the server: {exc}", file=sys.stderr)
        return 1
    try:
        resp = server.request("tools/call", {"name": tool_name, "arguments": args}, timeout=CALL_TIMEOUT)
    except (RuntimeError, OSError) as exc:  # ServerExited, TimeoutError
        print(str(exc), file=sys.stderr)
        return 1
    finally:
        server.close()

    result = resp.get("result")
    if result is None:
        print(json.dumps(resp.get("error", {"error": "no result"}), indent=2))
        return 1

    envelope, saved = None, 0
    for block in result.get("content", []):
        if block.get("type") == "text" and envelope is None:
            try:
                envelope = json.loads(block["text"])
            except json.JSONDecodeError:
                envelope = {"RAW_NON_JSON_TEXT": block["text"]}
        elif block.get("type") == "image" and img_dir:
            # Name by the REAL page number (render_preview_pages' envelope pages[] is in block order),
            # so a firstPage>1 render never overwrites page-1.png.
            pages = ((envelope or {}).get("data") or {}).get("pages") or []
            page_no = pages[saved]["pageNumber"] if saved < len(pages) else saved + 1
            saved += 1
            os.makedirs(img_dir, exist_ok=True)
            with open(os.path.join(img_dir, f"page-{page_no}.png"), "wb") as f:
                f.write(base64.b64decode(block["data"]))

    print(json.dumps(envelope, indent=2))
    if img_dir:
        print(f"[saved {saved} png(s) -> {img_dir}]", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
