"""One-shot MCP tool call against the Release-built BcWordLayout server.

Usage:  python tools/e2e/call.py TOOL_NAME 'JSON_ARGS' [--images OUTDIR]

Prints the envelope JSON (first text block) to stdout. If --images OUTDIR is given, any image
content blocks (render_preview_pages) are saved there as page-N.png. Keeps stdin open until the
response arrives (stdio EOF race - see the pilot-harness notes). Build first: dotnet build -c Release.
"""
import base64
import json
import os
import subprocess
import sys
import threading

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
HOST = os.path.join(REPO, "src", "BcWordLayout.McpHost", "bin", "Release", "net10.0", "BcWordLayout.McpHost.dll")


def main() -> int:
    tool_name = sys.argv[1]
    args = json.loads(sys.argv[2])
    img_dir = sys.argv[4] if len(sys.argv) > 4 and sys.argv[3] == "--images" else None

    proc = subprocess.Popen(["dotnet", HOST], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                            stderr=subprocess.DEVNULL, cwd=REPO)
    responses: dict[int, dict] = {}
    event = threading.Event()

    def reader():
        for line in proc.stdout:
            try:
                msg = json.loads(line.strip() or "{}")
            except json.JSONDecodeError:
                continue
            if "id" in msg:
                responses[msg["id"]] = msg
                event.set()

    threading.Thread(target=reader, daemon=True).start()

    def call(mid, method, params, timeout=400):
        proc.stdin.write((json.dumps({"jsonrpc": "2.0", "id": mid, "method": method, "params": params}) + "\n").encode())
        proc.stdin.flush()
        while mid not in responses:
            event.clear()
            if not event.wait(timeout):
                print(f"TIMEOUT on {method}", file=sys.stderr)
                proc.kill()
                sys.exit(1)
        return responses[mid]

    call(1, "initialize", {"protocolVersion": "2024-11-05", "capabilities": {},
                           "clientInfo": {"name": "one-shot-call", "version": "1"}})
    proc.stdin.write((json.dumps({"jsonrpc": "2.0", "method": "notifications/initialized"}) + "\n").encode())
    proc.stdin.flush()

    resp = call(2, "tools/call", {"name": tool_name, "arguments": args})
    result = resp.get("result")
    if result is None:
        print(json.dumps(resp.get("error", {"error": "no result"}), indent=2))
        proc.stdin.close()
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

    proc.stdin.close()
    try:
        proc.wait(timeout=10)
    except subprocess.TimeoutExpired:
        proc.kill()
    return 0


if __name__ == "__main__":
    sys.exit(main())
