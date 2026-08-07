"""Start an MCP server via any command and verify it answers initialize + tools/list.

Usage:  python tools/release/verify_server.py [--expect-tools N] -- <command> [args...]

Exit 0 iff the server starts, completes the MCP handshake, and tools/list returns the expected
number of tools (default 26). This is the release-channel smoke test: point it at `dnx ...`, a
published exe from a release zip, or `dotnet run` - the same check proves each install channel
actually starts a working server (release plan R3/G8).

Keeps stdin open until the responses arrive (stdio EOF race - see the pilot-harness notes).
"""
import json
import shutil
import subprocess
import sys
import threading

EXPECTED_TOOLS_DEFAULT = 26


def main() -> int:
    argv = sys.argv[1:]
    expected = EXPECTED_TOOLS_DEFAULT
    if argv and argv[0] == "--expect-tools":
        expected = int(argv[1])
        argv = argv[2:]
    if not argv or argv[0] != "--":
        print("usage: verify_server.py [--expect-tools N] -- <command> [args...]", file=sys.stderr)
        return 2
    command = argv[1:]
    # Resolve through PATH like a shell would: on Windows, CreateProcess only finds `dnx` if the
    # `.cmd` extension is spelled out, so `-- dnx ...` (the documented invocation) needs this.
    resolved = shutil.which(command[0])
    if resolved is None:
        print(f"command not found: {command[0]}", file=sys.stderr)
        return 2
    command[0] = resolved

    proc = subprocess.Popen(command, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE)
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
    stderr_tail: list[bytes] = []
    threading.Thread(target=lambda: stderr_tail.extend(proc.stderr), daemon=True).start()

    def call(mid, method, params, timeout=300):
        proc.stdin.write((json.dumps({"jsonrpc": "2.0", "id": mid, "method": method, "params": params}) + "\n").encode())
        proc.stdin.flush()
        while mid not in responses:
            event.clear()
            if not event.wait(timeout):
                print(f"TIMEOUT waiting for {method}", file=sys.stderr)
                for tail_line in stderr_tail[-15:]:
                    sys.stderr.buffer.write(tail_line)
                proc.kill()
                sys.exit(1)
        return responses[mid]

    try:
        init = call(1, "initialize", {"protocolVersion": "2024-11-05", "capabilities": {},
                                      "clientInfo": {"name": "release-verify", "version": "1"}})
        server_info = init.get("result", {}).get("serverInfo", {})
        proc.stdin.write((json.dumps({"jsonrpc": "2.0", "method": "notifications/initialized"}) + "\n").encode())
        proc.stdin.flush()

        tools = call(2, "tools/list", {}).get("result", {}).get("tools", [])
        names = sorted(t["name"] for t in tools)
        print(f"serverInfo: {server_info.get('name')} {server_info.get('version')}")
        print(f"tools ({len(names)}): {', '.join(names)}")
        if len(names) != expected:
            print(f"FAIL: expected {expected} tools, got {len(names)}", file=sys.stderr)
            return 1
        print("OK")
        return 0
    finally:
        try:
            proc.stdin.close()
        except OSError:
            pass
        proc.terminate()


if __name__ == "__main__":
    sys.exit(main())
