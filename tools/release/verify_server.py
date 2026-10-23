"""Start an MCP server via any command and verify it answers initialize + tools/list.

Usage:  python tools/release/verify_server.py [--expect-tools N] [--expect-version V] [--timeout S]
                                             -- <command> [args...]

Exit 0 iff the server starts, completes the MCP handshake, tools/list returns the expected number of
tools (default 27), and - with --expect-version - serverInfo.version is exactly V. Exit 1 on any
failure, printing the tail of the server's stderr with it; exit 2 on a usage error or a command that is
not found. This is the release-channel smoke test: point it at `dnx ...`, a published exe from a release
zip, or `dotnet run` - the same check proves each install channel actually starts a working server
(release plan R3/G8). Pass the tag's version (1.2.3, or 1.2.3-rc.1 for an RC) as --expect-version to
prove the channel serves THAT build: the server reports its informational version, so an RC and the
release it becomes no longer look alike.

A server that dies fails the check at once, naming its exit code, instead of after the per-request
timeout (--timeout, default 300 s). Keeps stdin open until the responses arrive (stdio EOF race - see
the pilot-harness notes).
"""
from __future__ import annotations

import argparse
import collections
import json
import shutil
import subprocess
import sys
import threading
import time

EXPECTED_TOOLS_DEFAULT = 28
DEFAULT_TIMEOUT = 300    # seconds per request
POLL_INTERVAL = 0.5      # seconds between liveness checks of the server while a request is outstanding
EXIT_GRACE = 2.0         # seconds a dead server's last output gets to drain before a wait gives up on it
STDERR_TAIL_LINES = 40   # how much of the server's stderr a failure prints
USAGE = ("verify_server.py [--expect-tools N] [--expect-version V] [--timeout SECONDS] "
         "-- <command> [args...]")


class ServerFailure(Exception):
    """The server died, timed out, or answered a request with an error."""


def format_exit_code(code: int) -> str:
    """An exit code as printed in failure messages; a Windows status code (a crashed .NET host exits
    with 0xE0434352) also in the hex form it is documented under."""
    return f"{code} (0x{code & 0xFFFFFFFF:08X})" if code < 0 or code > 255 else str(code)


class StdioServer:
    """The server under test, spoken to as newline-delimited JSON-RPC over its stdin and stdout.

    A request waits on a condition that the stdout reader notifies under the same lock, so a response
    landing between "not here yet" and "wait" cannot be missed - the Event clear()/wait() pair this
    replaces could lose that wakeup and report a TIMEOUT for a server that had answered. The wait also
    re-checks the process every POLL_INTERVAL, so a server that dies fails the request at once. stderr
    is drained into a bounded buffer for the failure report.
    """

    def __init__(self, command: list[str]):
        self.proc = subprocess.Popen(command, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                     stderr=subprocess.PIPE)
        self.responses: dict = {}
        self.stdout_closed = False
        self.cond = threading.Condition()
        self.stderr_lines: collections.deque = collections.deque(maxlen=STDERR_TAIL_LINES)
        self.stdout_reader = threading.Thread(target=self._read_stdout, daemon=True)
        self.stderr_reader = threading.Thread(target=self._read_stderr, daemon=True)
        self.stdout_reader.start()
        self.stderr_reader.start()

    def _read_stdout(self):
        try:
            for line in self.proc.stdout:
                try:
                    msg = json.loads(line.strip() or "{}")
                except json.JSONDecodeError:
                    continue
                # A response carries an id and no method; one with both is a server-to-client request.
                if isinstance(msg, dict) and "id" in msg and "method" not in msg:
                    with self.cond:
                        self.responses[msg["id"]] = msg
                        self.cond.notify_all()
        finally:
            with self.cond:
                self.stdout_closed = True
                self.cond.notify_all()

    def _read_stderr(self):
        for line in self.proc.stderr:
            self.stderr_lines.append(line.decode("utf-8", "replace").rstrip())

    def _send(self, payload: dict):
        try:
            self.proc.stdin.write((json.dumps(payload) + "\n").encode())
            self.proc.stdin.flush()
        except (OSError, ValueError) as exc:  # the pipe is gone with the process
            raise ServerFailure(
                f"{self.exit_description()} before {payload['method']} could be sent ({exc})") from exc

    def notify(self, method: str):
        self._send({"jsonrpc": "2.0", "method": method})

    def request(self, mid: int, method: str, params: dict, timeout: float) -> dict:
        """Sends one request and returns its result object.

        Raises ServerFailure when the server dies (or closes its stdout) first, when `timeout` seconds
        pass without a response, or when the response is a JSON-RPC error.
        """
        self._send({"jsonrpc": "2.0", "id": mid, "method": method, "params": params})
        deadline = time.monotonic() + timeout
        gone_since = None
        with self.cond:
            while mid not in self.responses:
                now = time.monotonic()
                if self.stdout_closed or self.proc.poll() is not None:
                    # A response written just before the server went is still a response: unless the
                    # reader already hit end of stream, give it EXIT_GRACE to drain the pipe.
                    gone_since = now if gone_since is None else gone_since
                    if self.stdout_closed or now - gone_since >= EXIT_GRACE:
                        break
                if now >= deadline:
                    break
                self.cond.wait(min(deadline - now, POLL_INTERVAL))
            response = self.responses.pop(mid, None)
        if response is None:
            if gone_since is not None:
                raise ServerFailure(f"{self.exit_description()} while waiting for the {method} response")
            raise ServerFailure(f"TIMEOUT: no {method} response within {timeout:g} s")
        if "error" in response:
            raise ServerFailure(f"{method} returned an error: {json.dumps(response['error'])}")
        result = response.get("result")
        if not isinstance(result, dict):
            raise ServerFailure(f"the {method} response carries no result object: {json.dumps(response)[:300]}")
        return result

    def exit_description(self) -> str:
        try:
            return f"server exited with code {format_exit_code(self.proc.wait(timeout=EXIT_GRACE))}"
        except subprocess.TimeoutExpired:
            return "server closed its pipes without exiting"

    def close(self):
        """Ends the server: end of input first (a stdio MCP server exits on it), then terminate, then kill."""
        try:
            self.proc.stdin.close()
        except OSError:
            pass
        try:
            self.proc.wait(timeout=EXIT_GRACE)
        except subprocess.TimeoutExpired:
            self.proc.terminate()
            try:
                self.proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self.proc.kill()
                try:
                    self.proc.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    pass
        # Each reader stops at end of stream; only then can its pipe be closed (closing one a thread is
        # still blocked reading would block too).
        for reader, stream in ((self.stdout_reader, self.proc.stdout), (self.stderr_reader, self.proc.stderr)):
            reader.join(timeout=1.0)
            if not reader.is_alive():
                stream.close()

    def stderr_tail(self) -> list[str]:
        """What the server wrote to stderr, up to STDERR_TAIL_LINES lines - all of it once close() has run."""
        self.stderr_reader.join(timeout=1.0)
        return list(self.stderr_lines)


def parse_args(argv: list[str]) -> argparse.Namespace:
    """The options before '--' and the server command after it; a usage error exits 2."""
    parser = argparse.ArgumentParser(
        prog="verify_server.py", usage=USAGE, allow_abbrev=False,
        description="Start an MCP server via any command and verify it answers initialize + tools/list.")
    parser.add_argument("--expect-tools", type=int, default=EXPECTED_TOOLS_DEFAULT, metavar="N",
                        help=f"number of tools tools/list must return (default {EXPECTED_TOOLS_DEFAULT})")
    parser.add_argument("--expect-version", metavar="V",
                        help="exact serverInfo.version the server must report, e.g. 1.2.3 or 1.2.3-rc.1")
    parser.add_argument("--timeout", type=float, default=DEFAULT_TIMEOUT, metavar="SECONDS",
                        help=f"per-request timeout in seconds (default {DEFAULT_TIMEOUT})")
    # Everything after the first '--' is the server command, verbatim - its own options included.
    split = argv.index("--") if "--" in argv else len(argv)
    args = parser.parse_args(argv[:split])
    args.command = argv[split + 1:]
    if split == len(argv) or not args.command:
        parser.error("the server command must follow '--'")
    if args.expect_tools < 0:
        parser.error("--expect-tools must be 0 or more")
    if args.timeout <= 0:
        parser.error("--timeout must be more than 0")
    return args


def verify(server: StdioServer, args: argparse.Namespace) -> list[str]:
    """Runs the handshake and tools/list against the server; the problems found (empty: all good)."""
    try:
        init = server.request(1, "initialize", {"protocolVersion": "2024-11-05", "capabilities": {},
                                                "clientInfo": {"name": "release-verify", "version": "1"}},
                              args.timeout)
        server.notify("notifications/initialized")
        listed = server.request(2, "tools/list", {}, args.timeout)
    except ServerFailure as failure:
        return [str(failure)]

    server_info = init.get("serverInfo") if isinstance(init.get("serverInfo"), dict) else {}
    tools = listed.get("tools")
    if not isinstance(tools, list):
        return [f"the tools/list result carries no tools array: {json.dumps(listed)[:300]}"]
    names = sorted(str(t.get("name")) if isinstance(t, dict) else "?" for t in tools)
    print(f"serverInfo: {server_info.get('name')} {server_info.get('version')}")
    print(f"tools ({len(names)}): {', '.join(names)}")

    problems = []
    if len(names) != args.expect_tools:
        problems.append(f"expected {args.expect_tools} tools, got {len(names)}")
    if args.expect_version is not None and server_info.get("version") != args.expect_version:
        problems.append(f"expected serverInfo.version {args.expect_version!r}, "
                        f"got {server_info.get('version')!r}")
    return problems


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    command = list(args.command)
    # Resolve through PATH like a shell would: on Windows, CreateProcess only finds `dnx` if the
    # `.cmd` extension is spelled out, so `-- dnx ...` (the documented invocation) needs this.
    resolved = shutil.which(command[0])
    if resolved is None:
        print(f"command not found: {command[0]}", file=sys.stderr)
        return 2
    command[0] = resolved

    try:
        server = StdioServer(command)
    except OSError as exc:
        print(f"FAIL: could not start {command[0]}: {exc}", file=sys.stderr)
        return 1
    try:
        problems = verify(server, args)
    finally:
        server.close()

    if not problems:
        print("OK")
        return 0
    sys.stdout.flush()  # what the server did report comes first, also when both streams are piped
    for problem in problems:
        print(f"FAIL: {problem}", file=sys.stderr)
    tail = server.stderr_tail()
    if tail:
        print(f"--- server stderr, last {len(tail)} line(s) ---", file=sys.stderr)
    else:
        print("--- the server wrote nothing to stderr ---", file=sys.stderr)
    for line in tail:
        print(line, file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
