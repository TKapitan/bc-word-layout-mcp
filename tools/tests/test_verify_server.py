"""Tests for tools/release/verify_server.py, the release-channel smoke test (GitHub issue #112).

Run from the repo root:  python -m unittest discover -s tools/tests -v

verify_server is pointed at a tiny FAKE MCP server - a Python script written into a temp folder that
speaks the same newline-delimited JSON-RPC over stdin/stdout - so nothing here needs a build, and each
misbehaviour (a crash, an error response, silence) is one command-line argument away.
"""
import contextlib
import io
import os
import subprocess
import sys
import tempfile
import time
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "release")))

import verify_server  # noqa: E402

# python fake_mcp_server.py MODE TOOL_COUNT VERSION
FAKE_SERVER = r'''
import json
import sys

mode, tool_count, version = sys.argv[1], int(sys.argv[2]), sys.argv[3]


def send(message):
    sys.stdout.write(json.dumps(message) + "\n")
    sys.stdout.flush()


def result(method):
    if method == "initialize":
        return {"protocolVersion": "2024-11-05", "capabilities": {"tools": {}},
                "serverInfo": {"name": "FakeServer", "version": version}}
    return {"tools": [{"name": "tool_%02d" % i, "inputSchema": {"type": "object"}} for i in range(tool_count)]}


sys.stderr.write("fake server: started in mode %s\n" % mode)
sys.stderr.flush()
if mode == "eager":
    # Both answers go out before a single request is read: they are in the pipe before the client waits.
    send({"jsonrpc": "2.0", "id": 1, "result": result("initialize")})
    send({"jsonrpc": "2.0", "id": 2, "result": result("tools/list")})

for line in sys.stdin:
    if not line.strip():
        continue
    message = json.loads(line)
    method = message.get("method")
    if "id" not in message or mode in ("eager", "silent"):
        continue
    if mode == "init-error" and method == "initialize":
        send({"jsonrpc": "2.0", "id": message["id"],
              "error": {"code": -32602, "message": "Unsupported protocol version"}})
        continue
    if mode == "list-error" and method == "tools/list":
        send({"jsonrpc": "2.0", "id": message["id"],
              "error": {"code": -32603, "message": "tool registry unavailable"}})
        continue
    send({"jsonrpc": "2.0", "id": message["id"], "result": result(method)})
    if mode == "crash-after-initialize":
        sys.stderr.write("fake server: crashing right after initialize\n")
        sys.stderr.flush()
        sys.exit(3)
'''


class VerifyServerTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        cls.fake_script = os.path.join(cls.temp.name, "fake_mcp_server.py")
        with open(cls.fake_script, "w", encoding="utf-8") as f:
            f.write(FAKE_SERVER)

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def fake(self, mode, tools=verify_server.EXPECTED_TOOLS_DEFAULT, version="1.2.0-rc.1"):
        """'--' plus the command line that starts the fake server."""
        return ["--", sys.executable, self.fake_script, mode, str(tools), version]

    @staticmethod
    def verify(*argv):
        """verify_server.main in process: (exit code, stdout, stderr, seconds taken)."""
        out, err = io.StringIO(), io.StringIO()
        start = time.monotonic()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            try:
                code = verify_server.main(list(argv))
            except SystemExit as exit_:  # argparse's usage errors
                code = exit_.code
        return code, out.getvalue(), err.getvalue(), time.monotonic() - start

    # ---- passing ----

    def test_a_healthy_server_with_the_expected_tools_and_version_passes(self):
        code, out, err, _ = self.verify("--expect-tools", "3", "--expect-version", "1.2.0-rc.1",
                                        *self.fake("healthy", tools=3))

        self.assertEqual(0, code, err)
        self.assertIn("serverInfo: FakeServer 1.2.0-rc.1", out)
        self.assertIn("tools (3): tool_00, tool_01, tool_02", out)
        self.assertEqual("OK", out.splitlines()[-1])

    def test_without_expect_tools_the_server_must_list_the_default_count(self):
        code, out, err, _ = self.verify(*self.fake("healthy"))

        self.assertEqual(0, code, err)
        self.assertIn(f"tools ({verify_server.EXPECTED_TOOLS_DEFAULT}):", out)

    def test_responses_written_before_the_requests_are_read_are_not_missed(self):
        code, _, err, _ = self.verify(*self.fake("eager"))

        self.assertEqual(0, code, err)

    def test_immediate_responses_are_never_missed(self):
        # The lost-wakeup race: a response landing between the client's "not here yet" check and its wait.
        # The old Event clear()/wait() pair could turn that into a 300 s TIMEOUT for a server that had
        # answered; an instant-answer handshake, repeated, must never stall.
        for attempt in range(10):
            code, _, err, seconds = self.verify(*self.fake("healthy"))
            self.assertEqual(0, code, f"attempt {attempt}: {err}")
            self.assertLess(seconds, 10, f"attempt {attempt} stalled")

    # ---- failing: exit 1, with the server's stderr tail ----

    def test_a_different_version_fails_naming_both_and_prints_the_stderr_tail(self):
        code, _, err, _ = self.verify("--expect-version", "1.2.0", *self.fake("healthy"))

        self.assertEqual(1, code)
        self.assertIn("FAIL: expected serverInfo.version '1.2.0', got '1.2.0-rc.1'", err)
        self.assertIn("fake server: started in mode healthy", err)

    def test_the_four_part_assembly_version_is_not_the_release_version(self):
        # What the server reported before GitHub issue #112: every build of 1.1.0, RCs included, said 1.1.0.0.
        code, _, err, _ = self.verify("--expect-version", "1.1.0", *self.fake("healthy", version="1.1.0.0"))

        self.assertEqual(1, code)
        self.assertIn("got '1.1.0.0'", err)

    def test_a_different_tool_count_fails(self):
        code, _, err, _ = self.verify("--expect-tools", "26", *self.fake("healthy", tools=25))

        self.assertEqual(1, code)
        self.assertIn("FAIL: expected 26 tools, got 25", err)

    def test_a_server_that_crashes_after_initialize_fails_at_once_naming_its_exit_code(self):
        code, _, err, seconds = self.verify(*self.fake("crash-after-initialize"))

        self.assertEqual(1, code)
        self.assertIn("server exited with code 3", err)
        self.assertIn("fake server: crashing right after initialize", err)
        self.assertLess(seconds, 15, "a dead server must fail the check, not cost the 300 s request timeout")

    def test_an_initialize_error_fails_naming_the_error(self):
        code, out, err, _ = self.verify(*self.fake("init-error"))

        self.assertEqual(1, code)
        self.assertIn("FAIL: initialize returned an error", err)
        self.assertIn("Unsupported protocol version", err)
        self.assertIn("-32602", err)
        self.assertNotIn("OK", out)

    def test_a_tools_list_error_fails_naming_the_error_rather_than_as_zero_tools(self):
        code, _, err, _ = self.verify(*self.fake("list-error"))

        self.assertEqual(1, code)
        self.assertIn("FAIL: tools/list returned an error", err)
        self.assertIn("tool registry unavailable", err)
        self.assertNotIn("got 0", err)

    def test_a_silent_server_fails_after_the_timeout(self):
        code, _, err, seconds = self.verify("--timeout", "1", *self.fake("silent"))

        self.assertEqual(1, code)
        self.assertIn("FAIL: TIMEOUT: no initialize response within 1 s", err)
        self.assertLess(seconds, 15)

    # ---- usage: exit 2 ----

    def test_usage_errors_exit_2(self):
        python = sys.executable
        for argv in ([], ["--"], ["--expect-tools", "3"], ["--expect-tools", "abc", "--", python],
                     ["--expect-tools", "-1", "--", python], ["--timeout", "0", "--", python],
                     ["--expect-version"], ["--no-such-option", "--", python], [python, self.fake_script]):
            with self.subTest(argv=argv):
                code, _, err, _ = self.verify(*argv)
                self.assertEqual(2, code)
                self.assertIn("usage: verify_server.py", err)

    def test_options_after_the_separator_belong_to_the_server_command(self):
        # `-- dnx BcWordLayout.Mcp@1.2.3 --yes`: the command's own options must reach the command verbatim.
        args = verify_server.parse_args(["--expect-tools", "5", "--", "dnx", "--expect-tools", "--yes"])

        self.assertEqual(5, args.expect_tools)
        self.assertEqual(["dnx", "--expect-tools", "--yes"], args.command)

    def test_a_command_that_is_not_found_exits_2(self):
        code, _, err, _ = self.verify("--", "no-such-command-bcwl-112")

        self.assertEqual(2, code)
        self.assertIn("command not found: no-such-command-bcwl-112", err)

    def test_the_command_line_entry_point_returns_the_exit_codes(self):
        script = verify_server.__file__
        passed = subprocess.run([sys.executable, script, "--expect-tools", "2", *self.fake("healthy", tools=2)],
                                capture_output=True, text=True, timeout=60)
        failed = subprocess.run([sys.executable, script, "--expect-tools", "3", *self.fake("healthy", tools=2)],
                                capture_output=True, text=True, timeout=60)
        usage = subprocess.run([sys.executable, script], capture_output=True, text=True, timeout=60)

        self.assertEqual(0, passed.returncode, passed.stderr)
        self.assertEqual(1, failed.returncode, failed.stderr)
        self.assertEqual(2, usage.returncode, usage.stderr)


if __name__ == "__main__":
    unittest.main()
