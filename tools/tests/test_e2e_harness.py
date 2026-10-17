"""Tests for the e2e harness in tools/e2e - its cache keys, failure handling and stdio client (GitHub
issue #112).

Run from the repo root:  python -m unittest discover -s tools/tests -v

Nothing here starts the real server, drives Word or needs a build: the cache keys are pointed at temp
folders, the runner at a scripted stand-in for McpServer, and McpServer itself at tiny fake MCP servers
written into a temp folder.
"""
import contextlib
import inspect
import io
import json
import os
import subprocess
import sys
import tempfile
import time
import unittest
from unittest import mock

E2E_DIR = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "e2e"))
sys.path.insert(0, E2E_DIR)

import bc_compare  # noqa: E402
import scenarios  # noqa: E402


def write(path, content, keep_timestamps=False):
    """Writes content (str); keep_timestamps restores the file's previous mtime, so only the bytes differ."""
    before = os.stat(path) if keep_timestamps else None
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
    if before is not None:
        os.utime(path, ns=(before.st_atime_ns, before.st_mtime_ns))


def ok(data):
    return {"ok": True, "data": data}


def failure(code, message, hint="Fix the call and retry."):
    return {"ok": False, "error": {"code": code, "message": message, "hint": hint}}


class ScriptedServer:
    """Stands in for McpServer: answers tool calls with healthy defaults unless `answers` (tool name ->
    envelope, or callable(args) -> envelope) says otherwise, and records every call."""

    def __init__(self, answers=None):
        self.answers = answers or {}
        self.calls = []
        self.dead = False
        self.closed = False

    def tool(self, name, args, image_dir=None):
        self.calls.append((name, args))
        answer = self.answers.get(name)
        if answer is not None:
            return answer(args) if callable(answer) else answer
        if name == "validate_layout":
            return ok({"passed": True, "errorCount": 0, "warningCount": 0, "findings": []})
        if name == "preview_layout":
            pdf = os.path.join(args["outputDir"], "preview-0001.pdf")
            os.makedirs(args["outputDir"], exist_ok=True)
            write(pdf, "%PDF-1.7 stand-in")
            return ok({"pdfPath": pdf, "conversionOk": True, "converterAvailable": True, "converterUsed": "fake",
                       "conversionError": None, "stats": {"fieldsFilled": 3, "rowsGenerated": 2, "unresolved": 0},
                       "warnings": []})
        if name == "render_preview_pages":
            return ok({"pagesRendered": 1, "pageCount": 1})
        return ok({"tables": [], "controls": []})

    def alive(self):
        return not self.dead

    def close(self):
        self.closed = True


def inspecting_scenario(ctx):
    ctx.expect(ctx.info() is not None, "the layout can be inspected")


def crashing_scenario(ctx):
    ctx.server.dead = True
    raise scenarios.ServerExited("server exited with code 7 while waiting for the tools/call response")


def key_error_scenario(ctx):
    return {}["data"]


def entry(name="t01-fake", layout="Fake.docx", fn=inspecting_scenario, **registration):
    return {"name": name, "layout": layout, "desc": "test scenario", "fn": fn,
            "allow_new_findings": False, "render_pages": 2, **registration}


class TempFolderTestCase(unittest.TestCase):
    """Points the harness's corpus, output and build folders at a fresh temp folder."""

    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = temp.name
        self.corpus = os.path.join(self.root, "corpus")
        self.out = os.path.join(self.root, "out")
        self.build = os.path.join(self.root, "build")
        for folder in (self.corpus, self.out, self.build):
            os.makedirs(folder)
        for name in ("Fake.docx", "Broken.docx"):
            write(os.path.join(self.corpus, name), f"stand-in for {name}")
        for name in ("BcWordLayout.McpHost.dll", "BcWordLayout.Domain.dll", "BcWordLayout.McpHost.deps.json"):
            write(os.path.join(self.build, name), f"original {name}")
        for attribute, value in (("CORPUS_DIR", self.corpus), ("OUT_ROOT", self.out),
                                 ("BASELINE_ROOT", os.path.join(self.out, "_baseline")),
                                 ("BUILD_DIR", self.build),
                                 ("HOST", os.path.join(self.build, "BcWordLayout.McpHost.dll"))):
            patcher = mock.patch.object(scenarios, attribute, value)
            patcher.start()
            self.addCleanup(patcher.stop)


# ---------------------------------------------------------------- cache keys

class BuildFingerprintTests(unittest.TestCase):
    BUILD_FILES = ("BcWordLayout.McpHost.dll", "BcWordLayout.Domain.dll", "BcWordLayout.Merge.dll",
                   "BcWordLayout.Render.dll", "BcWordLayout.McpHost.deps.json")
    OTHER_FILES = ("DocumentFormat.OpenXml.dll", "BcWordLayout.McpHost.exe", "BcWordLayout.McpHost.pdb",
                   "BcWordLayout.Domain.pdb", "BcWordLayout.McpHost.runtimeconfig.json")

    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.dir = temp.name
        for name in self.BUILD_FILES + self.OTHER_FILES:
            write(self.path(name), f"original {name}")

    def path(self, name):
        return os.path.join(self.dir, name)

    def fingerprint(self):
        return scenarios.build_fingerprint(self.dir)

    def test_it_is_stable_while_nothing_changes(self):
        first = self.fingerprint()

        self.assertNotEqual("no-build", first)
        self.assertEqual(first, self.fingerprint())

    def test_every_server_assembly_and_the_deps_json_is_part_of_it(self):
        # A same-size rewrite with the old timestamps restored: the change the mtime:size key used to miss,
        # on the assembly it used to ignore - a new LayoutValidator check only rebuilds BcWordLayout.Domain.dll.
        original = self.fingerprint()
        for name in self.BUILD_FILES:
            with self.subTest(name=name):
                write(self.path(name), f"modified {name}", keep_timestamps=True)
                self.assertNotEqual(original, self.fingerprint())
                write(self.path(name), f"original {name}")
                self.assertEqual(original, self.fingerprint())

    def test_files_outside_the_build_identity_do_not_change_it(self):
        original = self.fingerprint()
        for name in self.OTHER_FILES:
            write(self.path(name), f"modified {name}")

        self.assertEqual(original, self.fingerprint())

    def test_an_added_or_removed_server_assembly_changes_it(self):
        original = self.fingerprint()
        write(self.path("BcWordLayout.Extra.dll"), "new assembly")
        added = self.fingerprint()
        os.remove(self.path("BcWordLayout.Extra.dll"))
        os.remove(self.path("BcWordLayout.Render.dll"))

        self.assertNotEqual(original, added)
        self.assertNotIn(self.fingerprint(), (original, added))

    def test_a_missing_or_empty_build_folder_is_no_build(self):
        empty = os.path.join(self.dir, "empty")
        os.makedirs(empty)

        self.assertEqual("no-build", scenarios.build_fingerprint(os.path.join(self.dir, "missing")))
        self.assertEqual("no-build", scenarios.build_fingerprint(empty))

    def test_by_default_it_fingerprints_the_folder_the_host_is_in(self):
        self.assertEqual(os.path.dirname(scenarios.HOST), scenarios.BUILD_DIR)
        with mock.patch.object(scenarios, "BUILD_DIR", self.dir):
            self.assertEqual(self.fingerprint(), scenarios.build_fingerprint())


class CacheKeyTests(TempFolderTestCase):

    def test_a_domain_only_rebuild_invalidates_cached_passes_and_baselines(self):
        # GitHub issue #112: both keys used to hash BcWordLayout.McpHost.dll alone.
        scenario_before = scenarios.scenario_hash(entry())
        baseline_before = scenarios.baseline_stamp("Fake.docx")
        write(os.path.join(self.build, "BcWordLayout.Domain.dll"), "modified BcWordLayout.Domain.dll",
              keep_timestamps=True)

        self.assertNotEqual(scenario_before, scenarios.scenario_hash(entry()))
        self.assertNotEqual(baseline_before, scenarios.baseline_stamp("Fake.docx"))

    def test_the_keys_are_stable_for_the_same_inputs(self):
        self.assertEqual(scenarios.scenario_hash(entry(), "build-a"), scenarios.scenario_hash(entry(), "build-a"))
        self.assertEqual(scenarios.baseline_stamp("Fake.docx", "build-a"),
                         scenarios.baseline_stamp("Fake.docx", "build-a"))
        self.assertNotEqual(scenarios.scenario_hash(entry(), "build-a"), scenarios.scenario_hash(entry(), "build-b"))
        self.assertNotEqual(scenarios.baseline_stamp("Fake.docx", "build-a"),
                            scenarios.baseline_stamp("Fake.docx", "build-b"))

    def test_the_scenario_hash_is_keyed_on_the_shared_harness_code(self):
        with mock.patch.object(scenarios, "harness_source", return_value="harness, first version"):
            first = scenarios.scenario_hash(entry(), "build")
        with mock.patch.object(scenarios, "harness_source", return_value="harness, second version"):
            second = scenarios.scenario_hash(entry(), "build")

        self.assertNotEqual(first, second)

    def test_the_harness_source_covers_the_client_the_context_and_the_runner(self):
        source = scenarios.harness_source()

        for harness in (scenarios.McpServer, scenarios.Ctx, scenarios.tool_data, scenarios.render_preview,
                        scenarios.ensure_baseline, scenarios.run_scenario):
            self.assertIn(inspect.getsource(harness), source, harness.__name__)
        self.assertIn("def rows_off_grid", source)

    def test_the_scenario_hash_is_keyed_on_the_corpus_layout_it_edits_and_its_registration(self):
        original = scenarios.scenario_hash(entry(), "build")
        write(os.path.join(self.corpus, "Fake.docx"), "an edited corpus file")

        self.assertNotEqual(original, scenarios.scenario_hash(entry(), "build"))
        self.assertNotEqual(scenarios.scenario_hash(entry(), "build"),
                            scenarios.scenario_hash(entry(allow_new_findings=True), "build"))


# ---------------------------------------------------------------- failure envelopes

class EnvelopeHelperTests(unittest.TestCase):

    def test_tool_data_returns_the_data_of_a_success(self):
        self.assertEqual({"passed": True}, scenarios.tool_data("validate_layout", ok({"passed": True})))

    def test_tool_data_turns_a_failure_into_a_step_failure_naming_code_message_and_hint(self):
        # A failure envelope has no data key at all: this used to be KeyError: 'data'.
        envelope = failure("file_not_found", "Layout not found.", "Pass an existing path.")

        with self.assertRaises(scenarios.StepFailure) as raised:
            scenarios.tool_data("validate_layout", envelope)

        for fragment in ("validate_layout failed", "[file_not_found]", "Layout not found.",
                         "hint: Pass an existing path."):
            self.assertIn(fragment, str(raised.exception))

    def test_tool_data_handles_the_clients_own_failure_envelopes(self):
        # McpServer.tool's envelope for a JSON-RPC error: no hint, no data.
        envelope = {"ok": False, "error": {"code": "jsonrpc_error", "message": '{"code": -32603}'}}

        with self.assertRaises(scenarios.StepFailure) as raised:
            scenarios.tool_data("render_preview_pages", envelope)
        self.assertIn("[jsonrpc_error]", str(raised.exception))

    def test_a_success_without_data_is_a_step_failure_not_a_key_error(self):
        with self.assertRaises(scenarios.StepFailure):
            scenarios.tool_data("get_layout_info", {"ok": True})


# ---------------------------------------------------------------- runner

class RunScenarioTests(TempFolderTestCase):

    def run_scenario(self, server, scenario=None):
        with contextlib.redirect_stdout(io.StringIO()):
            return scenarios.run_scenario(server, scenario or entry(), "build")

    def baseline_report(self, layout="Fake"):
        return os.path.join(self.out, "_baseline", layout, "report.json")

    def test_a_passing_scenario_records_its_pages_and_the_baseline_is_captured_once(self):
        server = ScriptedServer()

        first = self.run_scenario(server)
        second = self.run_scenario(server)

        self.assertTrue(first["passed"], first.get("failure"))
        self.assertTrue(second["passed"], second.get("failure"))
        self.assertEqual("1/1", first["pages"])
        self.assertTrue(os.path.exists(self.baseline_report()))
        pristine = os.path.join(self.corpus, "Fake.docx")
        self.assertEqual(1, sum(1 for name, args in server.calls
                                if name == "validate_layout" and args["layoutPath"] == pristine))

    def test_a_baseline_that_cannot_be_captured_fails_only_that_scenario_and_is_not_cached(self):
        server = ScriptedServer({"validate_layout": failure("file_not_found", "Layout not found.")})

        report = self.run_scenario(server)

        self.assertFalse(report["passed"])
        self.assertIn("baseline of Fake.docx: validate_layout failed: [file_not_found] Layout not found.",
                      report["failure"])
        self.assertFalse(os.path.exists(self.baseline_report()))
        with open(os.path.join(self.out, "t01-fake", "done.json"), encoding="utf-8") as f:
            self.assertFalse(json.load(f)["passed"])

    def test_a_failed_render_fails_the_scenario(self):
        def render(args):
            return failure("rasterize_failed", "The PDF has no pages.") if "after" in args["pdfPath"] \
                else ok({"pagesRendered": 1, "pageCount": 1})

        report = self.run_scenario(ScriptedServer({"render_preview_pages": render}))

        self.assertFalse(report["passed"])
        self.assertIn("render_preview_pages failed: [rasterize_failed] The PDF has no pages.", report["failure"])

    def test_a_render_of_no_pages_fails_the_scenario(self):
        report = self.run_scenario(ScriptedServer({"render_preview_pages": ok({"pagesRendered": 0, "pageCount": 0})}))

        self.assertFalse(report["passed"])
        self.assertIn("rendered no page", report["failure"])

    def test_a_preview_whose_conversion_failed_fails_the_scenario(self):
        no_pdf = ok({"pdfPath": None, "conversionOk": False, "converterAvailable": False, "converterUsed": "word",
                     "conversionError": "Word is not installed.", "stats": {}, "warnings": []})

        report = self.run_scenario(ScriptedServer({"preview_layout": no_pdf}))

        self.assertFalse(report["passed"])
        self.assertIn("preview_layout produced no PDF", report["failure"])
        self.assertIn("Word is not installed.", report["failure"])

    def test_an_unexpected_exception_is_recorded_with_its_type_and_message(self):
        report = self.run_scenario(ScriptedServer(), entry(fn=key_error_scenario))

        self.assertFalse(report["passed"])
        self.assertEqual("KeyError: 'data'", report["failure"])
        self.assertIn("Traceback", report["trace"])


class MainTests(TempFolderTestCase):

    def main(self, scenario_list, factory, *argv):
        out = io.StringIO()
        with mock.patch.object(scenarios, "SCENARIOS", scenario_list), \
                mock.patch.object(scenarios, "McpServer", factory), contextlib.redirect_stdout(out):
            code = scenarios.main(list(argv))
        return code, out.getvalue()

    def test_a_failed_baseline_fails_one_scenario_and_the_run_goes_on_to_its_summary(self):
        def validate(args):
            return failure("file_not_found", "Layout not found.") if "Broken" in args["layoutPath"] \
                else ok({"passed": True, "errorCount": 0, "warningCount": 0, "findings": []})

        code, out = self.main([entry("t01-broken", "Broken.docx"), entry("t02-fine")],
                              lambda: ScriptedServer({"validate_layout": validate}))

        self.assertEqual(1, code)
        self.assertIn("FAIL  t01-broken: baseline of Broken.docx", out)
        self.assertIn("PASS  t02-fine", out)
        self.assertIn("2 selected, 0 skipped (cached), 1 failed ->", out)

    def test_a_dead_server_is_replaced_for_the_next_scenario(self):
        servers = []

        def factory():
            servers.append(ScriptedServer())
            return servers[-1]

        code, out = self.main([entry("t01-crash", fn=crashing_scenario), entry("t02-fine")], factory)

        self.assertEqual(1, code)
        self.assertEqual(2, len(servers))
        self.assertTrue(servers[0].closed)
        self.assertIn("FAIL  t01-crash: ServerExited: server exited with code 7", out)
        self.assertIn("PASS  t02-fine", out)

    def test_a_server_that_cannot_start_ends_the_run_with_a_summary(self):
        def factory():
            raise scenarios.ServerExited("server exited with code 1 while waiting for the initialize response")

        code, out = self.main([entry("t01-a"), entry("t02-b")], factory)

        self.assertEqual(1, code)
        self.assertIn("FAIL  t01-a: could not start the server: server exited with code 1", out)
        self.assertIn("2 selected, 0 skipped (cached), 1 failed, 1 not run", out)

    def test_a_cached_pass_is_skipped_until_a_server_assembly_changes(self):
        scenario_list = [entry()]
        _, first = self.main(scenario_list, ScriptedServer)
        _, second = self.main(scenario_list, ScriptedServer)
        write(os.path.join(self.build, "BcWordLayout.Domain.dll"), "modified BcWordLayout.Domain.dll",
              keep_timestamps=True)
        _, third = self.main(scenario_list, ScriptedServer)

        self.assertIn("PASS  t01-fake", first)
        self.assertIn("SKIP  t01-fake", second)
        self.assertIn("baseline for Fake.docx is stale", third)
        self.assertIn("PASS  t01-fake", third)


# ---------------------------------------------------------------- stdio client

# python fake_mcp_server.py MODE
FAKE_SERVER = r'''
import json
import sys

mode = sys.argv[1]


def send(message):
    sys.stdout.write(json.dumps(message) + "\n")
    sys.stdout.flush()


INITIALIZED = {"protocolVersion": "2024-11-05", "capabilities": {"tools": {}},
               "serverInfo": {"name": "FakeServer", "version": "0.0.1"}}

sys.stderr.write("fake server: started in mode %s\n" % mode)
sys.stderr.flush()
if mode == "eager":
    send({"jsonrpc": "2.0", "id": 1, "result": INITIALIZED})  # before the request has even been read

for line in sys.stdin:
    if not line.strip():
        continue
    message = json.loads(line)
    if "id" not in message:
        continue
    if message["method"] == "initialize":
        if mode == "init-error":
            send({"jsonrpc": "2.0", "id": message["id"],
                  "error": {"code": -32602, "message": "Unsupported protocol version"}})
        elif mode != "eager":
            send({"jsonrpc": "2.0", "id": message["id"], "result": INITIALIZED})
        continue
    name = message["params"]["name"]
    if mode == "crash-on-call":
        sys.stderr.write("fake server: dying on %s\n" % name)
        sys.stderr.flush()
        sys.exit(7)
    if mode == "silent-on-call":
        continue
    if mode == "server-request":
        # A server-to-client request that reuses the pending id must not be taken for the response.
        send({"jsonrpc": "2.0", "id": message["id"], "method": "roots/list"})
    envelope = {"ok": True, "data": {"tool": name, "arguments": message["params"]["arguments"]}}
    send({"jsonrpc": "2.0", "id": message["id"],
          "result": {"content": [{"type": "text", "text": json.dumps(envelope)}]}})
'''


class McpServerTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        cls.fake_script = os.path.join(cls.temp.name, "fake_mcp_server.py")
        with open(cls.fake_script, "w", encoding="utf-8") as f:
            f.write(FAKE_SERVER)

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def start(self, mode):
        server = scenarios.McpServer(command=[sys.executable, self.fake_script, mode])
        self.addCleanup(server.close)
        return server

    def test_a_tool_call_round_trips(self):
        server = self.start("healthy")

        self.assertEqual(ok({"tool": "get_layout_info", "arguments": {"layoutPath": "C:\\x.docx"}}),
                         server.tool("get_layout_info", {"layoutPath": "C:\\x.docx"}))
        self.assertTrue(server.alive())

    def test_immediate_responses_are_never_missed(self):
        # Each answer is written the moment the request is read; with the old Event clear()/wait() pair an
        # answer landing between the check and the clear was lost and the call sat out its timeout.
        server = self.start("healthy")
        for i in range(200):
            start = time.monotonic()
            self.assertEqual(i, server.tool("echo", {"i": i})["data"]["arguments"]["i"])
            self.assertLess(time.monotonic() - start, 5, f"call {i} stalled")

    def test_a_response_written_before_the_request_is_read_is_not_missed(self):
        self.assertTrue(self.start("eager").alive())

    def test_a_server_that_dies_mid_call_fails_it_at_once_with_its_exit_code_and_stderr(self):
        server = self.start("crash-on-call")
        start = time.monotonic()

        with self.assertRaises(scenarios.ServerExited) as raised:
            server.tool("preview_layout", {})

        self.assertLess(time.monotonic() - start, 15, "a dead server must not cost the 300 s timeout")
        self.assertIn("server exited with code 7 while waiting for the tools/call response", str(raised.exception))
        self.assertIn("fake server: dying on preview_layout", str(raised.exception))
        self.assertFalse(server.alive())

    def test_a_server_that_exits_before_answering_initialize_fails_the_start_at_once(self):
        # What `dotnet <missing dll>` does.
        crash = "import sys; sys.stderr.write('dll not found\\n'); sys.exit(1)"

        with self.assertRaises(scenarios.ServerExited) as raised:
            scenarios.McpServer(command=[sys.executable, "-c", crash])

        self.assertIn("server exited with code 1", str(raised.exception))
        self.assertIn("dll not found", str(raised.exception))

    def test_a_refused_handshake_fails_the_start(self):
        with self.assertRaises(RuntimeError) as raised:
            scenarios.McpServer(command=[sys.executable, self.fake_script, "init-error"])

        self.assertIn("Unsupported protocol version", str(raised.exception))

    def test_a_silent_server_times_out_with_its_stderr_tail(self):
        server = self.start("silent-on-call")

        with self.assertRaises(TimeoutError) as raised:
            server.request("tools/call", {"name": "echo", "arguments": {}}, timeout=0.5)

        self.assertIn("fake server: started in mode silent-on-call", str(raised.exception))
        self.assertTrue(server.alive())

    def test_a_server_to_client_request_is_not_taken_for_the_response(self):
        self.assertEqual("echo", self.start("server-request").tool("echo", {})["data"]["tool"])

    def test_importing_the_harness_modules_starts_nothing(self):
        # A fresh interpreter in which starting any process raises: the modules that reuse McpServer must
        # import cleanly with no server (and no build) behind them.
        code = ("import subprocess, sys\n"
                "def refuse(*args, **kwargs):\n"
                "    raise AssertionError('a process was started at import time')\n"
                "subprocess.Popen = refuse\n"
                f"sys.path.insert(0, {E2E_DIR!r})\n"
                "import scenarios, bc_compare, call, sandbox_pack\n")
        result = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, timeout=60)

        self.assertEqual(0, result.returncode, result.stderr)


# ---------------------------------------------------------------- bc_compare.py

class BcCompareTests(unittest.TestCase):

    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = temp.name
        self.pack = os.path.join(self.root, "sandbox-pack")
        self.host = os.path.join(self.root, "BcWordLayout.McpHost.dll")
        write(self.host, "stand-in host")

    def add_item(self, name, export=True, layout=True):
        item = os.path.join(self.pack, name)
        os.makedirs(os.path.join(item, "bc-output"))
        if layout:
            write(os.path.join(item, "Report.docx"), "stand-in layout")
        if export:
            write(os.path.join(item, "bc-output", "round1.xml"), "<ReportDataSet/>")

    def main(self, factory, *argv, host=None, pack=None):
        out, err = io.StringIO(), io.StringIO()
        with mock.patch.object(bc_compare, "HOST", host or self.host), \
                mock.patch.object(bc_compare, "PACK_ROOT", pack or self.pack), \
                mock.patch.object(bc_compare, "McpServer", factory), \
                contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = bc_compare.main(list(argv))
        return code, out.getvalue(), err.getvalue()

    @staticmethod
    def no_server():
        raise AssertionError("the server must not be started")

    def test_exit_code_is_1_only_for_a_failed_row(self):
        self.assertEqual(0, bc_compare.exit_code([]))
        self.assertEqual(0, bc_compare.exit_code([("p01", "ok", "pages=1/1"), ("p02", "skipped", "no layout")]))
        self.assertEqual(1, bc_compare.exit_code([("p01", "ok", "pages=1/1"), ("p02", "FAILED", "no PDF")]))

    def test_no_server_build_exits_2_before_starting_anything(self):
        code, _, err = self.main(self.no_server, host=os.path.join(self.root, "missing.dll"))

        self.assertEqual(2, code)
        self.assertIn("server not built:", err)
        self.assertIn("dotnet build -c Release", err)

    def test_no_pack_exits_2_before_starting_anything(self):
        code, _, err = self.main(self.no_server)

        self.assertEqual(2, code)
        self.assertIn(f"no sandbox pack at {self.pack}", err)

    def test_no_matching_item_exits_2(self):
        self.add_item("p01-something")

        code, _, err = self.main(self.no_server, "p09")

        self.assertEqual(2, code)
        self.assertIn("no pack item matches ['p09']", err)

    def test_any_failed_item_makes_the_run_exit_1_and_every_item_gets_its_row(self):
        for name, export in (("p01-good", True), ("p02-no-export", False), ("p03-bad-render", True),
                             ("p04-no-pdf", True), ("p05-preview-refused", True)):
            self.add_item(name, export=export)
        default = ScriptedServer()

        def preview(args):
            if "p04" in args["layoutPath"]:
                return ok({"pdfPath": None, "conversionOk": False, "converterUsed": "word",
                           "conversionError": "Word timed out.", "stats": {}, "warnings": []})
            if "p05" in args["layoutPath"]:
                return failure("invalid_argument", "dataOverridesPath is not a dataset export.")
            return default.tool("preview_layout", args)

        def render(args):
            return failure("rasterize_failed", "Damaged PDF.") if "p03" in args["pdfPath"] \
                else ok({"pagesRendered": 2, "pageCount": 2})

        code, out, _ = self.main(lambda: ScriptedServer({"preview_layout": preview, "render_preview_pages": render}))

        self.assertEqual(1, code)
        rows = dict(line.split(None, 1) for line in out.splitlines())
        self.assertTrue(rows["p01-good"].startswith("ok"), rows["p01-good"])
        self.assertIn("pages=2/2", rows["p01-good"])
        self.assertTrue(rows["p02-no-export"].startswith("skipped"))
        self.assertIn("render_preview_pages failed: [rasterize_failed] Damaged PDF.", rows["p03-bad-render"])
        self.assertIn("no PDF (converter word): Word timed out.", rows["p04-no-pdf"])
        self.assertIn("preview_layout failed: [invalid_argument]", rows["p05-preview-refused"])
        for failed in ("p03-bad-render", "p04-no-pdf", "p05-preview-refused"):
            self.assertTrue(rows[failed].startswith("FAILED"), rows[failed])
        self.assertTrue(os.path.exists(os.path.join(self.pack, "p01-good", "bc-data", "mock-preview.pdf")))

    def test_only_ok_and_skipped_items_exit_0(self):
        self.add_item("p01-good")
        self.add_item("p02-no-layout", layout=False)

        code, out, _ = self.main(ScriptedServer)

        self.assertEqual(0, code, out)

    def test_an_item_that_raises_is_a_failed_row_and_the_others_still_run(self):
        self.add_item("p01-crash")
        self.add_item("p02-good")
        default = ScriptedServer()

        def preview(args):
            if "p01" in args["layoutPath"]:
                raise scenarios.ServerExited("server exited with code 7 while waiting for the tools/call response")
            return default.tool("preview_layout", args)

        code, out, _ = self.main(lambda: ScriptedServer({"preview_layout": preview}))

        self.assertEqual(1, code)
        self.assertIn("FAILED   ServerExited: server exited with code 7", out)
        self.assertRegex(out, r"p02-good\s+ok")


if __name__ == "__main__":
    unittest.main()
