from __future__ import annotations

import os
import subprocess
import sys
import unittest
from pathlib import Path
from unittest import mock

from click.testing import CliRunner

from cli_anything.illustrator.illustrator_cli import cli
from cli_anything.illustrator.utils import ai_backend


HARNESS_ROOT = Path(__file__).resolve().parents[3]


class PackagingTests(unittest.TestCase):
    def test_setup_reports_illustrator_metadata(self):
        name = subprocess.run(
            [sys.executable, "setup.py", "--name"],
            cwd=HARNESS_ROOT,
            capture_output=True,
            text=True,
            check=True,
        )
        version = subprocess.run(
            [sys.executable, "setup.py", "--version"],
            cwd=HARNESS_ROOT,
            capture_output=True,
            text=True,
            check=True,
        )

        self.assertEqual(name.stdout.strip(), "cli-anything-illustrator")
        self.assertEqual(version.stdout.strip(), "1.0.0")

    def test_module_entrypoint_help(self):
        env = os.environ.copy()
        env["PYTHONPATH"] = str(HARNESS_ROOT) + os.pathsep + env.get("PYTHONPATH", "")
        result = subprocess.run(
            [sys.executable, "-m", "cli_anything.illustrator", "--help"],
            cwd=HARNESS_ROOT,
            env=env,
            capture_output=True,
            text=True,
        )

        self.assertEqual(result.returncode, 0, msg=result.stderr)
        self.assertIn("cli-anything-illustrator", result.stdout)
        self.assertIn("project", result.stdout)
        self.assertIn("export", result.stdout)


class CliSmokeTests(unittest.TestCase):
    def test_help_does_not_connect_to_illustrator(self):
        runner = CliRunner()
        result = runner.invoke(cli, ["--help"])

        self.assertEqual(result.exit_code, 0, msg=str(result.exception))
        self.assertIn("project", result.output)
        self.assertIn("shape", result.output)

    def test_backend_detect_reports_missing_pywin32(self):
        with mock.patch.object(ai_backend, "_load_com", side_effect=RuntimeError("pywin32 unavailable")):
            result = ai_backend.detect_illustrator()

        self.assertFalse(result["com_available"])
        self.assertEqual(result["error"], "pywin32 unavailable")


if __name__ == "__main__":
    unittest.main()
