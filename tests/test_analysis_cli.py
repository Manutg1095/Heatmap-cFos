"""Pruebas de portabilidad; no requieren los datos del estudio."""
from pathlib import Path
import os
import subprocess
import sys
import tempfile
import unittest

ANALYSIS = Path(__file__).resolve().parents[1] / "analysis"
SCRIPTS = ("cfos_final_analysis", "cfos_extra_analysis", "cfos_3d_support", "make_cfos_figures")


class AnalysisCLITests(unittest.TestCase):
    def test_scripts_import_without_running_analysis(self):
        with tempfile.TemporaryDirectory(prefix="cfos import ") as folder:
            env = dict(os.environ, PYTHONPATH=str(ANALYSIS), MPLBACKEND="Agg")
            result = subprocess.run(
                [sys.executable, "-c", "; ".join("import " + name for name in SCRIPTS)],
                cwd=folder, env=env, capture_output=True, text=True,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(result.stdout, "")
            self.assertEqual(list(Path(folder).iterdir()), [])

    def test_help_exposes_portable_paths_from_other_directory(self):
        with tempfile.TemporaryDirectory(prefix="cfos cli ") as folder:
            for script in SCRIPTS:
                with self.subTest(script=script):
                    result = subprocess.run(
                        [sys.executable, str(ANALYSIS / (script + ".py")), "--help"],
                        cwd=folder, capture_output=True, text=True,
                    )
                    self.assertEqual(result.returncode, 0, result.stderr)
                    self.assertIn("--input", result.stdout)
                    self.assertIn("--output", result.stdout)
                    if script == "make_cfos_figures":
                        self.assertIn("--results", result.stdout)


if __name__ == "__main__":
    unittest.main()
