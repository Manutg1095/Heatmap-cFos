import ast
import tempfile
import unittest
from pathlib import Path

import numpy as np
import pandas as pd

import app
import cfos_core as core


ROOT = Path(__file__).resolve().parent


class CompatibilityTests(unittest.TestCase):
    def test_version_and_splash_contract(self):
        self.assertEqual(core.APP_VERSION, "1.2.4")
        self.assertEqual(app.PRODUCT_NAME, "Heatmap V1.2")
        self.assertGreaterEqual(app.SPLASH_MIN_MS, 7000)
        self.assertEqual(
            app.OWNERSHIP_TEXT,
            "Propiedad del Laboratorio de neurobiología conductual VMTG UNAM",
        )
        self.assertTrue(app.SPLASH_IMAGE.exists())

    def test_splash_delay_is_used_by_the_application(self):
        tree = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
        names = {node.id for node in ast.walk(tree) if isinstance(node, ast.Name)}
        self.assertIn("SPLASH_MIN_MS", names)
        source = (ROOT / "app.py").read_text(encoding="utf-8")
        self.assertIn("SPLASH_MIN_MS - elapsed_ms", source)

    def test_name_parser_accepts_both_experiment_families(self):
        hb = core.parse_name("Statistics for HB_04_R1_C5_BLAa_20X.xlsx")
        scop = core.parse_name("Statistics for SCOP_LiCl_03_R81_C7_BLAp_20X.xlsx")
        self.assertEqual(hb["Stack_ID"], "HB_04_R1_C5_BLAa")
        self.assertEqual(scop["Stack_ID"], "SCOP_LiCl_03_R81_C7_BLAp")

    def test_catalogs_are_unique_and_every_cut_has_a_group(self):
        with tempfile.TemporaryDirectory() as directory:
            groups_path = Path(directory) / "groups.xlsx"
            bregmas_path = Path(directory) / "bregmas.xlsx"
            pd.DataFrame([{"Experimento": "HB_04", "Animal": "R1", "Grupo": "SCOP-HB"}]).to_excel(groups_path, index=False)
            pd.DataFrame([{"Experimento": "HB_04", "Animal y corte": "R1 C3", "Región": "BLAa", "Bregma": -2.12}]).to_excel(bregmas_path, index=False)
            groups = core.load_group_catalog(groups_path)
            bregmas = core.load_bregma_catalog(bregmas_path)
        joined = bregmas.merge(groups, on=["Experimento", "Animal"], how="left")
        self.assertTrue(joined["Grupo"].notna().all())
        parsed = core.parse_name("HB_04_R1_C3_BLAa_20X.xlsx")
        resolved = app.resolve_stack_metadata(parsed, groups, bregmas)
        self.assertEqual(resolved["Grupo"], "SCOP-HB")
        self.assertAlmostEqual(resolved["Bregma_mm"], -2.12)
        self.assertEqual(resolved["Fuente_bregma"], "Registro maestro")

    def test_area_without_bregma_remains_valid(self):
        with tempfile.TemporaryDirectory() as directory:
            sample = Path(directory) / "area_sin_bregma.xlsx"
            pd.DataFrame(
                {
                    "Area": [503_883.133],
                    "Bregma": [np.nan],
                    "Volume": [9_916_420.05744],
                }
            ).to_excel(sample, index=False)
            info = core.load_area(sample)
        self.assertGreater(info["Area_um2"], 0)
        self.assertGreater(info["Volumen_um3"], 0)
        self.assertNotIn("Bregma_area_mm", info)

    def test_area_bregma_is_preserved_when_present(self):
        with tempfile.TemporaryDirectory() as directory:
            sample = Path(directory) / "area_con_bregma.xlsx"
            pd.DataFrame(
                {"Area": [500_000], "Bregma": [-2.12], "Volume": [10_000_000]}
            ).to_excel(sample, index=False)
            info = core.load_area(sample)
        self.assertAlmostEqual(info["Bregma_area_mm"], -2.12)
        groups = core.load_group_catalog(ROOT / "catalogos" / "asignacion_grupos.xlsx")
        bregmas = core.load_bregma_catalog(ROOT / "catalogos" / "registro_bregma.xlsx")
        parsed = core.parse_name("HB_04_R1_C3_BLAa_20X.xlsx")
        resolved = app.resolve_stack_metadata(
            parsed, groups, bregmas, {"Bregma_area_mm": -2.3}
        )
        self.assertAlmostEqual(resolved["Bregma_mm"], -2.3)
        self.assertEqual(resolved["Fuente_bregma"], "Excel de área")


if __name__ == "__main__":
    unittest.main(verbosity=2)
