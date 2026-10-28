import tempfile
import unittest
from pathlib import Path

import pandas as pd

import batch_project as batch
import cfos_core as core


def record(**values):
    row = batch.blank_record()
    row.update(
        {
            "Stack_ID": "EXP_R1_C1_BLAa_20X_XYZ",
            "Experimento": "EXP",
            "Animal": "R1",
            "Corte": "C1",
            "Region": "BLAa",
            "Grupo": "Vehículo",
            "Bregma_mm": -2.3,
            "Bregma_area_mm": -2.3,
            "Fuente_bregma": "Excel de área",
            "QC": "Excluir",
            "Revision_ROI": "Pendiente",
        }
    )
    row.update(values)
    return {column: row.get(column, "") for column in core.META_COLUMNS}


class BregmaEditingTests(unittest.TestCase):
    def setUp(self):
        self.groups = pd.DataFrame(
            [{"Experimento": "EXP", "Animal": "R1", "Grupo": "Vehículo"}]
        )
        self.bregmas = pd.DataFrame(
            [
                {
                    "Experimento": "EXP",
                    "Animal": "R1",
                    "Corte": "C1",
                    "Region": "BLAa",
                    "Bregma_mm": -2.0,
                }
            ]
        )
        self.parsed = {
            "Experimento": "EXP",
            "Animal": "R1",
            "Corte": "C1",
            "Region": "BLAa",
        }

    def test_catalog_can_override_area_when_explicitly_updated(self):
        area = {"Bregma_area_mm": -2.3}
        legacy = batch.resolve_stack_metadata(
            self.parsed, self.groups, self.bregmas, area
        )
        refreshed = batch.resolve_stack_metadata(
            self.parsed, self.groups, self.bregmas, area, prefer_catalog=True
        )
        self.assertAlmostEqual(legacy["Bregma_mm"], -2.3)
        self.assertEqual(legacy["Fuente_bregma"], "Excel de área")
        self.assertAlmostEqual(refreshed["Bregma_mm"], -2.0)
        self.assertEqual(refreshed["Fuente_bregma"], "Registro maestro actualizado")

    def test_manual_bregma_is_validated_and_becomes_effective(self):
        original = record(Detalle_QC_base="Bregma no encontrado")
        edited = batch.set_bregma(original, "-2,15")
        self.assertAlmostEqual(edited["Bregma_mm"], -2.15)
        self.assertEqual(edited["Fuente_bregma"], "Edición manual")
        self.assertAlmostEqual(edited["Bregma_area_mm"], -2.3)
        self.assertNotIn("Bregma no encontrado", edited["Detalle_QC"])
        self.assertAlmostEqual(
            core.effective_bregma(edited, {"Bregma_area_mm": -2.3}), -2.15
        )
        with self.assertRaisesRegex(ValueError, "fuera del rango"):
            batch.set_bregma(original, -4.0)

    def test_catalog_sync_updates_saved_rows_but_preserves_manual_edits(self):
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            rows = pd.DataFrame(
                [
                    record(),
                    record(
                        Stack_ID="EXP_R2_C1_BLAa_20X_XYZ",
                        Animal="R2",
                        Bregma_mm=-2.1,
                        Fuente_bregma="Edición manual",
                    ),
                ],
                columns=core.META_COLUMNS,
            )
            batch.write_registry(project, rows)
            catalog = pd.DataFrame(
                [
                    {
                        "Experimento": "EXP",
                        "Animal": "R1",
                        "Corte": "C1",
                        "Region": "BLAa",
                        "Bregma_mm": -2.0,
                    },
                    {
                        "Experimento": "EXP",
                        "Animal": "R2",
                        "Corte": "C1",
                        "Region": "BLAa",
                        "Bregma_mm": -1.9,
                    },
                ]
            )

            result = batch.update_saved_bregmas_from_catalog(project, catalog)
            updated = result["Data"].set_index("Animal")
            self.assertEqual(result["Coincidencias"], 2)
            self.assertEqual(result["Bregmas_cambiados"], 1)
            self.assertEqual(result["Omitidos_manual"], 1)
            self.assertAlmostEqual(float(updated.loc["R1", "Bregma_mm"]), -2.0)
            self.assertEqual(
                updated.loc["R1", "Fuente_bregma"], "Registro maestro actualizado"
            )
            self.assertAlmostEqual(float(updated.loc["R2", "Bregma_mm"]), -2.1)
            experiment_file = project / "experimentos" / "EXP" / "registro_EXP.xlsx"
            self.assertTrue(experiment_file.is_file())
            saved = pd.read_excel(experiment_file, dtype=str).fillna("")
            r1 = saved[saved["Animal"].eq("R1")].iloc[0]
            self.assertAlmostEqual(float(r1["Bregma_mm"]), -2.0)

    def test_visual_row_position_never_appends_a_duplicate(self):
        staging = pd.DataFrame(
            [
                record(Stack_ID="EXP_R1_C1_BLAa", Animal="R1"),
                record(Stack_ID="EXP_R2_C1_BLAa", Animal="R2"),
            ],
            columns=core.META_COLUMNS,
            index=[111, 112],
        )
        edited = staging.iloc[1].to_dict()
        edited = batch.set_bregma(edited, -2.12)

        result = batch.replace_staging_row(staging, 1, edited)

        self.assertEqual(list(result.index), [0, 1])
        self.assertEqual(len(result), 2)
        self.assertEqual(result["Stack_ID"].nunique(), 2)
        self.assertAlmostEqual(float(result.iloc[1]["Bregma_mm"]), -2.12)

    def test_hb02_bregma_update_repairs_duplicate_and_preserves_roi(self):
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            inputs = project / "entradas_prueba"
            inputs.mkdir()
            objects = inputs / "objetos.xlsx"
            area = inputs / "area.xlsx"
            roi = inputs / "roi.roi"
            for path in (objects, area, roi):
                path.write_bytes(b"archivo de prueba")

            complete = record(
                Stack_ID="HB_02_R6_C4_BLAa",
                Experimento="HB_02",
                Animal="R6",
                Corte="C4",
                Region="BLAa",
                Bregma_mm=-2.0,
                Fuente_bregma="Excel de área",
                Archivo_objetos=str(objects),
                Archivo_area=str(area),
                Archivo_ROI=str(roi),
            )
            appended = record(
                Stack_ID="HB_02_R6_C4_BLAa",
                Experimento="HB_02",
                Animal="R6",
                Corte="C4",
                Region="BLAa",
                Bregma_mm=-2.12,
                Fuente_bregma="Registro maestro actualizado",
                Archivo_objetos="",
                Archivo_area="",
                Archivo_ROI="",
                QC="Revisar",
            )
            corrupt = pd.DataFrame([complete, appended], columns=core.META_COLUMNS)
            batch._write_workbook(
                project / batch.REGISTRY_NAME,
                {"Registro_stacks": corrupt},
            )
            catalog = pd.DataFrame(
                [
                    {
                        "Experimento": "HB_02",
                        "Animal": "R6",
                        "Corte": "C4",
                        "Region": "BLAa",
                        "Bregma_mm": -2.12,
                    }
                ]
            )

            result = batch.update_saved_bregmas_from_catalog(project, catalog)
            repaired = result["Data"]

            self.assertEqual(result["Duplicados_reparados"], 1)
            self.assertEqual(result["Stack_ID_consolidados"], ["HB_02_R6_C4_BLAa"])
            self.assertEqual(len(repaired), 1)
            self.assertEqual(repaired.iloc[0]["Stack_ID"], "HB_02_R6_C4_BLAa")
            self.assertAlmostEqual(float(repaired.iloc[0]["Bregma_mm"]), -2.12)
            self.assertTrue(Path(repaired.iloc[0]["Archivo_ROI"]).is_file())
            saved = batch.load_registry(project)
            self.assertEqual(len(saved), 1)
            self.assertAlmostEqual(float(saved.iloc[0]["Bregma_mm"]), -2.12)
            self.assertTrue(Path(saved.iloc[0]["Archivo_ROI"]).is_file())


class ExperimentDeletionTests(unittest.TestCase):
    def test_delete_archives_one_experiment_and_keeps_the_other_active(self):
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            rows = pd.DataFrame(
                [
                    record(Experimento="EXP_A", Stack_ID="EXP_A_R1_C1_BLAa_20X_XYZ"),
                    record(Experimento="EXP_B", Stack_ID="EXP_B_R2_C1_BLAa_20X_XYZ", Animal="R2"),
                ],
                columns=core.META_COLUMNS,
            )
            batch.write_registry(project, rows)
            active = project / "experimentos" / "EXP_A"
            active.mkdir(parents=True)
            (active / "dato.txt").write_text("respaldo", encoding="utf-8")

            result = batch.archive_experiment(project, "EXP_A")
            refreshed = batch.load_registry(project)
            self.assertEqual(result["N_stacks"], 1)
            self.assertEqual(set(refreshed["Experimento"]), {"EXP_B"})
            self.assertFalse(active.exists())
            self.assertTrue((result["Archivo"] / "experimento" / "dato.txt").is_file())
            self.assertTrue((result["Archivo"] / "ELIMINADO.txt").is_file())
            self.assertTrue((project / batch.COUNTS_NAME).is_file())


if __name__ == "__main__":
    unittest.main()
