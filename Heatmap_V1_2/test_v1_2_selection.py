from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pandas as pd

import batch_project as batch
import cfos_core as core


def registry_row(experiment, animal, include="Si", region="BLAa"):
    row = {column: "" for column in core.META_COLUMNS}
    row.update(
        {
            "Stack_ID": f"{experiment}_{animal}_C1_{region}",
            "Experimento": experiment,
            "Animal": animal,
            "Corte": "C1",
            "Region": region,
            "Grupo": "Vehículo",
            "QC": "OK",
            "Incluir_analisis": include,
            "Objetos_XYZ": 10,
            "Objetos_incluidos_filtro": 8,
            "Objetos_validos": 7,
            "Volumen_um3": 1000,
        }
    )
    return row


class AnalysisSelectionTests(unittest.TestCase):
    def test_blank_legacy_selection_means_selected(self):
        self.assertTrue(core.selected_for_analysis(""))
        self.assertTrue(core.selected_for_analysis("Si"))
        self.assertFalse(core.selected_for_analysis("No"))

    def test_toggle_is_persisted_for_every_stack_of_animal(self):
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            rows = [
                registry_row("SCOP_LiCl_02", "R61", region="BLAa"),
                registry_row("SCOP_LiCl_02", "R61", region="BLAp"),
                registry_row("SCOP_LiCl_02", "R62", region="BLAa"),
            ]
            pd.DataFrame(rows, columns=core.META_COLUMNS).to_excel(
                project / batch.REGISTRY_NAME,
                index=False,
                sheet_name="Registro_stacks",
            )
            result = batch.set_animal_analysis_selection(
                project, "SCOP_LiCl_02", "R61", False
            )
            raw = pd.read_excel(project / batch.REGISTRY_NAME, dtype=str).fillna("")
            r61 = raw[raw["Animal"].eq("R61")]
            r62 = raw[raw["Animal"].eq("R62")]
            self.assertTrue(r61["Incluir_analisis"].eq("No").all())
            self.assertTrue(r62["Incluir_analisis"].eq("Si").all())
            self.assertEqual(result["Stacks_actualizados"], 2)

    def test_unselected_animals_do_not_enter_count_summaries(self):
        rows = [
            registry_row("HB_02", "R8", include="Si"),
            registry_row("SCOP_LiCl_02", "R61", include="No"),
        ]
        tables = batch.count_tables(pd.DataFrame(rows, columns=core.META_COLUMNS))
        animals = tables["Conteo_por_animal"]
        self.assertEqual(animals[["Experimento", "Animal"]].values.tolist(), [["HB_02", "R8"]])
        stack_table = tables["Conteo_por_stack"].set_index("Animal")
        self.assertTrue(bool(stack_table.loc["R8", "Incluido_en_resumen"]))
        self.assertFalse(bool(stack_table.loc["R61", "Incluido_en_resumen"]))


if __name__ == "__main__":
    unittest.main(verbosity=2)
