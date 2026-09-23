import unittest
from io import StringIO
from pathlib import Path

import pandas as pd

from app import InputSchemaError, build_insumos_export, build_receitas_export


ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"


class AppTransformationsTest(unittest.TestCase):
    def test_build_insumos_export_accepts_new_csv_structure(self):
        source = StringIO(
            "New Column,INSUMOS,CUSTO,QUANTIDADE,UNIDADE DE MEDIDA\n"
            ".,Açaí,109,10000,GRAMAS\n"
            ".,Cereja,\"96,41\",4500,GRAMAS\n"
            ".,,,,\n"
        )

        result = build_insumos_export(
            email="Cliente@Example.com",
            csv_input=source,
            template_xlsx_path=ASSETS_DIR / "insumos.xlsx",
        )

        self.assertEqual(len(result), 2)
        self.assertEqual(list(result.columns), list(pd.read_excel(ASSETS_DIR / "insumos.xlsx").columns))
        self.assertEqual(result["App user email"].tolist(), ["cliente@example.com"] * 2)
        self.assertEqual(result["INSUMOS"].tolist(), ["Açaí", "Cereja"])
        self.assertEqual(result["CUSTO"].tolist(), [109, 96.41])
        self.assertTrue(result["TIPO DE EMBALAGEM"].isna().all())

    def test_build_insumos_export_normalizes_new_units_and_decimal_formats(self):
        source = StringIO(
            "INSUMOS,CUSTO,QUANTIDADE,UNIDADE DE MEDIDA\n"
            "Casquinha,0.22,1,UNIDADES\n"
            "Farinha,10,1,kg\n"
            "Leite,8,2,L\n"
            "Calda,5,250,ml\n"
        )

        result = build_insumos_export(
            email="user@test.com",
            csv_input=source,
            template_xlsx_path=ASSETS_DIR / "insumos.xlsx",
        )

        self.assertEqual(result["CUSTO"].tolist(), [0.22, 10, 8, 5])
        self.assertEqual(result["QUANTIDADE"].tolist(), [1, 1000, 2000, 250])
        self.assertEqual(
            result["UNIDADE DE MEDIDA"].tolist(),
            ["UNIDADES", "GRAMAS", "ML", "ML"],
        )

    def test_build_receitas_export_accepts_single_new_csv_and_pads_template_slots(self):
        source = StringIO(
            "RECEITA,STATUS,PREÇO DE VENDA,RENDIMENTO DA RECEITA,"
            "INSUMO 1,QND USADA 1,INSUMO 2,QNT USADA 2\n"
            "Sundae,true,\"7,50\",1,Sorvete,180,Colher,1\n"
        )

        result = build_receitas_export(
            email="Cliente@Example.com",
            csv_input=source,
            template_xlsx_path=ASSETS_DIR / "receitas modelo.xlsx",
        )

        self.assertEqual(len(result), 1)
        self.assertEqual(
            list(result.columns),
            list(pd.read_excel(ASSETS_DIR / "receitas modelo.xlsx").columns),
        )
        self.assertEqual(result.loc[0, "App user email"], "cliente@example.com")
        self.assertEqual(result.loc[0, "RECEITA"], "Sundae")
        self.assertEqual(result.loc[0, "STATUS"], 1)
        self.assertEqual(result.loc[0, "PREÇO DE VENDA"], 7.5)
        self.assertEqual(result.loc[0, "RENDIMENTO DA RECEITA"], 1)
        self.assertEqual(result.loc[0, "INSUMO 1"], "Sorvete")
        self.assertEqual(result.loc[0, "QND USADA 1"], 180)
        self.assertEqual(result.loc[0, "INSUMO 2"], "Colher")
        self.assertEqual(result.loc[0, "QNT USADA 2"], 1)
        self.assertTrue(pd.isna(result.loc[0, "INSUMO 25"]))

    def test_build_insumos_export_reports_missing_required_headers(self):
        source = StringIO("INSUMOS,CUSTO,QUANTIDADE\nFarinha,10,1\n")

        try:
            build_insumos_export(
                email="user@test.com",
                csv_input=source,
                template_xlsx_path=ASSETS_DIR / "insumos.xlsx",
            )
        except Exception as error:
            self.assertIsInstance(error, InputSchemaError)
            self.assertEqual(
                str(error),
                "Insumos: colunas obrigatórias ausentes: UNIDADE DE MEDIDA",
            )
        else:
            self.fail("build_insumos_export deveria rejeitar headers incompletos")

    def test_build_receitas_export_rejects_incomplete_ingredient_pair(self):
        source = StringIO(
            "RECEITA,STATUS,PREÇO DE VENDA,RENDIMENTO DA RECEITA,INSUMO 1\n"
            "Sundae,true,7,1,Sorvete\n"
        )

        with self.assertRaisesRegex(
            InputSchemaError,
            "Receitas: par incompleto: INSUMO 1 exige QND USADA 1",
        ):
            build_receitas_export(
                email="user@test.com",
                csv_input=source,
                template_xlsx_path=ASSETS_DIR / "receitas modelo.xlsx",
            )

    def test_build_receitas_export_rejects_unknown_status(self):
        source = StringIO(
            "RECEITA,STATUS,PREÇO DE VENDA,RENDIMENTO DA RECEITA,"
            "INSUMO 1,QND USADA 1\n"
            "Sundae,paused,7,1,Sorvete,180\n"
        )

        with self.assertRaisesRegex(
            InputSchemaError,
            "Receitas: STATUS inválido: paused",
        ):
            build_receitas_export(
                email="user@test.com",
                csv_input=source,
                template_xlsx_path=ASSETS_DIR / "receitas modelo.xlsx",
            )

    def test_find_insumos_warnings_reports_duplicates_without_removing_them(self):
        try:
            from app import find_insumos_warnings
        except ImportError:
            self.fail("find_insumos_warnings precisa existir")

        source_df = pd.DataFrame(
            [
                {
                    "INSUMOS": "Casquinha",
                    "CUSTO": "82",
                    "QUANTIDADE": "300",
                    "UNIDADE DE MEDIDA": "UNIDADES",
                },
                {
                    "INSUMOS": " casquinha ",
                    "CUSTO": "0,22",
                    "QUANTIDADE": "1",
                    "UNIDADE DE MEDIDA": "UNIDADES",
                },
            ]
        )

        warnings = find_insumos_warnings(source_df)

        self.assertEqual(warnings, ["Insumos duplicados mantidos: Casquinha."])

    def test_find_receitas_warnings_reports_blank_status(self):
        try:
            from app import find_receitas_warnings
        except ImportError:
            self.fail("find_receitas_warnings precisa existir")

        source_df = pd.DataFrame(
            [
                {
                    "RECEITA": "Casquinha",
                    "STATUS": pd.NA,
                    "PREÇO DE VENDA": "4",
                    "RENDIMENTO DA RECEITA": "1",
                    "INSUMO 1": "Casquinha",
                    "QND USADA 1": "1",
                }
            ]
        )

        warnings = find_receitas_warnings(source_df)

        self.assertEqual(warnings, ["Receitas com STATUS vazio: Casquinha."])

if __name__ == "__main__":
    unittest.main()
