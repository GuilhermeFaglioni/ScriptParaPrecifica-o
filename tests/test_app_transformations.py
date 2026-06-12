import unittest
from pathlib import Path

import pandas as pd

from app import build_insumos_export, build_receitas_export


ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"


class AppTransformationsTest(unittest.TestCase):
    def test_build_insumos_export_matches_template_columns_and_email(self):
        email = "isabelleite2020@outlook.com"

        result = build_insumos_export(
            email=email,
            csv_input=ASSETS_DIR / "insumos_subprodutos.csv",
            template_xlsx_path=ASSETS_DIR / "insumos.xlsx",
        )

        template_columns = list(pd.read_excel(ASSETS_DIR / "insumos.xlsx").columns)
        self.assertEqual(list(result.columns), template_columns)
        self.assertTrue((result["App user email"] == email.lower()).all())
        self.assertGreater(len(result), 0)
        self.assertIn("GRAMAS", set(result["UNIDADE DE MEDIDA"].dropna()))

    def test_build_receitas_export_matches_template_columns_and_fills_first_item_slots(self):
        email = "Vitoroliveiravfo@gmail.com"

        result = build_receitas_export(
            email=email,
            csv_input=ASSETS_DIR / "receita_itens.csv",
            template_xlsx_path=ASSETS_DIR / "receitas modelo.xlsx",
            receitas_master_path=ASSETS_DIR / "receita_valores.csv",
        )

        template_columns = list(pd.read_excel(ASSETS_DIR / "receitas modelo.xlsx").columns)
        self.assertEqual(list(result.columns), template_columns)
        self.assertTrue((result["App user email"] == email.lower()).all())
        self.assertGreater(len(result), 0)
        self.assertTrue(result["INSUMO 1"].notna().any())
        self.assertTrue(result["QND USADA 1"].notna().any())
        self.assertIn("test duplicar", set(result["RECEITA"].dropna()))
        self.assertIn(15, set(result["PREÇO DE VENDA"].dropna()))

    def test_build_receitas_export_uses_provided_master_csv(self):
        itens_df = pd.DataFrame(
            [
                {
                    "receita_id": "abc123",
                    "itens_id": "item-1",
                    "quantidade": "2",
                    "usuario_email": "user@test.com",
                    "nome": "Farinha",
                    "insumo_receita_itens": "Farinha",
                }
            ]
        )
        master_df = pd.DataFrame(
            [
                {
                    "🔒 Row ID": "abc123",
                    "usuario_email": "user@test.com",
                    "nome": "Bolo Teste",
                    "preco_venda": "19,9",
                    "rendimento_receita": "4",
                    "status": "1",
                }
            ]
        )
        itens_path = ASSETS_DIR / "tmp_receitas_itens.csv"
        master_path = ASSETS_DIR / "tmp_receitas_master.csv"
        itens_df.to_csv(itens_path, index=False)
        master_df.to_csv(master_path, index=False)

        try:
            result = build_receitas_export(
                email="user@test.com",
                csv_input=itens_path,
                template_xlsx_path=ASSETS_DIR / "receitas modelo.xlsx",
                receitas_master_path=master_path,
            )
        finally:
            itens_path.unlink(missing_ok=True)
            master_path.unlink(missing_ok=True)

        self.assertEqual(result.loc[0, "RECEITA"], "Bolo Teste")
        self.assertEqual(result.loc[0, "PREÇO DE VENDA"], 19.9)
        self.assertEqual(result.loc[0, "RENDIMENTO DA RECEITA"], 4)
        self.assertEqual(result.loc[0, "STATUS"], 1)

    def test_build_insumos_export_normalizes_unit_variants_to_canonical_uppercase(self):
        source_df = pd.DataFrame(
            [
                {
                    "usuario_email": "user@test.com",
                    "nome": "Item A",
                    "custo": "10,5",
                    "tipo_embalagem": "pacote",
                    "quantidade": "1000",
                    "tipo_unidade_medida": "g",
                },
                {
                    "usuario_email": "user@test.com",
                    "nome": "Item B",
                    "custo": "20",
                    "tipo_embalagem": "",
                    "quantidade": "1",
                    "tipo_unidade_medida": "Kilogramas (Kg)",
                },
                {
                    "usuario_email": "user@test.com",
                    "nome": "Item C",
                    "custo": "30",
                    "tipo_embalagem": "",
                    "quantidade": "250",
                    "tipo_unidade_medida": "ml",
                },
                {
                    "usuario_email": "user@test.com",
                    "nome": "Item D",
                    "custo": "40",
                    "tipo_embalagem": "",
                    "quantidade": "3",
                    "tipo_unidade_medida": "Unidades (pcs)",
                },
            ]
        )
        source_path = ASSETS_DIR / "tmp_insumos_units.csv"
        source_df.to_csv(source_path, index=False)

        try:
            result = build_insumos_export(
                email="user@test.com",
                csv_input=source_path,
                template_xlsx_path=ASSETS_DIR / "insumos.xlsx",
            )
        finally:
            source_path.unlink(missing_ok=True)

        self.assertEqual(result.loc[0, "UNIDADE DE MEDIDA"], "GRAMAS")
        self.assertEqual(result.loc[0, "QUANTIDADE"], 1000)
        self.assertEqual(result.loc[1, "UNIDADE DE MEDIDA"], "GRAMAS")
        self.assertEqual(result.loc[1, "QUANTIDADE"], 1000)
        self.assertEqual(result.loc[2, "UNIDADE DE MEDIDA"], "GRAMAS")
        self.assertEqual(result.loc[2, "QUANTIDADE"], 250)
        self.assertEqual(result.loc[3, "UNIDADE DE MEDIDA"], "UNIDADES")
        self.assertEqual(result.loc[3, "QUANTIDADE"], 3)


if __name__ == "__main__":
    unittest.main()
