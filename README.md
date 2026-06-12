# Conversor de Planilhas CSV

App Streamlit para buscar planilhas por email usando CSV enviado no upload:

- Receitas
- Insumos

## Rodar local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Fluxo atual

- input de email do usuário
- aba `Receitas` recebe 1 CSV no upload
- aba `Insumos` recebe 1 CSV no upload
- nomes esperados:
  - `receita_itens.csv`
  - `receita_valores.csv`
  - `insumos_subprodutos.csv`
- saída em `.xlsx` seguindo estrutura dos modelos:
  - `assets/receitas modelo.xlsx`
  - `assets/insumos.xlsx`
- campos ausentes no CSV ficam vazios no XLSX
