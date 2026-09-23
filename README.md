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
- estrutura esperada de insumos:
  - `INSUMOS`
  - `CUSTO`
  - `QUANTIDADE`
  - `UNIDADE DE MEDIDA`
- estrutura esperada de receitas:
  - `RECEITA`
  - `STATUS`
  - `PREÇO DE VENDA`
  - `RENDIMENTO DA RECEITA`
  - pares `INSUMO N` + `QND USADA 1`/`QNT USADA N`, até 25
- saída em `.xlsx` seguindo estrutura dos modelos:
  - `assets/receitas modelo.xlsx`
  - `assets/insumos.xlsx`
- email informado é aplicado em todas as linhas exportadas
- linhas vazias são ignoradas
- insumos duplicados e receitas sem status são mantidos com aviso
- formatos antigos com CSV mestre/IDs não são suportados
