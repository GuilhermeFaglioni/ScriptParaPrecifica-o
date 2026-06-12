from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"
RECEITAS_TEMPLATE_PATH = ASSETS_DIR / "receitas modelo.xlsx"
INSUMOS_TEMPLATE_PATH = ASSETS_DIR / "insumos.xlsx"
RECEITAS_MASTER_PATH = ASSETS_DIR / "receita_valores.csv"
MAX_RECEITA_ITEMS = 25

APP_TITLE = "Conversor de Planilhas CSV"
APP_DESCRIPTION = (
    "Busca por email. Recebe CSV no upload e gera planilhas XLSX "
    "na estrutura dos modelos."
)


def configure_page() -> None:
    st.set_page_config(
        page_title="Conversor de Planilhas",
        page_icon="📄",
        layout="wide",
    )


def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).replace("\n", " ").strip()
    return " ".join(text.split())


def normalize_email(value: str) -> str:
    return normalize_text(value).lower()


UNIT_ALIASES = {
    "g": "GRAMAS",
    "grama": "GRAMAS",
    "gramas": "GRAMAS",
    "gramas g": "GRAMAS",
    "kg": "QUILOGRAMAS",
    "quilo": "QUILOGRAMAS",
    "quilos": "QUILOGRAMAS",
    "kilograma": "QUILOGRAMAS",
    "kilogramas": "QUILOGRAMAS",
    "kilogramas kg": "QUILOGRAMAS",
    "quilograma": "QUILOGRAMAS",
    "quilogramas": "QUILOGRAMAS",
    "quilogramas kg": "QUILOGRAMAS",
    "ml": "MILILITROS",
    "mililitro": "MILILITROS",
    "mililitros": "MILILITROS",
    "mililitros ml": "MILILITROS",
    "l": "LITROS",
    "litro": "LITROS",
    "litros": "LITROS",
    "litros l": "LITROS",
    "un": "UNIDADES",
    "und": "UNIDADES",
    "unds": "UNIDADES",
    "unidade": "UNIDADES",
    "unidades": "UNIDADES",
    "unidades pcs": "UNIDADES",
    "pcs": "UNIDADES",
}


def normalize_unit(value) -> str:
    text = normalize_text(value)
    if not text:
        return ""

    simplified = (
        text.lower()
        .replace("(", " ")
        .replace(")", " ")
        .replace("/", " ")
        .replace("-", " ")
    )
    simplified = " ".join(simplified.split())

    if simplified in UNIT_ALIASES:
        return UNIT_ALIASES[simplified]

    parts = simplified.split()
    for size in range(len(parts), 0, -1):
        candidate = " ".join(parts[:size])
        if candidate in UNIT_ALIASES:
            return UNIT_ALIASES[candidate]
        candidate = " ".join(parts[-size:])
        if candidate in UNIT_ALIASES:
            return UNIT_ALIASES[candidate]

    return text.upper()


def normalize_insumo_measure(quantity_value, unit_value):
    unit = normalize_unit(unit_value)
    quantity = parse_number(quantity_value)

    if pd.isna(quantity):
        return quantity, unit

    if unit == "QUILOGRAMAS":
        return quantity * 1000, "GRAMAS"

    if unit in {"GRAMAS", "MILILITROS", "LITROS"}:
        if unit == "LITROS":
            return quantity * 1000, "GRAMAS"
        return quantity, "GRAMAS"

    return quantity, "UNIDADES"


def parse_number(value):
    if pd.isna(value):
        return pd.NA
    text = normalize_text(value)
    if not text:
        return pd.NA
    text = text.replace(".", "").replace(",", ".")
    try:
        number = float(text)
    except ValueError:
        return value
    if number.is_integer():
        return int(number)
    return number


def load_template_columns(template_xlsx_path: Path) -> list[str]:
    return list(pd.read_excel(template_xlsx_path, nrows=0).columns)


def to_xlsx_bytes(dataframe: pd.DataFrame, sheet_name: str = "Página1") -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer.getvalue()


def read_csv_input(csv_input) -> pd.DataFrame:
    if hasattr(csv_input, "seek"):
        csv_input.seek(0)
    return pd.read_csv(csv_input, dtype=str)


def ensure_columns(dataframe: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    for column in columns:
        if column not in dataframe.columns:
            dataframe[column] = pd.NA
    return dataframe


def build_insumos_export(
    email: str,
    csv_input,
    template_xlsx_path: Path = INSUMOS_TEMPLATE_PATH,
) -> pd.DataFrame:
    template_columns = load_template_columns(template_xlsx_path)
    source_df = read_csv_input(csv_input)
    source_df = ensure_columns(
        source_df,
        [
            "usuario_email",
            "nome",
            "custo",
            "tipo_embalagem",
            "quantidade",
            "tipo_unidade_medida",
        ],
    )

    filtered_df = source_df[
        source_df["usuario_email"].fillna("").str.lower() == normalize_email(email)
    ].copy()

    if filtered_df.empty:
        return pd.DataFrame(columns=template_columns)

    filtered_df = filtered_df[filtered_df["nome"].notna()].copy()
    filtered_df["App user email"] = normalize_email(email)
    filtered_df["INSUMOS"] = filtered_df["nome"].map(normalize_text)
    filtered_df["CUSTO"] = filtered_df["custo"].map(parse_number)
    filtered_df["TIPO DE EMBALAGEM"] = filtered_df["tipo_embalagem"].map(normalize_text)
    normalized_measure = filtered_df.apply(
        lambda row: normalize_insumo_measure(row["quantidade"], row["tipo_unidade_medida"]),
        axis=1,
    )
    filtered_df["QUANTIDADE"] = normalized_measure.map(lambda item: item[0])
    filtered_df["UNIDADE DE MEDIDA"] = normalized_measure.map(lambda item: item[1])

    export_df = filtered_df[
        [
            "App user email",
            "INSUMOS",
            "CUSTO",
            "TIPO DE EMBALAGEM",
            "QUANTIDADE",
            "UNIDADE DE MEDIDA",
        ]
    ].copy()

    export_df = export_df.reindex(columns=template_columns)
    return export_df.reset_index(drop=True)


def build_receitas_export(
    email: str,
    csv_input,
    template_xlsx_path: Path = RECEITAS_TEMPLATE_PATH,
    receitas_master_path: Path = RECEITAS_MASTER_PATH,
) -> pd.DataFrame:
    template_columns = load_template_columns(template_xlsx_path)
    source_df = read_csv_input(csv_input)
    receitas_master_df = pd.read_csv(receitas_master_path, dtype=str)
    source_df = ensure_columns(
        source_df,
        [
            "usuario_email",
            "receita_id",
            "insumo_receita_itens",
            "nome",
            "nome copy",
            "quantidade",
        ],
    )
    receitas_master_df = ensure_columns(
        receitas_master_df,
        ["🔒 Row ID", "nome", "preco_venda", "rendimento_receita", "status"],
    )

    filtered_df = source_df[
        source_df["usuario_email"].fillna("").str.lower() == normalize_email(email)
    ].copy()

    filtered_df = filtered_df[
        filtered_df["receita_id"].notna() & filtered_df["insumo_receita_itens"].notna()
    ].copy()

    if filtered_df.empty:
        return pd.DataFrame(columns=template_columns)

    receitas_master_df = receitas_master_df.rename(columns={"🔒 Row ID": "receita_id"})
    receitas_master_df = receitas_master_df.drop_duplicates(subset=["receita_id"], keep="first")
    filtered_df = filtered_df.merge(
        receitas_master_df[
            ["receita_id", "nome", "preco_venda", "rendimento_receita", "status"]
        ].rename(
            columns={
                "nome": "receita_nome",
                "preco_venda": "receita_preco_venda",
                "rendimento_receita": "receita_rendimento",
                "status": "receita_status",
            }
        ),
        on="receita_id",
        how="left",
    )

    filtered_df["item_name"] = filtered_df["nome copy"].where(
        filtered_df["nome copy"].notna() & (filtered_df["nome copy"].str.strip() != ""),
        filtered_df["insumo_receita_itens"],
    )
    filtered_df["item_name"] = filtered_df["item_name"].fillna(filtered_df["nome"]).map(normalize_text)
    filtered_df["item_quantity"] = filtered_df["quantidade"].map(parse_number)

    rows = []
    for _, group in filtered_df.groupby("receita_id", sort=False):
        row = {column: pd.NA for column in template_columns}
        row["App user email"] = normalize_email(email)
        first_item = group.iloc[0]
        row["RECEITA"] = normalize_text(first_item.get("receita_nome"))
        row["PREÇO DE VENDA"] = parse_number(first_item.get("receita_preco_venda"))
        row["RENDIMENTO DA RECEITA"] = parse_number(first_item.get("receita_rendimento"))
        status_value = normalize_text(first_item.get("receita_status"))
        row["STATUS"] = parse_number(status_value) if status_value else pd.NA

        for index, (_, item) in enumerate(group.head(MAX_RECEITA_ITEMS).iterrows(), start=1):
            row[f"INSUMO {index}"] = item["item_name"]
            quantity_column = "QND USADA 1" if index == 1 else f"QNT USADA {index}"
            row[quantity_column] = item["item_quantity"]

        rows.append(row)

    export_df = pd.DataFrame(rows, columns=template_columns)
    return export_df.reset_index(drop=True)


def render_download_section(
    title: str,
    description: str,
    dataframe: pd.DataFrame,
    output_name: str,
    download_label: str,
) -> None:
    st.subheader(title)
    st.write(description)

    if dataframe.empty:
        st.warning("Nenhum registro encontrado para este email.")
        return

    st.success(f"{len(dataframe)} linha(s) preparadas.")
    st.dataframe(dataframe, use_container_width=True)
    st.download_button(
        label=download_label,
        data=to_xlsx_bytes(dataframe),
        file_name=output_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def main() -> None:
    configure_page()

    st.title(APP_TITLE)
    st.write(APP_DESCRIPTION)
    st.caption(f"Pandas configurado: `{pd.__version__}`")
    st.info(
        "Cada aba funciona sozinha. `Receitas` precisa de 2 arquivos CSV. "
        "`Insumos` precisa de 1 arquivo CSV."
    )

    email = st.text_input("Email do usuário", placeholder="usuario@dominio.com")
    normalized_email = normalize_email(email)

    if not normalized_email:
        st.info("Informe um email para gerar as planilhas.")
        return

    receitas_tab, insumos_tab = st.tabs(["Receitas", "Insumos"])

    with receitas_tab:
        st.markdown("**Arquivos obrigatórios desta aba**")
        st.write("1. CSV de itens da receita. Nome esperado: `receita_itens.csv`")
        st.write("2. CSV mestre de receitas. Nome esperado: `receita_valores.csv`")

        receitas_file = st.file_uploader(
            "CSV de itens da receita (`receita_itens.csv`)",
            type=["csv"],
            accept_multiple_files=False,
            key="receitas-csv",
        )
        receitas_master_file = st.file_uploader(
            "CSV mestre de receitas (`receita_valores.csv`)",
            type=["csv"],
            accept_multiple_files=False,
            key="receitas-master-csv",
        )
        if receitas_file is None or receitas_master_file is None:
            st.warning("Para gerar `Receitas`, envie os 2 CSVs desta aba.")
        else:
            receitas_df = build_receitas_export(
                normalized_email,
                receitas_file,
                receitas_master_path=receitas_master_file,
            )
            render_download_section(
                title="Receitas",
                description=(
                    "CSVs enviados -> modelo `assets/receitas modelo.xlsx`"
                ),
                dataframe=receitas_df,
                output_name="receitas_revisadas.xlsx",
                download_label="Baixar receitas em XLSX",
            )

    with insumos_tab:
        st.markdown("**Arquivo obrigatório desta aba**")
        st.write("CSV de insumos/subprodutos. Nome esperado: `insumos_subprodutos.csv`")
        insumos_file = st.file_uploader(
            "CSV de insumos (`insumos_subprodutos.csv`)",
            type=["csv"],
            accept_multiple_files=False,
            key="insumos-csv",
        )
        if insumos_file is None:
            st.info("Envie CSV de insumos.")
        else:
            insumos_df = build_insumos_export(normalized_email, insumos_file)
            render_download_section(
                title="Insumos",
                description=("CSV enviado -> modelo `assets/insumos.xlsx`"),
                dataframe=insumos_df,
                output_name="insumos_revisados.xlsx",
                download_label="Baixar insumos em XLSX",
            )


if __name__ == "__main__":
    main()
