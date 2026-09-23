from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"
RECEITAS_TEMPLATE_PATH = ASSETS_DIR / "receitas modelo.xlsx"
INSUMOS_TEMPLATE_PATH = ASSETS_DIR / "insumos.xlsx"
MAX_RECEITA_ITEMS = 25

APP_TITLE = "Conversor de Planilhas CSV"
APP_DESCRIPTION = (
    "Busca por email. Recebe CSV no upload e gera planilhas XLSX "
    "na estrutura dos modelos."
)


class InputSchemaError(ValueError):
    """Raised when an uploaded CSV does not match the supported structure."""


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
    "ml": "ML",
    "mililitro": "ML",
    "mililitros": "ML",
    "mililitros ml": "ML",
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

    if unit == "LITROS":
        return quantity * 1000, "ML"

    if unit in {"GRAMAS", "ML", "UNIDADES"}:
        return quantity, unit

    return quantity, unit


def parse_number(value):
    if pd.isna(value):
        return pd.NA
    text = normalize_text(value)
    if not text:
        return pd.NA
    text = text.replace(" ", "")
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        number = float(text)
    except ValueError:
        return value
    if number.is_integer():
        return int(number)
    return number


def parse_status(value):
    text = normalize_text(value).lower()
    if not text:
        return pd.NA
    if text in {"true", "1"}:
        return 1
    if text in {"false", "0"}:
        return 0
    raise InputSchemaError(f"Receitas: STATUS inválido: {normalize_text(value)}")


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


def require_columns(
    dataframe: pd.DataFrame,
    required_columns: list[str],
    source_name: str,
) -> None:
    missing_columns = [
        column for column in required_columns if column not in dataframe.columns
    ]
    if missing_columns:
        missing_text = ", ".join(missing_columns)
        raise InputSchemaError(
            f"{source_name}: colunas obrigatórias ausentes: {missing_text}"
        )


def validate_receitas_columns(dataframe: pd.DataFrame) -> None:
    require_columns(
        dataframe,
        ["RECEITA", "STATUS", "PREÇO DE VENDA", "RENDIMENTO DA RECEITA"],
        "Receitas",
    )

    for index in range(1, MAX_RECEITA_ITEMS + 1):
        item_column = f"INSUMO {index}"
        quantity_column = "QND USADA 1" if index == 1 else f"QNT USADA {index}"
        has_item_column = item_column in dataframe.columns
        has_quantity_column = quantity_column in dataframe.columns
        if has_item_column != has_quantity_column:
            present_column = item_column if has_item_column else quantity_column
            missing_column = quantity_column if has_item_column else item_column
            raise InputSchemaError(
                f"Receitas: par incompleto: {present_column} exige {missing_column}"
            )

    require_columns(dataframe, ["INSUMO 1", "QND USADA 1"], "Receitas")


def find_insumos_warnings(dataframe: pd.DataFrame) -> list[str]:
    warnings = []
    names = dataframe["INSUMOS"].fillna("").map(normalize_text)
    normalized_names = names.str.casefold()
    duplicate_keys = normalized_names[
        (normalized_names != "") & normalized_names.duplicated(keep=False)
    ].drop_duplicates()
    if not duplicate_keys.empty:
        duplicate_names = [
            names[normalized_names == key].iloc[0] for key in duplicate_keys
        ]
        warnings.append(
            f"Insumos duplicados mantidos: {', '.join(duplicate_names)}."
        )

    business_columns = ["CUSTO", "QUANTIDADE", "UNIDADE DE MEDIDA"]
    populated_without_name = (names == "") & dataframe[business_columns].apply(
        lambda column: column.fillna("").map(normalize_text) != ""
    ).any(axis=1)
    invalid_count = int(populated_without_name.sum())
    if invalid_count:
        warnings.append(
            f"{invalid_count} linha(s) de insumos sem nome foram ignoradas."
        )

    return warnings


def find_receitas_warnings(dataframe: pd.DataFrame) -> list[str]:
    warnings = []
    names = dataframe["RECEITA"].fillna("").map(normalize_text)
    statuses = dataframe["STATUS"].fillna("").map(normalize_text)
    blank_status_names = names[(names != "") & (statuses == "")].tolist()
    if blank_status_names:
        warnings.append(
            f"Receitas com STATUS vazio: {', '.join(blank_status_names)}."
        )

    business_columns = [
        column for column in dataframe.columns if column != "RECEITA"
    ]
    populated_without_name = (names == "") & dataframe[business_columns].apply(
        lambda column: column.fillna("").map(normalize_text) != ""
    ).any(axis=1)
    invalid_count = int(populated_without_name.sum())
    if invalid_count:
        warnings.append(
            f"{invalid_count} linha(s) de receitas sem nome foram ignoradas."
        )

    return warnings


def build_insumos_export(
    email: str,
    csv_input,
    template_xlsx_path: Path = INSUMOS_TEMPLATE_PATH,
) -> pd.DataFrame:
    template_columns = load_template_columns(template_xlsx_path)
    source_df = read_csv_input(csv_input)
    require_columns(
        source_df,
        ["INSUMOS", "CUSTO", "QUANTIDADE", "UNIDADE DE MEDIDA"],
        "Insumos",
    )
    filtered_df = source_df[
        source_df["INSUMOS"].fillna("").map(normalize_text) != ""
    ].copy()

    filtered_df["App user email"] = normalize_email(email)
    filtered_df["INSUMOS"] = filtered_df["INSUMOS"].map(normalize_text)
    filtered_df["CUSTO"] = filtered_df["CUSTO"].map(parse_number)
    filtered_df["TIPO DE EMBALAGEM"] = pd.NA
    normalized_measure = filtered_df.apply(
        lambda row: normalize_insumo_measure(
            row["QUANTIDADE"], row["UNIDADE DE MEDIDA"]
        ),
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
) -> pd.DataFrame:
    template_columns = load_template_columns(template_xlsx_path)
    source_df = read_csv_input(csv_input)
    validate_receitas_columns(source_df)
    filtered_df = source_df[
        source_df["RECEITA"].fillna("").map(normalize_text) != ""
    ].copy()

    rows = []
    for _, source_row in filtered_df.iterrows():
        row = {column: pd.NA for column in template_columns}
        row["App user email"] = normalize_email(email)
        row["RECEITA"] = normalize_text(source_row.get("RECEITA"))
        row["STATUS"] = parse_status(source_row.get("STATUS"))
        row["PREÇO DE VENDA"] = parse_number(source_row.get("PREÇO DE VENDA"))
        row["RENDIMENTO DA RECEITA"] = parse_number(
            source_row.get("RENDIMENTO DA RECEITA")
        )

        for index in range(1, MAX_RECEITA_ITEMS + 1):
            quantity_column = "QND USADA 1" if index == 1 else f"QNT USADA {index}"
            item_name = normalize_text(source_row.get(f"INSUMO {index}"))
            if item_name:
                row[f"INSUMO {index}"] = item_name
            row[quantity_column] = parse_number(source_row.get(quantity_column))

        rows.append(row)

    export_df = pd.DataFrame(rows, columns=template_columns)
    return export_df.reset_index(drop=True)


def render_download_section(
    title: str,
    description: str,
    dataframe: pd.DataFrame,
    output_name: str,
    download_label: str,
    warnings: list[str] | None = None,
) -> None:
    st.subheader(title)
    st.write(description)

    for warning in warnings or []:
        st.warning(warning)

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
        "Cada aba funciona sozinha e recebe 1 arquivo CSV no formato novo."
    )

    email = st.text_input("Email do usuário", placeholder="usuario@dominio.com")
    normalized_email = normalize_email(email)

    if not normalized_email:
        st.info("Informe um email para gerar as planilhas.")
        return

    receitas_tab, insumos_tab = st.tabs(["Receitas", "Insumos"])

    with receitas_tab:
        st.markdown("**Arquivos obrigatórios desta aba**")
        st.write("CSV de receitas. Exemplo: `RECEITAS.csv`")

        receitas_file = st.file_uploader(
            "CSV de receitas (`RECEITAS.csv`)",
            type=["csv"],
            accept_multiple_files=False,
            key="receitas-csv",
        )
        if receitas_file is None:
            st.info("Envie CSV de receitas.")
        else:
            try:
                receitas_source_df = read_csv_input(receitas_file)
                receitas_df = build_receitas_export(normalized_email, receitas_file)
                receitas_warnings = find_receitas_warnings(receitas_source_df)
            except InputSchemaError as error:
                st.error(str(error))
            else:
                render_download_section(
                    title="Receitas",
                    description="CSV enviado -> modelo `assets/receitas modelo.xlsx`",
                    dataframe=receitas_df,
                    output_name="receitas_revisadas.xlsx",
                    download_label="Baixar receitas em XLSX",
                    warnings=receitas_warnings,
                )

    with insumos_tab:
        st.markdown("**Arquivo obrigatório desta aba**")
        st.write("CSV de insumos. Exemplo: `INSUMOS.csv`")
        insumos_file = st.file_uploader(
            "CSV de insumos (`INSUMOS.csv`)",
            type=["csv"],
            accept_multiple_files=False,
            key="insumos-csv",
        )
        if insumos_file is None:
            st.info("Envie CSV de insumos.")
        else:
            try:
                insumos_source_df = read_csv_input(insumos_file)
                insumos_df = build_insumos_export(normalized_email, insumos_file)
                insumos_warnings = find_insumos_warnings(insumos_source_df)
            except InputSchemaError as error:
                st.error(str(error))
            else:
                render_download_section(
                    title="Insumos",
                    description="CSV enviado -> modelo `assets/insumos.xlsx`",
                    dataframe=insumos_df,
                    output_name="insumos_revisados.xlsx",
                    download_label="Baixar insumos em XLSX",
                    warnings=insumos_warnings,
                )


if __name__ == "__main__":
    main()
