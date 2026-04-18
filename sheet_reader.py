import os
import re
import pandas as pd
from dotenv import load_dotenv

load_dotenv()


def build_sheet_csv_url() -> str:
    sheet_id = os.getenv("SHEET_ID")
    sheet_gid = os.getenv("SHEET_GID")

    if not sheet_id or not sheet_gid:
        raise ValueError("SHEET_ID ou SHEET_GID não encontrados no .env")

    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={sheet_gid}"


def load_sheet_dataframe() -> pd.DataFrame:
    csv_url = build_sheet_csv_url()
    df = pd.read_csv(csv_url, dtype=str)
    return df


def clean_process_id(value: str) -> str | None:
    if value is None:
        return None

    text = str(value).strip()

    if not text:
        return None

    lowered = text.lower()
    if lowered in {"nan", "none", "processos js", "processo js"}:
        return None

    # remove espaços
    text = text.replace(" ", "")

    # se vier como número com .0, remove
    if text.endswith(".0"):
        text = text[:-2]

    # pega só números
    text = re.sub(r"\D", "", text)

    if not text:
        return None

    return text


def get_process_ids(df: pd.DataFrame) -> list[str]:
    """
    Lê a coluna B (índice 1), limpa os IDs e remove duplicados.
    """
    if df.shape[1] < 2:
        raise ValueError("A planilha não tem pelo menos 2 colunas.")

    process_col = df.iloc[:, 1]

    ids = []
    for value in process_col.tolist():
        cleaned = clean_process_id(value)
        if cleaned:
            ids.append(cleaned)

    # remove duplicados mantendo ordem
    unique_ids = list(dict.fromkeys(ids))
    return unique_ids


def build_process_url(process_id: str) -> str:
    base_url = os.getenv(
        "BASE_URL",
        "https://www.jsaduaneiros.com.br/webjs/doWorkOcorrencias?CodProcesso="
    )
    return f"{base_url}{process_id}"