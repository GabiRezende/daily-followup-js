import os
import re
import json
from pathlib import Path
from datetime import datetime

import pandas as pd

from export_coordenadora import (
    latest_consolidated_json,
    build_row,
    extract_doc_rows,
)


HISTORY_DIR = Path("output/history")
HISTORY_DIR.mkdir(parents=True, exist_ok=True)

EXECUCOES_CSV = HISTORY_DIR / "execucoes.csv"
PROCESSOS_CSV = HISTORY_DIR / "processos_snapshot.csv"
MUDANCAS_CSV = HISTORY_DIR / "mudancas_snapshot.csv"
ALERTAS_CSV = HISTORY_DIR / "alertas_snapshot.csv"


FIELDS_TO_TRACK = [
    "Status",
    "Canal",
    "Data de Chegada",
    "Data Desembaraço",
    "Data Registro DI/DUIMP",
    "Numero DI/DUIMP",
    "CIF (R$)",
    "Recolhido II",
    "Recolhido IPI",
    "Recolhido PIS",
    "Recolhido COFINS",
    "Recolhido ICMS",
    "Taxa Siscomex",
]


def env_float(name: str, default: float) -> float:
    value = os.getenv(name)

    if not value:
        return default

    try:
        return float(value.replace(".", "").replace(",", "."))
    except ValueError:
        return default


def env_int(name: str, default: int) -> int:
    value = os.getenv(name)

    if not value:
        return default

    try:
        return int(value)
    except ValueError:
        return default


def clean_text(value) -> str:
    if value is None:
        return ""

    text = str(value).replace("\xa0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def parse_money_br(value) -> float:
    text = clean_text(value)

    if not text:
        return 0.0

    text = re.sub(r"[^\d,.-]", "", text)

    if not text:
        return 0.0

    text = text.replace(".", "").replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return 0.0


def parse_date_br(value):
    text = clean_text(value)

    if not text:
        return None

    match = re.search(r"\d{2}/\d{2}/\d{4}", text)

    if not match:
        return None

    try:
        return datetime.strptime(match.group(0), "%d/%m/%Y").date()
    except ValueError:
        return None


def read_csv(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()

    return pd.read_csv(path, dtype=str, encoding="utf-8-sig").fillna("")


def append_csv(path: Path, rows: list[dict]):
    if not rows:
        return

    new_df = pd.DataFrame(rows)

    if path.exists():
        old_df = read_csv(path)
        final_df = pd.concat([old_df, new_df], ignore_index=True)
    else:
        final_df = new_df

    final_df.to_csv(path, index=False, encoding="utf-8-sig")


def load_consolidated():
    consolidated_path = latest_consolidated_json()

    with open(consolidated_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    return consolidated_path, data


def flatten_processes(data: dict) -> list[dict]:
    processos = []

    for _, items in (data.get("clientes") or {}).items():
        processos.extend(items)

    return processos


def build_current_snapshot(data: dict) -> tuple[str, str, list[dict], list[dict]]:
    generated_at = data.get("generated_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    execution_id = datetime.now().strftime("%Y%m%d_%H%M%S")

    processos_raw = flatten_processes(data)

    snapshot_rows = []
    raw_by_process_id = {}

    for process_data in processos_raw:
        row = build_row(process_data)
        process_id = clean_text(row.get("Código JS") or process_data.get("process_id"))

        raw_by_process_id[process_id] = process_data

        snapshot_row = {
            "execution_id": execution_id,
            "generated_at": generated_at,
            **row,
        }

        snapshot_rows.append(snapshot_row)

    return execution_id, generated_at, snapshot_rows, raw_by_process_id


def get_previous_snapshot(current_execution_id: str) -> pd.DataFrame:
    history = read_csv(PROCESSOS_CSV)

    if history.empty:
        return pd.DataFrame()

    previous = history[history["execution_id"] != current_execution_id]

    if previous.empty:
        return pd.DataFrame()

    last_execution_id = previous["execution_id"].iloc[-1]
    return previous[previous["execution_id"] == last_execution_id].copy()


def generate_changes(
    execution_id: str,
    generated_at: str,
    current_rows: list[dict],
    previous_df: pd.DataFrame,
) -> list[dict]:
    if previous_df.empty:
        return []

    previous_by_process = {
        clean_text(row.get("Código JS")): row
        for row in previous_df.to_dict(orient="records")
    }

    changes = []

    for current in current_rows:
        process_id = clean_text(current.get("Código JS"))
        previous = previous_by_process.get(process_id)

        if not previous:
            changes.append({
                "execution_id": execution_id,
                "generated_at": generated_at,
                "Código JS": process_id,
                "Cliente": current.get("Cliente", ""),
                "campo": "NOVO_PROCESSO",
                "valor_anterior": "",
                "valor_atual": "Processo apareceu nesta execução",
            })
            continue

        for field in FIELDS_TO_TRACK:
            old_value = clean_text(previous.get(field, ""))
            new_value = clean_text(current.get(field, ""))

            if old_value != new_value:
                changes.append({
                    "execution_id": execution_id,
                    "generated_at": generated_at,
                    "Código JS": process_id,
                    "Cliente": current.get("Cliente", ""),
                    "campo": field,
                    "valor_anterior": old_value,
                    "valor_atual": new_value,
                })

    return changes


def get_status_stagnation_days(process_id: str, current_status: str, current_date) -> int:
    history = read_csv(PROCESSOS_CSV)

    if history.empty or not current_status or current_date is None:
        return 0

    history = history[history["Código JS"].astype(str) == str(process_id)]

    if history.empty or "Status" not in history.columns:
        return 0

    same_status = history[history["Status"].astype(str) == current_status]

    if same_status.empty:
        return 0

    dates = []

    for value in same_status.get("generated_at", []):
        try:
            dates.append(datetime.strptime(str(value)[:10], "%Y-%m-%d").date())
        except ValueError:
            continue

    if not dates:
        return 0

    first_seen = min(dates)
    return (current_date - first_seen).days


def extract_new_documents_today(process_data: dict, generated_at: str) -> list[str]:
    try:
        today = datetime.strptime(generated_at[:10], "%Y-%m-%d").date()
    except ValueError:
        today = datetime.now().date()

    documentos = process_data.get("documentos", {}) or {}
    adicionais = process_data.get("documentos_adicionais", {}) or {}

    rows = extract_doc_rows(documentos) + extract_doc_rows(adicionais)

    found = []

    for row in rows:
        data_entrega = parse_date_br(row.get("data_entrega"))

        if data_entrega != today:
            continue

        tipo = clean_text(row.get("tipo_documento"))
        numero = clean_text(row.get("nr_documento"))

        if tipo:
            found.append(f"{tipo}: {numero}" if numero else tipo)

    return found


def add_alert(alerts: list[dict], execution_id: str, generated_at: str, row: dict, tipo: str, gravidade: str, mensagem: str, valor_referencia: str = ""):
    alerts.append({
        "execution_id": execution_id,
        "generated_at": generated_at,
        "Código JS": row.get("Código JS", ""),
        "Cliente": row.get("Cliente", ""),
        "Referência Cliente": row.get("Referência Cliente", ""),
        "Status": row.get("Status", ""),
        "tipo_alerta": tipo,
        "gravidade": gravidade,
        "mensagem": mensagem,
        "valor_referencia": valor_referencia,
    })


def generate_alerts(
    execution_id: str,
    generated_at: str,
    current_rows: list[dict],
    raw_by_process_id: dict,
    changes: list[dict],
) -> list[dict]:
    alerts = []

    cif_limit = env_float("ALERTA_CIF_ALTO_LIMITE", 100000.0)
    status_parado_dias = env_int("ALERTA_STATUS_PARADO_DIAS", 5)

    try:
        current_date = datetime.strptime(generated_at[:10], "%Y-%m-%d").date()
    except ValueError:
        current_date = datetime.now().date()

    for row in current_rows:
        process_id = clean_text(row.get("Código JS"))
        modal = clean_text(row.get("Modal")).upper()
        status = clean_text(row.get("Status")).upper()

        numero_di_duimp = clean_text(row.get("Numero DI/DUIMP"))
        canal = clean_text(row.get("Canal"))
        data_chegada = clean_text(row.get("Data de Chegada"))
        data_desembaraco = clean_text(row.get("Data Desembaraço"))
        cif = parse_money_br(row.get("CIF (R$)"))

        is_import_process = modal not in {"ADMINISTRATIVO", "COMERCIAL"}

        if is_import_process and not numero_di_duimp:
            add_alert(
                alerts,
                execution_id,
                generated_at,
                row,
                "SEM_DI_DUIMP",
                "média",
                f"Processo {process_id} ainda está sem DI/DUIMP.",
            )

        if numero_di_duimp and not canal:
            add_alert(
                alerts,
                execution_id,
                generated_at,
                row,
                "SEM_CANAL",
                "média",
                f"Processo {process_id} tem DI/DUIMP, mas ainda está sem canal.",
                numero_di_duimp,
            )

        if data_chegada and not data_desembaraco:
            add_alert(
                alerts,
                execution_id,
                generated_at,
                row,
                "CHEGADA_SEM_DESEMBARACO",
                "alta",
                f"Processo {process_id} tem data de chegada, mas ainda não tem data de desembaraço.",
                data_chegada,
            )

        if cif_limit > 0 and cif >= cif_limit:
            add_alert(
                alerts,
                execution_id,
                generated_at,
                row,
                "CIF_ALTO",
                "atenção",
                f"Processo {process_id} tem CIF acima do limite configurado.",
                row.get("CIF (R$)", ""),
            )

        days_stopped = get_status_stagnation_days(process_id, clean_text(row.get("Status")), current_date)

        if days_stopped >= status_parado_dias:
            add_alert(
                alerts,
                execution_id,
                generated_at,
                row,
                "STATUS_PARADO",
                "atenção",
                f"Processo {process_id} está com o mesmo status há {days_stopped} dias.",
                row.get("Status", ""),
            )

        process_data = raw_by_process_id.get(process_id, {})
        new_docs = extract_new_documents_today(process_data, generated_at)

        for doc in new_docs:
            add_alert(
                alerts,
                execution_id,
                generated_at,
                row,
                "DOCUMENTO_NOVO_HOJE",
                "info",
                f"Processo {process_id} recebeu documento hoje: {doc}.",
                doc,
            )

    for change in changes:
        if change.get("campo") == "Status":
            fake_row = {
                "Código JS": change.get("Código JS", ""),
                "Cliente": change.get("Cliente", ""),
                "Referência Cliente": "",
                "Status": change.get("valor_atual", ""),
            }

            add_alert(
                alerts,
                execution_id,
                generated_at,
                fake_row,
                "STATUS_MUDOU",
                "info",
                f"Status mudou de '{change.get('valor_anterior')}' para '{change.get('valor_atual')}'.",
                change.get("valor_atual", ""),
            )

    return alerts


def main():
    consolidated_path, data = load_consolidated()

    execution_id, generated_at, snapshot_rows, raw_by_process_id = build_current_snapshot(data)

    previous_df = get_previous_snapshot(execution_id)
    changes = generate_changes(execution_id, generated_at, snapshot_rows, previous_df)
    alerts = generate_alerts(execution_id, generated_at, snapshot_rows, raw_by_process_id, changes)

    execucao_row = {
        "execution_id": execution_id,
        "generated_at": generated_at,
        "json_origem": str(consolidated_path),
        "total_processos": len(snapshot_rows),
        "total_mudancas": len(changes),
        "total_alertas": len(alerts),
    }

    append_csv(EXECUCOES_CSV, [execucao_row])
    append_csv(PROCESSOS_CSV, snapshot_rows)
    append_csv(MUDANCAS_CSV, changes)
    append_csv(ALERTAS_CSV, alerts)

    print("\nHistórico e alertas gerados com sucesso.")
    print(f"JSON origem: {consolidated_path}")
    print(f"Execução: {execution_id}")
    print(f"Processos: {len(snapshot_rows)}")
    print(f"Mudanças: {len(changes)}")
    print(f"Alertas: {len(alerts)}")
    print(f"Arquivos em: {HISTORY_DIR}")


if __name__ == "__main__":
    main()