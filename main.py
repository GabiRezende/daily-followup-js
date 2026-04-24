import json
import os
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from auth import ensure_logged_in, STATE_FILE
from playwright.sync_api import sync_playwright

from sheet_reader import load_sheet_dataframe, get_process_ids, build_process_url
from parser import (
    inspect_html,
    parse_summary,
    parse_dados_gerais,
    parse_generic_tab,
)

load_dotenv()

TAB_HASHES = {
    "dados_gerais": "#tabDadosGerais",
    "documentos": "#tabDocumentos",
    "documentos_adicionais": "#tabDocumentosAdicionais",
}


def str_to_bool(value: str | None, default: bool = False) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "sim"}


def save_html(process_id: str, html: str, suffix: str) -> str:
    output_dir = Path("output") / "html"
    output_dir.mkdir(parents=True, exist_ok=True)

    file_path = output_dir / f"processo_{process_id}_{suffix}.html"
    file_path.write_text(html, encoding="utf-8")
    return str(file_path)


def save_json(filename: str, data: dict | list) -> str:
    output_dir = Path("output") / "json"
    output_dir.mkdir(parents=True, exist_ok=True)

    file_path = output_dir / filename
    file_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    return str(file_path)

def save_daily_csv(rows: list[dict], filename: str) -> str:
    output_dir = Path("output") / "reports"
    output_dir.mkdir(parents=True, exist_ok=True)

    file_path = output_dir / filename
    df = pd.DataFrame(rows)
    df.to_csv(file_path, index=False, encoding="utf-8-sig")
    return str(file_path)


def activate_tab_by_hash(page, base_url: str, tab_hash: str):
    tab_url = f"{base_url}{tab_hash}"

    page.goto(tab_url, wait_until="load", timeout=60000)
    page.wait_for_timeout(1500)

    page.evaluate(
        """hash => {
            window.location.hash = hash;
            const link = document.querySelector(`a[href$="${hash}"]`);
            if (link) link.click();
        }""",
        tab_hash
    )
    page.wait_for_timeout(1500)


def capture_section(page, process_id: str, base_url: str, section_name: str) -> dict:
    tab_hash = TAB_HASHES[section_name]
    activate_tab_by_hash(page, base_url, tab_hash)

    html = page.content()
    info = inspect_html(html)

    save_html(process_id, html, section_name)
    
    if section_name == "dados_gerais":
        parsed = parse_dados_gerais(html)
    else:
        parsed = parse_generic_tab(html, section_name)

    merged = {
        "section_name": section_name,
        "tab_hash": tab_hash,
        "url": page.url,
        "title": page.title(),
        "preview": info["preview"],
        "indicators": info["indicators"],
        **parsed,
    }

    save_json(f"processo_{process_id}_{section_name}.json", merged)
    return merged


def process_one(page, process_id: str) -> dict:
    base_url = build_process_url(process_id)

    page.goto(base_url, wait_until="load", timeout=60000)
    page.wait_for_timeout(2000)

    html_inicial = page.content()
    summary = parse_summary(html_inicial, process_id, page.url)

    save_html(process_id, html_inicial, "resumo")
    save_json(f"processo_{process_id}_resumo.json", summary)

    resultado = {
        "process_id": process_id,
        "url_base": base_url,
        "resumo": summary,
    }

    sections = [
        "dados_gerais",
        "documentos",
        "documentos_adicionais",
    ]

    for section_name in sections:
        try:
            resultado[section_name] = capture_section(page, process_id, base_url, section_name)
        except Exception as e:
            resultado[section_name] = {
                "section_name": section_name,
                "erro": str(e),
            }

    save_json(f"processo_{process_id}_completo.json", resultado)
    return resultado


def build_flat_row(process_data: dict) -> dict:
    resumo = process_data.get("resumo", {})
    dados_gerais = process_data.get("dados_gerais", {})

    return {
        "process_id": process_data.get("process_id"),
        "cliente": resumo.get("cliente"),
        "referencia": resumo.get("referencia"),
        "responsavel_atual": resumo.get("responsavel_atual"),
        "status_processo": resumo.get("status_processo"),
        "vinculado_ao_processo": resumo.get("vinculado_ao_processo"),
        "tipo_processo": dados_gerais.get("tipo_processo"),
        "tipo_declaracao": dados_gerais.get("tipo_declaracao"),
        "nr_di": dados_gerais.get("nr_di"),
        "data_embarque": dados_gerais.get("data_embarque"),
        "incoterm": dados_gerais.get("incoterm"),
        "url_base": process_data.get("url_base"),
    }


def group_by_client(results: list[dict]) -> dict:
    grouped = {}

    for item in results:
        cliente = item.get("resumo", {}).get("cliente") or "CLIENTE_NAO_IDENTIFICADO"
        grouped.setdefault(cliente, []).append(item)

    return grouped


def build_daily_output(results: list[dict]) -> dict:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    grouped = group_by_client(results)

    return {
        "generated_at": now,
        "total_processos": len(results),
        "total_clientes": len(grouped),
        "clientes": grouped,
    }


def main():
    max_processes = int(os.getenv("MAX_PROCESSES", "0"))
    headless = str_to_bool(os.getenv("HEADLESS"), default=False)

    print("\nLendo planilha pública...")
    df = load_sheet_dataframe()
    process_ids = get_process_ids(df)

    print(f"Total de Processos JS encontrados na planilha: {len(process_ids)}")

    if not process_ids:
        print("Nenhum processo encontrado.")
        return

    if max_processes <= 0:
        selected_ids = process_ids
    else:
        selected_ids = process_ids[:max_processes]

    print(f"Processos que serão consultados agora: {len(selected_ids)}")

    results = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)

        if STATE_FILE.exists():
            context = browser.new_context(
                storage_state=str(STATE_FILE),
                ignore_https_errors=True,
            )
        else:
            context = browser.new_context(
                ignore_https_errors=True,
            )

        page = context.new_page()
        ensure_logged_in(page, context)

        for index, process_id in enumerate(selected_ids, start=1):
            print("\n" + "=" * 70)
            print(f"[{index}/{len(selected_ids)}] Consultando processo {process_id}")

            try:
                result = process_one(page, process_id)
                results.append(result)

                resumo = result.get("resumo", {})
                print("Cliente:", resumo.get("cliente"))
                print("Status:", resumo.get("status_processo"))
                print("Referência:", resumo.get("referencia"))

            except Exception as e:
                print(f"Erro no processo {process_id}: {e}")
                results.append({
                    "process_id": process_id,
                    "erro_geral": str(e),
                })

            time.sleep(1.5)

        browser.close()

    today = datetime.now().strftime("%Y%m%d_%H%M%S")

    consolidated = build_daily_output(results)
    consolidated_json_path = save_json(
        f"followup_consolidado_{today}.json",
        consolidated
    )

    flat_rows = [build_flat_row(item) for item in results if item.get("resumo")]
    consolidated_csv_path = save_daily_csv(
        flat_rows,
        f"followup_resumo_{today}.csv"
    )

    print("\n" + "=" * 70)
    print("FINALIZADO")
    print("JSON consolidado:", consolidated_json_path)
    print("CSV resumo:", consolidated_csv_path)


if __name__ == "__main__":
    main()