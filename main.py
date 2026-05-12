import json
import os
import time
from datetime import datetime
from pathlib import Path
from turtle import delay

import pandas as pd
from bs4 import BeautifulSoup
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
    "ocorrencia": "#tabOcorrencia",
    "demonstrativo_despesas": "#tabDemoDespesa",
}

TAB_IDS = {
    "dados_gerais": "tabDadosGerais",
    "documentos": "tabDocumentos",
    "documentos_adicionais": "tabDocumentosAdicionais",
    "ocorrencia": "tabOcorrencia",
    "demonstrativo_despesas": "tabDemoDespesa",
}


def extract_tab_html(full_html: str, tab_id: str) -> str:
    soup = BeautifulSoup(full_html, "html.parser")
    tab = soup.select_one(f"#{tab_id}")

    if not tab:
        return ""

    return str(tab)


def str_to_bool(value: str | None, default: bool = False) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "sim"}

def env_bool(name: str, default: bool = False) -> bool:
    return str_to_bool(os.getenv(name), default=default)


def env_int(name: str, default: int = 0) -> int:
    try:
        return int(os.getenv(name, str(default)))
    except ValueError:
        return default


def env_float(name: str, default: float = 0.0) -> float:
    try:
        return float(os.getenv(name, str(default)).replace(",", "."))
    except ValueError:
        return default


def save_debug_html(process_id: str, html: str, suffix: str) -> str:
    if not env_bool("SAVE_DEBUG_HTML", default=False):
        return ""

    return save_html(process_id, html, suffix)


def save_process_json(filename: str, data: dict | list) -> str:
    if not env_bool("SAVE_PROCESS_JSON", default=False):
        return ""

    return save_json(filename, data)


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

    page.goto(base_url, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_selector("#tabDadosGerais", timeout=20000)

    extra_wait_ms = env_int("WEBJS_EXTRA_WAIT_MS", default=700)
    if  extra_wait_ms > 0:
        page.wait_for_timeout(extra_wait_ms)

    html_completo = page.content()
    summary = parse_summary(html_completo, process_id, page.url)

    save_debug_html(process_id, html_completo, "completo_webjs")
    save_process_json(f"processo_{process_id}_resumo.json", summary)

    resultado = {
        "process_id": process_id,
        "url_base": base_url,
        "resumo": summary,
    }


    for section_name, tab_id in TAB_IDS.items():
        try:
            section_html = extract_tab_html(html_completo, tab_id)

            if not section_html:
                raise RuntimeError(f"Aba {tab_id} não encontrada no HTML.")

            save_debug_html(process_id, section_html, section_name)

            info = inspect_html(section_html)

            if section_name == "dados_gerais":
                parsed = parse_dados_gerais(section_html)
            else:
                parsed = parse_generic_tab(section_html, section_name)

            resultado[section_name] = {
                "section_name": section_name,
                "tab_id": tab_id,
                "url": f"{base_url}#{tab_id}",
                "title": page.title(),
                "preview": info["preview"],
                "indicators": info["indicators"],
                **parsed,
            }

            save_process_json(f"processo_{process_id}_{section_name}.json", resultado[section_name])

        except Exception as e:
            resultado[section_name] = {
                "section_name": section_name,
                "erro": str(e),
            }

    run_siscoweb_cif = str_to_bool(
        os.getenv("RUN_SISCOWEB_CIF", "true"),
        default=True,
    )

    if run_siscoweb_cif:
        try:
            print(f"Consultando CIF no SISCOWEB do processo {process_id}...")
            resultado["mercadorias"] = scrape_mercadorias_cif(page, process_id)
            save_process_json(f"processo_{process_id}_mercadorias.json", resultado["mercadorias"])
            print("CIF encontrado:", resultado["mercadorias"].get("valor_cif_rs") or "-")
        except Exception as e:
            print(f"Erro ao consultar CIF no SISCOWEB do processo {process_id}: {e}")
            resultado["mercadorias"] = {
                "section_name": "mercadorias",
                "url": f"{os.getenv('SISCOWEB_BASE_URL', '')}{process_id}",
                "valor_cif_rs": "",
                "preview": "",
                "erro": str(e),
            }

    save_process_json(f"processo_{process_id}_completo.json", resultado)
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

def login_siscoweb_if_needed(page):
    """
    Faz login no SISCOWEB apenas se aparecer a tela de login.

    Campos confirmados no HTML:
    usuário: input[name='vSR']
    senha: input[name='vSNH']
    """

    try:
        has_login = page.locator("input[name='vSR']").count() > 0
        has_password = page.locator("input[name='vSNH']").count() > 0
    except Exception:
        return

    if not has_login and not has_password:
        return

    username = (
        os.getenv("SISCOWEB_USERNAME")
        or ""
    )

    password = (
        os.getenv("SISCOWEB_PASSWORD")
        or ""
    )

    if not username or not password:
        raise RuntimeError(
            "SISCOWEB pediu login, mas SISCOWEB_USERNAME/SISCOWEB_PASSWORD "
            "não estão definidos no .env."
        )

    page.locator("input[name='vSR']").fill(username)
    page.locator("input[name='vSNH']").fill(password)

    page.locator("input[name='vSNH']").press("Enter")

    page.wait_for_load_state("networkidle", timeout=60000)

def extract_valor_cif_rs_from_page(page) -> str:
    try:
        page.wait_for_selector("#valorCif", timeout=15000)
    except Exception:
        return ""

    last_value = ""

    # Tenta por até 2 segundos, sem esperar 1s + 1.5s fixos sempre.
    for _ in range(10):
        value = page.evaluate(
            """
            () => {
                const input = document.querySelector("#valorCif");
                return input ? input.value : "";
            }
            """
        )

        value = str(value or "").strip()
        last_value = value

        if value and value not in {"0,00", "0,0", "0"}:
            return value

        page.wait_for_timeout(200)

    return last_value

def scrape_mercadorias_cif(page, process_id: str) -> dict:
    base_url = (os.getenv("SISCOWEB_BASE_URL") or "").strip()

    if not base_url:
        raise RuntimeError("SISCOWEB_BASE_URL não está definido no .env.")

    url = f"{base_url}{process_id}"

    page.goto(url, wait_until="domcontentloaded", timeout=60000)

    login_siscoweb_if_needed(page)

    if "doMercadorias" not in page.url:
        page.goto(url, wait_until="domcontentloaded", timeout=60000)

    valor_cif_rs = extract_valor_cif_rs_from_page(page)

    preview = ""
    if env_bool("SISCOWEB_SAVE_PREVIEW", default=False):
        try:
            preview = page.locator("body").inner_text(timeout=5000)[:5000]
        except Exception:
            preview = ""

    return {
        "section_name": "mercadorias",
        "url": url,
        "valor_cif_rs": valor_cif_rs,
        "preview": preview,
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
        try:
            browser = p.chromium.launch(headless=headless)
        except Exception as e:
            print(f"Erro ao abrir Chromium com headless={headless}: {e}")
            print("Tentando abrir Chromium em modo visível...")
            browser = p.chromium.launch(headless=False)

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

            delay = env_float("PROCESS_DELAY_SECONDS", default=0.2)
            if delay > 0:
                time.sleep(delay)

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