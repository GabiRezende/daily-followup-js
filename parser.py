import re
from bs4 import BeautifulSoup

SUMMARY_FIELD_MAP = {
    "Cliente": "cliente",
    "Referência": "referencia",
    "Responsável atual": "responsavel_atual",
    "Status Processo": "status_processo",
    "Vinculado ao Processo": "vinculado_ao_processo",
}

DADOS_GERAIS_FIELD_MAP = {
    "Embarcação/Plataforma": "embarcacao_plataforma",
    "Tipo Processo": "tipo_processo",
    "Contrato": "contrato",
    "Tipo Declaração": "tipo_declaracao",
    "Valor Declarado": "valor_declarado",
    "Moeda": "moeda",
    "Nr. DI": "nr_di",
    "Dt.Registro DI": "dt_registro_di",
    "Data Moeda Declarada": "data_moeda_declarada",
    "Valor da Taxa": "valor_taxa_cambio",
    "Tipo Transporte": "tipo_transporte",
    "Data Embarque": "data_embarque",
    "Data Chegada Carga": "data_chegada_carga",
    "Doc.Chegada de Carga": "doc_chegada_carga",
    "Identificação Chegada Carga": "identificacao_chegada_carga",
    "Nome Transportador": "nome_transportador",
    "Identificação": "identificacao_transportador",
    "País Procedência da Carga": "pais_procedencia_carga",
    "URF Entrada no País": "urf_entrada_pais",
    "Peso Líquido": "peso_liquido",
    "Peso Bruto": "peso_bruto",
    "Incoterm": "incoterm",
    "Taxa Siscomex": "taxa_siscomex",
    "Conhecimento": "conhecimento",
    "Local Embarque": "local_embarque",
    "Identificação Master": "identificacao_master",
    "Pack List": "pack_list",
    "País Transportador": "pais_transportador",
    "Moeda Frete": "moeda_frete",
    "Frete Prepaid": "frete_prepaid",
    "Frete Collect": "frete_collect",
    "Frete Nacional": "frete_nacional",
    "Moeda Seguro": "moeda_seguro",
    "Valor Seguro": "valor_seguro",
 
}


def normalize_text(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text)
    return text.strip()


def html_to_text(html: str) -> str:
    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text("\n", strip=True)
    return normalize_text(text)


def inspect_html(html: str) -> dict:
    soup = BeautifulSoup(html, "lxml")
    full_text = soup.get_text(" ", strip=True)

    title = soup.title.get_text(strip=True) if soup.title else "Sem título"

    indicators = {
        "has_processo_word": "Processo:" in full_text,
        "has_resultado_consulta": "Resultado de Consulta" in full_text,
        "has_login_word": "login" in full_text.lower(),
        "has_usuario_word": "usuário" in full_text.lower() or "usuario" in full_text.lower(),
        "has_senha_word": "senha" in full_text.lower(),
    }

    return {
        "title": title,
        "preview": full_text[:800],
        "indicators": indicators,
    }


def extract_field(text: str, label: str, all_labels: list[str]) -> str | None:
    escaped_label = re.escape(label)
    other_labels = [re.escape(x) for x in all_labels if x != label]

    if other_labels:
        stop_pattern = "|".join(other_labels)
        pattern = rf"{escaped_label}\s*:?\s*(.*?)(?=\n(?:{stop_pattern})\s*:|\Z)"
    else:
        pattern = rf"{escaped_label}\s*:?\s*(.*?)(?=\Z)"

    match = re.search(pattern, text, flags=re.DOTALL)

    if not match:
        return None

    value = normalize_text(match.group(1))
    return value if value else None


def parse_labeled_fields(text: str, field_map: dict[str, str]) -> dict:
    labels = list(field_map.keys())
    data = {}

    for label, key in field_map.items():
        data[key] = extract_field(text, label, labels)

    return data


def extract_tables_from_html(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "lxml")
    tables = []

    for idx, table in enumerate(soup.find_all("table"), start=1):
        rows = []
        for tr in table.find_all("tr"):
            cells = [normalize_text(td.get_text(" ", strip=True)) for td in tr.find_all(["th", "td"])]
            cells = [c for c in cells if c]
            if cells:
                rows.append(cells)

        if rows:
            tables.append({
                "table_index": idx,
                "rows": rows
            })

    return tables


def parse_summary(html: str, process_id: str, url: str) -> dict:
    text = html_to_text(html)
    data = {
        "process_id": process_id,
        "url": url,
    }

    process_match = re.search(r"Processo:\s*(\d+)", text)
    if process_match:
        data["processo_topo"] = process_match.group(1)

    data.update(parse_labeled_fields(text, SUMMARY_FIELD_MAP))
    return data


def parse_dados_gerais(html: str) -> dict:
    text = html_to_text(html)
    data = parse_labeled_fields(text, DADOS_GERAIS_FIELD_MAP)
    data["preview"] = text[:1500]
    data["tables"] = extract_tables_from_html(html)
    return data


def parse_generic_tab(html: str, tab_name: str) -> dict:
    text = html_to_text(html)
    return {
        "tab_name": tab_name,
        "preview": text[:2000],
        "tables": extract_tables_from_html(html),
    }