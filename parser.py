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
    "URF de Despacho": "urf_despacho",
    "Recinto Aduaneiro": "recinto_aduaneiro",
    "Armazém": "armazem",
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

SUMMARY_STOP_LABELS = [
    "Dados Gerais",
    "Documentos",
    "Informações Adicionais",
    "Ocorrência",
    "Demonstrativo de Despesas",
    "Embarcação/Plataforma",
    "Tipo Processo",
]

DADOS_GERAIS_STOP_LABELS = [
    "Documentos",
    "Informações Adicionais",
    "Ocorrência",
    "Demonstrativo de Despesas",
    "Taxa Cambio",
    "Conhecimento / Transporte",
    "Carga",
    "Volumes",
    "Chegada",
    "Invoice(s)",
    "Ocorrências",
    "Pagamentos na Conta do Cliente",
    "Nova Ocorrência",
    "Novo Documento",
]

NOISE_LINES = {
    "Dados Gerais",
    "Documentos",
    "Informações Adicionais",
    "Ocorrência",
    "Demonstrativo de Despesas",
    "Processing...",
    "Procurar:",
    "close",
    "Nova Ocorrência",
    "Novo Documento",
}

TAB_LABELS = {
    "Dados Gerais",
    "Documentos",
    "Informações Adicionais",
    "Ocorrência",
    "Demonstrativo de Despesas",
}


def is_label_or_group_header(value: str | None) -> bool:
    if value is None:
        return True

    text = normalize_text(str(value)).strip(" .:-")
    if not text:
        return True

    lowered = text.lower()

    if "has_processo_word" in lowered or "has_resultado_consulta" in lowered:
        return True

    known = set(SUMMARY_FIELD_MAP.keys())
    known.update(DADOS_GERAIS_FIELD_MAP.keys())
    known.update(SUMMARY_STOP_LABELS)
    known.update(DADOS_GERAIS_STOP_LABELS)
    known.update(NOISE_LINES)
    known.update(TAB_LABELS)
    known.update({
        "Taxa Cambio",
        "Taxa Câmbio",
        "Conhecimento / Transporte",
        "Carga",
        "Volumes",
        "Chegada",
        "Invoice(s)",
        "Nr.Invoice",
        "Moeda",
        "Valor na Moeda",
        "Exportador",
        "Ocorrências",
        "Pagamentos na Conta do Cliente",
        "Nova Ocorrência",
        "Novo Documento",
    })

    known_lower = {normalize_text(item).strip(" .:-").lower() for item in known}

    if lowered in known_lower:
        return True

    if lowered.startswith("invoice(s)") and ("nr.invoice" in lowered or "exportador" in lowered):
        return True

    return False


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


def remove_noise_lines(text: str) -> str:
    lines = [line.strip() for line in text.split("\n")]
    cleaned = []

    for line in lines:
        if not line:
            continue
        if line in NOISE_LINES:
            continue
        cleaned.append(line)

    return normalize_text("\n".join(cleaned))


def build_label_pattern(label: str) -> str:
    return rf"(?:^|\n)\s*{re.escape(label)}(?:\.+\s*)?:"


def clean_extracted_value(label: str, value: str | None) -> str | None:
    if value is None:
        return None

    value = normalize_text(value)
    value = remove_noise_lines(value)

    field_cut_markers = {
        "Responsável atual": [
            "Status Processo",
            "Vinculado ao Processo",
            "Dados Gerais",
            "Documentos",
        ],
        "Status Processo": [
            "Vinculado ao Processo",
            "Dados Gerais",
            "Documentos",
            "Informações Adicionais",
        ],
        "Vinculado ao Processo": [
            "Dados Gerais",
            "Documentos",
            "Informações Adicionais",
            "Ocorrência",
            "Demonstrativo de Despesas",
        ],
        "Embarcação/Plataforma": [
            "Tipo Processo",
            "Contrato",
        ],
        "Tipo Declaração": [
            "Taxa Cambio",
            "Valor Declarado",
            "Data Moeda Declarada",
        ],
        "Nr. DI": [
            "Contrato",
            "Dt.Registro DI",
            "Tipo Declaração",
        ],
        "Dt.Registro DI": [
            "Tipo Declaração",
            "Taxa Cambio",
        ],
        "Data Embarque": [
            "Local Embarque",
            "Doc.Chegada de Carga",
        ],
        "Data Chegada Carga": [
            "URF de Despacho",
            "Recinto Aduaneiro",
            "Setor de Armazenamento",
        ],
        "Local Embarque": [
            "Doc.Chegada de Carga",
            "Identificação Master",
        ],
        "Identificação Master": [
            "Identificação Chegada Carga",
            "Pack List",
        ],
        "Pack List": [
            "Nome Transportador",
            "País Transportador",
        ],
        "Moeda": [
            "Valor da Taxa",
            "Conhecimento / Transporte",
        ],
        "Valor da Taxa": [
            "Conhecimento / Transporte",
            "Tipo Transporte",
        ],
        "Moeda Frete": [
            "URF Entrada no País",
            "Frete Prepaid",
        ],
        "Frete Prepaid": [
            "Peso Líquido",
            "Frete Collect",
        ],
        "Frete Collect": [
            "Peso Bruto",
            "Frete Nacional",
        ],
        "Frete Nacional": [
            "Incoterm",
            "Moeda Seguro",
        ],
        "Moeda Seguro": [
            "Taxa Siscomex",
            "Valor Seguro",
        ],
        "Valor Seguro": [
            "Volumes",
            "Chegada",
        ],
        "Recinto Aduaneiro": [
            "Setor de Armazenamento",
            "Armazém",
            "Invoice(s)",
        ],
        "Armazém": [
            "Invoice(s)",
            "Ocorrências",
        ],
    }

    markers = field_cut_markers.get(label, [])

    cut_positions = []
    for marker in markers:
        pos = value.find(marker)
        if pos > 0:
            cut_positions.append(pos)

    if cut_positions:
        value = value[:min(cut_positions)]

    value = normalize_text(value)
    value = re.sub(r"^[\.\:\-\s]+", "", value)
    value = re.sub(r"[\.\:\-\s]+$", "", value)
    value = normalize_text(value)

    lines = [line.strip() for line in value.split("\n") if line.strip()]
    if lines and all(line in TAB_LABELS for line in lines):
        return None

    if label == "Vinculado ao Processo":
        return None

    if is_label_or_group_header(value):
        return None

    return value if value else None


def extract_field(text: str, label: str, all_labels: list[str]) -> str | None:
    start_pattern = build_label_pattern(label)
    other_patterns = [build_label_pattern(x) for x in all_labels if x != label]

    if other_patterns:
        stop_pattern = "|".join(other_patterns)
        pattern = rf"{start_pattern}\s*(.*?)(?=\n(?:{stop_pattern})|\Z)"
    else:
        pattern = rf"{start_pattern}\s*(.*?)(?=\Z)"

    match = re.search(pattern, text, flags=re.DOTALL)

    if not match:
        return None

    raw_value = match.group(1)
    return clean_extracted_value(label, raw_value)


def parse_labeled_fields(text: str, field_map: dict[str, str], extra_stop_labels: list[str] | None = None) -> dict:
    labels = list(field_map.keys())
    if extra_stop_labels:
        labels = labels + list(extra_stop_labels)

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


def cut_summary_block(text: str) -> str:
    start = text.find("Processo:")
    if start != -1:
        text = text[start:]

    stop_markers = [
        "\nDados Gerais",
        "\nDocumentos",
        "\nInformações Adicionais",
        "\nOcorrência",
        "\nDemonstrativo de Despesas",
        "\nEmbarcação/Plataforma",
        "\nTipo Processo",
    ]

    positions = [text.find(marker) for marker in stop_markers if text.find(marker) != -1]
    if positions:
        text = text[:min(positions)]

    return remove_noise_lines(text)


def cut_dados_gerais_block(text: str) -> str:
    start_markers = [
        "\nEmbarcação/Plataforma",
        "\nTipo Processo",
    ]

    start = -1
    for marker in start_markers:
        pos = text.find(marker)
        if pos != -1:
            start = pos
            break

    if start == -1:
        return remove_noise_lines(text)

    stop_markers = [
        "\nDocumentos",
        "\nInformações Adicionais",
        "\nOcorrência",
        "\nDemonstrativo de Despesas",
        "\nNova Ocorrência",
        "\nNovo Documento",
        "\nPagamentos na Conta do Cliente",
    ]

    end_positions = [text.find(marker, start) for marker in stop_markers if text.find(marker, start) != -1]
    if end_positions:
        text = text[start:min(end_positions)]
    else:
        text = text[start:]

    return remove_noise_lines(text)


def parse_summary(html: str, process_id: str, url: str) -> dict:
    text = html_to_text(html)
    text = cut_summary_block(text)

    data = {
        "process_id": process_id,
        "url": url,
    }

    process_match = re.search(r"Processo[\.\s]*:\s*(\d+)", text)
    if process_match:
        data["processo_topo"] = process_match.group(1)

    data.update(parse_labeled_fields(text, SUMMARY_FIELD_MAP, SUMMARY_STOP_LABELS))
    data["preview"] = text[:1000]
    data["vinculado_ao_processo"] = None

    return data


def parse_dados_gerais(html: str) -> dict:
    text = html_to_text(html)
    text = cut_dados_gerais_block(text)

    data = parse_labeled_fields(text, DADOS_GERAIS_FIELD_MAP, DADOS_GERAIS_STOP_LABELS)
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