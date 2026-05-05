import base64
import json
import mimetypes
import os
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path

from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

load_dotenv()

OUTPUT_CLIENTS_DIR = Path("output/clientes")
RECIPIENTS_FILE = Path(os.getenv("CLIENT_RECIPIENTS_FILE", "client_recipients.json"))

# Permissão mínima para enviar e-mail.
# Se você mudar o escopo depois de já ter criado token_gmail.json,
# apague o token_gmail.json e autorize de novo.
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]


def str_to_bool(value: str | None, default: bool = False) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "sim", "s"}


def split_env_list(value: str | None) -> set[str]:
    if not value:
        return set()
    return {item.strip().upper() for item in value.split(",") if item.strip()}


def latest_manifest() -> Path:
    manifests = sorted(OUTPUT_CLIENTS_DIR.glob("*/manifest.json"))
    if not manifests:
        raise FileNotFoundError(
            "Nenhum manifest.json encontrado em output/clientes. "
            "Rode primeiro: python export_clientes.py"
        )
    return manifests[-1]


def load_json(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def normalize_recipients_config(raw: dict) -> dict:
    normalized = {}

    for key, value in raw.items():
        clean_key = str(key).strip().upper()

        if isinstance(value, list):
            normalized[clean_key] = {"to": value, "cc": [], "bcc": []}
            continue

        if isinstance(value, dict):
            normalized[clean_key] = {
                "to": value.get("to", []) or [],
                "cc": value.get("cc", []) or [],
                "bcc": value.get("bcc", []) or [],
            }

    return normalized


def get_sender() -> tuple[str, str]:
    sender_email = os.getenv("MAIL_FROM_EMAIL") or os.getenv("GMAIL_USER")
    sender_name = os.getenv("MAIL_FROM_NAME", "JS Assessoria Aduaneira")

    if not sender_email:
        raise ValueError("Defina MAIL_FROM_EMAIL ou GMAIL_USER no .env")

    return sender_name, sender_email


def build_subject(client: str, date_br: str) -> str:
    template = os.getenv(
        "MAIL_SUBJECT_TEMPLATE",
        "Acompanhamento dos Processos em Aberto - {client} - {date}",
    )
    return template.format(client=client, date=date_br)


def build_body(client: str, row_count: int, date_br: str) -> str:
    template = os.getenv("MAIL_BODY_TEMPLATE")

    if template:
        return template.format(client=client, row_count=row_count, date=date_br)

    return (
        "Olá,\n\n"
        f"Segue em anexo o acompanhamento dos processos de {client}, referente a {date_br}.\n\n"
        "Qualquer dúvida, ficamos à disposição.\n\n"
        "Atenciosamente,\n"
        "JS Energy"
    )


def attach_file(message: EmailMessage, file_path: Path):
    if not file_path.exists():
        raise FileNotFoundError(f"Anexo não encontrado: {file_path}")

    content_type, _ = mimetypes.guess_type(file_path.name)
    if content_type is None:
        content_type = "application/octet-stream"

    maintype, subtype = content_type.split("/", 1)
    message.add_attachment(
        file_path.read_bytes(),
        maintype=maintype,
        subtype=subtype,
        filename=file_path.name,
    )


def build_message(client_item: dict, recipients: dict, date_br: str) -> EmailMessage:
    client_key = client_item["client_key"]
    client = client_item["client"]
    file_path = Path(client_item["path"])
    row_count = int(client_item.get("row_count", 0))

    config = recipients.get(client_key)
    if not config:
        raise KeyError(f"Cliente sem destinatário configurado: {client_key} ({client})")

    to_list = config.get("to", [])
    cc_list = config.get("cc", [])
    bcc_list = config.get("bcc", [])

    if not to_list:
        raise ValueError(f"Cliente sem e-mail em 'to': {client_key} ({client})")

    sender_name, sender_email = get_sender()

    msg = EmailMessage()
    msg["From"] = formataddr((sender_name, sender_email))
    msg["To"] = ", ".join(to_list)

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    # Com Gmail API não existe envelope separado como no SMTP.
    # Se usar Bcc, o Gmail usa esse header para entrega oculta.
    if bcc_list:
        msg["Bcc"] = ", ".join(bcc_list)

    msg["Subject"] = build_subject(client, date_br)
    msg.set_content(build_body(client, row_count, date_br))
    attach_file(msg, file_path)

    return msg


def get_gmail_service():
    credentials_path = Path(os.getenv("GMAIL_CREDENTIALS_FILE", "credentials_gmail.json"))
    token_path = Path(os.getenv("GMAIL_TOKEN_FILE", "token_gmail.json"))

    if not credentials_path.exists():
        raise FileNotFoundError(
            f"Arquivo de credenciais OAuth não encontrado: {credentials_path}\n"
            "Baixe o JSON do OAuth Client do Google Cloud e salve com esse nome."
        )

    creds = None

    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), SCOPES)
            creds = flow.run_local_server(port=0)

        token_path.write_text(creds.to_json(), encoding="utf-8")

    return build("gmail", "v1", credentials=creds)


def send_message_gmail_api(service, message: EmailMessage) -> dict:
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    body = {"raw": raw_message}

    return (
        service.users()
        .messages()
        .send(userId="me", body=body)
        .execute()
    )


def main():
    dry_run = str_to_bool(os.getenv("DRY_RUN"), default=True)
    send_only = split_env_list(os.getenv("SEND_ONLY"))

    manifest_path = Path(os.getenv("CLIENT_MANIFEST", "")) if os.getenv("CLIENT_MANIFEST") else latest_manifest()
    manifest = load_json(manifest_path)
    recipients = normalize_recipients_config(load_json(RECIPIENTS_FILE))

    date_br = manifest.get("date_br", "")
    files = manifest.get("files", [])

    print(f"Manifest: {manifest_path}")
    print(f"Modo dry-run: {dry_run}")

    messages = []

    for item in files:
        client_key = item.get("client_key", "").upper()

        if send_only and client_key not in send_only:
            continue

        if client_key not in recipients:
            print(f"PULANDO {client_key}: sem destinatário em {RECIPIENTS_FILE}")
            continue

        try:
            msg = build_message(item, recipients, date_br)
            messages.append((item, msg))
        except Exception as exc:
            print(f"ERRO preparando {client_key}: {exc}")

    if not messages:
        print("Nenhum e-mail para enviar.")
        return

    if dry_run:
        print("\nE-mails que seriam enviados:")
        for item, msg in messages:
            print("-" * 70)
            print(f"Cliente: {item['client']}")
            print(f"Para: {msg.get('To')}")
            print(f"Cc: {msg.get('Cc', '')}")
            print(f"Bcc: {msg.get('Bcc', '')}")
            print(f"Assunto: {msg.get('Subject')}")
            print(f"Anexo: {item['path']}")
        print("\nNada foi enviado porque DRY_RUN=true.")
        return

    service = get_gmail_service()

    for item, msg in messages:
        result = send_message_gmail_api(service, msg)
        print(f"ENVIADO: {item['client']} -> {msg.get('To')} | Gmail message id: {result.get('id')}")


if __name__ == "__main__":
    main()
