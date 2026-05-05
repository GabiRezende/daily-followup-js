import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()


LOGS_DIR = Path("output/logs")
LOGS_DIR.mkdir(parents=True, exist_ok=True)


def str_to_bool(value: str | None, default: bool = False) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "sim", "s"}


def run_step(step_name: str, command: list[str], log_file: Path) -> bool:
    print("\n" + "=" * 80)
    print(f"INICIANDO: {step_name}")
    print("=" * 80)

    with log_file.open("a", encoding="utf-8") as log:
        log.write("\n" + "=" * 80 + "\n")
        log.write(f"INICIANDO: {step_name}\n")
        log.write(f"COMANDO: {' '.join(command)}\n")
        log.write("=" * 80 + "\n")

        process = subprocess.run(
            command,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
        )

        output = process.stdout or ""

        print(output)
        log.write(output)
        log.write(f"\nCÓDIGO DE SAÍDA: {process.returncode}\n")

    if process.returncode != 0:
        print(f"\nERRO: etapa '{step_name}' falhou.")
        print(f"Veja o log em: {log_file}")
        return False

    print(f"OK: etapa '{step_name}' concluída.")
    return True


def main():
    timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
    log_file = LOGS_DIR / f"run_daily_{timestamp}.log"

    run_main = str_to_bool(os.getenv("RUN_MAIN", "true"), default=True)
    run_export_coordenadora = str_to_bool(os.getenv("RUN_EXPORT_COORDENADORA", "true"), default=True)
    run_export_clientes = str_to_bool(os.getenv("RUN_EXPORT_CLIENTES", "true"), default=True)
    run_mailer = str_to_bool(os.getenv("RUN_MAILER", "true"), default=True)

    python_exe = sys.executable

    print("Execução diária iniciada.")
    print(f"Python usado: {python_exe}")
    print(f"Log: {log_file}")

    with log_file.open("w", encoding="utf-8") as log:
        log.write("EXECUÇÃO DIÁRIA DO FOLLOW-UP\n")
        log.write(f"Início: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        log.write(f"Python usado: {python_exe}\n")
        log.write(f"RUN_MAIN={run_main}\n")
        log.write(f"RUN_EXPORT_COORDENADORA={run_export_coordenadora}\n")
        log.write(f"RUN_EXPORT_CLIENTES={run_export_clientes}\n")
        log.write(f"RUN_MAILER={run_mailer}\n")
        log.write(f"DRY_RUN={os.getenv('DRY_RUN', '')}\n")
        log.write(f"SEND_ONLY={os.getenv('SEND_ONLY', '')}\n")

    steps = []

    if run_main:
        steps.append(("Consulta dos processos e geração do JSON", [python_exe, "main.py"]))

    if run_export_coordenadora:
        steps.append(("Exportação da planilha da coordenação", [python_exe, "export_coordenadora.py"]))

    if run_export_clientes:
        steps.append(("Exportação das planilhas dos clientes", [python_exe, "export_clientes.py"]))

    if run_mailer:
        steps.append(("Envio dos e-mails via Gmail API", [python_exe, "mailer_gmail_api.py"]))

    for step_name, command in steps:
        success = run_step(step_name, command, log_file)

        if not success:
            with log_file.open("a", encoding="utf-8") as log:
                log.write("\nEXECUÇÃO INTERROMPIDA POR ERRO.\n")
                log.write(f"Etapa com erro: {step_name}\n")
                log.write(f"Fim: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")

            sys.exit(1)

    with log_file.open("a", encoding="utf-8") as log:
        log.write("\nEXECUÇÃO FINALIZADA COM SUCESSO.\n")
        log.write(f"Fim: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")

    print("\n" + "=" * 80)
    print("EXECUÇÃO FINALIZADA COM SUCESSO.")
    print(f"Log salvo em: {log_file}")
    print("=" * 80)


if __name__ == "__main__":
    main()