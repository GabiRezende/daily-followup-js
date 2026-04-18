import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

STATE_FILE = Path("auth_state.json")


def page_looks_logged_out(page) -> bool:
    return page.locator("#vSR").count() > 0 and page.locator("#vSNH").count() > 0


def do_login(page, context):
    login_url = os.getenv("JS_LOGIN_URL")
    username = os.getenv("JS_USERNAME")
    password = os.getenv("JS_PASSWORD")

    if not login_url or not username or not password:
        raise ValueError("JS_LOGIN_URL, JS_USERNAME ou JS_PASSWORD não encontrados no .env")

    page.goto(login_url, wait_until="load", timeout=60000)
    page.wait_for_timeout(1500)

    page.locator("#vSR").fill(username)
    page.locator("#vSNH").fill(password)

    try:
        page.locator("#Login").click(timeout=5000)
    except Exception:
        page.evaluate("""
            const btn = document.getElementById('Login');
            if (btn) btn.click();
        """)

    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2500)

    if page_looks_logged_out(page):
        raise RuntimeError("Login não funcionou. Verifique usuário, senha ou se a tela mudou.")

    context.storage_state(path=str(STATE_FILE))


def ensure_logged_in(page, context):
    login_url = os.getenv("JS_LOGIN_URL")
    if not login_url:
        raise ValueError("JS_LOGIN_URL não encontrado no .env")

    page.goto(login_url, wait_until="load", timeout=60000)
    page.wait_for_timeout(1500)

    if page_looks_logged_out(page):
        do_login(page, context)
    else:
        context.storage_state(path=str(STATE_FILE))