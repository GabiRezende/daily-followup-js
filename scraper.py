from pathlib import Path
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def fetch_process_page(url: str) -> tuple[str, int, str]:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/123.0.0.0 Safari/537.36"
        )
    }

    response = requests.get(
        url,
        headers=headers,
        timeout=30,
        allow_redirects=True,
        verify=False,   # <- só para teste
    )
    response.raise_for_status()

    return response.text, response.status_code, response.url


def save_html(process_id: str, html: str) -> str:
    output_dir = Path("output") / "html"
    output_dir.mkdir(parents=True, exist_ok=True)

    file_path = output_dir / f"processo_{process_id}.html"
    file_path.write_text(html, encoding="utf-8")

    return str(file_path)