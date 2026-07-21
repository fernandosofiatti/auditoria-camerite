"""Mantem o app do Streamlit Community Cloud acordado.

Abre o app com um navegador real (Chromium via Playwright) e, se o app
estiver dormindo, clica no botao "Yes, get this app back up!" para reativa-lo.
Faz varias tentativas com reload para lidar com a demora de renderizacao.

Rodado pelo GitHub Actions em um cron -- nao depende de nenhum computador local.
"""

import os
import sys

from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as PWTimeout

URL = os.environ["APP_URL"]
SLEEP_TEXT = "get this app back up"
SHOT = "estado_final.png"


def app_dormindo(page) -> bool:
    try:
        return SLEEP_TEXT in page.inner_text("body", timeout=5_000)
    except Exception:
        return False


def main() -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_context().new_page()
        print(f"Abrindo {URL}")
        page.goto(URL, wait_until="load", timeout=90_000)

        reativado = False
        for tentativa in range(1, 5):
            page.wait_for_timeout(6_000)  # deixa o JS renderizar a tela

            if not app_dormindo(page):
                print(f"Tentativa {tentativa}: app ja esta acordado.")
                reativado = True
                break

            print(f"Tentativa {tentativa}: app dormindo -- clicando para reativar...")
            btn = page.get_by_role("button", name="get this app back up")
            try:
                btn.click(timeout=20_000)
            except PWTimeout:
                print("  botao nao clicavel a tempo; recarregando...")
                page.reload(wait_until="load")
                continue

            # espera o texto de "dormindo" sumir = o clique funcionou e o app subiu
            try:
                page.wait_for_function(
                    "!document.body.innerText.includes('get this app back up')",
                    timeout=60_000,
                )
                print("  clique registrado -- app esta subindo.")
                reativado = True
                break
            except PWTimeout:
                print("  ainda na tela de sono; recarregando e tentando de novo...")
                page.reload(wait_until="load")

        # tempo para o app terminar de subir e registrar a sessao/atividade
        page.wait_for_timeout(30_000)
        try:
            page.screenshot(path=SHOT, full_page=True)
        except Exception:
            pass
        print(f"Titulo final: {page.title()!r} | reativado={reativado}")
        browser.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
