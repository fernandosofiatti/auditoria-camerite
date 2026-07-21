"""Mantem o app do Streamlit Community Cloud acordado.

Abre o app com um navegador real (estabelece a sessao/websocket que o
Streamlit usa para medir atividade) e, se o app estiver dormindo, clica no
botao "Yes, get this app back up!" para reativa-lo.

Rodado pelo GitHub Actions em um cron -- nao depende de nenhum computador local.
"""

import os
import sys

from playwright.sync_api import sync_playwright

URL = os.environ.get("APP_URL", "https://auditoria-camerite.streamlit.app/")


def main() -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        print(f"Abrindo {URL}")
        page.goto(URL, wait_until="domcontentloaded", timeout=90_000)

        # Se o app dormiu, aparece o botao de reativar.
        wake_button = page.get_by_role(
            "button", name="Yes, get this app back up!"
        )
        try:
            if wake_button.is_visible(timeout=8_000):
                print("App estava dormindo -- clicando para reativar...")
                wake_button.click()
                # Reativar leva ~1-2 min. Espera o app subir.
                page.wait_for_timeout(120_000)
            else:
                print("App ja estava acordado.")
        except Exception as exc:  # botao nunca apareceu = app acordado
            print(f"Botao de reativar nao encontrado (app acordado): {exc}")

        # Mantem a sessao aberta um pouco para registrar atividade.
        page.wait_for_timeout(20_000)
        print(f"Titulo final da pagina: {page.title()!r}")
        browser.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
