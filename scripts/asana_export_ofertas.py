#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ASANA EXPORTER - Oferta Zitron 2026 (proyecto unico)
Usa una sesion guardada (auth.json) para exportar SOLO el proyecto
"Oferta Zitron 2026" en formato XLSX a la carpeta data_ofertas/.

Mismo mecanismo que scripts/asana_export.py pero acotado a un solo
proyecto, para que esta automatizacion sea independiente.
"""

import time
from pathlib import Path
from playwright.sync_api import sync_playwright

CARPETA_SALIDA = Path(__file__).resolve().parent.parent / "data_ofertas"
CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

AUTH_FILE = Path(__file__).resolve().parent.parent / "auth.json"

NOMBRE = "Oferta Zitron 2026"
URL = "https://app.asana.com/1/402967058777498/project/1214299938971075"


def limpiar_nombre(nombre: str) -> str:
    for c in r'\/:*?"<>|':
        nombre = nombre.replace(c, "_")
    return nombre.strip()


def exportar_proyecto(page, nombre: str, url: str) -> bool:
    print(f"\nExportando: {nombre}")
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_timeout(4000)

        menu = page.locator('[role="button"][aria-label="Acciones"]')
        menu.first.wait_for(state="visible", timeout=20000)
        menu.first.click(timeout=15000)
        time.sleep(0.8)

        export_menu = page.get_by_text("Exportar o sincronizar", exact=False)
        export_menu.first.wait_for(state="visible", timeout=10000)
        export_menu.first.hover(timeout=10000)
        time.sleep(1.0)
        export_menu.first.hover(timeout=10000)
        time.sleep(0.8)

        opcion_csv = page.get_by_text("Tareas del proyecto en formato CSV/XLSX", exact=False)
        opcion_csv.first.click(timeout=15000)
        time.sleep(0.8)

        xlsx_radio = page.get_by_text("XLSX", exact=True)
        if xlsx_radio.count() > 0:
            xlsx_radio.first.click(timeout=5000)
            time.sleep(0.3)

        with page.expect_download(timeout=30000) as download_info:
            boton_exportar = page.get_by_role("button", name="Exportar", exact=True)
            boton_exportar.first.click(timeout=15000)

        download = download_info.value
        sufijo = Path(download.suggested_filename).suffix or ".xlsx"
        nombre_archivo = limpiar_nombre(nombre) + sufijo
        ruta_destino = CARPETA_SALIDA / nombre_archivo
        download.save_as(ruta_destino)
        print(f"  OK -> Guardado: {ruta_destino.name}")
        return True

    except Exception as e:
        print(f"  ERROR: {e}")
        debug_dir = CARPETA_SALIDA.parent / "debug_ofertas"
        debug_dir.mkdir(exist_ok=True)
        try:
            page.screenshot(path=str(debug_dir / "fallo.png"), full_page=True)
            debug_dir.joinpath("fallo.html").write_text(page.content(), encoding="utf-8")
            print("  -> Guardado screenshot/html de diagnostico en debug_ofertas/")
        except Exception as e2:
            print(f"  (no se pudo guardar diagnostico: {e2})")
        return False


def main():
    if not AUTH_FILE.exists():
        raise SystemExit(
            f"No se encontro {AUTH_FILE}. "
            "Genera la sesion con guardar_sesion.py y configura el secret ASANA_AUTH."
        )

    print("=" * 60)
    print("  ASANA EXPORTER - Oferta Zitron 2026")
    print(f"  Carpeta de salida: {CARPETA_SALIDA}")
    print("=" * 60)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True, storage_state=str(AUTH_FILE))
        page = context.new_page()
        ok = exportar_proyecto(page, NOMBRE, URL)
        browser.close()

    if not ok:
        raise SystemExit("No se pudo exportar el proyecto. La sesion (auth.json) probablemente expiro.")


if __name__ == "__main__":
    main()
