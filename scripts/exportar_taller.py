#!/usr/bin/env python3
"""
EXPORTADOR DEL PROYECTO DE TALLER (independiente del KPI)

Misma logica que exportar_asana.py, pero con UN SOLO proyecto y una
carpeta de salida propia: data-taller/.

Se mantiene aparte a proposito. El KPI exporta ~59 proyectos a data/;
si este usara la misma carpeta, generar_kpi.py tomaria el proyecto de
taller y generar_gantt.py tomaria los 59. Carpetas separadas, tableros
separados.

USO
---
  python3 exportar_taller.py
"""

import time
from pathlib import Path

from playwright.sync_api import sync_playwright

RAIZ = Path(__file__).resolve().parent.parent
CARPETA_SALIDA = RAIZ / "data-taller"
CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
AUTH_FILE = RAIZ / "auth.json"

# ── El unico proyecto de este tablero ────────────────────────────────
# Cambia el nombre por el que uses en Asana: es el que aparece como
# titulo de la fila en el Gantt si la tarea no trae OT.
PROYECTOS = [
    ("CARGA DE TALLER", "https://app.asana.com/1/402967058777498/project/1217946516695478"),
]


def limpiar_nombre(nombre: str) -> str:
    for c in r'\/:*?"<>|':
        nombre = nombre.replace(c, "_")
    return nombre.strip()


def exportar_proyecto(page, nombre: str, url: str, indice: int, total: int) -> bool:
    print(f"\n[{indice}/{total}] {nombre}")
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_timeout(6000)

        # Diagnostico: que pagina esta viendo realmente el navegador
        print(f"     URL final : {page.url}")
        print(f"     Titulo    : {page.title()}")
        if "login" in page.url or "account" in page.url:
            print("     -> Asana redirigio al login: la sesion no sirve.")
            raise RuntimeError("Sesion invalida")

        # El menu de acciones cambia de etiqueta segun el idioma de la
        # cuenta, asi que se prueban varias formas antes de rendirse.
        candidatos = [
            '[role="button"][aria-label="Acciones"]',
            '[role="button"][aria-label="Actions"]',
            '[aria-label="Acciones"]',
            '[aria-label="Actions"]',
            '[aria-label*="cciones"]',
            '[aria-label*="ctions"]',
        ]
        menu = None
        for sel in candidatos:
            loc = page.locator(sel)
            if loc.count() > 0:
                print(f"     Menu encontrado con: {sel}")
                menu = loc.first
                break
        if menu is None:
            # Ultimo recurso: listar que aria-labels hay en la cabecera
            etiquetas = page.eval_on_selector_all(
                "[aria-label]",
                "els => els.slice(0,40).map(e => e.getAttribute('aria-label'))")
            print(f"     No se hallo el menu. aria-labels visibles: {etiquetas}")
            raise RuntimeError("Menu de acciones no encontrado")

        menu.click(timeout=15000)
        time.sleep(0.8)

        export_menu = page.get_by_text("Exportar o sincronizar", exact=False)
        if export_menu.count() == 0:
            export_menu = page.get_by_text("Export", exact=False)
        export_menu.first.wait_for(state="visible", timeout=10000)
        export_menu.first.hover(timeout=10000)
        time.sleep(1.0)
        export_menu.first.hover(timeout=10000)
        time.sleep(0.8)

        opcion_csv = page.get_by_text("Tareas del proyecto en formato CSV/XLSX", exact=False)
        if opcion_csv.count() == 0:
            opcion_csv = page.get_by_text("CSV/XLSX", exact=False)
        opcion_csv.first.click(timeout=15000)
        time.sleep(0.8)

        xlsx_radio = page.get_by_text("XLSX", exact=True)
        if xlsx_radio.count() > 0:
            xlsx_radio.first.click(timeout=5000)
            time.sleep(0.3)

        with page.expect_download(timeout=30000) as download_info:
            boton = page.get_by_role("button", name="Exportar", exact=True)
            if boton.count() == 0:
                boton = page.get_by_role("button", name="Export", exact=True)
            boton.first.click(timeout=15000)

        download = download_info.value
        sufijo = Path(download.suggested_filename).suffix or ".xlsx"
        destino = CARPETA_SALIDA / (limpiar_nombre(nombre) + sufijo)
        download.save_as(destino)
        print(f"     ✓ Guardado: {destino.name}")
        return True

    except Exception as e:
        print(f"     ✗ ERROR: {e}")
        debug = RAIZ / "debug-taller"
        debug.mkdir(exist_ok=True)
        try:
            page.screenshot(path=str(debug / "fallo.png"), full_page=True)
            debug.joinpath("fallo.html").write_text(page.content(), encoding="utf-8")
            print(f"     -> Diagnostico en {debug}")
        except Exception as e2:
            print(f"     (no se pudo guardar diagnostico: {e2})")
        return False


def main():
    if not AUTH_FILE.exists():
        raise SystemExit(
            f"No se encontro {AUTH_FILE}. "
            "Genera la sesion con guardar_sesion.py y configura el secret ASANA_AUTH."
        )

    print("=" * 60)
    print("  EXPORT PROYECTO DE TALLER")
    print(f"  Carpeta de salida: {CARPETA_SALIDA}")
    print("=" * 60)

    # Limpia exports viejos: si el proyecto se renombra en Asana, el
    # archivo anterior quedaria y el Gantt mostraria datos duplicados.
    for viejo in CARPETA_SALIDA.glob("*.xlsx"):
        viejo.unlink()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True,
                                      storage_state=str(AUTH_FILE))
        page = context.new_page()

        exitosos = 0
        for i, (nombre, url) in enumerate(PROYECTOS, start=1):
            if exportar_proyecto(page, nombre, url, i, len(PROYECTOS)):
                exitosos += 1
            time.sleep(1.5)

        browser.close()

    print("\n" + "=" * 60)
    print(f"  COMPLETADO: {exitosos}/{len(PROYECTOS)}")
    print("=" * 60)

    if exitosos == 0:
        raise SystemExit("No se exporto el proyecto. La sesion (auth.json) probablemente expiro.")


if __name__ == "__main__":
    main()
