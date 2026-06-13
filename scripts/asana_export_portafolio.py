#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ASANA EXPORTER - Portafolio Zitron (todos los proyectos) + Board Servicios
----------------------------------------------------------------------------
Usa una sesion guardada (auth.json), igual que asana_export_ofertas.py,
pero en vez de un solo proyecto:

  1) Entra al PORTAFOLIO, detecta todos los proyectos que contiene y
     exporta cada uno como XLSX -> data_excels/
  2) Exporta el proyecto/board "Servicios y Mantencion" -> data_servicios/

Estos dos directorios son el input de generar_portafolio.py.

USO
---
  python3 asana_export_portafolio.py

Requiere auth.json (generado con guardar_sesion.py) en la carpeta padre.
"""

import re
import time
from pathlib import Path
from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent.parent
AUTH_FILE = BASE / "auth.json"

DATA_PROYECTOS = BASE / "data_excels"
DATA_SERVICIOS = BASE / "data_servicios"
DATA_PROYECTOS.mkdir(parents=True, exist_ok=True)
DATA_SERVICIOS.mkdir(parents=True, exist_ok=True)

# URL del portafolio "Proyectos" (el que tiene los ~42 proyectos)
PORTAFOLIO_URL = "https://app.asana.com/0/portfolio/1213511928397658/1213532266045644"

# URL del proyecto/board "Servicios y Mantencion"
SERVICIOS_URL = "https://app.asana.com/1/402967058777498/project/1213595645392940/board/1213596118449101"
SERVICIOS_NOMBRE = "Servicios y Mantencion"


def limpiar_nombre(nombre: str) -> str:
    nombre = re.sub(r"\s+", " ", nombre)
    for c in r'\/:*?"<>|':
        nombre = nombre.replace(c, "_")
    return nombre.strip()


# ---------------------------------------------------------------------
# Exportar un proyecto individual a XLSX (mismo flujo que asana_export_ofertas.py)
# ---------------------------------------------------------------------
def exportar_proyecto(page, nombre: str, url: str, carpeta_salida: Path) -> bool:
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
        ruta_destino = carpeta_salida / nombre_archivo
        download.save_as(ruta_destino)
        print(f"  OK -> Guardado: {ruta_destino.name}")
        return True

    except Exception as e:
        print(f"  ERROR: {e}")
        debug_dir = carpeta_salida.parent / "debug_portafolio"
        debug_dir.mkdir(exist_ok=True)
        nombre_debug = limpiar_nombre(nombre)
        try:
            page.screenshot(path=str(debug_dir / f"fallo_{nombre_debug}.png"), full_page=True)
            debug_dir.joinpath(f"fallo_{nombre_debug}.html").write_text(page.content(), encoding="utf-8")
            print("  -> Guardado screenshot/html de diagnostico en debug_portafolio/")
        except Exception as e2:
            print(f"  (no se pudo guardar diagnostico: {e2})")
        return False


# ---------------------------------------------------------------------
# Listar los proyectos contenidos en el portafolio
# ---------------------------------------------------------------------
def listar_proyectos_portafolio(page):
    """
    Entra al portafolio y devuelve una lista de (nombre, url) por cada
    proyecto que contiene.

    NOTA: el selector de las tarjetas/filas de proyecto dentro de un
    portafolio de Asana puede variar segun la version de la UI. El
    selector usado aqui (`a[href*="/project/"]` dentro del contenedor
    del portafolio) es el mas estable, pero si Asana cambia su DOM
    puede ser necesario ajustarlo. Si la lista sale vacia, revisar
    debug_portafolio/portafolio.png / .html.
    """
    page.goto(PORTAFOLIO_URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(5000)

    # Hacer scroll para forzar la carga de todas las filas (listas virtualizadas)
    for _ in range(40):
        page.mouse.wheel(0, 1200)
        page.wait_for_timeout(250)

    enlaces = page.locator('a[href*="/project/"]')
    n = enlaces.count()

    vistos = {}
    for i in range(n):
        try:
            href = enlaces.nth(i).get_attribute("href")
            texto = enlaces.nth(i).inner_text(timeout=2000).strip()
        except Exception:
            continue
        if not href or not texto:
            continue
        m = re.search(r"/project/(\d+)", href)
        if not m:
            continue
        gid = m.group(1)
        url = f"https://app.asana.com/0/{gid}/list"
        if gid not in vistos:
            vistos[gid] = (texto, url)

    proyectos = [(nombre, url) for gid, (nombre, url) in vistos.items()]

    if not proyectos:
        debug_dir = page.context.tracing  # placeholder, just to avoid unused import issues
        debug_dir2 = Path(__file__).resolve().parent.parent / "debug_portafolio"
        debug_dir2.mkdir(exist_ok=True)
        page.screenshot(path=str(debug_dir2 / "portafolio.png"), full_page=True)
        debug_dir2.joinpath("portafolio.html").write_text(page.content(), encoding="utf-8")
        print("  -> No se detectaron proyectos. Diagnostico guardado en debug_portafolio/")

    return proyectos


def main():
    if not AUTH_FILE.exists():
        raise SystemExit(
            f"No se encontro {AUTH_FILE}. "
            "Genera la sesion con guardar_sesion.py y configura el secret ASANA_AUTH."
        )

    print("=" * 60)
    print("  ASANA EXPORTER - Portafolio Zitron (todos los proyectos)")
    print(f"  Proyectos -> {DATA_PROYECTOS}")
    print(f"  Servicios -> {DATA_SERVICIOS}")
    print("=" * 60)

    ok_count = 0
    err_count = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True, storage_state=str(AUTH_FILE))
        page = context.new_page()

        # 1) Proyectos del portafolio
        proyectos = listar_proyectos_portafolio(page)
        print(f"\nProyectos detectados en el portafolio: {len(proyectos)}")
        for nombre, url in proyectos:
            ok = exportar_proyecto(page, nombre, url, DATA_PROYECTOS)
            if ok:
                ok_count += 1
            else:
                err_count += 1

        # 2) Board de servicios
        ok = exportar_proyecto(page, SERVICIOS_NOMBRE, SERVICIOS_URL, DATA_SERVICIOS)
        if ok:
            ok_count += 1
        else:
            err_count += 1

        browser.close()

    print("\n" + "=" * 60)
    print(f"  OK: {ok_count}   ERRORES: {err_count}")
    print("=" * 60)

    if err_count and ok_count == 0:
        raise SystemExit("No se pudo exportar nada. La sesion (auth.json) probablemente expiro.")


if __name__ == "__main__":
    main()
