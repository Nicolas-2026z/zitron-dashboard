#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ASANA EXPORTER - Portafolio Zitron (automatico) + Board Servicios
------------------------------------------------------------------
FLUJO:
  1) Descarga el CSV del portafolio "Consolidado proyectos" via Playwright
  2) Lee el CSV para extraer nombre y GID de cada proyecto
  3) Exporta cada proyecto como XLSX -> data/
  4) Exporta el board "Servicios y Mantencion" -> data/servicios/

Si agrega un proyecto nuevo al portafolio en Asana, se detecta
automaticamente sin tocar este script.

PROYECTOS_EXTRA: proyectos que NO aparecen en el CSV del portafolio
pero igual deben exportarse (ej: proyectos de subcarpetas).
"""

import csv
import io
import re
import time
from pathlib import Path
from playwright.sync_api import sync_playwright

BASE        = Path(__file__).resolve().parent.parent
AUTH_FILE   = BASE / "auth.json"
DATA_PROYECTOS = BASE / "data"
DATA_SERVICIOS = BASE / "data" / "servicios"
DATA_PROYECTOS.mkdir(parents=True, exist_ok=True)
DATA_SERVICIOS.mkdir(parents=True, exist_ok=True)

# GID del portafolio "Consolidado proyectos"
PORTAFOLIO_GID = "1213511928397658"
PORTAFOLIO_URL = f"https://app.asana.com/0/portfolio/{PORTAFOLIO_GID}/list"

# Board de servicios
SERVICIOS_URL    = "https://app.asana.com/1/402967058777498/project/1213595645392940/board/1213596118449101"
SERVICIOS_NOMBRE = "Servicios y Mantencion"

WORKSPACE_GID = "402967058777498"

# Proyectos extra que NO salen en el CSV del portafolio (agregar si es necesario)
PROYECTOS_EXTRA = [
    # ("NOMBRE", "https://app.asana.com/1/.../project/GID"),
]


# ─────────────────────────────────────────────────────────────────────────────
def limpiar_nombre(nombre: str) -> str:
    nombre = re.sub(r"\s+", " ", nombre)
    for c in r'\/:*?"<>|':
        nombre = nombre.replace(c, "_")
    return nombre.strip()


# ─────────────────────────────────────────────────────────────────────────────
def descargar_csv_portafolio(page) -> list[tuple[str, str]]:
    """
    Entra al portafolio, descarga el CSV y devuelve lista de (nombre, url).
    El CSV del portafolio contiene columna 'Name' y 'Linked projects' o
    directamente filas con el nombre y GID del proyecto.
    """
    print(f"\nDescargando CSV del portafolio...")
    page.goto(PORTAFOLIO_URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(5000)

    # Abrir menu del portafolio (los tres puntos o chevron del header)
    try:
        menu = page.locator('[aria-label="Acciones del portafolio"], [aria-label="Portfolio actions"]')
        if menu.count() == 0:
            menu = page.locator('[data-testid="portfolio-actions-button"]')
        if menu.count() == 0:
            # Fallback: boton de tres puntos en el header
            menu = page.locator('button').filter(has_text="").nth(0)
        menu.first.click(timeout=15000)
    except Exception:
        # Intentar con el menu contextual del portafolio via kebab
        page.keyboard.press("Escape")
        page.wait_for_timeout(500)
        # Click en el chevron del nombre del portafolio
        chevron = page.locator('[aria-label="Opciones del portafolio"], [aria-label="Portfolio options"]')
        chevron.first.click(timeout=10000)

    time.sleep(0.8)

    # Buscar "Exportar o sincronizar"
    export_btn = page.get_by_text("Exportar o sincronizar", exact=False)
    export_btn.first.hover(timeout=10000)
    time.sleep(0.8)

    # Click en CSV
    with page.expect_download(timeout=30000) as dl_info:
        page.get_by_text("CSV", exact=True).first.click(timeout=10000)

    download = dl_info.value
    csv_text = download.path()  # archivo temporal
    content = Path(csv_text).read_text(encoding="utf-8-sig", errors="replace")

    proyectos = _parsear_csv_portafolio(content)
    print(f"  Proyectos detectados en CSV: {len(proyectos)}")
    return proyectos


def _parsear_csv_portafolio(content: str) -> list[tuple[str, str]]:
    """
    Parsea el CSV del portafolio y devuelve (nombre, url_proyecto).
    Asana exporta el portafolio con columnas como:
      Name, Owner, Status, Start Date, Due Date, ...
    El GID del proyecto no viene directo, pero si viene la columna
    'Linked projects' o podemos inferirlo de otra columna.
    En algunos exports viene 'Project URL' o similar.
    Si no hay URL, construimos la URL desde el nombre buscando el GID
    en el contenido del CSV (a veces viene como ID).
    """
    reader = csv.DictReader(io.StringIO(content))
    rows = list(reader)
    if not rows:
        return []

    headers = [h.strip().lower() for h in rows[0].keys()]
    print(f"  Columnas CSV: {list(rows[0].keys())[:10]}")

    proyectos = []
    for row in rows:
        # Normalizar keys
        row_norm = {k.strip().lower(): v for k, v in row.items()}

        nombre = (row_norm.get("name") or row_norm.get("nombre") or "").strip()
        if not nombre:
            continue

        # Buscar URL o GID del proyecto
        url = ""
        # Intentar columna URL directa
        for col in ["url", "project url", "enlace", "link"]:
            if col in row_norm and row_norm[col].strip():
                url = row_norm[col].strip()
                break

        # Si no hay URL, buscar GID numerico en cualquier columna
        if not url:
            for col, val in row_norm.items():
                m = re.search(r'\b(\d{16,})\b', str(val))
                if m:
                    gid = m.group(1)
                    url = f"https://app.asana.com/1/{WORKSPACE_GID}/project/{gid}/list"
                    break

        if url:
            proyectos.append((nombre, url))
        else:
            print(f"  [AVISO] Sin URL para: {nombre}")

    return proyectos


# ─────────────────────────────────────────────────────────────────────────────
def exportar_proyecto(page, nombre: str, url: str, carpeta: Path,
                      indice: int = 0, total: int = 0) -> bool:
    prefix = f"[{indice}/{total}] " if total else ""
    print(f"\n{prefix}{nombre}")
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
        ruta_destino = carpeta / nombre_archivo
        download.save_as(ruta_destino)
        print(f"  ✓ Guardado: {ruta_destino.name}")
        return True

    except Exception as e:
        print(f"  ✗ ERROR: {e}")
        debug_dir = BASE / "debug_portafolio"
        debug_dir.mkdir(exist_ok=True)
        try:
            nombre_debug = limpiar_nombre(nombre)[:40]
            page.screenshot(path=str(debug_dir / f"fallo_{nombre_debug}.png"), full_page=True)
            debug_dir.joinpath(f"fallo_{nombre_debug}.html").write_text(
                page.content(), encoding="utf-8")
        except Exception:
            pass
        return False


# ─────────────────────────────────────────────────────────────────────────────
def main():
    if not AUTH_FILE.exists():
        raise SystemExit(f"No se encontro {AUTH_FILE}.")

    print("=" * 60)
    print("  ASANA EXPORTER - Portafolio Zitron (automatico)")
    print(f"  Proyectos -> {DATA_PROYECTOS}")
    print(f"  Servicios -> {DATA_SERVICIOS}")
    print("=" * 60)

    ok_count  = 0
    err_count = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True,
                                      storage_state=str(AUTH_FILE))
        page = context.new_page()

        # 1) Obtener lista de proyectos desde el CSV del portafolio
        try:
            proyectos = descargar_csv_portafolio(page)
        except Exception as e:
            print(f"\n[AVISO] No se pudo descargar el CSV del portafolio: {e}")
            print("  Usando lista de proyectos extra solamente.")
            proyectos = []

        # Agregar proyectos extra que no estan en el portafolio
        proyectos += PROYECTOS_EXTRA

        # Deduplicar por URL
        vistos = set()
        proyectos_unicos = []
        for nombre, url in proyectos:
            if url not in vistos:
                vistos.add(url)
                proyectos_unicos.append((nombre, url))

        total = len(proyectos_unicos)
        print(f"\nTotal proyectos a exportar: {total}")

        # 2) Exportar cada proyecto
        for i, (nombre, url) in enumerate(proyectos_unicos, 1):
            ok = exportar_proyecto(page, nombre, url, DATA_PROYECTOS, i, total)
            if ok:
                ok_count += 1
            else:
                err_count += 1
            time.sleep(1.5)

        # 3) Board de servicios
        ok = exportar_proyecto(page, SERVICIOS_NOMBRE, SERVICIOS_URL, DATA_SERVICIOS)
        if ok:
            ok_count += 1
        else:
            err_count += 1

        browser.close()

    print("\n" + "=" * 60)
    print(f"  OK: {ok_count}   ERRORES: {err_count}")
    print("=" * 60)

    if ok_count == 0:
        raise SystemExit("No se exporto nada. Sesion probablemente expirada.")


if __name__ == "__main__":
    main()
