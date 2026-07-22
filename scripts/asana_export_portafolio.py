#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ASANA EXPORTER - Portafolio Zitron (lista fija 42 proyectos) + Board Servicios
Para agregar un proyecto nuevo: agregar una linea a PROYECTOS con (nombre, url).
"""

import re
import time
from pathlib import Path
from playwright.sync_api import sync_playwright

BASE           = Path(__file__).resolve().parent.parent
AUTH_FILE      = BASE / "auth.json"
DATA_PROYECTOS = BASE / "data"
DATA_SERVICIOS = BASE / "data" / "servicios"
DATA_PROYECTOS.mkdir(parents=True, exist_ok=True)
DATA_SERVICIOS.mkdir(parents=True, exist_ok=True)

SERVICIOS_URL    = "https://app.asana.com/1/402967058777498/project/1213595645392940/board/1213596118449101"
SERVICIOS_NOMBRE = "Servicios y Mantencion"

# Portafolio "Consolidado proyectos" (vista con columna Fecha de entrega por proyecto)
PORTAFOLIO_URL    = "https://app.asana.com/0/portfolio/1213511928397658/1213532266045644"
PORTAFOLIO_NOMBRE = "PORTAFOLIO_dashboard_fechas"

# ── Lista de los 42 proyectos del portafolio consolidado ──────────────────────
PROYECTOS = [
    ("50001559 - FERROVIAL",                               "https://app.asana.com/1/402967058777498/project/1215504785832997"),
    ("50001558 - GESVIAL OT4342-4344",                    "https://app.asana.com/1/402967058777498/project/1215137857526702"),
    ("50001557 - ATACAMA KOZAN OT4340",                   "https://app.asana.com/1/402967058777498/project/1215137857526520"),
    ("50001557 - ATACAMA KOZAN OT4335-4337-4339",         "https://app.asana.com/1/402967058777498/project/1215139673478155"),
    ("50001554 - MAPIMI",                                  "https://app.asana.com/1/402967058777498/project/1215118022011721"),
    ("50001553 - ZITRON COLOMBIA ARIS OT4326",            "https://app.asana.com/1/402967058777498/project/1214969047721199"),
    ("50001553 - ZITRON COLOMBIA ARIS OT4327",            "https://app.asana.com/1/402967058777498/project/1214969047720950"),
    ("50001553 - ZITRON COLOMBIA ARIS OT4324-4325",       "https://app.asana.com/1/402967058777498/project/1214969284200793"),
    ("50001555 - ZITRON COLOMBIA ARIS MINING SEGOVIA",    "https://app.asana.com/1/402967058777498/project/1214922603742599"),
    ("50001544 - DMC MINING SERVICES",                    "https://app.asana.com/1/402967058777498/project/1214892476594825"),
    ("50001547 - XEMORTIZ AMERICAS GOLD",                 "https://app.asana.com/1/402967058777498/project/1214787972061369"),
    ("50001546 - XEMORTIZ",                               "https://app.asana.com/1/402967058777498/project/1214739697907723"),
    ("50001545 - XEMORTIZ MINERA FRISCO",                 "https://app.asana.com/1/402967058777498/project/1214563704105168"),
    ("50001543 - ZITRON COLOMBIA EGM",                    "https://app.asana.com/1/402967058777498/project/1214508463084754"),
    ("50001525 - ZITRON PERU METRO LIMA E07",             "https://app.asana.com/1/402967058777498/project/1213832650589314"),
    ("50001534 - ZITRON PERU METRO LIMA E04-05-06-R4",    "https://app.asana.com/1/402967058777498/project/1213881596172396"),
    ("50001541 - XEMORTIZ MINERA FRISCO B",               "https://app.asana.com/1/402967058777498/project/1214137717389412"),
    ("50001532 - XEMORTIZ AMERICAS GOLD B",               "https://app.asana.com/1/402967058777498/project/1213963037596622"),
    ("50001466 - ZITRON PERU PODEROSA",                   "https://app.asana.com/1/402967058777498/project/1213997266064436"),
    ("50001485 - CALABRESSE METRO STO DOMINGO",           "https://app.asana.com/1/402967058777498/project/1213377149548665"),
    ("50001498 - CALABRESSE METRO STO DOMINGO POZOS",     "https://app.asana.com/1/402967058777498/project/1213377149548798"),
    ("50001477 - TRITON MINERA",                          "https://app.asana.com/1/402967058777498/project/1213377149548590"),
    ("50001490 - ZITRON PERU METRO LIMA PV04-PV05",       "https://app.asana.com/1/402967058777498/project/1213377149548731"),
    ("50001504 - ZITRON PERU CODESTABLE A",               "https://app.asana.com/1/402967058777498/project/1213400421980747"),
    ("50001504 - ZITRON PERU CODESTABLE B",               "https://app.asana.com/1/402967058777498/project/1213377149548924"),
    ("50001415 - ZITRON PERU METRO LIMA E05",             "https://app.asana.com/1/402967058777498/project/1213391072478428"),
    ("50001506 - ZCO EPM FRIO AIRE",                      "https://app.asana.com/1/402967058777498/project/1213244147627519"),
    ("50001451 - ZITRON COLOMBIA",                        "https://app.asana.com/1/402967058777498/project/1213391072478496"),
    ("50001500 - EQ MIN LA HACIENDA A",                   "https://app.asana.com/1/402967058777498/project/1213234368822465"),
    ("50001497 - XEMORTIZ IMMSA STA BBR",                 "https://app.asana.com/1/402967058777498/project/1213193352022841"),
    ("50001499 - EQ CHAPARRAL A",                         "https://app.asana.com/1/402967058777498/project/1213244147627934"),
    ("50001508 - XEMORTIZ STOCK",                         "https://app.asana.com/1/402967058777498/project/1213362937195444"),
    ("50001473 - MAPIMI B",                               "https://app.asana.com/1/402967058777498/project/1213377149548525"),
    ("50001501 - EQ MIN LA HACIENDA B",                   "https://app.asana.com/1/402967058777498/project/1213377149548860"),
    ("50001446 - MINERA FRESNILLO",                       "https://app.asana.com/1/402967058777498/project/1213377149548988"),
    ("50001520 - XEMORTIZ DDG",                           "https://app.asana.com/1/402967058777498/project/1213396195711984"),
    ("50001518 - EQ CHAPARRAL B",                         "https://app.asana.com/1/402967058777498/project/1213458785834993"),
    ("50001533 - GESVIAL",                                "https://app.asana.com/1/402967058777498/project/1213955236952392"),
    ("50001524 - ATACAMA KOZAN",                          "https://app.asana.com/1/402967058777498/project/1213458788103499"),
    ("50001514 - CLIENTE POR DEFINIR A",                  "https://app.asana.com/1/402967058777498/project/1213185599076077"),
    ("50001515 - CLIENTE POR DEFINIR B",                  "https://app.asana.com/1/402967058777498/project/1213234368821945"),
    ("50001563 - ZITRON COLOMBIA ANTIOQUIA",              "https://app.asana.com/1/402967058777498/project/1215727332551578"),
    ("50001566 - XEMORTIZ",                               "https://app.asana.com/1/402967058777498/project/1215922055649550"),
    ("50001569 - EUROHINCA",                              "https://app.asana.com/1/402967058777498/project/1216208608300132"),
    ("50001579- XEMORTIZ C. MEXICANA",                    "https://app.asana.com/1/402967058777498/project/1216769310026482"),
    ("50001577- XEMORTIZ LA CANTERA",                     "https://app.asana.com/1/402967058777498/project/1216769310026176"),
    ("50001576- XEMORTIZ FIRST MAJESTIC",                 "https://app.asana.com/1/402967058777498/project/1216730497556427"),
    ("50001580- CIA. RIO MINERALES"",                     "https://app.asana.com/1/402967058777498/project/1216769310026525"),
]


def limpiar_nombre(nombre: str) -> str:
    nombre = re.sub(r"\s+", " ", nombre)
    for c in r'\/:*?"<>|':
        nombre = nombre.replace(c, "_")
    return nombre.strip()


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
        print(f"  OK -> {ruta_destino.name}")
        return True

    except Exception as e:
        print(f"  ERROR: {e}")
        debug_dir = BASE / "debug_portafolio"
        debug_dir.mkdir(exist_ok=True)
        try:
            nd = limpiar_nombre(nombre)[:40]
            page.screenshot(path=str(debug_dir / f"fallo_{nd}.png"), full_page=True)
            debug_dir.joinpath(f"fallo_{nd}.html").write_text(page.content(), encoding="utf-8")
        except Exception:
            pass
        return False


def exportar_portafolio(page, nombre: str, url: str, carpeta: Path) -> bool:
    """
    Exporta el portafolio (vista con columna 'Fecha de entrega') a XLSX.
    El menu de export de un portafolio puede estar en un lugar distinto al
    de un proyecto normal, asi que se prueban varias rutas conocidas de Asana
    antes de rendirse. Si ninguna funciona, se guarda screenshot + HTML de
    debug igual que exportar_proyecto(), para poder ajustar el selector.
    """
    print(f"\n[Portafolio] {nombre}")
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_timeout(4000)

        # Intento 1: mismo boton "Acciones" que usan los proyectos.
        menu = page.locator('[role="button"][aria-label="Acciones"]')
        abierto = False
        if menu.count() > 0:
            try:
                menu.first.wait_for(state="visible", timeout=8000)
                menu.first.click(timeout=8000)
                abierto = True
            except Exception:
                abierto = False

        # Intento 2: boton de opciones "..." generico (mas comun en portafolios).
        if not abierto:
            for sel in ['[aria-label="Mas opciones"]', '[aria-label="More actions"]',
                        '[aria-label="Opciones"]', 'button:has-text("...")']:
                loc = page.locator(sel)
                if loc.count() > 0:
                    try:
                        loc.first.click(timeout=8000)
                        abierto = True
                        break
                    except Exception:
                        continue

        if not abierto:
            raise RuntimeError("No se encontro el boton de menu/acciones del portafolio")

        time.sleep(0.8)

        # Buscar la opcion de exportar (puede llamarse distinto en portafolio vs proyecto)
        export_opcion = None
        for texto in ["Exportar o sincronizar", "Exportar", "Export"]:
            loc = page.get_by_text(texto, exact=False)
            if loc.count() > 0:
                export_opcion = loc.first
                break
        if export_opcion is None:
            raise RuntimeError("No se encontro la opcion 'Exportar' en el menu del portafolio")

        export_opcion.hover(timeout=10000)
        time.sleep(1.0)
        export_opcion.hover(timeout=10000)
        time.sleep(0.8)

        # Submenu: preferir CSV/XLSX; si el portafolio solo ofrece CSV, tomar esa.
        opcion_csv = None
        for texto in ["formato CSV/XLSX", "CSV/XLSX", "Exportar como CSV", "CSV"]:
            loc = page.get_by_text(texto, exact=False)
            if loc.count() > 0:
                opcion_csv = loc.first
                break
        if opcion_csv is not None:
            opcion_csv.click(timeout=15000)
            time.sleep(0.8)

        xlsx_radio = page.get_by_text("XLSX", exact=True)
        if xlsx_radio.count() > 0:
            xlsx_radio.first.click(timeout=5000)
            time.sleep(0.3)

        with page.expect_download(timeout=30000) as download_info:
            boton_exportar = page.get_by_role("button", name="Exportar", exact=True)
            if boton_exportar.count() > 0:
                boton_exportar.first.click(timeout=15000)
            elif opcion_csv is not None:
                pass  # el click a la opcion de CSV ya pudo haber disparado la descarga
            else:
                raise RuntimeError("No se encontro boton 'Exportar' para confirmar la descarga")

        download = download_info.value
        sufijo = Path(download.suggested_filename).suffix or ".xlsx"
        nombre_archivo = limpiar_nombre(nombre) + sufijo
        ruta_destino = carpeta / nombre_archivo
        download.save_as(ruta_destino)
        print(f"  OK -> {ruta_destino.name}")
        return True

    except Exception as e:
        print(f"  ERROR exportando portafolio: {e}")
        debug_dir = BASE / "debug_portafolio"
        debug_dir.mkdir(exist_ok=True)
        try:
            page.screenshot(path=str(debug_dir / "fallo_portafolio_fechas.png"), full_page=True)
            debug_dir.joinpath("fallo_portafolio_fechas.html").write_text(page.content(), encoding="utf-8")
            print(f"  Debug guardado en {debug_dir} (revisar para ajustar el selector)")
        except Exception:
            pass
        return False


def main():
    if not AUTH_FILE.exists():
        raise SystemExit(f"No se encontro {AUTH_FILE}.")

    print("=" * 60)
    print("  ASANA EXPORTER - Portafolio Zitron")
    print(f"  Proyectos ({len(PROYECTOS)}) -> {DATA_PROYECTOS}")
    print(f"  Servicios -> {DATA_SERVICIOS}")
    print("=" * 60)

    ok_count = 0
    err_count = 0
    total = len(PROYECTOS)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True,
                                      storage_state=str(AUTH_FILE))
        page = context.new_page()

        for i, (nombre, url) in enumerate(PROYECTOS, 1):
            ok = exportar_proyecto(page, nombre, url, DATA_PROYECTOS, i, total)
            if ok:
                ok_count += 1
            else:
                err_count += 1
            time.sleep(1.5)

        ok = exportar_proyecto(page, SERVICIOS_NOMBRE, SERVICIOS_URL, DATA_SERVICIOS)
        if ok:
            ok_count += 1
        else:
            err_count += 1

        ok = exportar_portafolio(page, PORTAFOLIO_NOMBRE, PORTAFOLIO_URL, DATA_PROYECTOS)
        if ok:
            ok_count += 1
        else:
            err_count += 1
            print("  [AVISO] No se pudo exportar el portafolio con fechas de entrega.")
            print("  El resto del portafolio se genera igual, solo faltaran esas fechas.")

        browser.close()

    print("\n" + "=" * 60)
    print(f"  OK: {ok_count}   ERRORES: {err_count}")
    print("=" * 60)

    if ok_count == 0:
        raise SystemExit("No se exporto nada. Sesion probablemente expirada.")


if __name__ == "__main__":
    main()
