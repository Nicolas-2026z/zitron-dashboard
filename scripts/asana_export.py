"""
ASANA BULK EXPORTER (version GitHub Actions)
Usa una sesion guardada (auth.json) en lugar de login manual.
Exporta cada proyecto en formato CSV/XLSX a la carpeta data/.
"""

import os
import time
from pathlib import Path
from playwright.sync_api import sync_playwright

# ── Carpeta de salida (dentro del repo) ──────────────────────────────────────
CARPETA_SALIDA = Path(__file__).resolve().parent.parent / "data"
CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

AUTH_FILE = Path(__file__).resolve().parent.parent / "auth.json"

# ── Lista de proyectos (nombre, URL) ─────────────────────────────────────────
PROYECTOS = [
    ("50001559 - FERROVIAL",                                    "https://app.asana.com/1/402967058777498/project/1215504785832997"),
    ("50001558 - GESVIAL OT4342-4344",                         "https://app.asana.com/1/402967058777498/project/1215137857526702"),
    ("50001557 - ATACAMA KOZAN OT4340",                        "https://app.asana.com/1/402967058777498/project/1215137857526520"),
    ("50001557 - ATACAMA KOZAN OT4335-4337-4339",              "https://app.asana.com/1/402967058777498/project/1215139673478155"),
    ("50001554 - MAPIMI",                                       "https://app.asana.com/1/402967058777498/project/1215118022011721"),
    ("50001553 - ZITRON COLOMBIA ARIS OT4326",                 "https://app.asana.com/1/402967058777498/project/1214969047721199"),
    ("50001553 - ZITRON COLOMBIA ARIS OT4327",                 "https://app.asana.com/1/402967058777498/project/1214969047720950"),
    ("50001553 - ZITRON COLOMBIA ARIS OT4324-4325",            "https://app.asana.com/1/402967058777498/project/1214969284200793"),
    ("50001555 - ZITRON COLOMBIA ARIS MINING SEGOVIA",         "https://app.asana.com/1/402967058777498/project/1214922603742599"),
    ("50001544 - DMC MINING SERVICES",                         "https://app.asana.com/1/402967058777498/project/1214892476594825"),
    ("50001547 - XEMORTIZ AMERICAS GOLD",                      "https://app.asana.com/1/402967058777498/project/1214787972061369"),
    ("50001546 - XEMORTIZ",                                    "https://app.asana.com/1/402967058777498/project/1214739697907723"),
    ("50001545 - XEMORTIZ MINERA FRISCO",                      "https://app.asana.com/1/402967058777498/project/1214563704105168"),
    ("50001543 - ZITRON COLOMBIA EGM",                         "https://app.asana.com/1/402967058777498/project/1214508463084754"),
    ("50001525 - ZITRON PERU METRO LIMA E07",                  "https://app.asana.com/1/402967058777498/project/1213832650589314"),
    ("50001534 - ZITRON PERU METRO LIMA E04-05-06-R4",         "https://app.asana.com/1/402967058777498/project/1213881596172396"),
    ("50001541 - XEMORTIZ MINERA FRISCO",                      "https://app.asana.com/1/402967058777498/project/1214137717389412"),
    ("50001532 - XEMORTIZ AMERICAS GOLD",                      "https://app.asana.com/1/402967058777498/project/1213963037596622"),
    ("50001466 - ZITRON PERU PODEROSA",                        "https://app.asana.com/1/402967058777498/project/1213997266064436"),
    ("50001485 - CALABRESSE METRO STO DOMINGO",                "https://app.asana.com/1/402967058777498/project/1213377149548665"),
    ("50001498 - CALABRESSE METRO STO DOMINGO POZOS",          "https://app.asana.com/1/402967058777498/project/1213377149548798"),
    ("50001477 - TRITON MINERA",                               "https://app.asana.com/1/402967058777498/project/1213377149548590"),
    ("50001490 - ZITRON PERU METRO LIMA PV04-PV05",            "https://app.asana.com/1/402967058777498/project/1213377149548731"),
    ("50001504 - ZITRON PERU CODESTABLE A",                    "https://app.asana.com/1/402967058777498/project/1213400421980747"),
    ("50001504 - ZITRON PERU CODESTABLE B",                    "https://app.asana.com/1/402967058777498/project/1213377149548924"),
    ("50001415 - ZITRON PERU METRO LIMA E05",                  "https://app.asana.com/1/402967058777498/project/1213391072478428"),
    ("50001506 - ZCO EPM FRIO AIRE",                           "https://app.asana.com/1/402967058777498/project/1213244147627519"),
    ("50001309 - ZITRON COLOMBIA BUGA BUENAVENTURA",           "https://app.asana.com/1/402967058777498/project/1213377149548388"),
    ("50001451 - ZITRON COLOMBIA",                             "https://app.asana.com/1/402967058777498/project/1213391072478496"),
    ("50001500 - EQ MIN LA HACIENDA",                          "https://app.asana.com/1/402967058777498/project/1213234368822465"),
    ("50001497 - XEMORTIZ IMMSA STA BBR",                      "https://app.asana.com/1/402967058777498/project/1213193352022841"),
    ("50001499 - EQ CHAPARRAL",                                "https://app.asana.com/1/402967058777498/project/1213244147627934"),
    ("50001508 - XEMORTIZ STOCK",                              "https://app.asana.com/1/402967058777498/project/1213362937195444"),
    ("50001473 - MAPIMI",                                       "https://app.asana.com/1/402967058777498/project/1213377149548525"),
    ("50001501 - EQ MIN LA HACIENDA B",                        "https://app.asana.com/1/402967058777498/project/1213377149548860"),
    ("50001446 - MINERA FRESNILLO",                            "https://app.asana.com/1/402967058777498/project/1213377149548988"),
    ("50001520 - XEMORTIZ DDG",                                "https://app.asana.com/1/402967058777498/project/1213396195711984"),
    ("50001518 - EQ CHAPARRAL B",                              "https://app.asana.com/1/402967058777498/project/1213458785834993"),
    ("50001533 - GESVIAL",                                     "https://app.asana.com/1/402967058777498/project/1213955236952392"),
    ("50001524 - ATACAMA KOZAN",                               "https://app.asana.com/1/402967058777498/project/1213458788103499"),
    ("50001514 - CLIENTE POR DEFINIR A",                       "https://app.asana.com/1/402967058777498/project/1213185599076077"),
    ("50001515 - CLIENTE POR DEFINIR B",                       "https://app.asana.com/1/402967058777498/project/1213234368821945"),
    ("50001563 - ZITRON COLOMBIA ANTIOQUIA",                   "https://app.asana.com/1/402967058777498/project/1215727332551578"),
]


def limpiar_nombre(nombre: str) -> str:
    for c in r'\/:*?"<>|':
        nombre = nombre.replace(c, "_")
    return nombre.strip()


def exportar_proyecto(page, nombre: str, url: str, indice: int, total: int) -> bool:
    print(f"\n[{indice}/{total}] {nombre}")

    try:
        page.goto(url, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_timeout(4000)

        # 1) Abrir menu "Acciones" del header del proyecto (icono chevron)
        menu = page.locator('[role="button"][aria-label="Acciones"]')
        menu.first.wait_for(state="visible", timeout=20000)
        menu.first.click(timeout=15000)
        time.sleep(0.8)

        # 2) Hover en "Exportar o sincronizar" para desplegar el submenu
        export_menu = page.get_by_text("Exportar o sincronizar", exact=False)
        export_menu.first.wait_for(state="visible", timeout=10000)
        export_menu.first.hover(timeout=10000)
        time.sleep(1.0)
        # algunos navegadores necesitan un segundo hover para abrir el submenu
        export_menu.first.hover(timeout=10000)
        time.sleep(0.8)

        # 3) Click en "Tareas del proyecto en formato CSV/XLSX" -> abre dialogo
        opcion_csv = page.get_by_text("Tareas del proyecto en formato CSV/XLSX", exact=False)
        opcion_csv.first.click(timeout=15000)
        time.sleep(0.8)

        # 4) En el dialogo, asegurar XLSX seleccionado y click en "Exportar"
        xlsx_radio = page.get_by_text("XLSX", exact=True)
        if xlsx_radio.count() > 0:
            xlsx_radio.first.click(timeout=5000)
            time.sleep(0.3)

        with page.expect_download(timeout=30000) as download_info:
            boton_exportar = page.get_by_role("button", name="Exportar", exact=True)
            boton_exportar.first.click(timeout=15000)

        download = download_info.value
        sufijo = Path(download.suggested_filename).suffix or ".csv"
        nombre_archivo = limpiar_nombre(nombre) + sufijo
        ruta_destino = CARPETA_SALIDA / nombre_archivo
        download.save_as(ruta_destino)
        print(f"     ✓ Guardado: {ruta_destino.name}")
        return True

    except Exception as e:
        print(f"     ✗ ERROR: {e}")
        # Guardar diagnostico solo para el primer fallo
        debug_dir = CARPETA_SALIDA.parent / "debug"
        debug_dir.mkdir(exist_ok=True)
        if not any(debug_dir.iterdir()):
            try:
                page.screenshot(path=str(debug_dir / "fallo.png"), full_page=True)
                debug_dir.joinpath("fallo.html").write_text(page.content(), encoding="utf-8")
                print("     -> Guardado screenshot/html de diagnostico")
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
    print("  ASANA BULK EXPORTER (modo CI)")
    print(f"  Carpeta de salida: {CARPETA_SALIDA}")
    print("=" * 60)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            accept_downloads=True,
            storage_state=str(AUTH_FILE),
        )
        page = context.new_page()

        exitosos = 0
        fallidos = []
        total = len(PROYECTOS)

        for i, (nombre, url) in enumerate(PROYECTOS, start=1):
            ok = exportar_proyecto(page, nombre, url, i, total)
            if ok:
                exitosos += 1
            else:
                fallidos.append(nombre)
            time.sleep(1.5)

        browser.close()

    print("\n" + "=" * 60)
    print(f"  COMPLETADO: {exitosos}/{total} proyectos exportados")
    if fallidos:
        print(f"\n  Fallidos ({len(fallidos)}):")
        for f in fallidos:
            print(f"    - {f}")
    print("=" * 60)

    if exitosos == 0:
        # Si nada funciono, probablemente la sesion expiro -> fallar el job
        raise SystemExit("Ningun proyecto se exporto. La sesion (auth.json) probablemente expiro.")


if __name__ == "__main__":
    main()
