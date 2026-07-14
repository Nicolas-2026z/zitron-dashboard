#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scripts/onedrive_export_ofertas.py
===================================
Reemplaza a asana_export_ofertas.py.
Descarga el Excel "Ofertas Pendientes" desde OneDrive / SharePoint
usando Microsoft Graph API y lo guarda en data_ofertas/
con el nombre "Oferta Zitron 2026.xlsx" que espera ofertas_dashboard.py.

SECRETOS GITHUB NECESARIOS
---------------------------
  MS_TENANT_ID      -> Directory (tenant) ID de Azure
  MS_CLIENT_ID      -> Application (client) ID del App Registration
  MS_CLIENT_SECRET  -> Client secret generado en Azure
  EXCEL_FILE_ID     -> 073a72d8-5c83-4521-9f05-c1b286ff409e
  EXCEL_OWNER       -> sergio.delafuente@zitron.com

CÓMO OBTENER LOS 3 PRIMEROS (una sola vez, ~10 min)
-----------------------------------------------------
1. Ir a https://portal.azure.com
2. "App registrations" -> New registration
   Nombre: Zitron Dashboard | Tipo: Single tenant | Registrar
3. Copiar "Application (client) ID" y "Directory (tenant) ID"
4. "Certificates & secrets" -> New client secret
   Copiar el VALUE antes de salir
5. "API permissions" -> Add permission -> Microsoft Graph
   -> Application permissions -> Files.Read.All + User.Read.All
   -> Grant admin consent (necesita ser admin del tenant)
"""

import os
import sys
from pathlib import Path
import msal
import requests

# ─────────────────────────────────────────────────────────────────
CARPETA_SALIDA = Path(__file__).resolve().parent.parent / "data_ofertas"
CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

TENANT_ID     = os.getenv("MS_TENANT_ID")
CLIENT_ID     = os.getenv("MS_CLIENT_ID")
CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
FILE_ID       = os.getenv("EXCEL_FILE_ID", "073a72d8-5c83-4521-9f05-c1b286ff409e")
OWNER_EMAIL   = os.getenv("EXCEL_OWNER",   "sergio.delafuente@zitron.com")

# Nombre de salida — debe coincidir con lo que busca ofertas_dashboard.py
OUTPUT_NAME = "Oferta Zitron 2026.xlsx"


def get_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise RuntimeError(
            f"Error de autenticación Microsoft: "
            f"{result.get('error_description', result)}"
        )
    return result["access_token"]


def download_file(token: str) -> Path:
    headers = {"Authorization": f"Bearer {token}"}

    # ── Intento 1: descarga directa por File ID ──────────────────
    url = (
        f"https://graph.microsoft.com/v1.0"
        f"/users/{OWNER_EMAIL}/drive/items/{FILE_ID}/content"
    )
    resp = requests.get(url, headers=headers, allow_redirects=True, timeout=30)

    if resp.status_code == 200:
        dest = CARPETA_SALIDA / OUTPUT_NAME
        dest.write_bytes(resp.content)
        print(f"  ✓ Descargado por File ID ({len(resp.content)//1024} KB) -> {dest.name}")
        return dest

    # ── Intento 2: buscar por nombre en el drive ──────────────────
    print(f"  File ID no disponible (HTTP {resp.status_code}), buscando por nombre...")
    search_url = (
        f"https://graph.microsoft.com/v1.0"
        f"/users/{OWNER_EMAIL}/drive/root/search(q='Ofertas Pendientes')"
    )
    search_resp = requests.get(search_url, headers=headers, timeout=30)

    if search_resp.status_code == 200:
        for item in search_resp.json().get("value", []):
            name = item.get("name", "")
            if "Ofertas Pendientes" in name or "Oferta" in name:
                dl_url = item.get("@microsoft.graph.downloadUrl")
                if dl_url:
                    content = requests.get(dl_url, timeout=30).content
                    dest = CARPETA_SALIDA / OUTPUT_NAME
                    dest.write_bytes(content)
                    print(f"  ✓ Encontrado por nombre '{name}' -> {dest.name}")
                    return dest

    # ── Intento 3: listar archivos recientes ──────────────────────
    print("  Buscando en archivos recientes del usuario...")
    recent_url = (
        f"https://graph.microsoft.com/v1.0"
        f"/users/{OWNER_EMAIL}/drive/recent"
    )
    recent_resp = requests.get(recent_url, headers=headers, timeout=30)
    if recent_resp.status_code == 200:
        for item in recent_resp.json().get("value", []):
            name = item.get("name", "")
            if "Ofertas" in name and name.endswith((".xlsx", ".xls")):
                dl_url = item.get("@microsoft.graph.downloadUrl")
                if dl_url:
                    content = requests.get(dl_url, timeout=30).content
                    dest = CARPETA_SALIDA / OUTPUT_NAME
                    dest.write_bytes(content)
                    print(f"  ✓ Encontrado en recientes '{name}' -> {dest.name}")
                    return dest

    raise RuntimeError(
        f"No se pudo descargar el archivo desde OneDrive.\n"
        f"  URL intentada: {url}\n"
        f"  HTTP status: {resp.status_code}\n"
        f"  Verifica: File ID, permisos Files.Read.All y admin consent en Azure."
    )


def main():
    print("=" * 60)
    print("  OneDrive Export — Ofertas Pendientes Zitron 2026")
    print(f"  Carpeta de salida: {CARPETA_SALIDA}")
    print("=" * 60)

    missing = [v for v in ("MS_TENANT_ID", "MS_CLIENT_ID", "MS_CLIENT_SECRET")
               if not os.getenv(v)]
    if missing:
        sys.exit(
            f"ERROR: Faltan secretos de GitHub: {', '.join(missing)}\n"
            "Configúralos en Settings -> Secrets -> Actions del repo."
        )

    print("  Autenticando con Microsoft Azure...")
    token = get_token()
    print("  ✓ Token obtenido")
    download_file(token)
    print("\n  ✓ Listo.")


if __name__ == "__main__":
    main()
