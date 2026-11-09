#!/usr/bin/env python3
"""Genera Burgos_Ofertas.xlsx con las ofertas de empleo de Burgos (datos abiertos de la Junta de CyL).

Sustituye al proceso manual del Excel con macros (Ofertas_Sepe_Auto.xlsm):
  1. baja las ofertas de Burgos
  2. quita las etiquetas HTML de la descripcion
  3. deja las columnas: Fecha, Origen, Puesto, Pr., Localidad, Enlace, Descripcion
"""
import html
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

BASE = "https://analisis.datosabiertos.jcyl.es/api/explore/v2.1/catalog/datasets/ofertas-de-empleo/exports/json"
SALIDA = "Burgos_Ofertas.xlsx"
DIAS = 75  # igual que el "Desde = HOY-75" del Excel original


def descargar():
    params = urllib.parse.urlencode(
        {"refine": 'provincia:"Burgos"', "order_by": "fecha_publicacion desc"}
    )
    req = urllib.request.Request(
        BASE + "?" + params, headers={"User-Agent": "ofertas-burgos-bot/1.0"}
    )
    ultimo_error = None
    for intento in range(4):
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.load(r)
        except Exception as e:  # noqa: BLE001
            ultimo_error = e
            time.sleep(5 * (intento + 1))
    raise SystemExit(f"No se pudo descargar: {ultimo_error}")


def limpiar(txt):
    """Quita las etiquetas HTML y deja texto normal."""
    if not txt:
        return ""
    txt = re.sub(
        r'<a [^>]*href="([^"]+)"[^>]*>(.*?)</a>', r"\2 (\1)", txt, flags=re.I | re.S
    )
    txt = re.sub(r"<br\s*/?>|</p>|</li>", " ", txt, flags=re.I)
    txt = re.sub(r"<[^>]+>", "", txt)
    txt = html.unescape(txt)
    return re.sub(r"\s+", " ", txt).strip()


def construir(registros, hoy=None):
    hoy = hoy or datetime.now(ZoneInfo("Europe/Madrid")).date()
    desde = hoy - timedelta(days=DIAS)
    filas = []
    for r in registros:
        f = r.get("fecha_publicacion")
        if not f:
            continue
        d = date.fromisoformat(f[:10])
        if d < desde:
            continue
        filas.append((d, r))
    filas.sort(key=lambda x: x[0], reverse=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Burgos Ofertas"
    ws["A1"] = "OFERTAS SEMANALES EMPLEO - Burgos Capital y Provincia"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Fecha act. datos:"
    ws["B2"] = hoy.strftime("%d-%m-%Y")
    ws["C2"] = "Desde"
    ws["D2"] = desde.strftime("%d-%m-%Y")
    ws["E2"] = "Hasta"
    ws["F2"] = hoy.strftime("%d-%m-%Y")
    ws["A3"] = "Fuente: ECyL y otros portales (datos abiertos Junta de Castilla y Leon)"

    cab = ["Fecha", "Origen", "Puesto", "Pr.", "Localidad", "Enlace", "Descripcion Requisitos"]
    for i, c in enumerate(cab, 1):
        celda = ws.cell(row=5, column=i, value=c)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F4E78")
    for n, (d, r) in enumerate(filas, start=6):
        enlace = r.get("enlace_al_contenido") or ""
        ws.cell(n, 1, d.strftime("%d-%m-%Y"))
        ws.cell(n, 2, r.get("fuente_contenido") or "")
        ws.cell(n, 3, r.get("titulo") or "")
        ws.cell(n, 4, r.get("localidad_asset2"))
        ws.cell(n, 5, r.get("localidad_asset0") or "-no informada")
        c = ws.cell(n, 6, enlace)
        if enlace:
            c.hyperlink = enlace
            c.font = Font(color="0563C1", underline="single")
        g = ws.cell(n, 7, limpiar(r.get("descripcion")))
        g.alignment = Alignment(wrap_text=True, vertical="top")

    for col, ancho in zip("ABCDEFG", [12, 18, 45, 6, 24, 40, 100]):
        ws.column_dimensions[col].width = ancho
    ws.freeze_panes = "A6"
    ws.auto_filter.ref = f"A5:G{5 + len(filas)}"
    return wb, len(filas)


def ya_esta_al_dia():
    """True si Burgos_Ofertas.xlsx ya se generó hoy (fecha en B2, hora de Madrid)."""
    try:
        from openpyxl import load_workbook

        hoy = datetime.now(ZoneInfo("Europe/Madrid")).strftime("%d-%m-%Y")
        return load_workbook(SALIDA).active["B2"].value == hoy
    except Exception:  # noqa: BLE001
        return False


def main():
    # En las ejecuciones programadas (hay varias cada madrugada por si GitHub
    # se salta alguna), si el Excel de hoy ya existe no se repite el trabajo.
    if os.environ.get("GITHUB_EVENT_NAME") == "schedule" and ya_esta_al_dia():
        print("El Excel de hoy ya estaba hecho: no se repite.")
        return
    if len(sys.argv) > 1:  # modo prueba: python generar_excel.py datos.json
        with open(sys.argv[1], encoding="utf-8") as f:
            registros = json.load(f)
    else:
        registros = descargar()
    if not registros:
        raise SystemExit("La Junta no devolvio ofertas: no se toca el archivo anterior.")
    wb, n = construir(registros)
    if n == 0:
        raise SystemExit("Ninguna oferta en el periodo: no se toca el archivo anterior.")
    wb.save(SALIDA)
    print(f"OK: {n} ofertas guardadas en {SALIDA}")


if __name__ == "__main__":
    main()
