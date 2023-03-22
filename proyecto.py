from datetime import datetime as dt, timedelta as td
from pathlib import Path
from tkinter import TRUE

import PySimpleGUI as sg
import pandas as pd

import sqlite3 as db

# Add some color to the window
sg.theme("LightBlue6")

EXCEL_FILE = Path.cwd() / "Cash Flow.xlsx"
conn = db.connect("cashflow.db")  # creates file
cur = conn.cursor()
TABLE_QUERY = """
    CREATE TABLE IF NOT EXISTS cashflows(
        fecha text,
        movimiento text check(movimiento in ('Ingreso', 'Egreso')),
        origen text,
        detalle text,
        importe real,
        medio text check(medio in ('Efectivo', 'Banco')),
        saldo real
    )"""

INGRESOS = [
    "Ingreso Mitre",
    "Ingreso Zona Sur",
    "Otros Ingresos",
]
EGRESOS = [
    "Egreso",
    "Rendición Anastacia",
]
cur.execute(TABLE_QUERY)

layout = [
    [sg.Text("Completar los siguientes datos:")],
    [
        sg.Text("Fecha", size=(16, 1)),
        sg.In(
            key="fecha",
            enable_events=True,
            visible=True,
            disabled=True,
            text_color="black",
            size=(12, 1),
        ),
        sg.CalendarButton("📅", target="fecha", format=("%d/%m/%Y"), key="cal_button"),
    ],
    [
        sg.Text("Movimiento", size=(16, 1)),
        sg.Combo(
            ["Ingreso", "Egreso"],
            key="movimiento",
            enable_events=True,
            size=(15, 1),
        ),
    ],
    [
        sg.Text("Origen", size=(16, 1)),
        sg.Combo(
            [],
            key="origen",
            size=(15, 1),
        ),
    ],
    [
        sg.Text("Detalle", size=(16, 1)),
        sg.InputText(
            key="detalle",
            size=(25, 1),
        ),
    ],
    [
        sg.Text("Importe", size=(16, 1)),
        sg.InputText(
            key="importe",
            size=(16, 1),
        ),
    ],
    [
        sg.Text("Medio de Pago/Cobro", size=(16, 1)),
        sg.Combo(
            ["Efectivo", "Banco", "MercadoPago"],
            key="medio",
            size=(15, 1),
        ),
    ],
    [sg.Submit("Guardar"), sg.Button("Borrar")],
    [
        sg.Text("Desde", size=(5, 1)),
        sg.In(
            key="fecha_desde",
            enable_events=True,
            visible=True,
            disabled=True,
            text_color="black",
            size=(12, 1),
        ),
        sg.CalendarButton(
            "📅", target="fecha_desde", format=("%d/%m/%Y"), key="cal_button_desde"
        ),
        sg.Text("Hasta", size=(5, 1)),
        sg.In(
            key="fecha_hasta",
            enable_events=True,
            visible=True,
            disabled=True,
            text_color="black",
            size=(12, 1),
        ),
        sg.CalendarButton(
            "📅", target="fecha_hasta", format=("%d/%m/%Y"), key="cal_button_hasta"
        ),
        sg.Button("Exportar"),
    ],
    [sg.Exit("Salir")],
]

window = sg.Window("Cash Flow IL SAPORE", layout)


def clear_input():
    for key in values:
        if "cal_button" not in key:
            window[key]("")
    return None


def limpiar_valores(df):
    df.importe = float(df.loc[0].importe.replace(",", "."))
    df.importe = (
        df.loc[0].importe
        if df.loc[0].movimiento == "Ingreso"
        else -1 * df.loc[0].importe
    )
    df.fecha = dt.strptime(df.loc[0].fecha, "%d/%m/%Y").strftime("%Y-%m-%d")
    df.drop(columns="cal_button", inplace=True)
    df.drop(columns="fecha_desde", inplace=True)
    df.drop(columns="fecha_hasta", inplace=True)
    df.drop(columns="cal_button_desde", inplace=True)
    df.drop(columns="cal_button_hasta", inplace=True)
    return df


def recalcular_saldo():
    df = pd.read_sql("select * from cashflows order by fecha asc", conn)
    df["saldo"] = df.loc[df["medio"] == "Efectivo", "importe"].cumsum()
    # df["saldo"] = df.importe.cumsum()
    save_to_database(df, "replace")


def save_to_database(df, method="append"):
    df.to_sql("cashflows", conn, if_exists=method, index=False)


def export_to_excel(values):
    if not values.get("fecha_desde") or not values.get("fecha_hasta"):
        fecha_hasta = dt.now()
        fecha_desde = fecha_hasta - td(days=30)
        values = {
            "fecha_desde": fecha_desde.strftime("%d/%m/%Y"),
            "fecha_hasta": fecha_hasta.strftime("%d/%m/%Y"),
        }
    fecha_desde = dt.strptime(values.get("fecha_desde"), "%d/%m/%Y")
    fecha_hasta = dt.strptime(values.get("fecha_hasta"), "%d/%m/%Y")
    if fecha_hasta <= fecha_desde:
        sg.popup("La fecha 'Hasta' debe ser mayor que 'Desde'!")
        return
    if (fecha_hasta - fecha_desde).days > 31:
        sg.popup("Solo puede exportar información de 30 días!")
        return
    query = f"""
        select
            STRFTIME('%d/%m/%Y', fecha) AS Dia, 
            movimiento AS Movimiento,
            origen AS Origen,
            detalle AS Detalle,
            importe AS Importe,
            medio AS Medio,
            saldo AS Saldo
        FROM
            cashflows 
        WHERE 
            fecha BETWEEN date({fecha_desde:'%Y-%m-%d'}) AND date({fecha_hasta:'%Y-%m-%d'})
        ORDER BY
            fecha ASC
    """
    df = pd.read_sql(query, conn)
    df.to_excel(EXCEL_FILE, index=False)
    sg.popup(f"Archivo exportado como {EXCEL_FILE}")


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Salir":
        conn.close()
        break
    if event == "Borrar":
        clear_input()
    if event == "Guardar":
        new_record = pd.DataFrame(values, index=[0])
        df = limpiar_valores(new_record)
        save_to_database(df)
        recalcular_saldo()
        sg.popup("Guardado!")
        clear_input()
    if event == "Exportar":
        export_to_excel(values)
    if event == "movimiento":
        if values.get("movimiento") == "Ingreso":
            window["origen"].update(values=INGRESOS)
        if values.get("movimiento") == "Egreso":
            window["origen"].update(values=EGRESOS)

window.close()
