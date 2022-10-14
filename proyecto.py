from pathlib import Path
from tkinter import TRUE

import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme("LightBlue6")

EXCEL_FILE = Path.cwd() / "Cash Flow.xlsx"
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text("Completar los siguientes datos:")],
    [
        sg.Text("Fecha", size=(16, 1)),
        sg.In(
            key="Fecha",
            enable_events=True,
            visible=True,
            disabled=True,
            text_color="black",
            size=(12, 1),
        ),
        sg.CalendarButton("📅", target="Fecha", format=("%d/%m/%Y"), key="cal_button"),
    ],
    [
        sg.Text("Movimiento", size=(16, 1)),
        sg.Combo(
            ["Ingreso", "Egreso"],
            key="Movimiento",
            size=(15, 1),
        ),
    ],
    [
        sg.Text("Origen", size=(16, 1)),
        sg.Combo(
            ["Ingreso Mitre", "Ingreso Zona Sur", "Otros Ingresos", "Egreso"],
            key="Origen",
            size=(15, 1),
        ),
    ],
    [
        sg.Text("Detalle", size=(16, 1)),
        sg.InputText(
            key="Detalle",
            size=(25, 1),
        ),
    ],
    [
        sg.Text("Importe", size=(16, 1)),
        sg.InputText(
            key="Importe",
            size=(16, 1),
        ),
    ],
    [
        sg.Text("Medio de Pago/Cobro", size=(16, 1)),
        sg.Combo(
            ["Efectivo", "BBVA"],
            key="Medio de Pago/Cobro",
            size=(15, 1),
        ),
    ],
    [sg.Submit("Guardar"), sg.Button("Borrar"), sg.Exit("Salir")],
]

window = sg.Window("Cash Flow IL SAPORE", layout)


def clear_input():
    for key in values:
        if key != "cal_button":
            window[key]("")
    return None


def limpiar_valores(df):
    # df.Fecha = pd.to_datetime(df['Fecha'], format="%d/%m/%y")
    df.Importe = df.Importe.apply(lambda x: float(x.replace(",", ".")))
    df.drop(columns="cal_button", inplace=True)
    return df


def saldo(row):
    ingreso_acum = df.loc[(df.Movimiento == "Ingreso") & (df.Fecha <= row.Fecha)][
        "Importe"
    ].sum()
    egreso_acum = df.loc[(df.Movimiento == "Egreso") & (df.Fecha <= row.Fecha)][
        "Importe"
    ].sum()
    return ingreso_acum - egreso_acum


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Salir":
        break
    if event == "Borrar":
        clear_input()
    if event == "Guardar":
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([limpiar_valores(new_record), df], ignore_index=True)
        df["Saldo"] = df.apply(saldo, axis=1)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup("Guardado!")
        clear_input()
window.close()
