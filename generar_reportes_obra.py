"""Genera datos sintéticos de obra, un archivo Excel y gráficos de seguimiento.

El escenario representa una obra de edificación de S/ 10 millones durante 2026.
Los datos son sintéticos, pero mantienen una secuencia constructiva y retrasos
plausibles en las partidas finales.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parent
EXCEL_FILE = ROOT / "datos_sinteticos_obra.xlsx"
CHARTS_DIR = ROOT / "graficos"
MONTHS = pd.date_range("2026-01-01", periods=12, freq="MS")

# Paleta solicitada: morado sobre fondo negro.
COLORS = {
    "background": "#08070D",
    "panel": "#12101A",
    "grid": "#3B2B4D",
    "planned": "#C084FC",
    "actual": "#A855F7",
    "accent": "#7E22CE",
    "text": "#F5F3FF",
    "muted": "#C4B5FD",
}


@dataclass(frozen=True)
class Partida:
    nombre: str
    presupuesto: int
    inicio_plan: int
    fin_plan: int
    inicio_real: int
    fin_real: int
    avance_real_final: float


PARTIDAS = [
    Partida("Obras preliminares", 350_000, 1, 2, 1, 2, 1.00),
    Partida("Movimiento de tierras", 650_000, 1, 3, 1, 3, 1.00),
    Partida("Obras provisionales y contención", 400_000, 1, 4, 1, 4, 1.00),
    Partida("Cimentaciones", 950_000, 2, 5, 3, 6, 1.00),
    Partida("Estructura de concreto", 2_100_000, 4, 9, 5, 10, 0.96),
    Partida("Albañilería", 800_000, 6, 10, 7, 11, 0.92),
    Partida("Cobertura y techos", 450_000, 5, 7, 6, 8, 0.95),
    Partida("Instalaciones eléctricas", 630_000, 7, 11, 8, 12, 0.86),
    Partida("Instalaciones sanitarias", 520_000, 7, 11, 8, 12, 0.82),
    Partida("Climatización y ventilación", 320_000, 8, 11, 9, 12, 0.78),
    Partida("Acabados arquitectónicos", 2_000_000, 9, 12, 10, 12, 0.55),
    Partida("Obras exteriores", 830_000, 10, 12, 11, 12, 0.50),
]


def perfil_suave(inicio: int, fin: int, total: float) -> np.ndarray:
    """Distribuye un total mensual con mayor intensidad en la fase central."""
    valores = np.zeros(len(MONTHS))
    cantidad_meses = fin - inicio + 1
    posicion = np.arange(cantidad_meses)
    pesos = np.sin(np.pi * (posicion + 0.5) / cantidad_meses) ** 1.35
    pesos = pesos / pesos.sum()
    valores[inicio - 1 : fin] = total * pesos
    return valores


def construir_datos() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Construye resumen mensual, avance por partida y detalle mensual."""
    detalle: list[dict[str, object]] = []
    resumen_partidas: list[dict[str, object]] = []

    for partida in PARTIDAS:
        plan_mensual = perfil_suave(partida.inicio_plan, partida.fin_plan, partida.presupuesto)
        real_mensual = perfil_suave(
            partida.inicio_real,
            partida.fin_real,
            partida.presupuesto * partida.avance_real_final,
        )

        resumen_partidas.append(
            {
                "Partida": partida.nombre,
                "Presupuesto (S/)": partida.presupuesto,
                "Inicio planificado": MONTHS[partida.inicio_plan - 1],
                "Fin planificado": MONTHS[partida.fin_plan - 1],
                "Avance planificado (%)": 1.0,
                "Avance real (%)": partida.avance_real_final,
                "Valor ganado planificado (S/)": partida.presupuesto,
                "Valor ganado real (S/)": partida.presupuesto * partida.avance_real_final,
            }
        )

        for indice, mes in enumerate(MONTHS):
            detalle.append(
                {
                    "Mes": mes,
                    "Partida": partida.nombre,
                    "Presupuesto de partida (S/)": partida.presupuesto,
                    "Avance planificado del mes (S/)": plan_mensual[indice],
                    "Avance real del mes (S/)": real_mensual[indice],
                    "Avance planificado acumulado (%)": plan_mensual[: indice + 1].sum()
                    / partida.presupuesto,
                    "Avance real acumulado (%)": real_mensual[: indice + 1].sum()
                    / partida.presupuesto,
                }
            )

    detalle_df = pd.DataFrame(detalle)
    partidas_df = pd.DataFrame(resumen_partidas)
    total_presupuesto = partidas_df["Presupuesto (S/)"].sum()

    resumen_df = (
        detalle_df.groupby("Mes", as_index=False)[
            ["Avance planificado del mes (S/)", "Avance real del mes (S/)"]
        ]
        .sum()
        .rename(
            columns={
                "Avance planificado del mes (S/)": "Planificado mensual (S/)",
                "Avance real del mes (S/)": "Real mensual (S/)",
            }
        )
    )
    resumen_df["Planificado acumulado (S/)"] = resumen_df["Planificado mensual (S/)"].cumsum()
    resumen_df["Real acumulado (S/)"] = resumen_df["Real mensual (S/)"].cumsum()
    resumen_df["Planificado acumulado (%)"] = (
        resumen_df["Planificado acumulado (S/)"].div(total_presupuesto)
    )
    resumen_df["Real acumulado (%)"] = resumen_df["Real acumulado (S/)"].div(total_presupuesto)

    return resumen_df, partidas_df, detalle_df


def escribir_excel(
    resumen_df: pd.DataFrame, partidas_df: pd.DataFrame, detalle_df: pd.DataFrame
) -> None:
    """Guarda los datos fuente en un libro de Excel legible y reutilizable."""
    with pd.ExcelWriter(EXCEL_FILE, engine="xlsxwriter", datetime_format="mmm-yy") as writer:
        resumen_df.to_excel(writer, sheet_name="Curva S", index=False)
        partidas_df.to_excel(writer, sheet_name="Avance por partidas", index=False)
        detalle_df.to_excel(writer, sheet_name="Detalle mensual", index=False)

        libro = writer.book
        encabezado = libro.add_format(
            {"bold": True, "font_color": "#FFFFFF", "bg_color": "#5B21B6", "border": 0}
        )
        moneda = libro.add_format({"num_format": '"S/" #,##0', "font_color": "#22133D"})
        porcentaje = libro.add_format({"num_format": "0.0%", "font_color": "#22133D"})
        fecha = libro.add_format({"num_format": "mmm-yy", "font_color": "#22133D"})

        for nombre_hoja, dataframe in {
            "Curva S": resumen_df,
            "Avance por partidas": partidas_df,
            "Detalle mensual": detalle_df,
        }.items():
            hoja = writer.sheets[nombre_hoja]
            hoja.freeze_panes(1, 0)
            hoja.autofilter(0, 0, len(dataframe), len(dataframe.columns) - 1)
            for columna, encabezado_texto in enumerate(dataframe.columns):
                hoja.write(0, columna, encabezado_texto, encabezado)
                ancho = min(max(len(encabezado_texto) + 2, 14), 34)
                hoja.set_column(columna, columna, ancho)

        writer.sheets["Curva S"].set_column("A:A", 14, fecha)
        writer.sheets["Curva S"].set_column("B:E", 25, moneda)
        writer.sheets["Curva S"].set_column("F:G", 26, porcentaje)

        writer.sheets["Avance por partidas"].set_column("A:A", 34)
        writer.sheets["Avance por partidas"].set_column("B:B", 20, moneda)
        writer.sheets["Avance por partidas"].set_column("C:D", 18, fecha)
        writer.sheets["Avance por partidas"].set_column("E:F", 24, porcentaje)
        writer.sheets["Avance por partidas"].set_column("G:H", 28, moneda)

        writer.sheets["Detalle mensual"].set_column("A:A", 14, fecha)
        writer.sheets["Detalle mensual"].set_column("B:B", 34)
        writer.sheets["Detalle mensual"].set_column("C:E", 27, moneda)
        writer.sheets["Detalle mensual"].set_column("F:G", 31, porcentaje)


def configurar_ejes(ax: plt.Axes) -> None:
    ax.set_facecolor(COLORS["panel"])
    ax.tick_params(colors=COLORS["text"], labelsize=10)
    for borde in ax.spines.values():
        borde.set_color(COLORS["grid"])
    ax.grid(axis="y", color=COLORS["grid"], linewidth=0.8, alpha=0.65)
    ax.set_axisbelow(True)


def graficar_curva_s(resumen_df: pd.DataFrame) -> None:
    fig, ax = plt.subplots(figsize=(14, 7.5), facecolor=COLORS["background"])
    configurar_ejes(ax)

    ax.plot(
        resumen_df["Mes"],
        resumen_df["Planificado acumulado (%)"] * 100,
        color=COLORS["planned"],
        marker="o",
        linewidth=2.2,
        linestyle="--",
        label="Planificado",
    )
    ax.plot(
        resumen_df["Mes"],
        resumen_df["Real acumulado (%)"] * 100,
        color=COLORS["actual"],
        marker="o",
        linewidth=3.2,
        label="Real",
    )
    ax.fill_between(
        resumen_df["Mes"],
        resumen_df["Real acumulado (%)"] * 100,
        resumen_df["Planificado acumulado (%)"] * 100,
        color=COLORS["accent"],
        alpha=0.16,
        label="Brecha de avance",
    )

    ultimo = resumen_df.iloc[-1]
    ax.annotate(
        f"Real: {ultimo['Real acumulado (%)']:.1%}",
        xy=(ultimo["Mes"], ultimo["Real acumulado (%)"] * 100),
        xytext=(-95, -30),
        textcoords="offset points",
        color=COLORS["text"],
        arrowprops={"arrowstyle": "->", "color": COLORS["actual"]},
    )
    ax.set_title("Curva S — avance acumulado de obra", color=COLORS["text"], fontsize=18, fontweight="bold", pad=18)
    ax.set_ylabel("Avance acumulado (%)", color=COLORS["text"])
    ax.set_xlabel("Mes de control", color=COLORS["text"])
    ax.set_ylim(0, 106)
    ax.yaxis.set_major_formatter(lambda valor, _: f"{valor:.0f}%")
    leyenda = ax.legend(facecolor=COLORS["panel"], edgecolor=COLORS["grid"], labelcolor=COLORS["text"], loc="upper left")
    for texto in leyenda.get_texts():
        texto.set_color(COLORS["text"])

    fig.tight_layout()
    fig.savefig(CHARTS_DIR / "curva_s.png", dpi=180, facecolor=fig.get_facecolor())
    plt.close(fig)


def graficar_partidas(partidas_df: pd.DataFrame) -> None:
    datos = partidas_df.sort_values("Avance real (%)", ascending=True).copy()
    fig, ax = plt.subplots(figsize=(14, 9), facecolor=COLORS["background"])
    configurar_ejes(ax)

    ax.barh(
        datos["Partida"],
        datos["Avance planificado (%)"] * 100,
        color="none",
        edgecolor=COLORS["planned"],
        linewidth=1.5,
        label="Planificado",
    )
    barras = ax.barh(
        datos["Partida"],
        datos["Avance real (%)"] * 100,
        color=COLORS["actual"],
        alpha=0.88,
        label="Real",
    )
    for barra, valor in zip(barras, datos["Avance real (%)"] * 100, strict=True):
        ax.text(
            min(valor + 1.5, 102),
            barra.get_y() + barra.get_height() / 2,
            f"{valor:.0f}%",
            va="center",
            color=COLORS["text"],
            fontsize=9,
            fontweight="bold",
        )

    ax.set_title("Avance físico por partidas", color=COLORS["text"], fontsize=18, fontweight="bold", pad=18)
    ax.set_xlabel("Avance físico (%)", color=COLORS["text"])
    ax.set_xlim(0, 108)
    ax.xaxis.set_major_formatter(lambda valor, _: f"{valor:.0f}%")
    leyenda = ax.legend(facecolor=COLORS["panel"], edgecolor=COLORS["grid"], labelcolor=COLORS["text"], loc="lower right")
    for texto in leyenda.get_texts():
        texto.set_color(COLORS["text"])

    fig.tight_layout()
    fig.savefig(CHARTS_DIR / "avance_por_partidas.png", dpi=180, facecolor=fig.get_facecolor())
    plt.close(fig)


def main() -> None:
    CHARTS_DIR.mkdir(exist_ok=True)
    resumen_df, partidas_df, detalle_df = construir_datos()
    escribir_excel(resumen_df, partidas_df, detalle_df)

    # Se vuelve a leer el Excel: los gráficos usan el archivo generado como fuente.
    resumen_excel = pd.read_excel(EXCEL_FILE, sheet_name="Curva S")
    partidas_excel = pd.read_excel(EXCEL_FILE, sheet_name="Avance por partidas")
    resumen_excel["Mes"] = pd.to_datetime(resumen_excel["Mes"])

    graficar_curva_s(resumen_excel)
    graficar_partidas(partidas_excel)

    print(f"Excel creado: {EXCEL_FILE}")
    print(f"Gráficos creados en: {CHARTS_DIR}")
    print(f"Avance real acumulado: {resumen_excel.iloc[-1]['Real acumulado (%)']:.1%}")


if __name__ == "__main__":
    main()
