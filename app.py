"""Panel web de control de obra construido con Streamlit."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st


ROOT = Path(__file__).resolve().parent
EXCEL_FILE = ROOT / "datos_sinteticos_obra.xlsx"
CHARTS_DIR = ROOT / "graficos"

st.set_page_config(page_title="Control de obra", page_icon="🏗️", layout="wide")


@st.cache_data
def cargar_datos() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Lee la fuente Excel entregada por el generador de reportes."""
    curva_s = pd.read_excel(EXCEL_FILE, sheet_name="Curva S")
    partidas = pd.read_excel(EXCEL_FILE, sheet_name="Avance por partidas")
    detalle = pd.read_excel(EXCEL_FILE, sheet_name="Detalle mensual")
    curva_s["Mes"] = pd.to_datetime(curva_s["Mes"])
    detalle["Mes"] = pd.to_datetime(detalle["Mes"])
    return curva_s, partidas, detalle


def moneda(valor: float) -> str:
    return f"S/ {valor:,.0f}"


def porcentaje(valor: float) -> str:
    return f"{valor:.1%}"


def aplicar_estilo() -> None:
    st.markdown(
        """
        <style>
            .stApp { background: #08070D; color: #F5F3FF; }
            [data-testid="stHeader"] { background: rgba(8, 7, 13, 0.92); }
            h1, h2, h3 { color: #E9D5FF !important; }
            [data-testid="stMetric"] {
                background: #12101A;
                border: 1px solid #5B21B6;
                border-radius: 12px;
                padding: 16px;
            }
            [data-testid="stMetricLabel"], [data-testid="stMetricValue"],
            [data-testid="stMetricDelta"] { color: #F5F3FF !important; }
            [data-baseweb="tab-list"] { gap: 12px; }
            [data-baseweb="tab"] { color: #C4B5FD; }
            [aria-selected="true"] { color: #E9D5FF !important; border-bottom-color: #A855F7 !important; }
            .stDownloadButton button { background: #7E22CE; color: white; border: 0; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    aplicar_estilo()

    if not EXCEL_FILE.exists():
        st.error("No se encontró el archivo Excel de datos. Ejecutá primero generar_reportes_obra.py.")
        st.stop()

    curva_s, partidas, detalle = cargar_datos()
    ultimo_mes = curva_s.iloc[-1]
    presupuesto_total = partidas["Presupuesto (S/)"].sum()
    avance_plan = ultimo_mes["Planificado acumulado (%)"]
    avance_real = ultimo_mes["Real acumulado (%)"]

    st.title("🏗️ Panel de control de obra")
    st.caption(f"Ejercicio sintético · Corte al 31 de diciembre de 2026 · Presupuesto total: {moneda(presupuesto_total)}")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Avance planificado", porcentaje(avance_plan))
    col2.metric("Avance real", porcentaje(avance_real), delta=f"{(avance_real - avance_plan):.1%}")
    col3.metric("Brecha", porcentaje(avance_plan - avance_real))
    col4.metric("Valor ganado real", moneda(ultimo_mes["Real acumulado (S/)"]))

    tab_curva, tab_partidas, tab_datos = st.tabs(["Curva S", "Partidas", "Datos fuente"])

    with tab_curva:
        st.subheader("Avance acumulado")
        curva_path = CHARTS_DIR / "curva_s.png"
        if curva_path.exists():
            st.image(curva_path, width="stretch")
        else:
            st.warning("No se encontró el gráfico. Ejecutá generar_reportes_obra.py para reconstruirlo.")
        tabla_curva = curva_s[["Mes", "Planificado acumulado (%)", "Real acumulado (%)"]].copy()
        tabla_curva["Mes"] = tabla_curva["Mes"].dt.strftime("%m/%Y")
        st.dataframe(
            tabla_curva.style.format(
                {"Planificado acumulado (%)": "{:.1%}", "Real acumulado (%)": "{:.1%}"}
            ),
            width="stretch",
            hide_index=True,
        )

    with tab_partidas:
        st.subheader("Avance físico por partidas")
        partidas_path = CHARTS_DIR / "avance_por_partidas.png"
        if partidas_path.exists():
            st.image(partidas_path, width="stretch")
        tabla_partidas = partidas[
            ["Partida", "Presupuesto (S/)", "Avance planificado (%)", "Avance real (%)"]
        ].sort_values("Avance real (%)", ascending=False)
        st.dataframe(
            tabla_partidas.style.format(
                {
                    "Presupuesto (S/)": "S/ {:,.0f}",
                    "Avance planificado (%)": "{:.1%}",
                    "Avance real (%)": "{:.1%}",
                }
            ),
            width="stretch",
            hide_index=True,
        )

    with tab_datos:
        st.subheader("Detalle mensual de una partida")
        partida = st.selectbox("Seleccioná una partida", sorted(detalle["Partida"].unique()))
        detalle_partida = detalle.loc[detalle["Partida"] == partida].copy()
        detalle_partida["Mes"] = detalle_partida["Mes"].dt.strftime("%m/%Y")
        st.dataframe(
            detalle_partida.style.format(
                {
                    "Presupuesto de partida (S/)": "S/ {:,.0f}",
                    "Avance planificado del mes (S/)": "S/ {:,.0f}",
                    "Avance real del mes (S/)": "S/ {:,.0f}",
                    "Avance planificado acumulado (%)": "{:.1%}",
                    "Avance real acumulado (%)": "{:.1%}",
                }
            ),
            width="stretch",
            hide_index=True,
        )

        with open(EXCEL_FILE, "rb") as archivo:
            st.download_button(
                "Descargar datos en Excel",
                data=archivo.read(),
                file_name=EXCEL_FILE.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()


