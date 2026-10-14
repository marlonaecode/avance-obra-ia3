# Curva S y avance de obra por partidas

Este ejercicio genera un escenario sintético de una obra de edificación valorizada en **S/ 10,000,000**. El archivo Excel incluye el presupuesto, la programación y el avance real de 12 partidas durante 2026. Los retrasos se concentran en las partidas finales, por lo que el avance real acumulado queda por debajo del planificado.

## Archivos generados

- `datos_sinteticos_obra.xlsx`: fuente de datos con las hojas `Curva S`, `Avance por partidas` y `Detalle mensual`.
- `graficos/curva_s.png`: curva S acumulada, comparando planificado contra real.
- `graficos/avance_por_partidas.png`: barras horizontales de avance físico por partida.

## Ejecución

El entorno virtual y las dependencias se instalan en la raíz del proyecto:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe .\generar_reportes_obra.py
```

## Cómo funciona el código

1. Define 12 partidas con presupuesto, calendario planificado, calendario real y porcentaje real de terminación.
2. Reparte el costo de cada partida con un perfil suave: menor al inicio y al cierre, mayor en la etapa productiva. La suma mensual preserva el presupuesto de cada partida.
3. Crea el Excel con los datos de control mensuales y de cada partida.
4. Lee nuevamente las hojas del Excel. Así los gráficos se construyen desde el archivo fuente, no desde los datos que permanecen en memoria.
5. Genera la curva S y el gráfico de avance por partidas con fondo negro y una paleta morada.

Para usar datos de una obra real, reemplazá la lista `PARTIDAS` o cargá las mismas columnas desde un Excel de control. Conservá la relación entre presupuesto, fechas y avance acumulado para que la curva S siga siendo consistente.

## Panel web con Streamlit

La aplicación `app.py` muestra los indicadores principales, los dos gráficos generados y las tablas del Excel. Mantiene el estilo de fondo negro y acentos morados.

```powershell
.\.venv\Scripts\python.exe -m streamlit run .\app.py
```

Luego abrí la dirección local que Streamlit indique, normalmente `http://localhost:8501`.
