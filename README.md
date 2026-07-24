# Town Hall 2026 — Ancora

Tablero Streamlit de **ingresos y venta nueva por área** para el Town Hall 2026.
Compara lo real de 2026 (año en curso) contra 2024, 2025 y el presupuesto anual.

## Contenido

| Página | Qué muestra |
|---|---|
| Panorama | Tarjetas por año (2024/2025/2026 con crecimiento), ingreso mensual comparado, mix por área y comparativo 2024 / 2025 / PTO 2026 / Real 2026 por área |
| Ingreso Semestral | Crecimiento por línea de negocio (1er semestre), barras por área, Real 2026 vs Presupuesto y detalle mensual |
| Venta Nueva | Acumulado vs años completos y presupuesto, por área |

## Correr en local

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Datos

`EXCEL/Town Hall Julio 2026.xlsx` — hojas `Ingresos` y `VN`. Cada hoja apila
bloques por año (2024, 2025, 2026 y PTO); el parser los localiza por el
encabezado `AREA`, no por filas fijas. El bloque de 2026 solo trae los meses
transcurridos (Ene–Jun).

Notas del parseo:

1. **El área `INTERNACIONAL` se excluye** de todo el tablero (`EXCLUIR_AREAS`).
2. Las comparaciones de crecimiento de 2026 se hacen contra el **mismo periodo**
   de 2025 (no contra el año completo), para que sean homogéneas.