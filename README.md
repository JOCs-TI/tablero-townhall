# Tablero Ejecutivo Ancora — Consejo (2T 2026)

Tablero Streamlit del estado de resultados de Ancora al cierre del segundo
trimestre 2026 (acumulado Ene–Jun), con identidad de marca.

> Nota: este repositorio se llama `tablero-townhall` por su origen; hoy sirve el
> tablero del **Consejo** (se reemplazó el Town Hall reusando el mismo link).

## Vistas

| Página | Qué muestra |
|---|---|
| Resumen | P&L 2T (PTO / Real / 2025 con % de integración), indicadores clave y EBITDA |
| Por Línea de Negocio | Comparativo por LdN, combo Ingresos (barras) / Utilidad Operativa (línea) de los 3 periodos, mix, tabla resumen con rentabilidad, detalle mensual y P&L por línea |

## Correr en local

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Datos

`EXCEL/Presentacion Consejo Junio 2026.xlsx` — hojas `General` (P&L 2T),
`LdN` (7 bloques en dos columnas) e `Ingresos` (mensual por área). El archivo
vive fuera de Google Drive a propósito: dentro de `Mi unidad` las fórmulas
pierden su referencia externa y quedan en `#REF!`. La hoja `Utilidad x LdN`
viene rota (`#REF!`) y no se usa.

## Confidencial

Contiene el P&L completo (utilidades, EBITDA, márgenes por línea de negocio).