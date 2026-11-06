# Town Hall 2026 — Ancora

Tablero Streamlit de ingresos y venta nueva por área, con identidad de marca.

> Nota: este repositorio se llama `tablero-townhall` por su origen. Entre
> agosto y noviembre de 2026 sirvió temporalmente el tablero del **Consejo**
> reusando el mismo link; volvió a servir Town Hall el 2026-11-06.

## Vistas

| Página | Qué muestra |
|---|---|
| Panorama | KPIs globales, comparativo 2024/2025/PTO/Real y proyección de cierre de año |
| Ingreso Semestral | Ingreso mensual por área |
| Venta Nueva | Venta nueva por área |

Excluye las áreas `INTERNACIONAL` y `ANI` (`EXCLUIR_AREAS`).

## Correr en local

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Datos

Google Sheet en vivo (`TOWNHALL_SHEET_ID`), descargado con
`export?format=xlsx` sin API key (mismo patrón que Facturación/Tarjetas de
Crédito). Se cachea 5 minutos (`@st.cache_data(ttl=300)`). Hojas `Ingresos`
y `VN`, cada una con 4 bloques apilados por año (2024, 2025, 2026 Real,
PTO 2026) — ver detalle en memoria del proyecto
(`town-hall-bloques-por-año`).

## Confidencial

Contiene ingreso y venta nueva por área. Uso interno.
