# Banco Central de Chile — Indicadores en Power BI

Reporte Power BI (formato PBIP) con indicadores macroeconómicos de Chile publicados por el
Banco Central, más un script en Python que actualiza los datos y calcula los análisis.

## Páginas del reporte

| Página | Contenido |
|---|---|
| IPC Chile | Variación mensual 1928-2026, inflación anual compuesta, inflación 12 meses, estacionalidad |
| Desempleo Chile | Tasa 2010-2026, tendencia STL robusta con punto de inflexión, estacionalidad, variación interanual |
| Desempleo vs Imacec | Correlación por rezago y ley de Okun |
| Política Monetaria | TPM, expectativas de inflación (EEE), inflación 12 meses, tasa real y rezagos |
| Dólar y Traspaso | Dólar observado y traspaso al IPC (regresión de rezagos distribuidos) |

## Actualizar los datos

```bash
pip install -r requirements.txt
python actualizar_datos.py            # actualiza banco_central.xlsx y recalcula los análisis
python actualizar_datos.py --simular  # muestra los cambios sin escribir
```

Luego, en Power BI Desktop: abrir `banco_central.pbip` → **Inicio → Actualizar**.

### Fuentes

- **API BDE del Banco Central** (IPC, Imacec, TPM, expectativas, dólar): requiere la variable de
  entorno `BCCH_TOKEN` con el *Apikey Token* de "Mi cuenta" en
  https://si3.bcentral.cl/Siete/ES/Siete/API (vigencia de un año).
- **mindicador.cl** (tasa de desempleo, y respaldo del IPC): sin credenciales.

El script solo agrega meses nuevos o corrige valores revisados, y guarda una copia del Excel en
`respaldos/` antes de escribir.

## Notas

- Las consultas de Power BI leen `banco_central.xlsx` con una ruta absoluta; si clonas el
  proyecto en otra carpeta, ajusta la ruta en Power Query.
- Los análisis (Okun, rezagos, traspaso) muestran asociaciones estadísticas, no causalidad.
