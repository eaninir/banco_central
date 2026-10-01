"""Actualiza banco_central.xlsx (hojas ipc y desempleo) con los últimos datos publicados.

Fuente: mindicador.cl, que republica las series del Banco Central de Chile sin requerir
credenciales. Opcionalmente, para el IPC se puede usar la API oficial del Banco Central
(BDE) definiendo la variable de entorno BCCH_TOKEN (Apikey Token de "Mi cuenta"), o bien
BCCH_USER y BCCH_PASS.

Solo agrega meses nuevos o corrige valores revisados; nunca borra datos existentes.
Antes de escribir guarda una copia en respaldos/. Además recalcula la hoja desempleo_stl
(descomposición STL robusta + punto de inflexión de la tendencia; requiere statsmodels) y,
con BCCH_TOKEN, las hojas imacec, desempleo_imacec, rezagos y okun (relación con el Imacec).

Uso:
    python actualizar_datos.py            # actualiza el Excel
    python actualizar_datos.py --simular  # muestra los cambios sin escribir
"""

import json
import os
import shutil
import sys
import urllib.parse
import urllib.request
from datetime import date, datetime

from openpyxl import load_workbook

CARPETA = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(CARPETA, "banco_central.xlsx")
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

# Código de la serie IPC (variación mensual) en la BDE del Banco Central. Solo se usa
# si hay credenciales; verificar en https://si3.bcentral.cl/siete si la API lo rechaza.
BCCH_SERIE_IPC = "F074.IPC.VAR.Z.Z.C.M"
# Imacec empalmado, serie original (índice 2018=100)
BCCH_SERIE_IMACEC = "F032.IMC.IND.Z.Z.EP18.Z.Z.0.M"
# Tasa de política monetaria (promedio mensual) y expectativas de inflación a 12 meses (EEE, mediana)
BCCH_SERIE_TPM = "F022.TPM.TIN.D001.NO.Z.M"
BCCH_SERIE_EXPECTATIVAS = "F089.IPC.V12.14.M"
# Dólar observado, promedio mensual ($ por dólar)
BCCH_SERIE_DOLAR = "F073.TCO.PRE.HIST.M"


def obtener_json(url):
    req = urllib.request.Request(url, headers={"User-Agent": "banco_central-pbi/1.0"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        crudo = resp.read()
    try:
        return json.loads(crudo.decode("utf-8"))
    except UnicodeDecodeError:  # la API del Banco Central a veces responde en latin-1
        return json.loads(crudo.decode("latin-1"))


def serie_mindicador(codigo, anios):
    """Devuelve {(año, mes): valor} para los años pedidos."""
    datos = {}
    for anio in anios:
        for punto in obtener_json(f"https://mindicador.cl/api/{codigo}/{anio}")["serie"]:
            # La fecha viene en UTC (ej. 2026-07-01T04:00:00Z = 1 de julio en Chile)
            f = datetime.fromisoformat(punto["fecha"].replace("Z", "+00:00"))
            datos[(f.year, f.month)] = round(float(punto["valor"]), 2)
    return datos


def variable_entorno(nombre):
    """Lee la variable de la sesión; si no está, de la configuración de Windows (usuario o sistema)."""
    valor = os.environ.get(nombre)
    if valor or os.name != "nt":
        return valor
    import winreg
    ubicaciones = [
        (winreg.HKEY_CURRENT_USER, r"Environment"),
        (winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"),
    ]
    for raiz, ruta in ubicaciones:
        try:
            with winreg.OpenKey(raiz, ruta) as clave:
                return winreg.QueryValueEx(clave, nombre)[0]
        except OSError:
            continue
    return None


def serie_bcch(codigo, desde):
    # Preferir el Apikey Token (Mi cuenta en la BDE, vigencia 1 año); si no, correo + contraseña
    token = variable_entorno("BCCH_TOKEN")
    usuario, clave = variable_entorno("BCCH_USER"), variable_entorno("BCCH_PASS")
    if token:
        credenciales = {"token": token}
    elif usuario and clave:
        credenciales = {"user": usuario, "pass": clave}
    else:
        return {}
    params = urllib.parse.urlencode({
        **credenciales, "function": "GetSeries", "timeseries": codigo,
        "firstdate": f"{desde}-01-01", "lastdate": date.today().isoformat(),
    })
    resp = obtener_json(f"https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx?{params}")
    if resp.get("Codigo") != 0:
        print(f"  ! API Banco Central respondió: {resp.get('Descripcion')}")
        return {}
    datos = {}
    for obs in resp["Series"]["Obs"]:
        if obs.get("statusCode") == "OK":
            d, m, a = (int(x) for x in obs["indexDateString"].split("-"))
            datos[(a, m)] = round(float(obs["value"]), 2)
    return datos


def leer_hoja(ws):
    """Lee la hoja en formato ancho (Año × meses) como {(año, mes): valor} y {año: fila}."""
    encabezado = [c.value for c in ws[1]]
    assert encabezado[0] == "Año" and encabezado[1:13] == MESES, f"Encabezado inesperado en {ws.title}"
    valores, filas = {}, {}
    for fila in ws.iter_rows(min_row=2):
        anio = fila[0].value
        if anio is None:
            continue
        filas[int(anio)] = fila[0].row
        for m in range(1, 13):
            v = fila[m].value
            if isinstance(v, (int, float)):
                valores[(int(anio), m)] = v
    return valores, filas


def actualizar_hoja(ws, nuevos, simular):
    actuales, filas = leer_hoja(ws)
    cambios = []
    for (anio, mes), valor in sorted(nuevos.items()):
        anterior = actuales.get((anio, mes))
        if anterior is not None and abs(anterior - valor) < 0.005:
            continue
        cambios.append((anio, mes, anterior, valor))
        if simular:
            continue
        if anio not in filas:
            fila = ws.max_row + 1
            ws.cell(row=fila, column=1, value=anio)
            filas[anio] = fila
        ws.cell(row=filas[anio], column=mes + 1, value=valor)
    for anio, mes, anterior, valor in cambios:
        tipo = "nuevo" if anterior is None else f"revisado (antes {anterior})"
        print(f"  {MESES[mes - 1]} {anio}: {valor} — {tipo}")
    if not cambios:
        ultimo = max(actuales)
        print(f"  Sin cambios (último dato: {MESES[ultimo[1] - 1]} {ultimo[0]})")
    return len(cambios)


def punto_inflexion(tendencia, excluir=("2020-03-01", "2021-12-01"), minimo=24):
    """Mes donde cambia la pendiente de la tendencia.

    Ajusta una recta continua con un solo quiebre (y = a + b·t + c·max(0, t - k)) y elige
    el k que minimiza el error, excluyendo la pandemia y dejando al menos `minimo` meses
    a cada lado. Devuelve (fecha, pendiente_antes, pendiente_despues) en pp por año.
    """
    import numpy as np

    t = np.arange(len(tendencia))
    y = tendencia.to_numpy()
    usar = ~((tendencia.index >= excluir[0]) & (tendencia.index <= excluir[1]))
    mejor = None
    for k in range(minimo, len(t) - minimo):
        X = np.column_stack([np.ones(len(t)), t, np.maximum(0, t - k)])[usar]
        coef = np.linalg.lstsq(X, y[usar], rcond=None)[0]
        error = ((X @ coef - y[usar]) ** 2).sum()
        if mejor is None or error < mejor[0]:
            mejor = (error, k, coef)
    _, k, coef = mejor
    return tendencia.index[k], coef[1] * 12, (coef[1] + coef[2]) * 12


def descomponer_desempleo(wb):
    """Escribe la hoja desempleo_stl: tendencia, estacionalidad y residuo (STL robusta)."""
    import pandas as pd
    from statsmodels.tsa.seasonal import STL

    datos, _ = leer_hoja(wb["desempleo"])
    serie = pd.Series({pd.Timestamp(a, m, 1): v for (a, m), v in datos.items()}).sort_index()
    serie = serie.asfreq("MS").interpolate()  # la STL no admite meses faltantes
    # robust=True: la pandemia queda en el residuo y no contamina la tendencia
    res = STL(serie, period=12, robust=True).fit()
    quiebre, antes, despues = punto_inflexion(res.trend)

    if "desempleo_stl" in wb.sheetnames:
        del wb["desempleo_stl"]
    ws = wb.create_sheet("desempleo_stl")
    ws.append(["Periodo", "Tasa", "Tendencia", "Estacional", "Residuo",
               "Desestacionalizada", "Quiebre"])
    for fecha in serie.index:
        ws.append([fecha.date(), round(serie[fecha], 2), round(res.trend[fecha], 3),
                   round(res.seasonal[fecha], 3), round(res.resid[fecha], 3),
                   round(serie[fecha] - res.seasonal[fecha], 3), 1 if fecha == quiebre else 0])
    for celda in ws["A"][1:]:
        celda.number_format = "yyyy-mm-dd"
    print(f"  {len(serie)} meses. Punto de inflexión de la tendencia: "
          f"{MESES[quiebre.month - 1]} {quiebre.year} "
          f"({antes:+.2f} pp/año antes, {despues:+.2f} pp/año después)")


def serie_bcch_completa(codigo, desde="1996-01-01"):
    """Serie mensual completa desde la API del Banco Central como pandas.Series (requiere BCCH_TOKEN)."""
    import pandas as pd

    params = urllib.parse.urlencode({
        "token": variable_entorno("BCCH_TOKEN"), "function": "GetSeries", "timeseries": codigo,
        "firstdate": desde, "lastdate": date.today().isoformat(),
    })
    resp = obtener_json(f"https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx?{params}")
    if resp.get("Codigo") != 0:
        raise RuntimeError(resp.get("Descripcion"))
    return pd.Series({
        pd.to_datetime(o["indexDateString"], format="%d-%m-%Y"): float(o["value"])
        for o in resp["Series"]["Obs"] if o.get("statusCode") == "OK"
    }).sort_index()


def escribir_hoja(wb, nombre, encabezado, filas):
    if nombre in wb.sheetnames:
        del wb[nombre]
    ws = wb.create_sheet(nombre)
    ws.append(encabezado)
    for fila in filas:
        ws.append(fila)
    if encabezado[0] == "Periodo":
        for celda in ws["A"][1:]:
            celda.number_format = "yyyy-mm-dd"


def analizar_imacec(wb, max_rezago=24, pandemia=("2020-03-01", "2021-12-01")):
    """Relación desempleo–Imacec: correlación por rezago y ley de Okun.

    - Variación del desempleo: pp respecto del mismo mes del año anterior.
    - Crecimiento del Imacec: % anual del promedio móvil de 3 meses (la tasa del INE es
      un trimestre móvil, así ambas series miden el mismo período).
    - Rezago óptimo: el que da la correlación más negativa (el empleo reacciona después).
    - Okun: variación del desempleo = intercepto + pendiente × crecimiento rezagado.
    """
    import numpy as np
    import pandas as pd

    if not variable_entorno("BCCH_TOKEN"):
        print("  Omitido: falta la variable BCCH_TOKEN (Imacec viene de la API del Banco Central)")
        return

    imacec = serie_bcch_completa(BCCH_SERIE_IMACEC)
    datos, _ = leer_hoja(wb["desempleo"])
    tasa = pd.Series({pd.Timestamp(a, m, 1): v for (a, m), v in datos.items()}).sort_index()

    var_desempleo = (tasa - tasa.shift(12)).dropna()
    trimestre = imacec.rolling(3).mean()
    crecimiento = ((trimestre / trimestre.shift(12) - 1) * 100).dropna()

    correlaciones = []
    for k in range(max_rezago + 1):
        par = pd.concat([var_desempleo, crecimiento.shift(k)], axis=1, sort=True).dropna()
        correlaciones.append((k, par.corr().iloc[0, 1]))
    rezago, correlacion = min(correlaciones, key=lambda x: x[1])

    par = pd.concat([var_desempleo.rename("du"), crecimiento.rename("g"),
                     crecimiento.shift(rezago).rename("g_rez")], axis=1, sort=True).dropna()
    pendiente, intercepto = np.polyfit(par.g_rez, par.du, 1)
    r2 = np.corrcoef(par.g_rez, par.du)[0, 1] ** 2
    en_pandemia = (par.index >= pandemia[0]) & (par.index <= pandemia[1])
    sin = par[~en_pandemia]
    pendiente_sp, intercepto_sp = np.polyfit(sin.g_rez, sin.du, 1)
    r2_sp = np.corrcoef(sin.g_rez, sin.du)[0, 1] ** 2

    escribir_hoja(wb, "imacec", ["Periodo", "Imacec", "CrecimientoAnual"], [
        [f.date(), round(v, 2), round(crecimiento[f], 3) if f in crecimiento.index else None]
        for f, v in imacec.items()])
    escribir_hoja(wb, "desempleo_imacec",
                  ["Periodo", "VarDesempleo", "CrecImacec", "CrecImacecRezagado", "VarDesempleoPredicha", "Pandemia"],
                  [[f.date(), round(r.du, 3), round(r.g, 3), round(r.g_rez, 3),
                    round(intercepto + pendiente * r.g_rez, 3), int(p)]
                   for (f, r), p in zip(par.iterrows(), en_pandemia)])
    escribir_hoja(wb, "rezagos", ["Rezago", "Correlacion", "Optimo"],
                  [[k, round(c, 4), int(k == rezago)] for k, c in correlaciones])
    escribir_hoja(wb, "okun",
                  ["Rezago", "Correlacion", "Pendiente", "Intercepto", "R2", "CrecimientoNeutral",
                   "PendienteSinPandemia", "R2SinPandemia", "Meses"],
                  [[rezago, round(correlacion, 4), round(pendiente, 4), round(intercepto, 4), round(r2, 4),
                    round(-intercepto / pendiente, 3), round(pendiente_sp, 4), round(r2_sp, 4), len(par)]])
    print(f"  Imacec hasta {MESES[imacec.index[-1].month - 1]} {imacec.index[-1].year}. "
          f"Rezago óptimo: {rezago} mes(es), correlación {correlacion:+.2f}")
    print(f"  Okun: {pendiente:+.3f} pp de desempleo por cada punto de crecimiento (R² {r2:.2f}); "
          f"sin pandemia {pendiente_sp:+.3f} (R² {r2_sp:.2f}). Crecimiento neutral: {-intercepto / pendiente:.1f}%")


def analizar_politica_monetaria(wb, desde="2001-01-01", max_rezago=36):
    """TPM, expectativas de inflación e inflación efectiva a 12 meses.

    - Tasa real ex ante = TPM − expectativas de inflación a 12 meses.
    - Reacción: correlación entre la inflación de hoy y la TPM k meses después
      (cuánto tarda el Banco Central en responder).
    - Efecto: correlación entre el cambio de la TPM en 12 meses y el cambio de la
      inflación en los k meses siguientes (negativa = las alzas bajan la inflación).
    """
    import numpy as np
    import pandas as pd

    if not variable_entorno("BCCH_TOKEN"):
        print("  Omitido: falta la variable BCCH_TOKEN")
        return

    tpm = serie_bcch_completa(BCCH_SERIE_TPM, desde)
    expectativas = serie_bcch_completa(BCCH_SERIE_EXPECTATIVAS, desde)
    datos, _ = leer_hoja(wb["ipc"])
    variacion = pd.Series({pd.Timestamp(a, m, 1): v for (a, m), v in datos.items()}).sort_index()
    inflacion = ((1 + variacion / 100).rolling(12).apply(np.prod, raw=True) - 1) * 100

    d = pd.concat([tpm.rename("tpm"), expectativas.rename("exp"), inflacion.rename("pi")],
                  axis=1, sort=True)
    d = d[d.index >= desde].dropna(how="all")
    d["real"] = d.tpm - d.exp

    completos = d.dropna(subset=["tpm", "pi"])
    cambio_tpm = completos.tpm - completos.tpm.shift(12)
    filas_rezago = []
    for k in range(max_rezago + 1):
        reaccion = completos.pi.corr(completos.tpm.shift(-k))
        efecto = cambio_tpm.corr(completos.pi.shift(-k) - completos.pi) if k else None
        filas_rezago.append([k, round(reaccion, 4), None if efecto is None or pd.isna(efecto) else round(efecto, 4)])
    k_reaccion, r_reaccion = max(((f[0], f[1]) for f in filas_rezago if f[0] <= 24), key=lambda x: x[1])
    k_efecto, r_efecto = min(((f[0], f[2]) for f in filas_rezago if f[2] is not None), key=lambda x: x[1])

    def r(x, n=3):
        return None if pd.isna(x) else round(float(x), n)

    escribir_hoja(wb, "politica_monetaria",
                  ["Periodo", "Año", "TPM", "Expectativas", "Inflacion12m", "TasaReal", "Meta"],
                  [[f.date(), f.year, r(x.tpm, 2), r(x.exp, 2), r(x.pi), r(x.real), 3]
                   for f, x in d.iterrows()])
    escribir_hoja(wb, "pm_rezagos", ["Rezago", "Reaccion", "Efecto"], filas_rezago)
    escribir_hoja(wb, "pm_resumen", ["RezagoReaccion", "CorrReaccion", "RezagoEfecto", "CorrEfecto"],
                  [[k_reaccion, round(r_reaccion, 4), k_efecto, round(r_efecto, 4)]])
    ultimo = d.dropna(subset=["tpm"]).index[-1]
    print(f"  TPM {d.tpm[ultimo]:.2f}% ({MESES[ultimo.month - 1]} {ultimo.year}), "
          f"tasa real {d.real.dropna().iloc[-1]:+.2f}%. "
          f"Reacción: {k_reaccion} meses (r {r_reaccion:+.2f}); efecto máximo: {k_efecto} meses (r {r_efecto:+.2f})")


def traspaso_rezagos_distribuidos(variacion_ipc, variacion_dolar, desde, rezagos=12):
    """Regresión: IPC mensual = a + Σ b_j · dólar mensual(t−j) + dummies de mes.

    Devuelve (meses, b_j, acumulado_h, error_estandar_h) con el traspaso acumulado a cada
    horizonte h = b_0 + … + b_h (pp de IPC por 1% de alza del dólar).
    """
    import numpy as np
    import pandas as pd

    df = pd.DataFrame({"pi": variacion_ipc})
    for j in range(rezagos + 1):
        df[f"e{j}"] = variacion_dolar.shift(j)
    for mes in range(2, 13):
        df[f"m{mes}"] = (df.index.month == mes).astype(float)
    df = df[df.index >= desde].dropna()
    X = np.column_stack([np.ones(len(df))] + [df[c].to_numpy() for c in df.columns if c != "pi"])
    y = df.pi.to_numpy()
    coef = np.linalg.lstsq(X, y, rcond=None)[0]
    resid = y - X @ coef
    cov = resid @ resid / (len(y) - X.shape[1]) * np.linalg.inv(X.T @ X)
    b = coef[1:rezagos + 2]
    cov_b = cov[1:rezagos + 2, 1:rezagos + 2]
    acumulado = np.cumsum(b)
    error = np.array([np.sqrt(np.ones(h + 1) @ cov_b[:h + 1, :h + 1] @ np.ones(h + 1))
                      for h in range(rezagos + 1)])
    return len(df), b, acumulado, error


def analizar_dolar(wb, desde="2010-01-01", desde_comparacion="2001-01-01"):
    """Dólar observado y su traspaso al IPC (efecto de un alza de 10% del dólar, en pp)."""
    import numpy as np
    import pandas as pd

    if not variable_entorno("BCCH_TOKEN"):
        print("  Omitido: falta la variable BCCH_TOKEN")
        return

    dolar = serie_bcch_completa(BCCH_SERIE_DOLAR, "1999-01-01")
    datos, _ = leer_hoja(wb["ipc"])
    variacion_ipc = pd.Series({pd.Timestamp(a, m, 1): v for (a, m), v in datos.items()}).sort_index()
    inflacion = ((1 + variacion_ipc / 100).rolling(12).apply(np.prod, raw=True) - 1) * 100
    variacion_dolar = np.log(dolar).diff() * 100  # % mensual
    var12_dolar = (dolar / dolar.shift(12) - 1) * 100

    meses, b, acumulado, error = traspaso_rezagos_distribuidos(variacion_ipc, variacion_dolar, desde)
    _, _, acumulado_comp, _ = traspaso_rezagos_distribuidos(variacion_ipc, variacion_dolar, desde_comparacion)

    def r(x, n=3):
        return None if pd.isna(x) else round(float(x), n)

    periodo = dolar[dolar.index >= "2001-01-01"].index
    escribir_hoja(wb, "dolar", ["Periodo", "Año", "Dolar", "VarDolar12m", "Inflacion12m"],
                  [[f.date(), f.year, r(dolar[f], 2), r(var12_dolar.get(f)), r(inflacion.get(f))]
                   for f in periodo])
    escribir_hoja(wb, "traspaso", ["Horizonte", "EfectoMes", "Acumulado", "LimInf", "LimSup"],
                  [[h, r(10 * b[h], 4), r(10 * acumulado[h], 4),
                    r(10 * (acumulado[h] - 1.96 * error[h]), 4), r(10 * (acumulado[h] + 1.96 * error[h]), 4)]
                   for h in range(len(b))])
    escribir_hoja(wb, "traspaso_resumen",
                  ["Traspaso3m", "Traspaso12m", "LimInf12m", "LimSup12m", "Traspaso12mDesde2001", "Meses", "Desde"],
                  [[r(10 * acumulado[3], 3), r(10 * acumulado[-1], 3), r(10 * (acumulado[-1] - 1.96 * error[-1]), 3),
                    r(10 * (acumulado[-1] + 1.96 * error[-1]), 3), r(10 * acumulado_comp[-1], 3), meses,
                    int(desde[:4])]])
    print(f"  Dólar {dolar.iloc[-1]:.2f} ({MESES[dolar.index[-1].month - 1]} {dolar.index[-1].year}), "
          f"{var12_dolar.iloc[-1]:+.1f}% en 12 meses. Alza de 10% → IPC +{10 * acumulado[3]:.2f} pp a 3 meses, "
          f"+{10 * acumulado[-1]:.2f} pp a 12 meses (IC95% {10 * (acumulado[-1] - 1.96 * error[-1]):.2f} a "
          f"{10 * (acumulado[-1] + 1.96 * error[-1]):.2f}; desde 2001: {10 * acumulado_comp[-1]:+.2f})")


def main():
    sys.stdout.reconfigure(encoding="utf-8")  # acentos en la consola de Windows
    simular = "--simular" in sys.argv
    hoy = date.today()
    anios = [hoy.year - 1, hoy.year]

    wb = load_workbook(EXCEL)
    total = 0

    print("IPC (variación mensual, %):")
    ipc = serie_mindicador("ipc", anios)
    ipc.update(serie_bcch(BCCH_SERIE_IPC, anios[0]))  # la fuente oficial tiene prioridad
    total += actualizar_hoja(wb["ipc"], ipc, simular)

    print("Desempleo (tasa, %):")
    total += actualizar_hoja(wb["desempleo"], serie_mindicador("tasa_desempleo", anios), simular)

    if simular:
        print("\nNo se escribió nada.")
        return

    print("Descomposición STL del desempleo:")
    descomponer_desempleo(wb)
    print("Desempleo vs. Imacec:")
    analizar_imacec(wb)
    print("Política monetaria:")
    analizar_politica_monetaria(wb)
    print("Dólar y traspaso al IPC:")
    analizar_dolar(wb)

    if total:
        respaldos = os.path.join(CARPETA, "respaldos")
        os.makedirs(respaldos, exist_ok=True)
        copia = os.path.join(respaldos, f"banco_central_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        shutil.copy2(EXCEL, copia)
        print(f"\n{total} valor(es) actualizados. Respaldo: {copia}")
    else:
        print("\nLos datos ya estaban al día; se recalcularon las hojas de análisis.")
    wb.save(EXCEL)
    print("Siguiente paso: en Power BI, Inicio → Actualizar.")


if __name__ == "__main__":
    main()
