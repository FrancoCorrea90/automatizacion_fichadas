"""
Automatización de fichadas por quincena
Autor: Py / Franco

Objetivo:
- Tomar el archivo descargado desde el reloj (Excel .xls o .xlsx)
- Normalizar fichadas de Entrada / Salida
- Detectar duplicados, registros faltantes y jornadas atípicas
- Calcular horas trabajadas por persona
- Guardar salidas separadas por quincena y archivar el archivo original

Instalación sugerida:
    pip install pandas openpyxl xlrd watchdog

Uso manual:
    python automatizar_fichadas.py --archivo "data"

Uso monitor automático:
    python automatizar_fichadas.py --monitorear "data"
"""

from __future__ import annotations

import argparse
import re
import shutil
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import calendar
import pandas as pd


# ============================================================
# 1) CONFIGURACIÓN GENERAL
# ============================================================

@dataclass
class ConfigFichadas:
    # Carpeta donde quedan los reportes procesados
    carpeta_salida: Path = Path(
        r"data/carpeta_salida"
    )

    # Carpeta donde se guardan los Excel originales, para trazabilidad
    carpeta_historico_originales: Path = Path(
        r"data/carpeta_historico_originales"
    )

    # Día de corte de la primera quincena
    dia_corte_primera_quincena: int = 15

    # Hora límite de cómputo para cierre de quincena
    hora_corte_quincena: str = "23:59:59"

    # Activar o desactivar corte automático de turnos nocturnos
    aplicar_corte_quincenal: bool = True

    # Ventana para considerar duplicados cercanos
    minutos_duplicado: float = 2.0

    # Jornadas sospechosas por duración baja
    jornada_minima_minutos: float = 10.0

    # Jornada larga: no necesariamente error. Se marca para revisión.
    jornada_larga_horas: float = 15.0

    # Jornada demasiado larga: muy probablemente error de fichada.
    jornada_critica_horas: float = 18.0

    # Alerta si una jornada supera el promedio del operario en este porcentaje.
    porcentaje_sobre_promedio: float = 0.30


CONFIG = ConfigFichadas()


# ============================================================
# 2) LECTURA Y NORMALIZACIÓN DEL ARCHIVO
# ============================================================

def normalizar_texto(valor) -> str:
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def buscar_fila_encabezado(df_sin_header: pd.DataFrame) -> int:
    """
    Busca automáticamente la fila donde están los encabezados.
    En el archivo ejemplo, el encabezado aparece como:
    AC Nº | Nombre | Día/Hora | [Hora] | Estado
    """
    for idx, fila in df_sin_header.iterrows():
        valores = [normalizar_texto(x).lower() for x in fila.tolist()]
        contiene_legajo = any("ac" in v and ("n" in v or "nº" in v or "n°" in v) for v in valores)
        contiene_nombre = any("nombre" in v for v in valores)
        contiene_estado = any("estado" in v for v in valores)
        if contiene_legajo and contiene_nombre and contiene_estado:
            return idx
    raise ValueError("No se encontró la fila de encabezados con AC Nº, Nombre y Estado.")




def leer_excel_seguro(ruta_archivo: Path, header=None) -> pd.DataFrame:
    """
    Lee archivos de fichadas con mayor tolerancia.

    Soporta:
    - .xlsx reales mediante openpyxl
    - .xls reales mediante xlrd
    - archivos exportados por relojes como HTML con extensión .xls

    Esto evita el error:
    Excel file format cannot be determined, you must specify an engine manually.
    """
    ruta_archivo = Path(ruta_archivo)
    extension = ruta_archivo.suffix.lower()
    errores = []

    # 1) Intento según extensión
    if extension == ".xlsx":
        try:
            return pd.read_excel(ruta_archivo, header=header, engine="openpyxl")
        except Exception as exc:
            errores.append(f"openpyxl: {exc}")

    if extension == ".xls":
        try:
            return pd.read_excel(ruta_archivo, header=header, engine="xlrd")
        except Exception as exc:
            errores.append(f"xlrd: {exc}")

    # 2) Intento automático
    try:
        return pd.read_excel(ruta_archivo, header=header)
    except Exception as exc:
        errores.append(f"read_excel automático: {exc}")

    # 3) Muchos relojes exportan HTML con extensión .xls
    try:
        tablas = pd.read_html(ruta_archivo, header=header)
        if tablas:
            return tablas[0]
    except Exception as exc:
        errores.append(f"read_html: {exc}")

    raise ValueError(
        "No se pudo leer el archivo como Excel ni como HTML. "
        "Probá abrirlo en Excel y guardarlo como .xlsx.\n\n"
        "Errores detectados:\n" + "\n".join(errores)
    )


def leer_excel_fichadas(ruta_archivo: Path) -> pd.DataFrame:
    """
    Lee archivo .xls o .xlsx de fichadas con estructura:
    AC N° | Nombre | Día/Hora | Estado
    """
    ruta_archivo = Path(ruta_archivo)

    # Lee el archivo sin asumir encabezado
    df_raw = leer_excel_seguro(ruta_archivo, header=None)

    # Busca la fila donde están los encabezados reales
    fila_header = buscar_fila_encabezado(df_raw)

    # Lee nuevamente usando esa fila como encabezado
    df = leer_excel_seguro(ruta_archivo, header=fila_header)
    df = df.dropna(how="all")

    # Nos quedamos con las primeras 4 columnas reales
    columnas = list(df.columns)

    print("Columnas detectadas:")
    print(columnas)
    print("Cantidad de columnas:", len(columnas))

    if len(columnas) < 4:
        raise ValueError(
            "El archivo no tiene las 4 columnas esperadas: AC N°, Nombre, Día/Hora, Estado."
        )

    df = df.iloc[:, :4].copy()
    df.columns = ["legajo", "nombre", "fecha_hora", "estado"]

    # Limpieza básica
    df["legajo"] = df["legajo"].astype(str).str.strip()
    df["nombre"] = df["nombre"].astype(str).str.strip()
    df["estado"] = df["estado"].astype(str).str.strip().str.title()

    # Convertir Día/Hora a datetime
    df["timestamp"] = pd.to_datetime(df["fecha_hora"], errors="coerce", dayfirst=True)

    # Eliminar registros incompletos
    df = df.dropna(subset=["legajo", "timestamp", "estado"])

    # Quedarse solo con Entrada / Salida
    df = df[df["estado"].isin(["Entrada", "Salida"])]

    # Ordenar por empleado y fecha/hora
    df = df[["legajo", "nombre", "timestamp", "estado"]].sort_values(
        ["legajo", "timestamp", "estado"]
    ).reset_index(drop=True)

    return df


# ============================================================
# 3) DETECCIÓN DE DUPLICADOS CERCANOS
# ============================================================

def detectar_duplicados_cercanos(df: pd.DataFrame, config: ConfigFichadas) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Marca como duplicado cercano una segunda fichada del mismo legajo,
    mismo estado, en una ventana de pocos minutos.
    Se conserva la primera marca y se excluye la repetida del cálculo.
    """
    df = df.copy().sort_values(["legajo", "timestamp", "estado"]).reset_index(drop=True)

    df["timestamp_anterior"] = df.groupby("legajo")["timestamp"].shift()
    df["estado_anterior"] = df.groupby("legajo")["estado"].shift()
    df["minutos_desde_anterior"] = (
        df["timestamp"] - df["timestamp_anterior"]
    ).dt.total_seconds() / 60

    mascara_duplicado = (
        (df["estado"] == df["estado_anterior"]) &
        (df["minutos_desde_anterior"].between(0, config.minutos_duplicado, inclusive="both"))
    )

    duplicados = df.loc[mascara_duplicado, [
        "legajo", "nombre", "timestamp", "estado", "minutos_desde_anterior"
    ]].copy()
    duplicados["observacion"] = "Duplicado cercano: se excluye del cálculo"

    df_valido = df.loc[~mascara_duplicado, ["legajo", "nombre", "timestamp", "estado"]].copy()
    return df_valido, duplicados

# ============================================================
# 4) CORTE QUINCENAL DE TURNOS
# ============================================================

import calendar
from datetime import datetime


def obtener_fin_de_quincena(
    timestamp: pd.Timestamp,
    config: ConfigFichadas = CONFIG
) -> pd.Timestamp:
    """
    Devuelve el límite de cómputo de la quincena correspondiente
    al timestamp recibido.

    Primera quincena:
        desde el día 1 hasta el día 15 a las 23:59:59

    Segunda quincena:
        desde el día 16 hasta el último día del mes a las 23:59:59
    """

    fecha = pd.Timestamp(timestamp)

    anio = fecha.year
    mes = fecha.month
    dia = fecha.day

    hora_corte = datetime.strptime(
        config.hora_corte_quincena,
        "%H:%M:%S"
    ).time()

    if dia <= config.dia_corte_primera_quincena:
        dia_corte = config.dia_corte_primera_quincena
    else:
        dia_corte = calendar.monthrange(anio, mes)[1]

    return pd.Timestamp(
        datetime.combine(
            datetime(anio, mes, dia_corte).date(),
            hora_corte
        )
    )


def aplicar_corte_a_turno(
    entrada: pd.Timestamp,
    salida_real: pd.Timestamp,
    config: ConfigFichadas = CONFIG
) -> tuple[pd.Timestamp, pd.Timestamp, bool]:
    """
    Aplica el corte quincenal al turno.

    Si un turno empieza dentro de una quincena pero termina después
    del cierre de esa quincena, se computa solo hasta el límite.

    Ejemplo:
        Entrada: 15/06/2026 22:00
        Salida real: 16/06/2026 06:00
        Salida computable: 15/06/2026 23:59:59
    """

    entrada = pd.Timestamp(entrada)
    salida_real = pd.Timestamp(salida_real)

    if not config.aplicar_corte_quincenal:
        return entrada, salida_real, False

    limite_quincena = obtener_fin_de_quincena(entrada, config)

    if salida_real > limite_quincena:
        return entrada, limite_quincena, True

    return entrada, salida_real, False


# ============================================================
# 5) ARMADO DE TURNOS Y DETECCIÓN DE ERRORES
# ============================================================

def generar_turnos(
    df_valido: pd.DataFrame,
    config: ConfigFichadas = CONFIG
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Recorre cada legajo ordenado por fecha/hora.

    Regla base:
    - Entrada abre turno.
    - Salida cierra turno.
    - Entrada con otra entrada abierta => se analiza.
    - Salida sin entrada abierta => error.
    - Entrada abierta al final => puede ser error o corte automático de quincena.

    Regla especial de corte quincenal:
    - Si un empleado entra el día 15 y sigue trabajando luego de las 00:00 del día 16,
      se computa desde la entrada hasta el día 15 a las 23:59:59.
    - Si un empleado entra el último día del mes y sigue trabajando luego de las 00:00
      del mes siguiente, se computa hasta el último día del mes a las 23:59:59.
    """

    turnos = []
    errores = []

    def registrar_error(legajo, nombre, timestamp, tipo_error, detalle):
        errores.append({
            "legajo": legajo,
            "nombre": nombre,
            "timestamp": timestamp,
            "tipo_error": tipo_error,
            "detalle": detalle,
            "contabilizada_en_total_horas": "No",
            "observacion_calculo": "Fichada errónea: se excluye del cálculo",
        })

    def registrar_turno(
        legajo,
        nombre,
        entrada,
        salida_real,
        salida_computable,
        turno_recortado,
        observacion_calculo
    ):
        entrada = pd.Timestamp(entrada)
        salida_computable = pd.Timestamp(salida_computable)

        if pd.isna(salida_real):
            horas_reales = None
        else:
            salida_real = pd.Timestamp(salida_real)
            horas_reales = (salida_real - entrada).total_seconds() / 3600

        horas_computables = (salida_computable - entrada).total_seconds() / 3600

        turnos.append({
            "legajo": legajo,
            "nombre": nombre,
            "entrada": entrada,
            "salida_real": salida_real,
            "salida_computable": salida_computable,
            "horas_reales": horas_reales,
            "horas": horas_computables,
            "turno_recortado_por_cierre_quincena": "Sí" if turno_recortado else "No",
            "contabilizada_en_total_horas": "Sí",
            "observacion_calculo": observacion_calculo,
        })

    def entrada_corresponde_a_cierre_quincena(entrada):
        """
        Determina si una entrada abierta corresponde al día de cierre:
        - día 15 para primera quincena
        - último día del mes para segunda quincena
        """
        entrada = pd.Timestamp(entrada)
        limite_quincena = obtener_fin_de_quincena(entrada, config)

        return entrada.date() == limite_quincena.date()

    for (legajo, nombre), grupo in df_valido.groupby(["legajo", "nombre"], sort=False):
        entrada_abierta = None

        for _, row in grupo.sort_values("timestamp").iterrows():
            ts = row["timestamp"]
            estado = row["estado"]

            if estado == "Entrada":

                if entrada_abierta is not None:
                    limite_quincena = obtener_fin_de_quincena(entrada_abierta, config)

                    # Caso especial:
                    # Había una entrada abierta en el día de cierre de quincena
                    # y aparece otra marca posterior al cierre.
                    # En vez de excluir la entrada anterior, se computa hasta 23:59:59.
                    if (
                        config.aplicar_corte_quincenal
                        and entrada_corresponde_a_cierre_quincena(entrada_abierta)
                        and ts > limite_quincena
                    ):
                        registrar_turno(
                            legajo=legajo,
                            nombre=nombre,
                            entrada=entrada_abierta,
                            salida_real=pd.NaT,
                            salida_computable=limite_quincena,
                            turno_recortado=True,
                            observacion_calculo=(
                                "Turno abierto en cierre de quincena: "
                                "se computa hasta las 23:59:59 del día de corte"
                            )
                        )

                    else:
                        registrar_error(
                            legajo=legajo,
                            nombre=nombre,
                            timestamp=ts,
                            tipo_error="Entrada sin salida previa",
                            detalle=f"Entrada anterior abierta: {entrada_abierta}",
                        )

                # Se toma la última entrada como nueva entrada abierta.
                entrada_abierta = ts

            elif estado == "Salida":

                if entrada_abierta is None:
                    registrar_error(
                        legajo=legajo,
                        nombre=nombre,
                        timestamp=ts,
                        tipo_error="Salida sin entrada previa",
                        detalle="No existe entrada abierta para cerrar turno",
                    )

                else:
                    entrada = entrada_abierta
                    salida_real = ts

                    entrada_computable, salida_computable, turno_recortado = aplicar_corte_a_turno(
                        entrada,
                        salida_real,
                        config
                    )

                    registrar_turno(
                        legajo=legajo,
                        nombre=nombre,
                        entrada=entrada,
                        salida_real=salida_real,
                        salida_computable=salida_computable,
                        turno_recortado=turno_recortado,
                        observacion_calculo=(
                            "Turno válido: se computa con corte quincenal"
                            if turno_recortado
                            else "Turno válido: se incluye en el cálculo"
                        )
                    )

                    entrada_abierta = None

        if entrada_abierta is not None:
            limite_quincena = obtener_fin_de_quincena(entrada_abierta, config)

            # Caso clave:
            # Si la entrada abierta quedó en el día de cierre de quincena,
            # no se manda a errores. Se liquida hasta 23:59:59.
            if (
                config.aplicar_corte_quincenal
                and entrada_corresponde_a_cierre_quincena(entrada_abierta)
            ):
                registrar_turno(
                    legajo=legajo,
                    nombre=nombre,
                    entrada=entrada_abierta,
                    salida_real=pd.NaT,
                    salida_computable=limite_quincena,
                    turno_recortado=True,
                    observacion_calculo=(
                        "Turno abierto en cierre de quincena: "
                        "se computa hasta las 23:59:59 del día de corte"
                    )
                )

            else:
                registrar_error(
                    legajo=legajo,
                    nombre=nombre,
                    timestamp=entrada_abierta,
                    tipo_error="Entrada sin salida al cierre de quincena",
                    detalle="Quedó una entrada abierta sin salida posterior",
                )

    return pd.DataFrame(turnos), pd.DataFrame(errores)


# ============================================================
# 6) ALERTAS SOBRE JORNADAS CORTAS, LARGAS Y PROMEDIO
# ============================================================

def clasificar_turnos(turnos: pd.DataFrame, config: ConfigFichadas) -> pd.DataFrame:
    if turnos.empty:
        return turnos

    turnos = turnos.copy()
    turnos["fecha_entrada"] = turnos["entrada"].dt.date
    turnos["horas"] = turnos["horas"].round(4)

    promedio_por_legajo = turnos.groupby("legajo")["horas"].transform("mean")
    turnos["promedio_operario"] = promedio_por_legajo.round(4)
    turnos["porcentaje_sobre_promedio"] = ((turnos["horas"] / promedio_por_legajo) - 1).round(4)

    def evaluar(row):
        if row["horas"] < config.jornada_minima_minutos / 60:
            return "ERROR: jornada demasiado corta"
        if row["horas"] > config.jornada_critica_horas:
            return "ERROR: jornada crítica muy larga"
        if row["horas"] > config.jornada_larga_horas:
            return "REVISAR: jornada larga, puede estar autorizada"
        if row["porcentaje_sobre_promedio"] > config.porcentaje_sobre_promedio:
            return "REVISAR: supera su promedio habitual"
        return "OK"

    turnos["estado_control"] = turnos.apply(evaluar, axis=1)
    return turnos


# ============================================================
# 7) RESUMEN POR EMPLEADO
# ============================================================

def generar_resumen(turnos_clasificados: pd.DataFrame, errores: pd.DataFrame, duplicados: pd.DataFrame) -> pd.DataFrame:
    if turnos_clasificados.empty:
        return pd.DataFrame(columns=[
            "legajo", "nombre", "total_horas", "cantidad_turnos", "promedio_horas_turno",
            "max_horas_turno", "turnos_a_revisar", "errores_fichada", "duplicados_detectados"
        ])

    resumen = turnos_clasificados.groupby(["legajo", "nombre"], as_index=False).agg(
        total_horas=("horas", "sum"),
        cantidad_turnos=("horas", "count"),
        promedio_horas_turno=("horas", "mean"),
        max_horas_turno=("horas", "max"),
        turnos_a_revisar=("estado_control", lambda x: (x != "OK").sum()),
    )

    if not errores.empty:
        err = errores.groupby(["legajo", "nombre"]).size().reset_index(name="errores_fichada")
        resumen = resumen.merge(err, on=["legajo", "nombre"], how="left")
    else:
        resumen["errores_fichada"] = 0

    if not duplicados.empty:
        dup = duplicados.groupby(["legajo", "nombre"]).size().reset_index(name="duplicados_detectados")
        resumen = resumen.merge(dup, on=["legajo", "nombre"], how="left")
    else:
        resumen["duplicados_detectados"] = 0

    resumen[["errores_fichada", "duplicados_detectados"]] = resumen[[
        "errores_fichada", "duplicados_detectados"
    ]].fillna(0).astype(int)

    for col in ["total_horas", "promedio_horas_turno", "max_horas_turno"]:
        resumen[col] = resumen[col].round(2)

    resumen["estado_final"] = resumen.apply(
        lambda r: "REVISAR" if (r["errores_fichada"] > 0 or r["turnos_a_revisar"] > 0) else "OK",
        axis=1
    )

    return resumen.sort_values(["estado_final", "legajo"])


# ============================================================
# 8) NOMBRE DE QUINCENA Y GUARDADO DE SALIDAS
# ============================================================

def obtener_nombre_quincena(df: pd.DataFrame, archivo: Path) -> str:
    """
    Usa las fechas internas para generar un nombre estándar:
    fichadas_2026-06_Q1 o fichadas_2026-06_Q2.
    """
    if df.empty:
        base = archivo.stem
        return re.sub(r"[^A-Za-z0-9_-]+", "_", base)

    fecha_min = df["timestamp"].min()
    anio_mes = fecha_min.strftime("%Y-%m")
    quincena = "Q1" if fecha_min.day <= 15 else "Q2"
    return f"fichadas_{anio_mes}_{quincena}"


def guardar_reportes(
    archivo: Path,
    df_original_normalizado: pd.DataFrame,
    turnos: pd.DataFrame,
    resumen: pd.DataFrame,
    errores: pd.DataFrame,
    duplicados: pd.DataFrame,
    config: ConfigFichadas,
) -> Path:
    nombre_quincena = obtener_nombre_quincena(df_original_normalizado, archivo)
    carpeta_reporte = config.carpeta_salida / nombre_quincena
    carpeta_reporte.mkdir(parents=True, exist_ok=True)
    config.carpeta_historico_originales.mkdir(parents=True, exist_ok=True)

    # Guarda CSVs separados para trazabilidad y uso simple.
    df_original_normalizado.to_csv(carpeta_reporte / "01_fichadas_normalizadas.csv", index=False, encoding="utf-8-sig")
    turnos.to_csv(carpeta_reporte / "02_turnos_calculados.csv", index=False, encoding="utf-8-sig")
    resumen.to_csv(carpeta_reporte / "03_resumen_por_empleado.csv", index=False, encoding="utf-8-sig")
    errores.to_csv(carpeta_reporte / "04_errores_fichada.csv", index=False, encoding="utf-8-sig")
    duplicados.to_csv(carpeta_reporte / "05_duplicados_detectados.csv", index=False, encoding="utf-8-sig")

    # También genera un Excel consolidado con varias hojas.
    excel_salida = carpeta_reporte / f"{nombre_quincena}_reporte.xlsx"
    with pd.ExcelWriter(excel_salida, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        turnos.to_excel(writer, sheet_name="Turnos", index=False)
        errores.to_excel(writer, sheet_name="Errores", index=False)
        duplicados.to_excel(writer, sheet_name="Duplicados", index=False)
        df_original_normalizado.to_excel(writer, sheet_name="Normalizado", index=False)

    # Archiva el original, sin pisar si ya existe.
    marca_tiempo = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino_original = config.carpeta_historico_originales / f"{nombre_quincena}_{marca_tiempo}_{archivo.name}"
    shutil.copy2(archivo, destino_original)

    return carpeta_reporte


# ============================================================
# 9) PROCESO PRINCIPAL
# ============================================================

def procesar_archivo(archivo: str | Path, config: ConfigFichadas = CONFIG) -> Path:
    archivo = Path(archivo)
    if not archivo.exists():
        raise FileNotFoundError(f"No existe el archivo: {archivo}")

    df = leer_excel_fichadas(archivo)
    df_valido, duplicados = detectar_duplicados_cercanos(df, config)
    turnos, errores = generar_turnos(df_valido)
    turnos = clasificar_turnos(turnos, config)
    resumen = generar_resumen(turnos, errores, duplicados)

    carpeta_reporte = guardar_reportes(
        archivo=archivo,
        df_original_normalizado=df,
        turnos=turnos,
        resumen=resumen,
        errores=errores,
        duplicados=duplicados,
        config=config,
    )

    print("Proceso finalizado.")
    print(f"Archivo procesado: {archivo}")
    print(f"Reporte generado en: {carpeta_reporte.resolve()}")
    print(f"Fichadas leídas: {len(df)}")
    print(f"Duplicados cercanos detectados: {len(duplicados)}")
    print(f"Turnos calculados: {len(turnos)}")
    print(f"Errores de fichada: {len(errores)}")
    return carpeta_reporte


# ============================================================
# 10) MONITOREO DE CARPETA LAN
# ============================================================

def monitorear_carpeta(carpeta_entrada: str | Path, config: ConfigFichadas = CONFIG):
    """
    Modo automático: procesa cada nuevo .xls/.xlsx que RRHH copie en la carpeta.
    Requiere: pip install watchdog
    """
    from watchdog.events import FileSystemEventHandler
    from watchdog.observers import Observer

    carpeta_entrada = Path(carpeta_entrada)
    carpeta_entrada.mkdir(parents=True, exist_ok=True)

    class HandlerFichadas(FileSystemEventHandler):
        def on_created(self, event):
            if event.is_directory:
                return
            archivo = Path(event.src_path)
            if archivo.suffix.lower() not in [".xls", ".xlsx"]:
                return

            # Espera breve para asegurar que Windows termine de copiar el archivo.
            time.sleep(3)
            try:
                procesar_archivo(archivo, config)
            except Exception as exc:
                print(f"ERROR procesando {archivo}: {exc}")

    observer = Observer()
    observer.schedule(HandlerFichadas(), str(carpeta_entrada), recursive=False)
    observer.start()

    print(f"Monitoreando carpeta: {carpeta_entrada.resolve()}")
    print("Presioná CTRL+C para detener.")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


# ============================================================
# 11) CLI
# ============================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Automatización de fichadas por quincena")
    parser.add_argument("--archivo", help="Ruta del Excel de fichadas a procesar")
    parser.add_argument("--monitorear", help="Carpeta LAN a monitorear para nuevos Excel")
    args = parser.parse_args()

    if args.archivo:
        procesar_archivo(args.archivo)
    elif args.monitorear:
        monitorear_carpeta(args.monitorear)
    else:
        parser.print_help()
