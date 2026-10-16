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
    python automatizar_fichadas.py --archivo "C:/Fichadas/entrada/Copia de 1° QUINCENA 062026.xls"

Uso monitor automático:
    python automatizar_fichadas.py --monitorear "C:/Fichadas/entrada"
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

# Carpeta donde esta ubicado este archivo Python
BASE_DIR = Path(__file__).resolve().parent


@dataclass
class ConfigFichadas:
    # Carpeta donde quedan los reportes procesados
    carpeta_salida: Path = BASE_DIR / "salidas_fichadas"

    # Carpeta donde se guardan los originales para trazabilidad
    carpeta_historico_originales: Path = BASE_DIR / "Archivos de origen"

    # Día de corte de la primera quincena
    dia_corte_primera_quincena: int = 15

    # Hora límite de cómputo para cierre de quincena
    hora_corte_quincena: str = "23:59:59"

    # Si un operario no tiene fichada de Entrada, se puede computar una Salida inicial
    computar_salida_inicial_sin_entrada: bool = True
    hora_maxima_salida_inicial: str = "08:00:00"

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

    # Ventana para detectar correcciones inmediatas de estado.
    # Ejemplo: marca Salida por error y al minuto marca Entrada.
    minutos_correccion_estado: float = 5.0

    # Si una entrada queda abierta más horas que este límite y aparece otra Entrada,
    # se considera que la entrada anterior quedó incompleta y se reemplaza.
    horas_maximas_entrada_abierta: float = 18.0

   

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
# 5) DETECCION DE DUPLICADOS CERCANOS
# ============================================================
def detectar_correcciones_estado(
    df_valido: pd.DataFrame,
    config: ConfigFichadas = CONFIG
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Detecta fichadas donde el operario marca un estado incorrecto
    y lo corrige inmediatamente.

    Casos contemplados:

    Caso 1:
        Salida incorrecta -> Entrada correcta
        Ejemplo:
            06:01 Salida
            06:02 Entrada

        Se excluye la Salida incorrecta y se conserva la Entrada.

    Caso 2:
        Entrada incorrecta -> Salida correcta
        Ejemplo:
            06:00 Entrada
            16:55 Entrada  <- incorrecta
            16:56 Salida   <- correcta

        Si ya existía una entrada abierta anteriormente,
        se excluye la Entrada incorrecta y se conserva la Salida.
    """

    df = df_valido.copy().sort_values(["legajo", "timestamp"]).reset_index(drop=True)

    indices_a_excluir = set()
    correcciones = []

    for (legajo, nombre), grupo in df.groupby(["legajo", "nombre"], sort=False):
        grupo = grupo.sort_values("timestamp")
        registros = list(grupo.iterrows())

        entrada_abierta = None

        i = 0
        while i < len(registros):
            idx_actual, fila_actual = registros[i]

            ts_actual = fila_actual["timestamp"]
            estado_actual = fila_actual["estado"]

            idx_siguiente = None
            fila_siguiente = None

            if i + 1 < len(registros):
                idx_siguiente, fila_siguiente = registros[i + 1]
                ts_siguiente = fila_siguiente["timestamp"]
                estado_siguiente = fila_siguiente["estado"]
                minutos_diferencia = (ts_siguiente - ts_actual).total_seconds() / 60
            else:
                ts_siguiente = None
                estado_siguiente = None
                minutos_diferencia = None

            # ============================================================
            # CASO 1:
            # Salida sin entrada abierta y enseguida Entrada.
            # Se asume que la Salida fue marcada por error.
            # ============================================================
            if (
                entrada_abierta is None
                and estado_actual == "Salida"
                and estado_siguiente == "Entrada"
                and minutos_diferencia is not None
                and 0 <= minutos_diferencia <= config.minutos_correccion_estado
            ):
                indices_a_excluir.add(idx_actual)

                correcciones.append({
                    "legajo": legajo,
                    "nombre": nombre,
                    "timestamp_excluido": ts_actual,
                    "estado_excluido": estado_actual,
                    "timestamp_conservado": ts_siguiente,
                    "estado_conservado": estado_siguiente,
                    "minutos_diferencia": round(minutos_diferencia, 2),
                    "tipo_correccion": "Salida incorrecta corregida por Entrada",
                    "contabilizada_en_total_horas": "No",
                    "observacion_calculo": (
                        "Corrección inmediata de estado: "
                        "se excluye la Salida incorrecta y se conserva la Entrada siguiente"
                    ),
                })

                # No abre/cierra turno con la Salida incorrecta.
                # La Entrada siguiente queda para ser procesada normalmente.
                i += 1
                continue

            # ============================================================
            # CASO 2:
            # Hay una entrada abierta, el operario marca Entrada por error
            # y enseguida corrige con Salida.
            # Se excluye la Entrada incorrecta y se conserva la Salida.
            # ============================================================
            if (
                entrada_abierta is not None
                and estado_actual == "Entrada"
                and estado_siguiente == "Salida"
                and minutos_diferencia is not None
                and 0 <= minutos_diferencia <= config.minutos_correccion_estado
            ):
                indices_a_excluir.add(idx_actual)

                correcciones.append({
                    "legajo": legajo,
                    "nombre": nombre,
                    "timestamp_excluido": ts_actual,
                    "estado_excluido": estado_actual,
                    "timestamp_conservado": ts_siguiente,
                    "estado_conservado": estado_siguiente,
                    "minutos_diferencia": round(minutos_diferencia, 2),
                    "tipo_correccion": "Entrada incorrecta corregida por Salida",
                    "contabilizada_en_total_horas": "No",
                    "observacion_calculo": (
                        "Corrección inmediata de estado: "
                        "se excluye la Entrada incorrecta y se conserva la Salida siguiente"
                    ),
                })

                # No reemplazamos la entrada_abierta.
                # La Salida siguiente queda para cerrar la entrada original.
                i += 1
                continue

            # ============================================================
            # Actualización del estado lógico para detectar casos futuros
            # ============================================================
            if estado_actual == "Entrada":
                entrada_abierta = ts_actual

            elif estado_actual == "Salida":
                if entrada_abierta is not None:
                    entrada_abierta = None

            i += 1

    df_limpio = df.drop(index=list(indices_a_excluir)).reset_index(drop=True)
    df_correcciones = pd.DataFrame(correcciones)

    return df_limpio, df_correcciones

# ============================================================
# 6) IDENTIFICACIÓN DE TURNOS QUE INICIAN QUINCENA
# ============================================================

def es_inicio_de_quincena(timestamp: pd.Timestamp, config: ConfigFichadas = CONFIG) -> bool:
    """
    Determina si una fecha corresponde al primer día computable de una quincena:
    - día 1 del mes
    - día 16 del mes
    """

    fecha = pd.Timestamp(timestamp)

    return fecha.day in [1, config.dia_corte_primera_quincena + 1]


def obtener_inicio_de_quincena(timestamp: pd.Timestamp, config: ConfigFichadas = CONFIG) -> pd.Timestamp:
    """
    Devuelve el inicio computable de la quincena correspondiente.

    Si la fecha está entre el día 1 y 15:
        inicio = día 1 a las 00:00:00

    Si la fecha está entre el día 16 y fin de mes:
        inicio = día 16 a las 00:00:00
    """

    fecha = pd.Timestamp(timestamp)

    anio = fecha.year
    mes = fecha.month

    if fecha.day <= config.dia_corte_primera_quincena:
        dia_inicio = 1
    else:
        dia_inicio = config.dia_corte_primera_quincena + 1

    return pd.Timestamp(datetime(anio, mes, dia_inicio, 0, 0, 0))


def salida_inicial_es_computable(
    salida: pd.Timestamp,
    es_primera_fichada_del_empleado: bool,
    config: ConfigFichadas = CONFIG
) -> bool:
    """
    Determina si una Salida sin entrada previa puede computarse desde
    el inicio de la quincena.

    Caso típico:
        01/07/2026 05:52 Salida

    Se asume que la entrada real fue el día anterior, pero para esta
    quincena corresponde computar desde 00:00 hasta la salida.
    """

    if not config.computar_salida_inicial_sin_entrada:
        return False

    salida = pd.Timestamp(salida)

    if not es_primera_fichada_del_empleado:
        return False

    if not es_inicio_de_quincena(salida, config):
        return False

    hora_maxima = datetime.strptime(
        config.hora_maxima_salida_inicial,
        "%H:%M:%S"
    ).time()

    if salida.time() > hora_maxima:
        return False

    return True

# ============================================================
# 7) ARMADO DE TURNOS Y DETECCIÓN DE ERRORES
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
    - Salida sin entrada abierta => error, salvo salida inicial computable.
    - Entrada abierta al final => puede ser error o corte automático de quincena.

    Reglas especiales:
    - Si un empleado entra el día 15 y sigue trabajando luego de las 00:00 del día 16,
      se computa desde la entrada hasta el día 15 a las 23:59:59.
    - Si un empleado entra el último día del mes y sigue trabajando luego de las 00:00
      del mes siguiente, se computa hasta el último día del mes a las 23:59:59.
    - Si la primera fichada del día 1 o día 16 es una Salida de madrugada,
      se puede computar desde las 00:00:00 del inicio de quincena hasta esa salida.
    - Si una entrada queda abierta más de la jornada crítica, no se arrastra durante días.
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
        entrada_real,
        entrada_computable,
        salida_real,
        salida_computable,
        turno_recortado,
        observacion_calculo
    ):
        entrada_computable = pd.Timestamp(entrada_computable)
        salida_computable = pd.Timestamp(salida_computable)

        if pd.isna(entrada_real) or pd.isna(salida_real):
            horas_reales = None
        else:
            entrada_real = pd.Timestamp(entrada_real)
            salida_real = pd.Timestamp(salida_real)
            horas_reales = (salida_real - entrada_real).total_seconds() / 3600

        horas_computables = (
            salida_computable - entrada_computable
        ).total_seconds() / 3600

        turnos.append({
            "legajo": legajo,
            "nombre": nombre,
            "entrada_real": entrada_real,
            "entrada_computable": entrada_computable,
            "salida_real": salida_real,
            "salida_computable": salida_computable,
            "horas_reales": horas_reales,
            "horas_computables": horas_computables,
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

        grupo_ordenado = grupo.sort_values("timestamp").reset_index(drop=True)

        for posicion, row in grupo_ordenado.iterrows():
            ts = row["timestamp"]
            estado = row["estado"]

            # ============================================================
            # CASO ENTRADA
            # ============================================================
            if estado == "Entrada":

                if entrada_abierta is not None:
                    limite_quincena = obtener_fin_de_quincena(entrada_abierta, config)
                    horas_desde_entrada_abierta = (
                        ts - entrada_abierta
                    ).total_seconds() / 3600

                    # Caso 1:
                    # Entrada abierta justo en día de cierre de quincena.
                    # Se computa hasta 23:59:59 y luego se abre la nueva entrada.
                    if (
                        config.aplicar_corte_quincenal
                        and entrada_corresponde_a_cierre_quincena(entrada_abierta)
                        and ts > limite_quincena
                    ):
                        registrar_turno(
                            legajo=legajo,
                            nombre=nombre,
                            entrada_real=entrada_abierta,
                            entrada_computable=entrada_abierta,
                            salida_real=pd.NaT,
                            salida_computable=limite_quincena,
                            turno_recortado=True,
                            observacion_calculo=(
                                "Turno abierto en cierre de quincena: "
                                "se computa hasta las 23:59:59 del día de corte"
                            )
                        )

                        entrada_abierta = ts

                    # Caso 2:
                    # Entrada abierta demasiado antigua.
                    # Evita turnos absurdos de varios días.
                    elif horas_desde_entrada_abierta > config.horas_maximas_entrada_abierta:
                        registrar_error(
                            legajo=legajo,
                            nombre=nombre,
                            timestamp=entrada_abierta,
                            tipo_error="Entrada abierta demasiado antigua",
                            detalle=(
                                f"Entrada abierta desde {entrada_abierta}. "
                                f"Aparece nueva Entrada el {ts}. "
                                f"Transcurrieron {horas_desde_entrada_abierta:.2f} horas. "
                                "Se excluye la entrada anterior y se toma la nueva."
                            ),
                        )

                        entrada_abierta = ts

                    # Caso 3:
                    # Dos entradas dentro de un rango no crítico.
                    # Se informa error, pero se conserva la primera entrada.
                    else:
                        registrar_error(
                            legajo=legajo,
                            nombre=nombre,
                            timestamp=ts,
                            tipo_error="Entrada sin salida previa",
                            detalle=f"Entrada anterior abierta: {entrada_abierta}",
                        )

                        continue

                else:
                    entrada_abierta = ts

            # ============================================================
            # CASO SALIDA
            # ============================================================
            elif estado == "Salida":

                es_primera_fichada = posicion == 0

                # --------------------------------------------------------
                # Caso 1:
                # Salida sin entrada abierta
                # --------------------------------------------------------
                if entrada_abierta is None:

                    # Primera fichada del período es una salida de madrugada.
                    # Se computa desde 00:00:00 del inicio de quincena.
                    if salida_inicial_es_computable(
                        salida=ts,
                        es_primera_fichada_del_empleado=es_primera_fichada,
                        config=config
                    ):
                        entrada_computable = obtener_inicio_de_quincena(ts, config)
                        salida_real = ts
                        salida_computable = ts

                        registrar_turno(
                            legajo=legajo,
                            nombre=nombre,
                            entrada_real=pd.NaT,
                            entrada_computable=entrada_computable,
                            salida_real=salida_real,
                            salida_computable=salida_computable,
                            turno_recortado=True,
                            observacion_calculo=(
                                "Salida inicial sin entrada previa: "
                                "se computa desde las 00:00:00 del inicio de quincena"
                            )
                        )

                    else:
                        registrar_error(
                            legajo=legajo,
                            nombre=nombre,
                            timestamp=ts,
                            tipo_error="Salida sin entrada previa",
                            detalle="No existe entrada abierta para cerrar turno",
                        )

                    continue

                # --------------------------------------------------------
                # Caso 2:
                # Salida con entrada abierta
                # --------------------------------------------------------
                entrada = entrada_abierta
                salida_real = ts

                horas_reales_preliminares = (
                    salida_real - entrada
                ).total_seconds() / 3600

                # Si el turno supera el límite crítico, no se computa.
                if horas_reales_preliminares > config.jornada_critica_horas:
                    registrar_error(
                        legajo=legajo,
                        nombre=nombre,
                        timestamp=salida_real,
                        tipo_error="Turno crítico no computado",
                        detalle=(
                            f"El turno desde {entrada} hasta {salida_real} "
                            f"arroja {horas_reales_preliminares:.2f} horas. "
                            "Se excluye del cálculo por superar la jornada crítica."
                        ),
                    )

                    entrada_abierta = None
                    continue

                entrada_computable, salida_computable, turno_recortado = aplicar_corte_a_turno(
                    entrada,
                    salida_real,
                    config
                )

                registrar_turno(
                    legajo=legajo,
                    nombre=nombre,
                    entrada_real=entrada,
                    entrada_computable=entrada_computable,
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

        # ============================================================
        # ENTRADA ABIERTA AL FINAL DEL ARCHIVO
        # ============================================================
        if entrada_abierta is not None:
            limite_quincena = obtener_fin_de_quincena(entrada_abierta, config)

            # Si la entrada abierta quedó justo en día de cierre,
            # se computa hasta 23:59:59.
            if (
                config.aplicar_corte_quincenal
                and entrada_corresponde_a_cierre_quincena(entrada_abierta)
            ):
                registrar_turno(
                    legajo=legajo,
                    nombre=nombre,
                    entrada_real=entrada_abierta,
                    entrada_computable=entrada_abierta,
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
# 8) ALERTAS SOBRE JORNADAS CORTAS, LARGAS Y PROMEDIO
# ============================================================

def clasificar_turnos(
    turnos: pd.DataFrame,
    config: ConfigFichadas = CONFIG
) -> pd.DataFrame:
    """
    Clasifica los turnos calculados según reglas de control:
    - Jornada demasiado corta
    - Jornada larga
    - Jornada crítica
    - Jornada superior al promedio habitual del operario

    Importante:
    La columna 'horas' representa las horas computables para liquidación.
    """

    if turnos.empty:
        return turnos

    turnos = turnos.copy()

    # Convertir columnas de fecha/hora a datetime
    for col in ["entrada_real", "entrada_computable", "salida_real", "salida_computable"]:
        if col in turnos.columns:
            turnos[col] = pd.to_datetime(turnos[col], errors="coerce")

    # Fecha de entrada para análisis.
    # Se usa entrada_computable porque en algunos casos no existe entrada_real,
    # por ejemplo cuando la primera fichada de la quincena es una Salida.
    turnos["fecha_entrada"] = turnos["entrada_computable"].dt.date

    # La columna 'horas' es la que se usa para liquidar
    turnos["horas"] = pd.to_numeric(
        turnos["horas"],
        errors="coerce"
    ).fillna(0).round(4)

    # Si existe horas_computables, la mantenemos alineada con horas
    if "horas_computables" in turnos.columns:
        turnos["horas_computables"] = pd.to_numeric(
            turnos["horas_computables"],
            errors="coerce"
        ).fillna(0).round(4)

    if "horas_reales" in turnos.columns:
        turnos["horas_reales"] = pd.to_numeric(
            turnos["horas_reales"],
            errors="coerce"
        ).round(4)

    promedio_por_legajo = turnos.groupby("legajo")["horas"].transform("mean")

    turnos["promedio_operario"] = promedio_por_legajo.round(4)

    turnos["porcentaje_sobre_promedio"] = (
        (turnos["horas"] / promedio_por_legajo) - 1
    ).replace([float("inf"), -float("inf")], 0).fillna(0).round(4)

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
# 9) RESUMEN POR EMPLEADO
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
# 10) NOMBRE DE QUINCENA Y GUARDADO DE SALIDAS
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
    correcciones_estado: pd.DataFrame,
    config: ConfigFichadas,
) -> Path:
    nombre_quincena = obtener_nombre_quincena(df_original_normalizado, archivo)
    carpeta_reporte = config.carpeta_salida / nombre_quincena

    carpeta_reporte.mkdir(parents=True, exist_ok=True)
    config.carpeta_historico_originales.mkdir(parents=True, exist_ok=True)

    # Guarda CSVs separados para trazabilidad y uso simple.
    df_original_normalizado.to_csv(
        carpeta_reporte / "01_fichadas_normalizadas.csv",
        index=False,
        encoding="utf-8-sig"
    )

    turnos.to_csv(
        carpeta_reporte / "02_turnos_calculados.csv",
        index=False,
        encoding="utf-8-sig"
    )

    resumen.to_csv(
        carpeta_reporte / "03_resumen_por_empleado.csv",
        index=False,
        encoding="utf-8-sig"
    )

    errores.to_csv(
        carpeta_reporte / "04_errores_fichada.csv",
        index=False,
        encoding="utf-8-sig"
    )

    duplicados.to_csv(
        carpeta_reporte / "05_duplicados_detectados.csv",
        index=False,
        encoding="utf-8-sig"
    )

    correcciones_estado.to_csv(
        carpeta_reporte / "06_correcciones_estado.csv",
        index=False,
        encoding="utf-8-sig"
    )

    # También genera un Excel consolidado con varias hojas.
    excel_salida = carpeta_reporte / f"{nombre_quincena}_reporte.xlsx"

    with pd.ExcelWriter(excel_salida, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        turnos.to_excel(writer, sheet_name="Turnos", index=False)
        errores.to_excel(writer, sheet_name="Errores", index=False)
        duplicados.to_excel(writer, sheet_name="Duplicados", index=False)
        correcciones_estado.to_excel(writer, sheet_name="Correcciones", index=False)
        df_original_normalizado.to_excel(writer, sheet_name="Normalizado", index=False)

    # Archiva el original, sin pisar si ya existe.
    marca_tiempo = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino_original = (
        config.carpeta_historico_originales
        / f"{nombre_quincena}_{marca_tiempo}_{archivo.name}"
    )

    shutil.copy2(archivo, destino_original)

    return carpeta_reporte


# ============================================================
# 11) PROCESO PRINCIPAL
# ============================================================

def procesar_archivo(archivo: str | Path, config: ConfigFichadas = CONFIG) -> Path:
    archivo = Path(archivo)

    if not archivo.exists():
        raise FileNotFoundError(f"No existe el archivo: {archivo}")

    # 1) Leer archivo original
    df = leer_excel_fichadas(archivo)

    # 2) Detectar y excluir duplicados cercanos
    df_valido, duplicados = detectar_duplicados_cercanos(df, config)

    # 3) Detectar y excluir correcciones inmediatas de estado
    # Ejemplo:
    # - Salida incorrecta seguida de Entrada correcta
    # - Entrada incorrecta seguida de Salida correcta
    df_valido, correcciones_estado = detectar_correcciones_estado(df_valido, config)

    # 4) Generar turnos con las fichadas ya depuradas
    turnos, errores = generar_turnos(df_valido, config)

    # 5) Clasificar turnos: OK, revisar, jornada corta, jornada larga, etc.
    turnos = clasificar_turnos(turnos, config)

    # 6) Generar resumen por empleado
    resumen = generar_resumen(turnos, errores, duplicados)

    # 7) Guardar reportes
    carpeta_reporte = guardar_reportes(
        archivo=archivo,
        df_original_normalizado=df,
        turnos=turnos,
        resumen=resumen,
        errores=errores,
        duplicados=duplicados,
        correcciones_estado=correcciones_estado,
        config=config,
    ) 

    print("Proceso finalizado.")
    print(f"Archivo procesado: {archivo}")
    print(f"Reporte generado en: {carpeta_reporte.resolve()}")
    print(f"Fichadas leídas: {len(df)}")
    print(f"Duplicados cercanos detectados: {len(duplicados)}")
    print(f"Correcciones inmediatas detectadas: {len(correcciones_estado)}")
    print(f"Turnos calculados: {len(turnos)}")
    print(f"Errores de fichada: {len(errores)}")

    return carpeta_reporte


# ============================================================
# 12) MONITOREO DE CARPETA LAN
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
# 13) CLI
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
