# Automatización de fichadas por quincena

Herramienta en Python para procesar las marcaciones de un reloj de asistencia, calcular horas por empleado y generar reportes de control por quincena. Desarrollada por Franco Correa.

## Funcionalidades

- Lectura de Excel `.xls`, `.xlsx` y tablas HTML exportadas con extensión `.xls`.
- Detección del encabezado y normalización de entradas y salidas.
- Exclusión de duplicados cercanos y detección de correcciones inmediatas de estado.
- Armado de turnos, control de fichadas incompletas y cálculo de horas computables.
- Corte de turnos al cierre de quincena y tratamiento de salidas iniciales de madrugada.
- Alertas por jornadas cortas, largas y superiores al promedio del empleado.
- Exportación a CSV y Excel; copia del archivo de origen para trazabilidad.
- Procesamiento manual o monitoreo de una carpeta de entrada.

## Requisitos e instalación

Python 3.10 o superior. Desde PowerShell en Windows:

```powershell
git clone https://github.com/FrancoCorrea90/automatizacion_fichadas.git
cd automatizacion_fichadas
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Los comandos usan el ejecutable del entorno virtual directamente; no requieren activarlo.

## Formato de entrada

El archivo debe tener estas cuatro columnas **en este orden**:

| Columna | Contenido |
| --- | --- |
| AC Nº | Legajo del empleado |
| Nombre | Nombre del empleado |
| Día/Hora | Fecha y hora de la marcación |
| Estado | Entrada o Salida |

Puede haber filas previas al encabezado. La fecha se interpreta con día primero. El programa toma las primeras cuatro columnas; un archivo con fecha y hora separadas necesita adaptación previa. Las fechas no interpretables y los estados distintos de Entrada/Salida se descartan durante la normalización.

## Uso manual

```powershell
.\.venv\Scripts\python.exe automatizar_fichadas.py --archivo "C:\Fichadas\entrada\quincena.xlsx"
```

`--archivo` recibe un archivo, no una carpeta.

## Monitoreo de carpeta

```powershell
.\.venv\Scripts\python.exe automatizar_fichadas.py --monitorear "C:\Fichadas\entrada"
```

Procesa nuevos archivos `.xls` y `.xlsx` al detectar su creación. Para detenerlo, presionar `Ctrl+C`. No procesa automáticamente los archivos ya existentes ni los eventos de modificación o renombrado. Espera tres segundos antes de leer: si una copia tarda más, puede fallar y requerir ejecución manual. El comportamiento en una carpeta de red debe validarse en el entorno de uso.

## Configuración y reglas de cálculo

Los parámetros se editan en `ConfigFichadas`, al inicio de `automatizar_fichadas.py`.

| Parámetro | Valor predeterminado | Efecto |
| --- | --- | --- |
| `dia_corte_primera_quincena` | 15 | Cierre de la primera quincena |
| `hora_corte_quincena` | 23:59:59 | Hora computable de cierre |
| `aplicar_corte_quincenal` | True | Recorta turnos en el límite del período |
| `computar_salida_inicial_sin_entrada` | True | Permite inferir el inicio de un turno |
| `hora_maxima_salida_inicial` | 08:00:00 | Límite de la salida inicial inferida |
| `minutos_duplicado` | 2 | Ventana de duplicados cercanos |
| `minutos_correccion_estado` | 5 | Ventana de corrección de estado |
| `horas_maximas_entrada_abierta` | 18 | Reemplazo de entradas antiguas ante una nueva entrada |
| `jornada_minima_minutos` | 10 | Alerta por turno corto |
| `jornada_larga_horas` | 15 | Alerta por turno largo |
| `jornada_critica_horas` | 18 | Excluye turnos cerrados que superan este límite |
| `porcentaje_sobre_promedio` | 0.30 | Alerta por superar un 30 % el promedio calculado |

La primera salida del empleado puede computarse desde las 00:00 si ocurre el día 1 o 16 y hasta las 08:00. Una entrada sin salida al final del archivo puede computarse hasta el cierre si corresponde al día de corte. Son inferencias que requieren revisión del responsable de asistencia.

Las alertas de jornadas cortas o largas no excluyen por sí solas las horas del total. El promedio se calcula con los turnos del archivo procesado, sin consultar un historial externo.

## Reportes

Las carpetas se crean junto al script:

- `salidas_fichadas/fichadas_AAAA-MM_Q1/` o `Q2`: reportes.
- `Archivos de origen/`: copias de los originales con marca de tiempo.

| Archivo | Contenido |
| --- | --- |
| `01_fichadas_normalizadas.csv` | Marcaciones normalizadas |
| `02_turnos_calculados.csv` | Turnos, horas y controles |
| `03_resumen_por_empleado.csv` | Totales e indicadores por empleado con turnos calculados |
| `04_errores_fichada.csv` | Incidencias excluidas |
| `05_duplicados_detectados.csv` | Marcaciones duplicadas excluidas |
| `06_correcciones_estado.csv` | Correcciones inmediatas detectadas |
| `fichadas_AAAA-MM_Qn_reporte.xlsx` | Hojas Resumen, Turnos, Errores, Duplicados, Correcciones y Normalizado |

## Alcance y limitaciones

- Usar un archivo por quincena. El nombre de salida se obtiene de la fecha mínima del archivo; no se separan automáticamente varios períodos. El nombre Q1/Q2 usa el día 15 aunque se cambie el parámetro de corte.
- Si se reprocesa una quincena, se reemplazan sus reportes. Las copias de origen usan una marca de tiempo con precisión de segundos; dos ejecuciones en el mismo segundo pueden coincidir.
- Un turno que cruza el cierre se recorta: el tramo posterior no se genera automáticamente como otro turno en ese procesamiento.
- Los empleados sin turnos calculados no aparecen en el resumen; revisar también la hoja Errores.
- El programa no calcula salarios, adicionales, licencias ni reglas de convenio. Los resultados necesitan revisión antes de su uso administrativo.

## Estructura del repositorio

- `automatizar_fichadas.py`: programa principal y parámetros.
- `requirements.txt`: dependencias de ejecución.
- `.gitignore`: exclusión de datos operativos y archivos locales.
- `README.md`: documentación de instalación y uso.

## Datos personales

Los archivos de entrada y los reportes contienen información de asistencia. Mantenerlos fuera del repositorio y usar datos ficticios para ejemplos. El `.gitignore` previene nuevas incorporaciones, pero no elimina archivos ya registrados en el historial de Git.

## Verificación básica

```powershell
.\.venv\Scripts\python.exe -m py_compile automatizar_fichadas.py
.\.venv\Scripts\python.exe automatizar_fichadas.py --help
```

Estas comprobaciones verifican sintaxis y acceso a la ayuda. Para validar las reglas del negocio se requieren casos de prueba y contraste con resultados revisados manualmente.
