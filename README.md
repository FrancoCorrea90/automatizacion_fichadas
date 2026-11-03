# Automatización de fichadas laborales | Python

**Proyecto de automatización de procesos administrativos y control de asistencia.**  
**Autor:** Franco Correa · **Estado:** versión funcional en validación operativa · **Plataforma:** Windows / Python 3.10+

## Objetivo

Transformar exportaciones quincenales del reloj de asistencia (`.xls` / `.xlsx`) en reportes trazables por empleado. El proceso normaliza registros, identifica inconsistencias, calcula jornadas computables, aplica cortes quincenales y entrega un Excel con hojas de control para Recursos Humanos.

> **Importante:** la herramienta **no liquida salarios** ni reemplaza el control de RRHH. Las horas calculadas son una propuesta de cómputo que debe validarse antes de usarse en una liquidación.

## Flujo de trabajo

```text
Reloj de asistencia → Exportación Excel → data/entrada/
    → Lectura y normalización → Duplicados y correcciones
    → Armado de turnos → Controles de duración y corte quincenal
    → Reporte de RRHH (Excel + CSV) → Validación humana
```

## Reglas principales

| Condición | Regla implementada | Tratamiento |
|---|---|---|
| Turno hasta 14 h | Duración admisible, sujeto a otros controles | Computa |
| Más de 14 h y hasta 17 h | Jornada extensa | Computa y se marca **REVISAR** |
| Más de 17 h | Turno crítico | **No computa** y se registra en Errores |
| Entrada abierta > 17 h y aparece otra Entrada | Marca antigua inconsistente | Se excluye la antigua y se toma la nueva |
| Duplicados del mismo tipo dentro de 2 min | Se conserva la primera marca | Segunda marca excluida |
| Salida sin entrada y seguida de Entrada dentro de 5 min | Corrección inmediata candidata | Se excluye la Salida incorrecta |
| Turno inferior a 10 min | Jornada extremadamente corta | Computa, pero se marca con alerta de error para revisión |
| Turno > 30 % del promedio del empleado | Atípico respecto del archivo procesado | Computa y se marca **REVISAR** |
| Salida inicial, día 1/16 hasta 08:00 | Inferencia desde las 00:00 del período | Computa con observación; validar |
| Entrada sin salida en día de corte | Inferencia hasta 23:59:59, si no excede 17 h | Computa con observación; validar |

**Orden de controles:** el límite crítico de 17 h se evalúa sobre la duración real al cerrar un par Entrada–Salida; los cortes de quincena tienen reglas específicas de inferencia. Las alertas no equivalen a una autorización de horas extras. Los límites de 14/17 horas son parámetros operativos del proyecto, no una interpretación legal del régimen laboral.

Para la explicación completa, ver [Reglas de negocio para RRHH](docs/REGLAS_NEGOCIO_RRHH.md).

## Estructura del proyecto

```text
automatizacion_fichadas/
├── automatizar_fichadas.py         # Procesamiento principal y reglas
├── requirements.txt                # Dependencias
├── scripts/
│   ├── instalar.bat                # Crea entorno e instala dependencias
│   ├── procesar_ultimo.bat         # Procesa exportación más reciente
│   └── monitorear_entrada.bat      # Observa archivos nuevos
├── docs/
│   ├── REGLAS_NEGOCIO_RRHH.md      # Manual funcional
│   ├── GUIA_TECNICA.md             # Rutas, operación y limitaciones
│   └── PRESENTACION_ACADEMICA.md   # Resumen profesional y académico
├── tests/
│   └── test_reglas.py              # Pruebas con datos ficticios
└── data/
    ├── entrada/                    # Archivo original descargado
    ├── salidas_fichadas/           # Excel y CSV producidos
    └── historico_originales/       # Copias auditables de originales
```

La carpeta `data/` contiene directorios vacíos versionados con `.gitkeep`; los archivos de personal y los reportes están **excluidos por `.gitignore`**.

## Instalación en Windows

1. Instalar Python 3.10 o superior y Git.
2. Clonar el proyecto y entrar en su carpeta:

   ```powershell
   git clone https://github.com/FrancoCorrea90/automatizacion_fichadas.git
   cd automatizacion_fichadas
   ```
3. Ejecutar `scripts\instalar.bat` con doble clic (o desde Terminal). Genera `.venv`, instala dependencias y prepara carpetas.
4. Copiar **un archivo por quincena** a `data\entrada`.
5. Ejecutar `scripts\procesar_ultimo.bat` para procesar el archivo de fecha de modificación más reciente.

**Alternativa por terminal** (sin activar entorno virtual):

```powershell
.\.venv\Scripts\python.exe automatizar_fichadas.py --archivo ".\data\entrada\1_QUINCENA_072026.xls"
```

**Modo monitor:** `scripts\monitorear_entrada.bat`. Este modo vigila **archivos nuevos** mientras la ventana está abierta; no reevalúa automáticamente archivos ya existentes ni garantiza que una copia lenta haya terminado.

## Formato de entrada

Se identifican por nombre las columnas `AC Nº`, `Nombre`, `Día/Hora` y `Estado` (`Entrada` / `Salida`). Se toleran columnas adicionales, como `Unnamed`, y encabezados precedidos por filas de información. Fechas con formato día/mes/año. Un Excel con fecha y hora separadas requiere adaptación.

## Reportes

Cada quincena produce una carpeta `data/salidas_fichadas/fichadas_AAAA-MM_Q1` o `Q2`, con un Excel consolidado y seis CSV. El Excel incluye:

| Hoja | Uso por RRHH |
|---|---|
| **Resumen** | Horas por persona, turnos, promedio, máximo y alertas |
| **Turnos** | Entrada/salida reales y computables, horas y estado de control |
| **Errores** | Fichadas y turnos excluidos; revisar antes de liquidar |
| **Duplicados** | Marcas descartadas por repetición cercana |
| **Correcciones** | Marcas excluidas por corrección inmediata |
| **Normalizado** | Base de fichadas leída, previa al filtrado de duplicados |

La columna `turnos_a_revisar` cuenta turnos computados cuyo `estado_control` es distinto de `OK`. No equivale a horas excluidas. Los turnos críticos no computados se registran en `Errores`.

## Validación y calidad

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe -m py_compile automatizar_fichadas.py
```

Las pruebas automatizadas cubren casos simulados. **Antes de usar el resultado en liquidaciones reales**, contrastar una quincena completa con el control manual de RRHH, especialmente salidas iniciales, cortes, turnos nocturnos y correcciones.

## Alcance y límites

- Un archivo debe contener un solo período quincenal; el sistema rechaza archivos con fichadas de distintas quincenas.
- El corte a las 23:59:59 puede generar una diferencia de un segundo respecto de sistemas que usan las 00:00 como extremo exclusivo.
- La parte posterior de un turno que cruza de período no se recrea automáticamente en este archivo: se requiere la exportación y validación de la siguiente quincena.
- El resumen agrupa solo empleados con turnos calculados: revisar la hoja **Errores** para los demás.
- No aplica por sí mismo convenios colectivos, feriados, nocturnidad, horas extras ni salarios.
- Al reprocesar la misma quincena, los reportes se actualizan; se conserva otra copia del original en el histórico.

## Privacidad

No subir planillas reales de asistencia, nombres, legajos ni reportes con información personal a un repositorio público. `.gitignore` previene nuevas incorporaciones, pero **no borra archivos que ya hayan sido versionados**. Revisar el historial de Git antes de publicar.

