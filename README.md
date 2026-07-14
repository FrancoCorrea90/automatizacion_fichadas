# Automatización del procesamiento de fichadas laborales con Python

## Descripción del proyecto

Desarrollé una solución en Python para automatizar el procesamiento de fichadas laborales descargadas desde un software de reloj biométrico/en la nube. El objetivo del proyecto fue transformar un proceso manual de control horario en un flujo automatizado, trazable y escalable para el área de Recursos Humanos.

El sistema permite leer archivos Excel generados por el reloj de asistencia, normalizar los registros de entrada y salida, detectar inconsistencias operativas y calcular las horas trabajadas por empleado para cada período quincenal.

## Problema detectado

El proceso original requería que el área de RRHH descargara manualmente los archivos de fichadas y luego realizara controles sobre los registros de cada empleado. Durante este procesamiento podían presentarse errores frecuentes, como:

* Olvido de registrar entrada o salida.
* Fichadas duplicadas en pocos segundos o minutos.
* Salidas sin entrada previa.
* Entradas sin salida posterior.
* Jornadas demasiado cortas.
* Jornadas excesivamente largas.
* Turnos nocturnos que cruzan el cierre de quincena o fin de mes.

Estos casos generaban riesgo de errores en la liquidación de horas y demandaban tiempo de revisión manual.

## Solución implementada

Se desarrolló un pipeline automatizado en Python que procesa los archivos Excel de fichadas y genera un reporte consolidado por quincena. El sistema realiza las siguientes etapas:

1. Lectura del archivo Excel descargado desde el reloj.
2. Detección automática de columnas relevantes: legajo, nombre, fecha/hora y estado.
3. Normalización de fichadas de entrada y salida.
4. Ordenamiento cronológico por empleado.
5. Detección de fichadas duplicadas cercanas.
6. Armado de turnos válidos mediante pares Entrada → Salida.
7. Detección de errores operativos.
8. Cómputo de horas trabajadas por empleado.
9. Aplicación de reglas especiales de corte quincenal.
10. Generación de reportes separados por período.
11. Resguardo del archivo original para trazabilidad.

## Reglas de negocio aplicadas

El sistema contempla criterios específicos para el control horario:

* Los duplicados cercanos se excluyen del cálculo de horas.
* Las fichadas erróneas se identifican y se excluyen del cómputo.
* Los turnos válidos se computan en la hoja de resumen.
* Las jornadas muy cortas o muy largas se marcan para revisión.
* Los turnos que superan el promedio habitual del empleado se identifican como casos a revisar.
* Cuando un empleado trabaja durante el día 15 y su jornada continúa luego de las 00:00 del día 16, el sistema computa únicamente las horas hasta las 23:59:59 del día 15.
* La misma lógica se aplica para el último día del mes, computando solo hasta las 23:59:59 del cierre de la segunda quincena.

Esta regla permite que los turnos nocturnos no sean excluidos injustamente, sino computados correctamente hasta el límite de la quincena correspondiente.

## Reportes generados

El sistema genera un libro Excel con distintas hojas de análisis:

* **Resumen:** total de horas por empleado, cantidad de turnos, promedio de horas, máximos, errores y turnos a revisar.
* **Turnos:** detalle de cada turno computado, incluyendo entrada, salida real, salida computable, horas reales y horas liquidadas.
* **Errores:** fichadas inconsistentes excluidas del cálculo.
* **Duplicados:** registros repetidos detectados y excluidos.
* **Normalizado:** base limpia de fichadas procesadas.

Además, se generan archivos CSV independientes para facilitar auditoría, revisión o integración futura con otros sistemas.

## Resultado obtenido

La automatización permite reducir el trabajo manual del área de RRHH, mejorar la trazabilidad del proceso de liquidación horaria y detectar rápidamente inconsistencias en los registros de asistencia.

El proyecto deja una estructura escalable para futuras liquidaciones, permitiendo procesar nuevos archivos quincenales sin modificar el código base.

## Tecnologías utilizadas

* Python
* Pandas
* OpenPyXL
* xlrd
* pathlib
* argparse
* Automatización de lectura y procesamiento de archivos Excel
* Exportación de reportes en Excel y CSV

## Valor agregado del proyecto

Este desarrollo convierte un proceso administrativo manual en una herramienta automatizada, parametrizable y auditable. Permite mejorar el control interno, reducir errores de liquidación, ahorrar tiempo operativo y dejar evidencia clara de qué fichadas fueron computadas, excluidas o marcadas para revisión.

