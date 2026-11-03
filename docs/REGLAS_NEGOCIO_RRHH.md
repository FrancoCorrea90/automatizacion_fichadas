# Guía funcional para Recursos Humanos

**Propósito:** explicar qué computa el programa y qué debe verificar el personal de RRHH antes de aceptar los totales de la quincena.

## Cómo se construyen las horas

Un turno comienza con `Entrada` y termina con `Salida` del mismo legajo, ordenadas cronológicamente. `horas` y `horas_computables` corresponden al tiempo computado para el período; `horas_reales` solo está completo cuando existen ambas marcas reales. El programa no descuenta automáticamente pausas o refrigerios.

## Tabla de reglas

| N.º | Situación detectada | Acción del sistema | ¿Se suma a `total_horas`? |
|---|---|---|---|
| 01 | Entrada seguida de Salida, duración ≤14 h | Forma turno y evalúa otras alertas | Sí |
| 02 | Duración >14 h y ≤17 h | Marca `REVISAR: jornada larga` | Sí |
| 03 | Duración >17 h entre Entrada y Salida | Marca `Turno crítico no computado` en Errores | No |
| 04 | Mismo estado repetido dentro de 2 minutos | Descarta repetición y la informa en Duplicados | Solo marca conservada |
| 05 | Salida sin entrada abierta seguida de Entrada dentro de 5 minutos | Descarta Salida presuntamente errónea | No para la Salida descartada |
| 06 | Entrada con otra entrada abierta y >17 h de distancia | Informa entrada anterior como antigua y toma la nueva | Solo turnos válidos posteriores |
| 07 | Entrada repetida con una entrada abierta ≤17 h | Señala error y conserva la primera entrada | Se suma el eventual turno válido |
| 08 | Salida sin entrada previa | Informa en Errores | No, salvo regla 09 |
| 09 | Primera marca de un legajo: Salida del día 1/16 a más tardar a las 08:00 | Infere Entrada computable desde 00:00 del período | Sí, revisar inferencia |
| 10 | Entrada abierta el día 15 o último día del mes sin Salida | Infere Salida a las 23:59:59, siempre que el tramo sea ≤17 h | Sí, revisar inferencia |
| 11 | Turno <10 minutos | Marca `ERROR: jornada demasiado corta` | Sí, **a revisar** |
| 12 | Turno >30 % del promedio del mismo empleado en el archivo | Marca alerta de duración atípica | Sí |

**Importante:** el umbral crítico de 17 h se comprueba sobre el turno real antes de aplicar el recorte de cierre. Una inferencia por corte se controla por sus horas hasta las 23:59:59. Las reglas automáticas pueden identificar candidatos, pero no demuestran que la presencia haya ocurrido.

## Ejemplos prácticos

**Jornada normal:** Entrada 06:00, Salida 17:00 → 11 h computadas.  
**Jornada extensa:** Entrada 06:00, Salida 21:30 → 15 h 30 min computados; queda `REVISAR`.  
**Jornada crítica:** Entrada lunes 06:00, Salida martes 03:30 → 21 h 30 min; se excluye y se informa en Errores.  
**Entrada vieja:** Entrada 07/07 16:55; nueva Entrada 13/07 05:56; Salida 13/07 17:03 → se descarta la primera Entrada y se forma un turno de aproximadamente 11 h el día 13.  
**Inicio de período:** primera marca 01/07 05:52 `Salida` → se computa desde 01/07 00:00 hasta 05:52 (inferencia).  
**Corte de período:** Entrada 15/07 19:00, sin Salida en ese archivo → se computa hasta 15/07 23:59:59 (inferencia).

## Qué debe revisar RRHH

1. Abrir `Resumen` y ordenar por `estado_final` para localizar `REVISAR`.
2. Buscar a cada empleado con incidencias en las hojas `Turnos`, `Errores` y `Correcciones`.
3. Confirmar turnos superiores a 14 horas, turnos muy cortos y los que superan el promedio.
4. Revisar especialmente cortes a las 23:59:59 y salidas iniciales inferidas.
5. Contrastar el total con registro manual, planificación de turnos y novedades del período.
6. Documentar los ajustes autorizados fuera de esta herramienta antes de realizar la liquidación definitiva.

## Interpretación de indicadores

- `total_horas`: suma de turnos computados; **no** incluye casos excluidos en `Errores`.
- `cantidad_turnos`: número de turnos incluidos.
- `max_horas_turno`: máximo de horas entre turnos incluidos.
- `turnos_a_revisar`: turnos incluidos cuyo control no es `OK` (las horas continúan sumadas).
- `errores_fichada`: cantidad de incidencias excluidas (no equivale a horas no trabajadas).
- `duplicados_detectados`: marcas repetidas descartadas.
- `estado_final`: `REVISAR` si tiene errores o turnos marcados; de lo contrario, `OK`.

## Límites y responsabilidad

Se trata de un sistema de apoyo, no de una decisión automática sobre derechos laborales. Las reglas de 14 y 17 h son umbrales configurados para detectar anomalías, no normas legales. Un turno excluido podría ser real; RRHH deberá investigarlo y efectuar la corrección autorizada. Un turno computado puede requerir ajuste por descansos, permisos, novedades, convenio o error de reloj.
