# Guía técnica y rutas

## Rutas portables

El proyecto obtiene su ubicación desde `Path(__file__).resolve().parent`. Las salidas e históricos quedan dentro de la carpeta clonada, **no en una ruta personal fija**.

| Función | Ruta predeterminada |
|---|---|
| Archivo descargado por RRHH | `data/entrada/` |
| Reportes CSV/Excel | `data/salidas_fichadas/` |
| Copias de originales | `data/historico_originales/` |
| Instalación de librerías | `.venv/` |
| Lanzadores | `scripts/*.bat` |

Las tres rutas pueden cambiarse en `ConfigFichadas`, pero es recomendable conservar el diseño relativo o configurar carpetas externas de acceso restringido. **Nunca publicar rutas de red privadas ni datos personales**.

## Archivos .bat

- **`instalar.bat`**: ubica la raíz del proyecto con `%~dp0..`, crea el entorno virtual Python, instala librerías y prepara carpetas. Se ejecuta una vez por equipo o después de cambios de dependencias.
- **`procesar_ultimo.bat`**: obtiene el último `.xls/.xlsx` por **fecha de modificación**, no por nombre de quincena, y llama al script con `--archivo`. Al finalizar muestra el código de salida. Si se copiaron otros archivos después, puede no elegir el período esperado.
- **`monitorear_entrada.bat`**: inicia un proceso de observación que captura archivos nuevos mientras permanezca abierto. No es un servicio de Windows; verificar copias completas y permisos de red.

## Ejecución manual

```powershell
.\.venv\Scripts\python.exe automatizar_fichadas.py --archivo ".\data\entrada\archivo.xls"
```

## Controles técnicos recomendados

- Validar período único por archivo y encabezados `AC Nº`, `Nombre`, `Día/Hora`, `Estado`.
- Mantener resultados bajo control de acceso de RRHH y respaldos apropiados.
- No abrir el Excel de salida mientras se vuelve a generar: Windows podría bloquear la escritura.
- Registrar quién procesó/revisó el informe y la versión del código utilizada.
- Realizar conciliación por legajo contra una quincena validada manualmente antes de implementar en producción.
- En una mejora posterior: agregar un registro de ejecución con hash del origen, criterios de revisión, conciliación de tramos entre períodos y un estado explícito `pendiente_aprobacion`.

## Limitaciones conocidas

- Infiere algunas marcas: estas inferencias necesitan revisión de RRHH.
- No reconstruye automáticamente la porción de un turno nocturno perteneciente a la quincena siguiente; depende del siguiente export.
- El promedio se obtiene solo de este archivo, no de un historial de jornadas.
- El monitor puede detectar una copia antes de que termine; en tal caso reprocesar manualmente.
- No calcula remuneraciones, descuentos, recargos ni otras reglas de convenio.
