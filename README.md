# Verificador de Liquidaciones CENASE v8.3 – Control legal de períodos

Cambios principales:
- Fecha de ingreso: se toma del ciclo correcto de la BDD de Personal.
- Fecha de salida: manda la fecha indicada en la liquidación de RR.HH.
- Días IESS: informativos; no generan observación por sí solos.
- Base IESS: se ajusta a los días correctos según ingreso BDD + salida RR.HH.
- Tolerancia monetaria: diferencias hasta USD 1,50 pasan como OK.
- Décimo tercero: valida el período legal 1 de diciembre a 30 de noviembre, recortado por ingreso/salida. La APP indica PERÍODO COMPLETO o PROPORCIONAL.
- Décimo cuarto: valida el período regional (Costa/Insular: 1-mar a último día de febrero; Sierra/Amazonía: 1-ago a 31-jul), recortado por ingreso/salida.
- Vacaciones: usa ciclos individuales por aniversario de la fecha real de ingreso. La APP indica AÑO COMPLETO o PROPORCIONAL AL CESE.
- La lista de meses usada por RR.HH. se compara por separado contra los meses que legalmente corresponden a D13 y al ciclo vacacional vigente.
- Ciclos completos de vacaciones anteriores se muestran como informativos para validar si ya fueron gozados/pagados; no se suman automáticamente para evitar duplicidad.

## Archivos para Streamlit
Subir a la raíz del repositorio:
- `main.py`
- `engine.py`
- `requirements.txt`

## Archivos de entrada
1. Liquidaciones RR.HH. masivas.
2. `BDD de personal.xlsx`.
3. `IESS BASE.xlsx`.
