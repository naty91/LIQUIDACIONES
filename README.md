# Verificador de Liquidaciones CENASE v8.2

Ajuste principal de esta versión:

- La fecha de ingreso se toma de la BDD de Personal.
- La fecha de salida que manda para días y beneficios es la indicada en la liquidación de RR.HH.
- Los días IESS son informativos y no generan observación por sí solos.
- Para la base legal usada en décimo tercero, vacaciones y otros cálculos basados en IESS, la APP conserva el valor diario implícito del IESS pero reemplaza los días por los días correctos del vínculo laboral.
- Ejemplo: IESS reporta $482 por 30 días, pero la liquidación indica salida el día 27: $482 / 30 x 27 = $433,80 de base IESS ajustada.
- La comparación RR.HH. vs IESS se realiza contra la BASE IESS AJUSTADA, no contra la base mensual cruda cuando los días reportados por IESS no coinciden con el período correcto.
- Tolerancia monetaria: diferencias de hasta USD 1,50 pasan como OK.
- La revisión de días se hace contra ingreso BDD + salida RR.HH.; no contra días IESS.

## Archivos para Streamlit
Subir a la raíz del repositorio:
- main.py
- engine.py
- requirements.txt

