# Verificador Liquidaciones CENASE v7

Archivos en la raíz de GitHub: `main.py`, `engine.py`, `requirements.txt`.

Carga 3 archivos:
1. Liquidaciones masivas RR.HH.
2. BDD de personal.xlsx (ACTIVOS/REINGRESOS/INACTIVOS)
3. IESS BASE.xlsx (hoja IESS, encabezados en fila 2)

Cambios v7:
- IESS se reconoce sin mapeo manual.
- Cruce por cédula + período, nombre como respaldo.
- Fecha real de ingreso/ciclo desde BDD Personal.
- Mes completo = 30 días, incluido febrero.
- Vacaciones por aniversario de ingreso; muestra años completos y proporcional vigente.
- Décimo tercero por período legal 1-dic/30-nov usando las bases IESS del período.
- Décimo cuarto por período regional y SBU.
- Fondos de reserva desde el primer aniversario como control general.
- Aportes IESS: 9,45% personal y 11,15% patronal para relación privada ordinaria.

Nota: los ciclos completos de vacaciones deben contrastarse con registro de vacaciones gozadas/pagadas para determinar saldo pendiente real. La app no inventa ese dato.
