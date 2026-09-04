# Verificador Integral de Liquidaciones CENASE v6

## Cambios principales
- Lee automáticamente la BDD real de CENASE: hojas **ACTIVOS**, **REINGRESOS** e **INACTIVOS**.
- Prioriza cruce por cédula; usa nombre como respaldo.
- En reingresos selecciona el ciclo laboral que corresponde a la fecha de la liquidación, no el primer ingreso histórico.
- Muestra estado BDD, ciclo, ingreso, salida y motivo de salida.
- Febrero completo = 30 días para control de nómina.
- Mes incompleto: sueldo fijo proporcional = salario básico mensual / 30 × días.
- Como la BDD suministrada no contiene una columna de sueldo, el salario básico de control queda configurable (por defecto SBU 2026 = USD 482).
- Mantiene el control RR.HH. vs APP y APP vs IESS.

## Streamlit
Subir a la raíz del repositorio:
- main.py
- engine.py
- requirements.txt
