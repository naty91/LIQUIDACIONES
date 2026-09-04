# Verificador Integral de Liquidaciones CENASE v5

## Flujo
1. Subir **Liquidaciones RR.HH.** (archivo masivo actual de CENASE).
2. Subir **Base de Personal** y mapear cédula/nombre, fecha real de ingreso y salario básico mensual.
3. Subir **reporte IESS** y mapear cédula/nombre, período, días, materia gravada, aporte personal y aporte patronal.
4. La APP realiza tres controles:
   - RR.HH. vs APP (fechas, días, décimo tercero, vacaciones, desahucio y total del formato cargado).
   - APP vs IESS (días, materia gravada y aportes 9,45% / 11,15%).
   - RR.HH./IESS vs Base de Personal (fecha real de ingreso y salario básico).

## Regla de días
- Mes completo = 30 días.
- Febrero completo = 30 días aunque el calendario tenga 28/29.
- Mes incompleto: sueldo fijo proporcional = salario básico mensual / 30 × días aplicables.

## IESS
Para trabajador privado bajo relación de dependencia: 9,45% aporte personal + 11,15% patronal = 20,60%.
La materia gravada puede incluir rubros adicionales al salario fijo; por eso se muestra separada del sueldo proporcional básico.

## Streamlit
```bash
pip install -r requirements.txt
streamlit run main.py
```
