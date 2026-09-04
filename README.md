# Verificador de Liquidaciones CENASE v8.1

Versión masiva para cargar Liquidaciones RR.HH., BDD de Personal e IESS BASE.

Cambios de control:
- Tolerancia monetaria: diferencias absolutas de hasta USD 1,50 se consideran aceptables y NO generan revisión.
- Diferencias mayores a USD 1,50 se muestran en una sección independiente de observaciones monetarias.
- El control de días se realiza únicamente entre RR.HH. y las fechas/ciclo laboral de la BDD de Personal.
- Los días del IESS son informativos y NO generan observaciones de días ni afectan el estado por días.
- Mes completo = 30 días, incluido febrero; meses parciales se calculan sobre base 30.
- Exporta Excel masivo y PDF con observaciones monetarias y revisión de días separadas.

Para Streamlit coloque `main.py`, `engine.py` y `requirements.txt` en la raíz del repositorio.
