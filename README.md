# Verificador de Liquidaciones CENASE v8.4

Cambios principales:
- Décimo tercero de guardias tratado como **MENSUALIZADO**: se muestra diferencia RR.HH. vs APP/IESS, pero **no genera estado REVISAR**, ni por período ni por diferencia monetaria.
- Vacaciones: único beneficio del bloque que genera control de período en esta fase. Se determina **AÑO COMPLETO** o **PROPORCIONAL AL CESE** según fecha de ingreso real de BDD y fecha de salida de RR.HH.
- Se mantiene tolerancia monetaria de USD 1,50 para vacaciones y demás controles monetarios que sí aplican.
- Se mantiene fecha de salida RR.HH. como fecha efectiva para ajustar la base IESS del último mes.
- Días IESS continúan solo como informativos.

Subir `main.py`, `engine.py` y `requirements.txt` juntos a la raíz del repositorio de Streamlit.


## v8.5 — Regla de tolerancia
- Diferencia absoluta **hasta e incluyendo USD 1,50** en décimo tercero y vacaciones = **✅ CORRECTO**.
- Vacaciones: solo diferencia monetaria **> USD 1,50** o período vacacional incorrecto = **⚠️ REVISAR**.
- Décimo tercero mensualizado: si supera USD 1,50 se informa la diferencia, pero no cambia el dictamen a REVISAR.
