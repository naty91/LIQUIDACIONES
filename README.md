# Verificador de Liquidaciones CENASE v8.6

Corrección del estado masivo:
- Décimo tercero mensualizado: siempre informativo; nunca genera REVISAR.
- Vacaciones: diferencia absoluta <= USD 1,50 = CORRECTO; > USD 1,50 = REVISAR.
- El período vacacional sí genera REVISAR si no corresponde al ciclo según ingreso BDD + salida RR.HH.
- Diferencias mensuales de base RR.HH. vs IESS ajustada quedan informativas y no cambian el estado.
- Diferencias de aportes IESS quedan informativas y no cambian el estado.
- Días/fechas RR.HH. vs BDD Personal sí pueden generar revisión.
- Si no se encuentra al trabajador en IESS, se mantiene como revisión porque no existe base monetaria para validar vacaciones.
