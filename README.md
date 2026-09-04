# Verificador de Liquidaciones CENASE v8.7

## Regla de días corregida
- Para nómina, un mes completo siempre equivale a 30 días, incluso febrero (28/29) y meses de 31 días.
- La fecha de ingreso de la BDD y la fecha de salida de RR.HH. determinan el **máximo** de días posibles del mes.
- Si RR.HH./IESS reportan menos días por trabajo efectivo, faltas u otra novedad, **no se completan ni se inflan**.
- La base IESS se conserva si sus días son menores o iguales al máximo permitido.
- Solo se recorta la base IESS cuando IESS reporta más días que los posibles según la salida de RR.HH.

### Ejemplos validados
- MORAN ITURRALDE JAEL BYRON: enero 14 días / $238,22 y febrero 17 días / $290,13 se conservan. Vacaciones: $3.665,53 / 24 = $152,73.
- REYES CIRINO JORGE ALFREDO: agosto IESS $482 / 30 días, salida 27/08 -> base utilizable $433,80.

## Criterios vigentes
- Décimo tercero de guardias: informativo; no genera REVISAR.
- Vacaciones: se revisan por ciclo desde fecha de ingreso BDD hasta salida RR.HH.
- Diferencias de vacaciones y décimos hasta USD 1,50 inclusive = CORRECTO.
- Días IESS son informativos.
- Salida RR.HH. manda; una diferencia con la salida BDD se informa, pero no bloquea.
