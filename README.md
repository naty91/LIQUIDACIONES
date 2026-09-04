# Verificador Masivo de Liquidaciones CENASE v3

## Uso
1. Suba a Streamlit un archivo XLSX/XLSM con varias liquidaciones en una sola hoja, consecutivas.
2. Cada bloque debe contener: nombre, fecha ingreso, fecha salida, encabezado `mes / tercer decimo / dias / vacac`, detalle mensual, `a recibir`, `Vacaciones`, `Desahucio x N años`, `TOTAL A RECIBIR`.
3. La app detecta cada trabajador automáticamente y revisa uno por uno.
4. Descarga consolidado Excel y PDF con detalle individual.

## Ejecutar
```bash
pip install -r requirements.txt
streamlit run main.py
```

## Control automático con este archivo
- suma de remuneración computable
- suma de días
- décimo tercero = base / 12
- vacaciones = base / 24
- desahucio indicado = última remuneración x 25% x años escritos en el archivo
- total a recibir del bloque = vacaciones + desahucio indicado

Los rubros que no están presentes en la carga (causal legal, décimo cuarto, despido, fondos de reserva, sueldos pendientes, descuentos, etc.) requieren datos/soportes adicionales y no se aprueban por inferencia.
