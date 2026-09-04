# Verificador de Liquidaciones CENASE v2

Aplicación Streamlit para auditar liquidaciones laborales de CENASE contra un cálculo independiente basado en la normativa ecuatoriana y la guía oficial del Ministerio del Trabajo.

## Qué hace
- Lee directamente el archivo CENASE XLSX/XLSM cuando contiene la hoja `VERIFICADOR`.
- Detecta nombre, fechas, base mensual y valores de décimo tercero/vacaciones/desahucio cuando el formato lo permite.
- Recalcula décimo tercero, décimo cuarto, vacaciones, bonificación por desahucio e indemnización por despido intempestivo.
- Permite controlar remuneración pendiente, fondos de reserva y otros valores.
- Compara APP vs RR.HH. y marca: CORRECTO, DIFERENCIA, RUBRO OMITIDO o REVISAR APLICACIÓN.
- Emite dictamen `APTO PARA PAGO` o `REQUIERE CORRECCIÓN / REVISIÓN`.
- Descarga verificación en Excel y PDF.
- Descarga/carga respaldo JSON de la revisión.

## Referencia principal
https://calculadoras.trabajo.gob.ec/liquidaciones

## Subir a Streamlit Community Cloud
1. Cree un repositorio en GitHub.
2. Suba `main.py`, `engine.py` y `requirements.txt` en la raíz.
3. En Streamlit Community Cloud seleccione el repositorio y `main.py`.

## Ejecutar localmente
```bash
pip install -r requirements.txt
streamlit run main.py
```

## Criterios importantes
- El archivo de RR.HH. es un dato a verificar, no la fuente de verdad del cálculo.
- La causal documentada de terminación determina desahucio/indemnización.
- Vacaciones acumuladas, gozadas o previamente pagadas deben estar soportadas; el campo de ajuste permite reflejar esos casos sin alterar la base original.
- El verificador no sustituye el acta de finiquito del SUT.
