# Verificador de Liquidaciones CENASE v5.1

Corrección del ImportError de v5.0.

## Archivos para Streamlit
Sube **estos tres archivos a la raíz del repositorio**:
- `main.py`
- `engine.py`
- `requirements.txt`

No los pongas dentro de otra carpeta si tu archivo principal configurado en Streamlit es `main.py`.

## Controles
1. RR.HH. vs APP.
2. APP vs IESS.
3. Datos vs Base de Personal.

Criterio de días: mes completo = 30 días para nómina, incluido febrero. Para mes incompleto, sueldo fijo proporcional = salario básico mensual / 30 × días aplicables.
