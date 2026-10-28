# Flujo: viabilidad financiera (TFG, ADE, Universidad de Zaragoza)

Código, datos y modelo financiero del Trabajo Fin de Grado *Viabilidad financiera de Flujo: una start-up de ayuda a jóvenes emprendedores* (Andrés Leach Sagarra, curso 2026-2027).

## Contenido

- `modelo/hipotesis.py`: todas las hipótesis del plan, con su justificación.
- `modelo/modelo.py`: modelo mensual a 36 meses (PyG, tesorería con IVA e IRPF, balance, VAN, TIR, payback, ratios, escenarios y sensibilidad). Comprueba que el balance cuadra.
- `modelo/graficos.py`: genera las figuras de `figuras/`.
- `modelo/exportar_excel.py`: vuelca los resultados a `Modelo_financiero_Flujo.xlsx`.
- `datos/`: datos de mercado usados en el Anexo I, con la fuente de cada cifra.
- `documento/`: script que genera el documento Word del TFG.

## Reproducir

```bash
cd modelo
python modelo.py
python graficos.py
python exportar_excel.py
cd ../documento && node build.js
```

Requisitos: Python 3 con `matplotlib` y `openpyxl`; Node.js con `docx`.
