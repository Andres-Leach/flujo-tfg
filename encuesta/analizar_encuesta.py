"""Analiza las respuestas de la encuesta (CSV exportado de Google Forms).

Uso:  python analizar_encuesta.py respuestas.csv

En Google Forms: pestaña "Respuestas" > icono de Hojas de cálculo > en la hoja,
Archivo > Descargar > Valores separados por comas (.csv).

Genera resultados_encuesta.json (para el documento), figuras en ../figuras y un
resumen por pantalla. Solo cuenta las respuestas válidas: consentimiento, 20-35
años y trabajo por cuenta propia.
"""
import csv
import json
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

RAIZ = Path(__file__).resolve().parent.parent
EDADES_VALIDAS = {"De 20 a 24 años", "De 25 a 29 años", "De 30 a 35 años"}


def columna(fila, prefijo):
    """Valor de la pregunta cuyo título empieza por 'P7.' (por ejemplo)."""
    for k, v in fila.items():
        if k.strip().startswith(prefijo + "."):
            return (v or "").strip()
    return ""


def numero(x):
    try:
        return float(x.replace(",", ".").replace("€", "").strip())
    except ValueError:
        return None


def pct(n, total):
    return round(100 * n / total, 1) if total else 0.0


def van_westendorp(muy_barato, barato, caro, muy_caro):
    """Puntos de cruce del medidor de sensibilidad al precio de Van Westendorp."""
    precios = [p / 2 for p in range(0, 2 * int(max(muy_caro + caro + [1])) + 2)]
    n = len(muy_barato)
    curva = lambda xs, p, desc: sum((x >= p) if desc else (x <= p) for x in xs) / n
    mb = [curva(muy_barato, p, True) for p in precios]   # demasiado barato (decreciente)
    b = [curva(barato, p, True) for p in precios]        # barato
    c = [curva(caro, p, False) for p in precios]         # caro (creciente)
    mc = [curva(muy_caro, p, False) for p in precios]    # demasiado caro

    def cruce(f, g):
        for i in range(1, len(precios)):
            if (f[i - 1] - g[i - 1]) * (f[i] - g[i]) <= 0:
                return precios[i]
        return None

    no_barato = [1 - x for x in b]
    no_caro = [1 - x for x in c]
    return dict(
        precio_minimo=cruce(mb, no_barato),          # PMC: punto de barato marginal
        precio_maximo=cruce(mc, no_caro),            # PME: punto de caro marginal
        precio_optimo=cruce(mb, mc),                 # OPP
        precio_indiferencia=cruce(b, c),             # IPP
        curvas=dict(precios=precios, muy_barato=mb, no_barato=no_barato, no_caro=no_caro, muy_caro=mc),
    )


def main(ruta):
    with open(ruta, encoding="utf-8-sig") as f:
        filas = list(csv.DictReader(f))
    validas = [r for r in filas
               if columna(r, "P0").startswith("Sí")
               and columna(r, "P1") in EDADES_VALIDAS
               and columna(r, "P2") and not columna(r, "P2").startswith("No trabajo")]
    n = len(validas)
    if n == 0:
        sys.exit("No hay respuestas válidas en el archivo.")

    cuenta = lambda pre, cond: sum(cond(columna(r, pre)) for r in validas)
    res = dict(
        respuestas_totales=len(filas), respuestas_validas=n,
        no_sabe_neto_exacto=pct(cuenta("P7", lambda v: v != "Sí, con exactitud"), n),
        no_sabe_neto=pct(cuenta("P7", lambda v: v == "No"), n),
        no_separa=pct(cuenta("P8", lambda v: v in ("No", "A medias")), n),
        no_reserva_siempre=pct(cuenta("P11", lambda v: v != "Sí, siempre"), n),
        sorpresa_hacienda=pct(cuenta("P12", lambda v: v == "Sí"), n),
        tiene_gestoria=pct(cuenta("P10", lambda v: v.startswith("Sí")), n),
        interes_alto=pct(cuenta("P15", lambda v: v in ("4", "5")), n),
        contrataria_75=pct(cuenta("P21", lambda v: v in ("Seguro que sí", "Probablemente sí")), n),
        capacidad_media=round(sum(numero(columna(r, "P13")) or 0 for r in validas) / n, 2),
        estres_medio=round(sum(numero(columna(r, "P14")) or 0 for r in validas) / n, 2),
    )
    for pre, clave in [("P3", "actividad"), ("P4", "antiguedad"), ("P6", "ubicacion"), ("P10", "gestoria")]:
        dist = {}
        for r in validas:
            dist[columna(r, pre) or "Sin respuesta"] = dist.get(columna(r, pre) or "Sin respuesta", 0) + 1
        res[clave] = dist
    formatos = {}   # Google Forms separa las casillas marcadas con ", " en el CSV
    for r in validas:
        for op in filter(None, (x.strip() for x in columna(r, "P16").split(", "))):
            formatos[op] = formatos.get(op, 0) + 1
    res["formatos"] = formatos

    precios = [[numero(columna(r, p)) for p in ("P17", "P18", "P19", "P20")] for r in validas]
    coherentes = [x for x in precios if None not in x and x[0] <= x[1] <= x[2] <= x[3]]
    res["respuestas_precio_coherentes"] = len(coherentes)
    if len(coherentes) >= 5:
        vw = van_westendorp(*[[x[i] for x in coherentes] for i in range(4)])
        res["van_westendorp"] = {k: v for k, v in vw.items() if k != "curvas"}
        c = vw["curvas"]
        fig, ax = plt.subplots(figsize=(6.2, 3.2))
        for clave, color, etiqueta in [("muy_barato", "#1baf7a", "Demasiado barato"),
                                       ("no_barato", "#9cc3ef", "No es barato"),
                                       ("no_caro", "#f3a582", "No es caro"),
                                       ("muy_caro", "#eb6834", "Demasiado caro")]:
            ax.plot(c["precios"], c[clave], color=color, lw=2, label=etiqueta)
        ax.axvline(75, color="#52514e", ls="--", lw=1)
        ax.text(76, 0.95, "Precio del plan (75 €)", fontsize=8, color="#52514e")
        ax.set_xlim(0, max(150, vw["precio_maximo"] or 0))
        ax.set_xlabel("Precio mensual (€)")
        ax.set_ylabel("Proporción de encuestados")
        ax.spines[["top", "right"]].set_visible(False)
        ax.legend(frameon=False, fontsize=8, loc="center right")
        fig.savefig(RAIZ / "figuras" / "fig_van_westendorp.png", dpi=200, bbox_inches="tight")

    salida = RAIZ / "encuesta" / "resultados_encuesta.json"
    salida.write_text(json.dumps(res, indent=1, ensure_ascii=False))
    print(json.dumps(res, indent=1, ensure_ascii=False))
    print("guardado", salida)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
