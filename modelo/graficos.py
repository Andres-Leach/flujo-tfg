"""Gráficos del TFG a partir de resultados.json y datos/*.csv."""
import csv
import json
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

RAIZ = Path(__file__).resolve().parent.parent
FIG = RAIZ / "figuras"
FIG.mkdir(exist_ok=True)
R = json.loads((RAIZ / "modelo" / "resultados.json").read_text())

AZUL, NARANJA, AGUA, GRIS = "#2a78d6", "#eb6834", "#1baf7a", "#8a8984"
TINTA, TINTA2 = "#0b0b0b", "#52514e"
plt.rcParams.update({
    "font.family": "DejaVu Serif", "font.size": 10, "axes.edgecolor": "#c9c8c2",
    "axes.labelcolor": TINTA2, "xtick.color": TINTA2, "ytick.color": TINTA2,
    "axes.spines.top": False, "axes.spines.right": False, "axes.grid": True,
    "grid.color": "#ecebe6", "grid.linewidth": 0.6, "axes.axisbelow": True,
    "figure.dpi": 200, "savefig.bbox": "tight",
})
eur = matplotlib.ticker.FuncFormatter(lambda x, _: f"{x:,.0f} €".replace(",", "."))


def guardar(fig, nombre):
    fig.savefig(FIG / nombre)
    plt.close(fig)


def leer(nombre):
    with open(RAIZ / "datos" / nombre, encoding="utf-8") as f:
        return list(csv.DictReader(f))


# --- Anexo: datos de mercado --------------------------------------------------
edad = leer("reta_por_edad_2025.csv")
fig, ax = plt.subplots(figsize=(6, 3))
ax.barh([r["tramo"] for r in edad][::-1], [float(r["porcentaje"]) for r in edad][::-1],
        color=[GRIS, GRIS, AZUL], height=0.55)
for i, r in enumerate(edad[::-1]):
    ax.text(float(r["porcentaje"]) + 1, i, f'{float(r["porcentaje"]):.1f} %'.replace(".", ","),
            va="center", color=TINTA)
ax.set_xlim(0, 60)
ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, _: f"{x:.0f} %"))
ax.grid(axis="y", visible=False)
guardar(fig, "fig_anexo_reta_edad.png")

big3 = leer("big3_por_edad_2021.csv")
fig, ax = plt.subplots(figsize=(6, 3))
colores = [AZUL if r["grupo"] == "18-34" else (TINTA2 if "Media" in r["grupo"] else GRIS) for r in big3]
ax.bar([r["grupo"] for r in big3], [float(r["big3"]) for r in big3], color=colores, width=0.55)
for i, r in enumerate(big3):
    ax.text(i, float(r["big3"]) + 1, f'{r["big3"]} %', ha="center", color=TINTA)
ax.set_ylim(0, 65)
ax.set_ylabel("Aciertos medios (%)")
ax.grid(axis="x", visible=False)
guardar(fig, "fig_anexo_big3.png")

crec = leer("crecimiento_reta_2025.csv")
fig, ax = plt.subplots(figsize=(6, 2.4))
izq = 0
for r, c in zip(crec, [NARANJA, "#f3a582", GRIS]):
    v = float(r["porcentaje"])
    ax.barh([0], [v], left=izq, color=c, height=0.5, edgecolor="white", linewidth=2)
    ax.text(izq + v / 2, 0, f'{r["grupo"]}\n{v:.1f} %'.replace(".", ","), ha="center", va="center",
            color=TINTA, fontsize=9)
    izq += v
ax.set_xlim(0, 100)
ax.set_yticks([])
ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, _: f"{x:.0f} %"))
ax.grid(False)
guardar(fig, "fig_anexo_crecimiento.png")

# --- Plan financiero ---------------------------------------------------------
mb = R["base"]["mensual"]
meses = range(1, 37)
fig, ax = plt.subplots(figsize=(6.2, 3))
for n, c in [("pesimista", NARANJA), ("base", AZUL), ("optimista", AGUA)]:
    ax.plot(meses, R[n]["mensual"]["activos"], color=c, lw=2, label=n.capitalize())
ax.axhline(45, color=TINTA2, lw=1, ls="--")
ax.text(1, 46, "Límite de capacidad a jornada completa (≈45 clientes)", color=TINTA2, fontsize=8)
ax.set_xlabel("Mes")
ax.set_ylabel("Clientes de suscripción activos")
ax.set_xticks([1, 6, 12, 18, 24, 30, 36])
ax.set_ylim(0, 52)
ax.legend(frameon=False, loc="lower right")
guardar(fig, "fig_clientes.png")

A = R["base"]["anual"]
lineas = [("Servicio: setup", "ing_setup", "#9cc3ef"), ("Servicio: suscripción", "ing_suscripcion", AZUL),
          ("Plantillas y cursos", None, NARANJA), ("Herramienta digital", "ing_app", AGUA)]
fig, ax = plt.subplots(figsize=(6, 3.2))
abajo = [0, 0, 0]
for nombre, k, c in lineas:
    v = [A["ing_plantillas"][a] + A["ing_cursos"][a] for a in range(3)] if k is None else A[k]
    ax.bar(["Año 1", "Año 2", "Año 3"], v, bottom=abajo, color=c, width=0.5, label=nombre,
           edgecolor="white", linewidth=1.5)
    abajo = [abajo[i] + v[i] for i in range(3)]
for i, t in enumerate(abajo):
    ax.text(i, t + 900, f"{t:,.0f} €".replace(",", "."), ha="center", color=TINTA)
ax.yaxis.set_major_formatter(eur)
ax.set_ylim(0, 66000)
ax.grid(axis="x", visible=False)
ax.legend(frameon=False, fontsize=8, loc="upper left")
guardar(fig, "fig_ingresos.png")

fig, ax = plt.subplots(figsize=(6.2, 3))
for n, c in [("pesimista", NARANJA), ("base", AZUL), ("optimista", AGUA)]:
    ax.plot(meses, R[n]["mensual"]["caja"], color=c, lw=2, label=n.capitalize())
ax.axhline(0, color=TINTA2, lw=1)
ax.yaxis.set_major_formatter(eur)
ax.set_xlabel("Mes")
ax.set_xticks([1, 6, 12, 18, 24, 30, 36])
ax.legend(frameon=False, loc="upper left")
guardar(fig, "fig_tesoreria.png")

S = R["sensibilidad"]
fig, axs = plt.subplots(1, 2, figsize=(6.4, 2.8), sharey=True)
p = S["precio"]
axs[0].plot([int(k) for k in p], list(p.values()), color=AZUL, lw=2, marker="o", ms=5)
axs[0].set_xlabel("Precio de la suscripción (€/mes)")
f = S["altas"]
axs[1].plot([float(k) * 100 for k in f], list(f.values()), color=AZUL, lw=2, marker="o", ms=5)
axs[1].set_xlabel("Captación sobre el escenario base (%)")
axs[0].yaxis.set_major_formatter(eur)
axs[0].set_ylabel("VAN a 3 años")
for ax in axs:
    ax.set_ylim(0, 50000)
guardar(fig, "fig_sensibilidad.png")
print("gráficos generados en", FIG)
