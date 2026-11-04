"""Genera Modelo_financiero_Flujo.xlsx con fórmulas vivas.

El libro reproduce el modelo de modelo.py: cambiando cualquier celda azul de la hoja
Hipotesis (o el escenario y las medidas del panel de control) se recalculan la cuenta
de resultados, la tesorería, el balance, el VAN, la TIR y el payback.

Uso:  python excel_formulas.py [ruta.xlsx] [escenario] [palancas S/N x5]
"""
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

import hipotesis as H

RAIZ = Path(__file__).resolve().parent.parent
FUENTE = "Arial"
AZUL = Font(name=FUENTE, color="0000FF")
NEGRO = Font(name=FUENTE)
NEGRITA = Font(name=FUENTE, bold=True)
TITULO = Font(name=FUENTE, bold=True, size=14)
VERDE = Font(name=FUENTE, color="008000")
AMARILLO = PatternFill("solid", fgColor="FFFF00")
CABECERA = PatternFill("solid", fgColor="DCE6F2")
GRIS = PatternFill("solid", fgColor="F2F2F2")
EUR = '#,##0 €;(#,##0 €);"-"'
EUR2 = '#,##0.00 €;(#,##0.00 €);"-"'
PCT = '0.0%;(0.0%);"-"'
NUM = '#,##0.0;(#,##0.0);"-"'
MESES = 36
C0 = 3                      # primera columna de meses en la hoja Mensual (C)
COLS = [L(C0 + i) for i in range(MESES)]
ULT = COLS[-1]
ANIOS = ["C", "D", "E"]


def build(ruta, escenario="base", palancas=("No",) * 5):
    wb = Workbook()
    wb.remove(wb.active)
    ins = wb.create_sheet("Instrucciones")
    hip = wb.create_sheet("Hipotesis")
    men = wb.create_sheet("Mensual")
    anu = wb.create_sheet("Anual")
    val = wb.create_sheet("Valoracion")
    res = wb.create_sheet("Resumen")
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False

    # ------------------------------------------------------------------ Hipotesis
    N = {}                  # nombre -> referencia absoluta 'Hipotesis!$C$n'
    fila = [1]

    def titulo(texto):
        hip.cell(fila[0], 2, texto).font = NEGRITA
        hip.cell(fila[0], 2).fill = CABECERA
        for c in (3, 4):
            hip.cell(fila[0], c).fill = CABECERA
        fila[0] += 1

    def dato(clave, etiqueta, valor, fmt=None, nota="", clave_input=True, amarillo=False):
        r = fila[0]
        hip.cell(r, 2, etiqueta).font = NEGRO
        c = hip.cell(r, 3, valor)
        c.font = AZUL if clave_input and not (isinstance(valor, str) and valor.startswith("=")) else NEGRO
        if amarillo:
            c.fill = AMARILLO
        if fmt:
            c.number_format = fmt
        if nota:
            hip.cell(r, 4, nota).font = Font(name=FUENTE, italic=True, color="666666")
        N[clave] = f"Hipotesis!$C${r}"
        fila[0] += 1
        return r

    hip.cell(1, 2, "Hipótesis del plan económico-financiero de Flujo").font = TITULO
    fila[0] = 3
    titulo("PANEL DE CONTROL")
    dato("esc", "Escenario (pesimista / base / optimista)", escenario, amarillo=True,
         nota="Cambia el ritmo de altas, las bajas y las ventas de productos y de la herramienta")
    nombres_pal = ["Medida 1: actualizar precios con el IPC", "Medida 2: estandarizar el servicio (2 h por cliente)",
                   "Medida 3: talleres grupales para instituciones", "Medida 4: derivación a gestorías",
                   "Medida 5: herramienta no-code sin microcrédito"]
    for i, (nombre, v) in enumerate(zip(nombres_pal, palancas)):
        dato(f"pal{i + 1}", nombre, v, amarillo=True, nota="Sí / No (capítulo 8 del TFG)")
    dv_esc = DataValidation(type="list", formula1='"pesimista,base,optimista"', allow_blank=False)
    dv_sn = DataValidation(type="list", formula1='"Sí,No"', allow_blank=False)
    hip.add_data_validation(dv_esc)
    hip.add_data_validation(dv_sn)
    dv_esc.add(N["esc"].split("!")[1].replace("$", ""))
    for i in range(5):
        dv_sn.add(N[f"pal{i + 1}"].split("!")[1].replace("$", ""))
    fila[0] += 1

    titulo("ESCENARIOS (multiplicadores sobre el escenario base)")
    r_esc = fila[0]
    for j, cab in enumerate(["Escenario", "Altas", "Baja mensual", "Productos", "Herramienta"]):
        hip.cell(r_esc, 2 + j, cab).font = NEGRITA
    fila[0] += 1
    esc_vals = {"pesimista": (0.6, 0.06, 0.5, 0.5), "base": (1.0, H.BAJA_MENSUAL, 1.0, 1.0),
                "optimista": (1.3, 0.03, 1.3, 1.5)}
    for nombre, vs in esc_vals.items():
        hip.cell(fila[0], 2, nombre).font = NEGRO
        for j, v in enumerate(vs):
            c = hip.cell(fila[0], 3 + j, v)
            c.font = AZUL
            c.number_format = PCT if j == 1 else '0.00'
        fila[0] += 1
    rango_esc = f"Hipotesis!$B${r_esc + 1}:$B${r_esc + 3}"
    for j, clave in enumerate(["f_altas", "f_baja", "f_prod", "f_app"]):
        col = L(3 + j)
        N[clave] = f"INDEX(Hipotesis!${col}${r_esc + 1}:${col}${r_esc + 3},MATCH({N['esc']},{rango_esc},0))"
    fila[0] += 1

    titulo("PRECIOS (sin IVA)")
    dato("p_setup", "Setup inicial (€)", H.PRECIO_SETUP, EUR, "Centro del rango 100-300 €")
    dato("p_sub", "Suscripción mensual (€)", H.PRECIO_SUSCRIPCION, EUR, "Centro del rango 50-100 €")
    dato("p_pl", "Plantilla (€)", H.PRECIO_PLANTILLA, EUR)
    dato("p_cur", "Curso (€)", H.PRECIO_CURSO, EUR)
    dato("p_app", "Herramienta digital (€/mes por usuario)", H.PRECIO_APP, EUR)
    dato("p_taller", "Taller para instituciones (€/sesión)", H.PRECIO_TALLER, EUR, "Hipótesis por validar")
    dato("comision_der", "Comisión por cliente derivado a gestoría (€)", H.COMISION_DERIVACION, EUR, "Hipótesis por validar")
    dato("tasa_der_on", "Parte de las altas derivadas (si la medida 4 está activa)", 0.3, PCT)
    dato("tasa_der", "Parte de las altas derivadas aplicada", f'=IF({N["pal4"]}="Sí",C{fila[0] - 1},0)', PCT)
    fila[0] += 1

    titulo("CAPACIDAD Y OPERACIONES")
    dato("h_cli1", "Horas por cliente y mes, año 1", H.HORAS_CLIENTE_MES, NUM)
    dato("h_std", "Horas por cliente con el servicio estandarizado", 2.0, NUM)
    dato("h_cli2", "Horas por cliente y mes, años 2 y 3", f'=IF({N["pal2"]}="Sí",C{fila[0] - 1},C{fila[0] - 2})', NUM)
    dato("h_setup", "Horas por alta (diagnóstico y montaje)", H.HORAS_SETUP, NUM)
    dato("h_taller", "Horas por taller", H.HORAS_TALLER, NUM)
    dato("com_cobro", "Comisión de cobro (servicio y herramienta)", H.COMISION_COBRO, PCT)
    dato("com_plat", "Comisión de plataforma (plantillas y cursos)", H.COMISION_PLATAFORMA, PCT)
    dato("cv_app", "Coste variable de la herramienta (€/usuario y mes)", H.COSTE_VARIABLE_APP, EUR2)
    dato("mant_app", "Mantenimiento de la herramienta (€/mes)", H.MANTENIMIENTO_APP, EUR)
    dato("mes_lanz", "Mes de lanzamiento de la herramienta", H.MES_LANZAMIENTO_APP, "0")
    dato("inflacion", "Inflación anual de costes desde el año 2", H.INFLACION_COSTES, PCT, "Banco de España, proyecciones T3 2026")
    dato("indexar", "¿Se actualizan también los precios?", f'=IF({N["pal1"]}="Sí",1,0)', "0")
    fila[0] += 1

    titulo("COSTES FIJOS MENSUALES (€, precios del año 1)")
    r_cf0 = fila[0]
    hip.cell(r_cf0 - 1, 4, "¿Soporta IVA? (1 = sí)").font = NEGRITA
    for k, v in H.COSTES_FIJOS.items():
        hip.cell(fila[0], 2, k).font = NEGRO
        c = hip.cell(fila[0], 3, v)
        c.font, c.number_format = AZUL, EUR
        d = hip.cell(fila[0], 4, 1 if k in H.CON_IVA else 0)
        d.font = AZUL
        fila[0] += 1
    r_cf1 = fila[0] - 1
    N["cf_total"] = f"SUM(Hipotesis!$C${r_cf0}:$C${r_cf1})"
    N["cf_iva"] = f"SUMPRODUCT(Hipotesis!$C${r_cf0}:$C${r_cf1},Hipotesis!$D${r_cf0}:$D${r_cf1})"
    dato("puesta", "Gastos de puesta en marcha (mes 1)", H.GASTOS_PUESTA_EN_MARCHA, EUR, "Logo, plantillas y textos legales")
    fila[0] += 1

    titulo("INVERSIONES")
    dato("dev_base", "Desarrollo a medida de la herramienta (€)", H.DESARROLLO_APP, EUR)
    dato("dev_nc", "Desarrollo no-code de la herramienta (€)", 5_000, EUR)
    dato("dev", "Desarrollo aplicado", f'=IF({N["pal5"]}="Sí",C{fila[0] - 1},C{fila[0] - 2})', EUR)
    r_inv_cab = fila[0]
    for j, cab in enumerate(["Concepto", "Importe (€)", "Mes de compra", "Vida (años)", "Tipo", "Amortiza desde el mes"]):
        hip.cell(r_inv_cab, 2 + j, cab).font = NEGRITA
    fila[0] += 1
    r_inv0 = fila[0]
    filas_inv = [("Página web con agenda y pagos", 350, 0, 3, "intangible", 1),
                 ("Registro de marca en la OEPM", 125, 0, 10, "intangible", 1),
                 ("Equipo de grabación para cursos", 450, 13, 4, "material", 14)]
    filas_inv += [(f"Desarrollo de la herramienta (mes {m})", f"={N['dev']}/6", m, 3, "intangible", f"={N['mes_lanz']}")
                  for m in range(25, 31)]
    for concepto, imp, mes, vida, tipo, desde in filas_inv:
        vals = [concepto, imp, mes, vida, tipo, desde]
        for j, v in enumerate(vals):
            c = hip.cell(fila[0], 2 + j, v)
            c.font = AZUL if j in (1, 2, 3) and not str(v).startswith("=") else NEGRO
            if j == 1:
                c.number_format = EUR
        fila[0] += 1
    r_inv1 = fila[0] - 1
    INV_IMP = f"Hipotesis!$C${r_inv0}:$C${r_inv1}"
    INV_MES = f"Hipotesis!$D${r_inv0}:$D${r_inv1}"
    INV_VIDA = f"Hipotesis!$E${r_inv0}:$E${r_inv1}"
    INV_TIPO = f"Hipotesis!$F${r_inv0}:$F${r_inv1}"
    INV_DESDE = f"Hipotesis!$G${r_inv0}:$G${r_inv1}"
    fila[0] += 1

    titulo("FINANCIACIÓN")
    dato("aport", "Aportación inicial del titular (€)", H.APORTACION_INICIAL, EUR)
    dato("prest_base", "Microcrédito previsto (€)", H.PRESTAMO, EUR, "Referencia: MicroBank")
    dato("prest", "Microcrédito aplicado (€)", f'=IF({N["pal5"]}="Sí",0,C{fila[0] - 1})', EUR)
    dato("mes_prest", "Mes en que se recibe", H.MES_PRESTAMO, "0")
    dato("tipo_prest", "Tipo de interés anual (TIN)", H.TIPO_PRESTAMO, PCT)
    dato("plazo_prest", "Plazo (meses)", H.PLAZO_PRESTAMO, "0")
    dato("cuota_prest", "Cuota mensual (sistema francés)",
         f'=IF({N["prest"]}>0,PMT({N["tipo_prest"]}/12,{N["plazo_prest"]},-{N["prest"]}),0)', EUR2)
    fila[0] += 1

    titulo("FISCALIDAD Y SEGURIDAD SOCIAL")
    dato("iva", "IVA", H.IVA, PCT)
    dato("tp", "Tarifa plana de autónomos (€/mes)", H.TARIFA_PLANA, EUR, "RDL 13/2022")
    dato("smi", "SMI anual 2026 (€)", H.SMI_ANUAL, EUR, "RD 126/2026")
    dato("ret", "Retención IRPF de inicio de actividad", H.RETENCION, PCT)
    dato("ret_share", "Parte de clientes que practican retención", H.CUOTA_CLIENTES_RETENCION, PCT)
    dato("pf", "Pago fraccionado (modelo 130)", H.PAGO_FRACCIONADO, PCT)
    dato("dj", "Gastos de difícil justificación", H.DIFICIL_JUSTIFICACION, PCT)
    dato("dj_max", "Límite de gastos de difícil justificación (€)", 2_000, EUR)
    dato("minimo", "Mínimo personal (€)", H.MINIMO_PERSONAL, EUR)
    dato("tipo_reta", "Tipo de cotización de autónomos", H.TIPO_RETA, PCT, "Orden PJC/297/2026")
    fila[0] += 1

    def tabla_escala(nombre, tramos):
        hip.cell(fila[0], 2, nombre).font = NEGRITA
        for j, cab in enumerate(["Desde (€)", "Hasta (€)", "Tipo"]):
            hip.cell(fila[0], 3 + j, cab).font = NEGRITA
        fila[0] += 1
        r0 = fila[0]
        anterior = 0
        for limite, tipo in tramos:
            for j, v in enumerate([anterior, 1e12 if limite == float("inf") else limite, tipo]):
                c = hip.cell(fila[0], 3 + j, v)
                c.font = AZUL
                c.number_format = PCT if j == 2 else EUR
            anterior = limite
            fila[0] += 1
        return r0, fila[0] - 1

    titulo("ESCALAS DEL IRPF (base liquidable general)")
    e0, e1 = tabla_escala("Escala estatal", H.ESCALA_ESTATAL)
    a0, a1 = tabla_escala("Escala autonómica de Aragón", H.ESCALA_ARAGON)
    fila[0] += 1

    def escala(base, r0, r1):
        lo, hi, t = (f"Hipotesis!${c}${r0}:${c}${r1}" for c in "CDE")
        b = f"MAX({base},0)"
        return f"SUMPRODUCT(({b}>{lo})*(({b}>{hi})*({hi}-{lo})+({b}<={hi})*({b}-{lo}))*{t})"

    def irpf(base):
        return (f"MAX({escala(base, e0, e1)}+{escala(base, a0, a1)}"
                f"-{escala(N['minimo'], e0, e1)}-{escala(N['minimo'], a0, a1)},0)")

    titulo("CUOTAS DE AUTÓNOMOS POR TRAMO DE RENDIMIENTO MENSUAL (2026)")
    for j, cab in enumerate(["Hasta (€/mes)", "Base mínima (€)", "Cuota (€/mes)"]):
        hip.cell(fila[0], 3 + j, cab).font = NEGRITA
    fila[0] += 1
    t0 = fila[0]
    for limite, base in H.BASES_RETA:
        for j, v in enumerate([1e12 if limite == float("inf") else limite, base, f"=D{fila[0]}*{N['tipo_reta']}"]):
            c = hip.cell(fila[0], 3 + j, v)
            c.font = AZUL if j < 2 else NEGRO
            c.number_format = EUR2
        fila[0] += 1
    t1 = fila[0] - 1
    LIM = f"Hipotesis!$C${t0}:$C${t1}"
    CUO = f"Hipotesis!$E${t0}:$E${t1}"

    def tramo(rend_anual):
        return f'INDEX({CUO},COUNTIF({LIM},"<"&MAX({rend_anual},0)*0.93/12)+1)'
    fila[0] += 1

    titulo("VALORACIÓN")
    dato("rf", "Rentabilidad sin riesgo (Obligaciones a 10 años)", H.TASA_LIBRE_RIESGO, PCT, "Tesoro, subasta 1/10/2026: 4,176 %")
    dato("erp", "Prima de riesgo de mercado", H.PRIMA_MERCADO, PCT, "Damodaran (2026)")
    dato("beta", "Beta", H.BETA, "0.00")
    dato("prima", "Prima específica (tamaño, iliquidez, fundador)", H.PRIMA_ESPECIFICA, PCT)
    dato("tasa", "Tasa de descuento", f'=ROUND({N["rf"]}+{N["beta"]}*{N["erp"]}+{N["prima"]},3)', PCT)
    dato("g", "Crecimiento perpetuo (solo valor residual)", H.CRECIMIENTO_PERPETUO, PCT)
    dato("ss_trab", "Cotización del trabajador asalariado", 0.065, PCT)
    dato("sal_ine", "Salario bruto alternativo: 25-29 años (€)", H.SALARIO_INE_25_29, EUR, "INE, Encuesta de Estructura Salarial 2022")

    hip.column_dimensions["A"].width = 2
    hip.column_dimensions["B"].width = 58
    for c in "CDEFG":
        hip.column_dimensions[c].width = 16
    hip.column_dimensions["D"].width = 22

    # ------------------------------------------------------------------ Mensual
    men.cell(1, 1, "Cálculo mensual (36 meses)").font = TITULO
    men.cell(3, 1, "Mes").font = NEGRITA
    men.cell(4, 1, "Año").font = NEGRITA
    for i, c in enumerate(COLS):
        men[f"{c}3"] = i + 1
        men[f"{c}4"] = f"=ROUNDUP({c}3/12,0)"
        men[f"{c}3"].font = men[f"{c}4"].font = NEGRITA
        men[f"{c}3"].fill = CABECERA
    R = {}
    fila_m = [5]

    def mrow(clave, etiqueta, f, fmt=EUR, entrada=False, negrita=False):
        r = fila_m[0]
        men.cell(r, 1, etiqueta).font = NEGRITA if negrita else NEGRO
        for i, c in enumerate(COLS):
            prev = COLS[i - 1] if i else None
            v = f(i, c, prev) if callable(f) else f[i]
            cell = men[f"{c}{r}"]
            cell.value = v
            cell.number_format = fmt
            cell.font = AZUL if entrada else (NEGRITA if negrita else NEGRO)
        R[clave] = r
        fila_m[0] += 1
        return r

    def sec(texto):
        men.cell(fila_m[0], 1, texto).font = NEGRITA
        men.cell(fila_m[0], 1).fill = GRIS
        fila_m[0] += 1

    sec("ENTRADAS MENSUALES (escenario base, editables)")
    mrow("altas_base", "Altas previstas (escenario base)", H.ALTAS, "0.0", entrada=True)
    mrow("h_disp", "Horas de servicio disponibles", H.HORAS_DISPONIBLES, "0", entrada=True)
    mrow("uds_pl", "Plantillas vendidas (base)", H.UDS_PLANTILLAS, "0.0", entrada=True)
    mrow("uds_cur", "Cursos vendidos (base)", H.UDS_CURSOS, "0.0", entrada=True)
    mrow("usu_base", "Usuarios de la herramienta (base)", H.USUARIOS_APP, "0", entrada=True)
    mrow("mk_base", "Marketing (€, precios del año 1)", H.MARKETING, EUR, entrada=True)
    mrow("tall_base", "Talleres al mes si la medida 3 está activa", [0] * 12 + [1] * 12 + [2] * 12, "0", entrada=True)
    mrow("retir", "Retiradas del titular (€)", H.RETIRADAS, EUR, entrada=True)

    sec("CLIENTES Y CAPACIDAD")
    ref = lambda k, c: f"{c}${R[k]}" if False else f"{c}{R[k]}"
    mrow("ipc", "Índice de costes", lambda i, c, p: f"=(1+{N['inflacion']})^({c}4-1)", "0.000")
    mrow("ipp", "Índice de precios", lambda i, c, p: f"=IF({N['indexar']}=1,{c}{R['ipc']},1)", "0.000")
    mrow("hc", "Horas por cliente", lambda i, c, p: f"=IF({c}4=1,{N['h_cli1']},{N['h_cli2']})", "0.0")
    mrow("tall", "Talleres", lambda i, c, p: f'=IF({N["pal3"]}="Sí",{c}{R["tall_base"]},0)', "0")
    mrow("cap", "Horas disponibles para clientes", lambda i, c, p: f"={c}{R['h_disp']}-{c}{R['tall']}*{N['h_taller']}", "0")
    r_act = fila_m[0] + 2   # fila de clientes activos (se define dos filas más abajo)
    mrow("reten", "Clientes que siguen del mes anterior",
         lambda i, c, p: "=0" if p is None else f"={p}{r_act}*(1-{N['f_baja']})", "0.0")
    mrow("altas", "Altas de clientes",
         lambda i, c, p: f"=MIN({c}{R['altas_base']}*{N['f_altas']},MAX(({c}{R['cap']}-{c}{R['reten']}*{c}{R['hc']})/({c}{R['hc']}+{N['h_setup']}),0))", "0.0")
    mrow("act", "Clientes activos", lambda i, c, p: f"={c}{R['reten']}+{c}{R['altas']}", "0.0", negrita=True)
    assert R["act"] == r_act
    mrow("horas", "Horas ocupadas", lambda i, c, p: f"={c}{R['act']}*{c}{R['hc']}+{c}{R['altas']}*{N['h_setup']}+{c}{R['tall']}*{N['h_taller']}", "0")

    sec("INGRESOS (€)")
    mrow("i_setup", "Setup", lambda i, c, p: f"={c}{R['altas']}*{N['p_setup']}*{c}{R['ipp']}")
    mrow("i_sub", "Suscripción", lambda i, c, p: f"={c}{R['act']}*{N['p_sub']}*{c}{R['ipp']}")
    mrow("i_pl", "Plantillas", lambda i, c, p: f"={c}{R['uds_pl']}*{N['f_prod']}*{N['p_pl']}*{c}{R['ipp']}")
    mrow("i_cur", "Cursos", lambda i, c, p: f"={c}{R['uds_cur']}*{N['f_prod']}*{N['p_cur']}*{c}{R['ipp']}")
    mrow("usu", "Usuarios de la herramienta", lambda i, c, p: f"={c}{R['usu_base']}*{N['f_app']}", "0.0")
    mrow("i_app", "Herramienta digital", lambda i, c, p: f"={c}{R['usu']}*{N['p_app']}*{c}{R['ipp']}")
    mrow("i_tall", "Talleres", lambda i, c, p: f"={c}{R['tall']}*{N['p_taller']}*{c}{R['ipp']}")
    mrow("i_der", "Comisiones de derivación", lambda i, c, p: f"={c}{R['altas']}*{N['tasa_der']}*{N['comision_der']}")
    mrow("ing", "Total ingresos", lambda i, c, p: f"=SUM({c}{R['i_setup']}:{c}{R['i_cur']})+SUM({c}{R['i_app']}:{c}{R['i_der']})", negrita=True)

    sec("GASTOS (€)")
    mrow("cvar", "Costes variables",
         lambda i, c, p: f"=({c}{R['i_setup']}+{c}{R['i_sub']})*{N['com_cobro']}+({c}{R['i_pl']}+{c}{R['i_cur']})*{N['com_plat']}"
                         f"+{c}{R['usu']}*{N['cv_app']}*{c}{R['ipc']}+{c}{R['i_app']}*{N['com_cobro']}")
    mrow("cfij", "Otros gastos fijos", lambda i, c, p: f"={N['cf_total']}*{c}{R['ipc']}")
    mrow("cfij_iva", "  de ellos, con IVA", lambda i, c, p: f"={N['cf_iva']}*{c}{R['ipc']}")
    mrow("mk", "Marketing", lambda i, c, p: f"={c}{R['mk_base']}*{c}{R['ipc']}")
    mrow("reta", "Cuota de autónomos", lambda i, c, p: f"=INDEX(Anual!$C$RETA:$E$RETA,{c}4)")
    mrow("mant", "Mantenimiento de la herramienta",
         lambda i, c, p: f"=IF({c}3>={N['mes_lanz']},{N['mant_app']}*{c}{R['ipc']},0)")
    mrow("pm", "Gastos de puesta en marcha", lambda i, c, p: f"=IF({c}3=1,{N['puesta']},0)")
    mrow("amort", "Amortización",
         lambda i, c, p: f"=SUMPRODUCT(({INV_DESDE}<={c}3)*{INV_IMP}/({INV_VIDA}*12))")

    sec("MICROCRÉDITO (€)")
    r_saldo = fila_m[0]
    mrow("saldo", "Saldo vivo",
         lambda i, c, p: (f"=IF({c}3<{N['mes_prest']},0,IF({c}3={N['mes_prest']},{N['prest']},"
                          + ("0" if p is None else f"{p}{r_saldo}") + f"-{c}{r_saldo + 2}))"))
    mrow("int", "Intereses",
         lambda i, c, p: "=0" if p is None else f"=IF({c}3>{N['mes_prest']},{p}{r_saldo}*{N['tipo_prest']}/12,0)")
    mrow("amp", "Devolución del principal",
         lambda i, c, p: f"=IF(AND({c}3>{N['mes_prest']},{N['prest']}>0),{N['cuota_prest']}-{c}{r_saldo + 1},0)")
    mrow("rn_mes", "Resultado antes de impuestos del mes",
         lambda i, c, p: f"={c}{R['ing']}-{c}{R['cvar']}-{c}{R['cfij']}-{c}{R['mk']}-{c}{R['reta']}-{c}{R['mant']}-{c}{R['pm']}-{c}{R['amort']}-{c}{R['int']}",
         negrita=True)

    sec("TESORERÍA (€, IVA incluido)")
    mrow("inv", "Inversiones del mes (sin IVA)", lambda i, c, p: f"=SUMIF({INV_MES},{c}3,{INV_IMP})")
    mrow("retn", "Retenciones soportadas",
         lambda i, c, p: f"=({c}{R['i_setup']}+{c}{R['i_sub']})*{N['ret_share']}*{N['ret']}")
    mrow("cob", "Cobros", lambda i, c, p: f"={c}{R['ing']}*(1+{N['iva']})-{c}{R['retn']}")
    mrow("g_iva", "Gastos con IVA (base)",
         lambda i, c, p: f"={c}{R['cfij_iva']}+{c}{R['mk']}+{c}{R['mant']}+{c}{R['pm']}+{c}{R['cvar']}")
    mrow("pag", "Pagos de explotación",
         lambda i, c, p: f"={c}{R['g_iva']}*(1+{N['iva']})+{c}{R['cfij']}-{c}{R['cfij_iva']}+{c}{R['reta']}+{c}{R['int']}")
    iva0 = f"-SUMIF({INV_MES},0,{INV_IMP})*{N['iva']}"
    r_trim_post = fila_m[0] + 1
    mrow("iva_pre", "IVA acumulado del trimestre",
         lambda i, c, p: f"={iva0 if p is None else p + str(r_trim_post)}+({c}{R['ing']}-{c}{R['g_iva']}-{c}{R['inv']})*{N['iva']}")
    mrow("iva_post", "IVA arrastrado tras el cierre",
         lambda i, c, p: f"=IF(MOD({c}3,3)=0,MIN({c}{R['iva_pre']},0),{c}{R['iva_pre']})")
    assert R["iva_post"] == r_trim_post
    r_pend = fila_m[0]
    mrow("iva_pend", "IVA pendiente de ingresar",
         lambda i, c, p: f"=IF(MOD({c}3,3)=0,MAX({c}{R['iva_pre']},0),IF(MOD({c}3,3)=1,0,"
                         + ("0" if p is None else f"{p}{r_pend}") + "))")
    mrow("iva_pag", "Liquidación de IVA",
         lambda i, c, p: "=0" if p is None else f"=IF(AND(MOD({c}3,3)=1,{c}3>1),{p}{r_pend},0)")
    r_reta_acum = fila_m[0]
    mrow("ret_acum", "Retenciones acumuladas en el año",
         lambda i, c, p: f"=IF(MOD({c}3,12)=1,0,{p}{r_reta_acum})+{c}{R['retn']}" if p else f"={c}{R['retn']}")
    r_rn_acum = fila_m[0]
    mrow("rn_acum", "Resultado acumulado en el año",
         lambda i, c, p: f"=IF(MOD({c}3,12)=1,0,{p}{r_rn_acum})+{c}{R['rn_mes']}" if p else f"={c}{R['rn_mes']}")
    r_p130_post = fila_m[0] + 1
    mrow("p130", "Pago fraccionado (modelo 130)",
         lambda i, c, p: f"=IF(MOD({c}3,3)=0,MAX({N['pf']}*{c}{R['rn_acum']}-{c}{R['ret_acum']}-"
                         + ("0" if p is None else f"IF(MOD({c}3,12)=1,0,{p}{r_p130_post})") + ",0),0)")
    mrow("p130_acum", "Pagos fraccionados acumulados en el año",
         lambda i, c, p: (f"=IF(MOD({c}3,12)=1,0,{p}{r_p130_post})+{c}{R['p130']}" if p else f"={c}{R['p130']}"))
    assert R["p130_acum"] == r_p130_post
    mrow("liq", "Liquidación de la renta del año anterior (junio)",
         lambda i, c, p: f"=IF(AND(MOD({c}3,12)=6,{c}4>1),INDEX(Anual!$C$LIQ:$E$LIQ,{c}4-1),0)")
    r_caja = fila_m[0]
    caja0 = f"{N['aport']}-SUMIF({INV_MES},0,{INV_IMP})*(1+{N['iva']})"
    mrow("caja", "Saldo de tesorería",
         lambda i, c, p: (f"={caja0 if p is None else p + str(r_caja)}+{c}{R['cob']}-{c}{R['pag']}-{c}{R['inv']}*(1+{N['iva']})"
                          f"-{c}{R['retir']}+IF({c}3={N['mes_prest']},{N['prest']},0)-{c}{R['amp']}"
                          f"-{c}{R['iva_pag']}-{c}{R['p130']}-{c}{R['liq']}"), negrita=True)

    men.column_dimensions["A"].width = 48
    men.column_dimensions["B"].width = 2
    for c in COLS:
        men.column_dimensions[c].width = 11
    men.freeze_panes = "C5"

    # ------------------------------------------------------------------ Anual
    anu.cell(1, 1, "Estados financieros anuales").font = TITULO
    for j, c in enumerate(ANIOS):
        anu[f"{c}3"] = f"Año {j + 1}"
        anu[f"{c}3"].font = NEGRITA
        anu[f"{c}3"].fill = CABECERA
    A = {}
    fila_a = [4]

    def suma(k, c, j):
        return f"SUMIF(Mensual!${COLS[0]}$4:${ULT}$4,{j + 1},Mensual!${COLS[0]}${R[k]}:${ULT}${R[k]})"

    def fin(k, j):
        return f"INDEX(Mensual!${COLS[0]}${R[k]}:${ULT}${R[k]},1,{12 * (j + 1)})"

    def arow(clave, etiqueta, f, fmt=EUR, negrita=False):
        r = fila_a[0]
        anu.cell(r, 1, etiqueta).font = NEGRITA if negrita else NEGRO
        for j, c in enumerate(ANIOS):
            cell = anu[f"{c}{r}"]
            cell.value = f(j, c)
            cell.number_format = fmt
            cell.font = NEGRITA if negrita else NEGRO
            if negrita:
                cell.fill = GRIS
        if negrita:
            anu.cell(r, 1).fill = GRIS
        A[clave] = r
        fila_a[0] += 1

    def asec(texto):
        anu.cell(fila_a[0], 1, texto).font = NEGRITA
        anu.cell(fila_a[0], 1).fill = CABECERA
        fila_a[0] += 1

    asec("CUENTA DE RESULTADOS")
    arow("altas", "Altas de clientes", lambda j, c: "=" + suma("altas", c, j), "0")
    arow("act", "Clientes activos a final de año", lambda j, c: "=" + fin("act", j), "0")
    for k, et in [("i_setup", "Ingresos por setup"), ("i_sub", "Ingresos por suscripción"), ("i_pl", "Ingresos por plantillas"),
                  ("i_cur", "Ingresos por cursos"), ("i_app", "Ingresos por la herramienta"), ("i_tall", "Ingresos por talleres"),
                  ("i_der", "Comisiones de derivación")]:
        arow(k, et, lambda j, c, k=k: "=" + suma(k, c, j))
    arow("ing", "Importe neto de la cifra de negocios", lambda j, c: f"=SUM({c}{A['i_setup']}:{c}{A['i_der']})", negrita=True)
    arow("cvar", "Costes variables", lambda j, c: "=-" + suma("cvar", c, j))
    arow("mc", "Margen de contribución", lambda j, c: f"={c}{A['ing']}+{c}{A['cvar']}", negrita=True)
    arow("reta", "Cuota de autónomos", lambda j, c: f"=-12*{c}RETAROW")
    arow("cfij", "Otros gastos fijos", lambda j, c: "=-" + suma("cfij", c, j))
    arow("mk", "Marketing", lambda j, c: "=-" + suma("mk", c, j))
    arow("mant", "Mantenimiento de la herramienta", lambda j, c: "=-" + suma("mant", c, j))
    arow("pm", "Gastos de puesta en marcha", lambda j, c: "=-" + suma("pm", c, j))
    arow("ebitda", "EBITDA", lambda j, c: f"=SUM({c}{A['mc']}:{c}{A['pm']})-{c}{A['mc']}+{c}{A['mc']}", negrita=True)
    arow("amort", "Amortización", lambda j, c: "=-" + suma("amort", c, j))
    arow("ebit", "Resultado de explotación (EBIT)", lambda j, c: f"={c}{A['ebitda']}+{c}{A['amort']}", negrita=True)
    arow("int", "Gastos financieros", lambda j, c: "=-" + suma("int", c, j))
    arow("bai", "Resultado antes de impuestos", lambda j, c: f"={c}{A['ebit']}+{c}{A['int']}", negrita=True)
    arow("djv", "Gastos de difícil justificación", lambda j, c: f"=MIN(MAX({c}{A['bai']},0)*{N['dj']},{N['dj_max']})")
    arow("rn", "Rendimiento neto (IRPF)", lambda j, c: f"={c}{A['bai']}-{c}{A['djv']}")
    arow("irpf", "IRPF imputable a la actividad", lambda j, c: "=-" + irpf(f"{c}{A['rn']}"))
    arow("neto", "Resultado neto", lambda j, c: f"={c}{A['bai']}+{c}{A['irpf']}", negrita=True)
    # EBITDA = margen + gastos de explotación (todos negativos)
    for c in ANIOS:
        anu[f"{c}{A['ebitda']}"] = f"={c}{A['mc']}+SUM({c}{A['reta']}:{c}{A['pm']})"

    asec("CUOTA DE AUTÓNOMOS (rendimiento computable = rendimiento neto + cuotas, menos 7 %)")
    arow("rn_sin", "Resultado antes de la cuota de autónomos",
         lambda j, c: f"={c}{A['mc']}+SUM({c}{A['cfij']}:{c}{A['pm']})+{c}{A['amort']}+{c}{A['int']}")

    def computable(c, cuota):
        # rendimiento neto fiscal + cuotas pagadas (art. 308 LGSS)
        x = f"({c}{A['rn_sin']}-12*{cuota})"
        return f"({x}-MIN(MAX({x},0)*{N['dj']},{N['dj_max']})+12*{cuota})"
    arow("c1", "Primera estimación de la cuota", lambda j, c: "=" + tramo(computable(c, N["tp"])), EUR2)
    arow("c2", "Segunda estimación de la cuota", lambda j, c: "=" + tramo(computable(c, f"{c}{A['c1']}")), EUR2)
    arow("cuota", "Cuota mensual aplicada",
         lambda j, c: (f"={N['tp']}" if j == 0 else
                       f"=IF({computable(c, N['tp'])}<{N['smi']},{N['tp']},{c}{A['c2']})" if j == 1 else f"={c}{A['c2']}"),
         EUR2, negrita=True)
    for c in ANIOS:
        anu[f"{c}{A['reta']}"] = f"=-12*{c}{A['cuota']}"

    asec("IMPUESTOS PENDIENTES")
    arow("retn", "Retenciones soportadas", lambda j, c: "=" + suma("retn", c, j))
    arow("p130", "Pagos fraccionados (modelo 130)", lambda j, c: "=" + suma("p130", c, j))
    arow("liq", "Liquidación de la renta pendiente (+ a pagar / − a devolver)",
         lambda j, c: f"=-{c}{A['irpf']}-{c}{A['retn']}-{c}{A['p130']}")
    arow("iva_fin", "IVA pendiente a 31 de diciembre", lambda j, c: "=" + fin("iva_pend", j))

    asec("ESTADO DE FLUJOS DE TESORERÍA")
    arow("f_cob", "Cobros de explotación", lambda j, c: "=" + suma("cob", c, j))
    arow("f_pag", "Pagos de explotación", lambda j, c: "=-" + suma("pag", c, j))
    arow("f_iva", "Liquidaciones de IVA", lambda j, c: "=-" + suma("iva_pag", c, j))
    arow("f_irpf", "Pagos de IRPF", lambda j, c: f"=-({suma('p130', c, j)}+{suma('liq', c, j)})")
    arow("f_expl", "Flujo de explotación", lambda j, c: f"=SUM({c}{A['f_cob']}:{c}{A['f_irpf']})", negrita=True)
    arow("f_inv", "Flujo de inversión",
         lambda j, c: f"=-({suma('inv', c, j)}" + (f"+SUMIF({INV_MES},0,{INV_IMP})" if j == 0 else "") + f")*(1+{N['iva']})",
         negrita=True)
    arow("f_aport", "Aportación del titular", lambda j, c: f"={N['aport']}" if j == 0 else "=0")
    arow("f_prest", "Microcrédito recibido",
         lambda j, c: f"=IF(ROUNDUP({N['mes_prest']}/12,0)={j + 1},{N['prest']},0)")
    arow("f_dev", "Devolución del principal", lambda j, c: "=-" + suma("amp", c, j))
    arow("f_ret", "Retiradas del titular", lambda j, c: "=-" + suma("retir", c, j))
    arow("f_fin", "Flujo de financiación", lambda j, c: f"=SUM({c}{A['f_aport']}:{c}{A['f_ret']})", negrita=True)
    arow("caja_ini", "Tesorería inicial", lambda j, c: "=0" if j == 0 else f"={ANIOS[j - 1]}CAJAFIN")
    arow("caja", "Tesorería final", lambda j, c: "=" + fin("caja", j), negrita=True)
    arow("caja_chk", "Comprobación (flujos − variación de caja)",
         lambda j, c: f"=ROUND({c}{A['f_expl']}+{c}{A['f_inv']}+{c}{A['f_fin']}+{c}{A['caja_ini']}-{c}CAJAFIN,2)")
    for j, c in enumerate(ANIOS):
        for k in ("caja_ini", "caja_chk"):
            anu[f"{c}{A[k]}"] = anu[f"{c}{A[k]}"].value.replace("CAJAFIN", str(A["caja"]))
    arow("caja_min", "Saldo mínimo de tesorería (36 meses)",
         lambda j, c: f"=MIN(Mensual!${COLS[0]}${R['caja']}:${ULT}${R['caja']})" if j == 0 else '=""')

    asec("BALANCE A 31 DE DICIEMBRE")
    arow("b_int", "Inmovilizado intangible",
         lambda j, c: f'=SUMPRODUCT(({INV_TIPO}="intangible")*({INV_MES}<={12 * (j + 1)})*{INV_IMP})')
    arow("b_mat", "Inmovilizado material",
         lambda j, c: f'=SUMPRODUCT(({INV_TIPO}="material")*({INV_MES}<={12 * (j + 1)})*{INV_IMP})')
    arow("b_am", "Amortización acumulada", lambda j, c: f"=SUM($C${A['amort']}:{c}{A['amort']})")
    arow("b_anc", "Activo no corriente", lambda j, c: f"=SUM({c}{A['b_int']}:{c}{A['b_am']})", negrita=True)
    arow("b_hpd", "Hacienda Pública deudora", lambda j, c: f"=MAX(-{c}{A['liq']},0)")
    arow("b_tes", "Tesorería", lambda j, c: f"={c}{A['caja']}")
    arow("b_act", "TOTAL ACTIVO", lambda j, c: f"={c}{A['b_anc']}+{c}{A['b_hpd']}+{c}{A['b_tes']}", negrita=True)
    arow("b_apo", "Aportaciones del titular", lambda j, c: f"={N['aport']}")
    arow("b_res", "Resultados acumulados", lambda j, c: f"=SUM($C${A['neto']}:{c}{A['neto']})")
    arow("b_ret", "Retiradas del titular", lambda j, c: f"=SUM($C${A['f_ret']}:{c}{A['f_ret']})")
    arow("b_pn", "PATRIMONIO NETO", lambda j, c: f"=SUM({c}{A['b_apo']}:{c}{A['b_ret']})", negrita=True)

    def k_pag(j):
        return f"({12 * (j + 1)}-{N['mes_prest']})"
    arow("b_dcp", "Deudas con entidades de crédito a corto plazo",
         lambda j, c: (f"=IF({fin('saldo', j)}>0,-CUMPRINC({N['tipo_prest']}/12,{N['plazo_prest']},{N['prest']},"
                       f"{k_pag(j)}+1,MIN({k_pag(j)}+12,{N['plazo_prest']}),0),0)"))
    arow("b_dlp", "Deudas con entidades de crédito a largo plazo",
         lambda j, c: f"=MAX({fin('saldo', j)}-{c}{A['b_dcp']},0)")
    arow("b_hpa", "Hacienda Pública acreedora (IVA e IRPF)", lambda j, c: f"={c}{A['iva_fin']}+MAX({c}{A['liq']},0)")
    arow("b_pas", "TOTAL PASIVO", lambda j, c: f"=SUM({c}{A['b_dcp']}:{c}{A['b_hpa']})", negrita=True)
    arow("b_pnp", "TOTAL PATRIMONIO NETO Y PASIVO", lambda j, c: f"={c}{A['b_pn']}+{c}{A['b_pas']}", negrita=True)
    arow("b_chk", "Comprobación (activo − PN − pasivo)", lambda j, c: f"=ROUND({c}{A['b_act']}-{c}{A['b_pnp']},2)")

    # sustituir marcadores pendientes
    for row in men.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                cell.value = (cell.value.replace("$C$RETA:$E$RETA", f"$C${A['cuota']}:$E${A['cuota']}")
                              .replace("$C$LIQ:$E$LIQ", f"$C${A['liq']}:$E${A['liq']}"))
    anu.column_dimensions["A"].width = 58
    for c in ANIOS:
        anu.column_dimensions[c].width = 15
    anu.freeze_panes = "B4"

    # ------------------------------------------------------------------ Valoracion
    val.cell(1, 1, "Valoración: VAN, TIR y periodo de recuperación").font = TITULO
    cols_v = ["B", "C", "D", "E"]
    for j, c in enumerate(cols_v):
        val[f"{c}3"] = f"Año {j}"
        val[f"{c}3"].font = NEGRITA
        val[f"{c}3"].fill = CABECERA
    V = {}
    fila_v = [4]

    def vrow(clave, etiqueta, valores, fmt=EUR, negrita=False):
        r = fila_v[0]
        val.cell(r, 1, etiqueta).font = NEGRITA if negrita else NEGRO
        for c, v in zip(cols_v, valores):
            if v is None:
                continue
            cell = val[f"{c}{r}"]
            cell.value = v
            cell.number_format = fmt
            cell.font = NEGRITA if negrita else NEGRO
        V[clave] = r
        fila_v[0] += 1

    inv_mes0 = f"SUMIF({INV_MES},0,{INV_IMP})"
    vrow("inv0", "Inversión inicial (activos + puesta en marcha)", [f"={inv_mes0}+{N['puesta']}", None, None, None])
    vrow("fm", "Colchón de tesorería inicial",
         [f"={N['aport']}-B{fila_v[0] - 1}-{inv_mes0}*{N['iva']}", None, None, None])
    a = lambda k, j: f"Anual!{ANIOS[j]}{A[k]}"
    vrow("ebitda", "EBITDA sin gastos de puesta en marcha",
         [None] + [f"={a('ebitda', j)}" + (f"+{N['puesta']}" if j == 0 else "") for j in range(3)])
    def ebit_fiscal(j):
        e = a("ebit", j)
        return f"({e}-MIN(MAX({e},0)*{N['dj']},{N['dj_max']}))"
    vrow("irpf_op", "IRPF sobre el resultado de explotación",
         [None] + [f"=-{irpf(ebit_fiscal(j))}" for j in range(3)])
    vrow("inv", "Inversiones",
         [None] + [f'=-SUMIFS({INV_IMP},{INV_MES},">"&{12 * j},{INV_MES},"<="&{12 * (j + 1)})' for j in range(3)])
    vrow("nof", "Necesidades operativas de fondos (− deudas con Hacienda)",
         [None] + [f"=-({a('iva_fin', j)}+{a('liq', j)})" for j in range(3)])
    vrow("dnof", "Variación del capital circulante",
         [None, f"=-C{fila_v[0] - 1}", f"=-(D{fila_v[0] - 1}-C{fila_v[0] - 1})", f"=-(E{fila_v[0] - 1}-D{fila_v[0] - 1})"])
    vrow("rec", "Recuperación del colchón de tesorería", [None, "=0", "=0", f"=B{V['fm']}"])
    vrow("fcl", "FLUJO DE CAJA LIBRE",
         [f"=-(B{V['inv0']}+B{V['fm']})"] + [f"=SUM({c}{V['ebitda']}:{c}{V['rec']})-{c}{V['nof']}" for c in "CDE"],
         negrita=True)
    fila_v[0] += 1
    vrow("van", "VAN a 3 años", [f"=B{V['fcl']}+NPV({N['tasa']},C{V['fcl']}:E{V['fcl']})"], negrita=True)
    vrow("tir", "TIR", [f"=IRR(B{V['fcl']}:E{V['fcl']},0.5)"], PCT)
    vrow("vr", "Valor residual (crecimiento perpetuo, solo referencia)",
         [f"=(E{V['fcl']}-B{V['fm']})*(1+{N['g']})/({N['tasa']}-{N['g']})"])
    vrow("van_vr", "VAN con valor residual", [f"=B{V['van']}+B{V['vr']}/(1+{N['tasa']})^3"])

    def payback(fila_f, desc):
        d = (lambda c, n: f"{c}{fila_f}/(1+{N['tasa']})^{n}") if desc else (lambda c, n: f"{c}{fila_f}")
        a0 = f"B{fila_f}"
        a1 = f"({a0}+{d('C', 1)})"
        a2 = f"({a1}+{d('D', 2)})"
        a3 = f"({a2}+{d('E', 3)})"
        return (f'=IF({a1}>=0,-{a0}/({d("C", 1)}),IF({a2}>=0,1-{a1}/({d("D", 2)}),'
                f'IF({a3}>=0,2-{a2}/({d("E", 3)}),"no se recupera")))')
    vrow("pb", "Payback simple (años)", [payback(V["fcl"], False)], "0.00")
    vrow("pbd", "Payback descontado (años)", [payback(V["fcl"], True)], "0.00")
    fila_v[0] += 1
    val.cell(fila_v[0], 1, "COSTE DE OPORTUNIDAD DEL FUNDADOR").font = NEGRITA
    val.cell(fila_v[0], 1).fill = CABECERA
    fila_v[0] += 1
    vrow("co_smi", "Sueldo neto alternativo: SMI (dedicación parcial el año 1)",
         [None, f"=0.5*{N['smi']}*(1-{N['ss_trab']})", f"={N['smi']}*(1-{N['ss_trab']})", f"={N['smi']}*(1-{N['ss_trab']})"])
    vrow("fcl_smi", "Flujo menos coste de oportunidad (SMI)",
         [f"=B{V['fcl']}"] + [f"={c}{V['fcl']}-{c}{V['co_smi']}" for c in "CDE"])
    vrow("van_smi", "VAN con coste de oportunidad (SMI)", [f"=B{V['fcl_smi']}+NPV({N['tasa']},C{V['fcl_smi']}:E{V['fcl_smi']})"], negrita=True)
    vrow("tir_smi", "TIR con coste de oportunidad (SMI)", [f'=IFERROR(IRR(B{V["fcl_smi"]}:E{V["fcl_smi"]},0.3),"n.d.")'], PCT)
    rn_ine = f"({N['sal_ine']}*(1-{N['ss_trab']}))"
    red = (f"IF({rn_ine}<=14852,7302,IF({rn_ine}<=17673.52,7302-1.75*({rn_ine}-14852),"
           f"IF({rn_ine}<=19747.5,2364.34-1.14*({rn_ine}-17673.52),0)))")
    vrow("neto_ine", "Sueldo neto alternativo: salario medio de 25 a 29 años (INE)",
         [f"={rn_ine}-{irpf(f'({rn_ine}-2000-{red})')}"])
    vrow("co_ine", "Coste de oportunidad (INE)",
         [None, f"=0.5*$B${V['neto_ine']}", f"=$B${V['neto_ine']}", f"=$B${V['neto_ine']}"])
    vrow("fcl_ine", "Flujo menos coste de oportunidad (INE)",
         [f"=B{V['fcl']}"] + [f"={c}{V['fcl']}-{c}{V['co_ine']}" for c in "CDE"])
    vrow("van_ine", "VAN con coste de oportunidad (INE)", [f"=B{V['fcl_ine']}+NPV({N['tasa']},C{V['fcl_ine']}:E{V['fcl_ine']})"], negrita=True)
    val.column_dimensions["A"].width = 62
    for c in cols_v:
        val.column_dimensions[c].width = 15

    # ------------------------------------------------------------------ Resumen
    res.cell(1, 1, "Resumen del plan").font = TITULO
    res.cell(2, 1, "Cambia el escenario y las medidas en la hoja Hipotesis (celdas amarillas).").font = Font(name=FUENTE, italic=True)
    res["A4"], res["B4"] = "Escenario", f"={N['esc']}"
    for i in range(5):
        res.cell(5 + i, 1, nombres_pal[i]).font = NEGRO
        res.cell(5 + i, 2, f"={N[f'pal{i + 1}']}").font = VERDE
    res["B4"].font = VERDE
    r = 11
    for j, c in enumerate(["Año 1", "Año 2", "Año 3"]):
        res.cell(r, 2 + j, c).font = NEGRITA
        res.cell(r, 2 + j).fill = CABECERA
    filas_res = [("Ingresos", "ing"), ("EBITDA", "ebitda"), ("Resultado neto", "neto"), ("Clientes activos a final de año", "act"),
                 ("Tesorería final", "caja"), ("Cuota mensual de autónomos", "cuota")]
    for k, (et, clave) in enumerate(filas_res):
        res.cell(r + 1 + k, 1, et).font = NEGRO
        for j, c in enumerate(ANIOS):
            cell = res.cell(r + 1 + k, 2 + j, f"=Anual!{c}{A[clave]}")
            cell.font = VERDE
            cell.number_format = "0" if clave == "act" else EUR
    r = r + 2 + len(filas_res)
    indicadores = [("Tasa de descuento", f"={N['tasa']}", PCT), ("VAN a 3 años", f"=Valoracion!B{V['van']}", EUR),
                   ("TIR", f"=Valoracion!B{V['tir']}", PCT), ("Payback (años)", f"=Valoracion!B{V['pb']}", "0.00"),
                   ("VAN con coste de oportunidad (SMI)", f"=Valoracion!B{V['van_smi']}", EUR),
                   ("VAN con coste de oportunidad (salario INE)", f"=Valoracion!B{V['van_ine']}", EUR),
                   ("Saldo mínimo de tesorería", f"=Anual!C{A['caja_min']}", EUR),
                   ("Comprobación del balance (debe ser 0)", f"=SUM(Anual!C{A['b_chk']}:E{A['b_chk']})", EUR2)]
    for k, (et, f, fmt) in enumerate(indicadores):
        res.cell(r + k, 1, et).font = NEGRITA
        cell = res.cell(r + k, 2, f)
        cell.font, cell.number_format = Font(name=FUENTE, bold=True, color="008000"), fmt
    res.column_dimensions["A"].width = 52
    for c in "BCD":
        res.column_dimensions[c].width = 15

    graf = LineChart()
    graf.title = "Saldo mensual de tesorería"
    graf.y_axis.title = "€"
    graf.x_axis.title = "Mes"
    graf.add_data(Reference(men, min_col=C0, max_col=C0 + MESES - 1, min_row=R["caja"]), from_rows=True, titles_from_data=False)
    graf.set_categories(Reference(men, min_col=C0, max_col=C0 + MESES - 1, min_row=3))
    graf.legend = None
    graf.height, graf.width = 7, 16
    res.add_chart(graf, "F3")
    barras = BarChart()
    barras.title = "Ingresos y resultado neto"
    barras.add_data(Reference(anu, min_col=3, max_col=5, min_row=A["ing"]), from_rows=True, titles_from_data=False)
    barras.add_data(Reference(anu, min_col=3, max_col=5, min_row=A["neto"]), from_rows=True, titles_from_data=False)
    barras.series[0].tx = SeriesLabel(v="Ingresos")
    barras.series[1].tx = SeriesLabel(v="Resultado neto")
    barras.set_categories(Reference(anu, min_col=3, max_col=5, min_row=3))
    barras.height, barras.width = 7, 16
    res.add_chart(barras, "F18")

    # ------------------------------------------------------------------ Instrucciones
    ins.column_dimensions["A"].width = 110
    lineas = [
        ("Modelo financiero de Flujo (TFG, Universidad de Zaragoza)", TITULO),
        ("", NEGRO),
        ("Este libro reproduce con fórmulas el modelo del trabajo. Al cambiar una hipótesis se recalcula todo.", NEGRO),
        ("", NEGRO),
        ("Cómo usarlo", NEGRITA),
        ("1. En la hoja Hipotesis, las celdas amarillas del panel de control eligen el escenario y activan o desactivan las cinco medidas del capítulo 8.", NEGRO),
        ("2. Las cifras en azul son hipótesis que puedes cambiar (precios, costes, tipos, plazos...). Las negras son fórmulas: no las modifiques.", NEGRO),
        ("3. En la hoja Mensual, las primeras filas en azul (altas previstas, marketing, retiradas...) también se pueden editar mes a mes.", NEGRO),
        ("4. Los resultados están en Resumen; el detalle, en Anual (PyG, tesorería y balance) y en Valoracion (VAN, TIR y payback).", NEGRO),
        ("5. Las celdas de comprobación del balance y de la tesorería deben valer 0.", NEGRO),
        ("", NEGRO),
        ("Convenciones: azul = dato editable · negro = fórmula · verde = enlace a otra hoja · amarillo = control de escenario.", NEGRO),
        ("La cuota de autónomos se calcula en dos pasadas para evitar referencias circulares; la simulación de Montecarlo está en el código Python (repositorio del Anexo II).", NEGRO),
        ("Ejemplo: pon el escenario en «pesimista» y la medida 3 en «Sí» y mira cómo cambian el VAN y la tesorería en la hoja Resumen.", NEGRO),
    ]
    for i, (t, f) in enumerate(lineas, start=1):
        ins.cell(i, 1, t).font = f
        ins.cell(i, 1).alignment = Alignment(wrap_text=True)
    # Hoja estática con la simulación de Montecarlo (calculada en Python)
    rj = RAIZ / "modelo" / "resultados.json"
    if rj.exists():
        Rj = json.loads(rj.read_text())
        mc = wb.create_sheet("Montecarlo")
        mc.sheet_view.showGridLines = False
        mc["A1"] = "Simulación de Montecarlo (2.000 escenarios, calculada con modelo.py; valores fijos, no se recalcula)"
        mc["A1"].font = NEGRITA
        filas_mc = [("Probabilidad de VAN positivo", "p_van", PCT),
                    ("Probabilidad de VAN con coste de oportunidad (SMI) positivo", "p_van_op", PCT),
                    ("Probabilidad de VAN con coste de oportunidad (INE) positivo", "p_van_ine", PCT),
                    ("VAN con coste de oportunidad (SMI): percentil 5", "p5_op", EUR),
                    ("VAN con coste de oportunidad (SMI): mediana", "p50_op", EUR),
                    ("VAN con coste de oportunidad (SMI): percentil 95", "p95_op", EUR)]
        for j, cab in enumerate(["Indicador", "Plan base", "Plan mejorado"]):
            mc.cell(3, 1 + j, cab).font = NEGRITA
            mc.cell(3, 1 + j).fill = CABECERA
        for k, (et, clave, fmt) in enumerate(filas_mc):
            mc.cell(4 + k, 1, et).font = NEGRO
            for j, bloque in enumerate(("montecarlo", "montecarlo_opt")):
                cell = mc.cell(4 + k, 2 + j, Rj[bloque][clave])
                cell.number_format, cell.font = fmt, NEGRO
        mc.column_dimensions["A"].width = 62
        mc.column_dimensions["B"].width = mc.column_dimensions["C"].width = 16
    orden = ["Instrucciones", "Resumen", "Hipotesis", "Mensual", "Anual", "Valoracion", "Montecarlo"]
    wb._sheets.sort(key=lambda ws: orden.index(ws.title))
    wb.active = 0
    wb.save(ruta)
    return dict(N=N, R=R, A=A, V=V)


if __name__ == "__main__":
    ruta = sys.argv[1] if len(sys.argv) > 1 else str(RAIZ / "Modelo_financiero_Flujo.xlsx")
    esc = sys.argv[2] if len(sys.argv) > 2 else "base"
    pal = tuple(sys.argv[3:8]) if len(sys.argv) > 7 else ("No",) * 5
    mapa = build(ruta, esc, pal)
    Path(ruta).with_suffix(".mapa.json").write_text(json.dumps({k: mapa[k] for k in ("A", "V")}))
    print("guardado", ruta)
