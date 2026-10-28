"""Motor del plan económico-financiero de Flujo (36 meses).

Uso:  python modelo.py      -> imprime el resumen y guarda resultados.json
"""
import json
from pathlib import Path

import hipotesis as H

ESCENARIOS = {
    "pesimista": dict(altas=0.6, baja=0.06, productos=0.5, app=0.5),
    "base": dict(altas=1.0, baja=H.BAJA_MENSUAL, productos=1.0, app=1.0),
    "optimista": dict(altas=1.3, baja=0.03, productos=1.3, app=1.5),
}


def escala(base, tramos):
    cuota, anterior = 0.0, 0.0
    for limite, tipo in tramos:
        if base <= anterior:
            break
        cuota += (min(base, limite) - anterior) * tipo
        anterior = limite
    return cuota


def irpf(rendimiento):
    """Cuota íntegra estatal + autonómica (Aragón) sobre el rendimiento."""
    base = max(rendimiento, 0)
    cuota = escala(base, H.ESCALA_ESTATAL) + escala(base, H.ESCALA_ARAGON)
    minimo = escala(H.MINIMO_PERSONAL, H.ESCALA_ESTATAL) + escala(H.MINIMO_PERSONAL, H.ESCALA_ARAGON)
    return max(cuota - minimo, 0)


def cuota_reta(rendimiento_anual):
    mensual = max(rendimiento_anual, 0) * 0.93 / 12   # 7 % de deducción genérica
    return next(c for limite, c in H.TRAMOS_RETA if mensual <= limite)


def cuota_prestamo(capital, tipo, n):
    i = tipo / 12
    return capital * i / (1 - (1 + i) ** -n)


def simular(nombre="base", precio=None, tasa=None, factor_altas=None, indexar_precios=None):
    e = dict(ESCENARIOS[nombre])
    if factor_altas is not None:
        e["altas"] = factor_altas
    precio = precio or H.PRECIO_SUSCRIPCION
    e["indexar"] = H.INDEXAR_PRECIOS if indexar_precios is None else indexar_precios
    M = H.MESES
    m = {k: [0.0] * M for k in [
        "altas", "activos", "ing_setup", "ing_suscripcion", "ing_plantillas",
        "ing_cursos", "ing_app", "ingresos", "c_variables", "c_fijos",
        "marketing", "reta", "puesta_marcha", "mant_app", "amortizacion",
        "intereses", "horas", "usuarios_app"]}

    # Cuota de autónomos: tarifa plana el año 1; el año 2 se prorroga si el
    # rendimiento del año 1 no supera el SMI; el año 3 va por tramos.  El
    # rendimiento depende de la cuota, así que se itera hasta que converge.
    reta_anual = [H.TARIFA_PLANA, H.TARIFA_PLANA, H.TARIFA_PLANA]
    for _ in range(10):
        res = _calcular(m, e, precio, reta_anual)
        rn = res["anual"]["rendimiento_neto"]
        nueva = [H.TARIFA_PLANA,
                 H.TARIFA_PLANA if rn[0] < H.SMI_ANUAL else cuota_reta(rn[1]),
                 cuota_reta(rn[2])]
        if nueva == reta_anual:
            break
        reta_anual = nueva
    res["reta_mensual"] = reta_anual
    res["valoracion"] = valorar(res, tasa or H.TASA_DESCUENTO)
    res["escenario"] = nombre
    return res


def _calcular(m, e, precio, reta_anual):
    M = H.MESES
    activos = 0.0
    for t in range(M):
        ipc = (1 + H.INFLACION_COSTES) ** (t // 12)
        ipp = ipc if e["indexar"] else 1.0
        cap = H.HORAS_DISPONIBLES[t]
        altas = H.ALTAS[t] * e["altas"]
        retenidos = activos * (1 - e["baja"])
        # límite de capacidad: si no hay horas, no se aceptan más clientes
        max_altas = max((cap - retenidos * H.HORAS_CLIENTE_MES) / (H.HORAS_CLIENTE_MES + H.HORAS_SETUP), 0)
        altas = min(altas, max_altas)
        activos = retenidos + altas
        m["altas"][t], m["activos"][t] = altas, activos
        m["horas"][t] = activos * H.HORAS_CLIENTE_MES + altas * H.HORAS_SETUP
        m["ing_setup"][t] = altas * H.PRECIO_SETUP * ipp
        m["ing_suscripcion"][t] = activos * precio * ipp
        m["ing_plantillas"][t] = H.UDS_PLANTILLAS[t] * e["productos"] * H.PRECIO_PLANTILLA * ipp
        m["ing_cursos"][t] = H.UDS_CURSOS[t] * e["productos"] * H.PRECIO_CURSO * ipp
        usuarios = H.USUARIOS_APP[t] * e["app"]
        m["usuarios_app"][t] = usuarios
        m["ing_app"][t] = usuarios * H.PRECIO_APP * ipp
        servicio = m["ing_setup"][t] + m["ing_suscripcion"][t]
        productos = m["ing_plantillas"][t] + m["ing_cursos"][t]
        m["ingresos"][t] = servicio + productos + m["ing_app"][t]
        m["c_variables"][t] = (servicio * H.COMISION_COBRO + productos * H.COMISION_PLATAFORMA
                               + usuarios * H.COSTE_VARIABLE_APP * ipc + m["ing_app"][t] * H.COMISION_COBRO)
        m["c_fijos"][t] = sum(H.COSTES_FIJOS.values()) * ipc
        m["marketing"][t] = H.MARKETING[t] * ipc
        m["reta"][t] = reta_anual[t // 12]
        m["mant_app"][t] = H.MANTENIMIENTO_APP * ipc if t + 1 >= H.MES_LANZAMIENTO_APP else 0
        m["puesta_marcha"][t] = H.GASTOS_PUESTA_EN_MARCHA if t == 0 else 0

    m["amortizacion"] = [0.0] * M
    m["intereses"] = [0.0] * M
    # amortizaciones (lineales, desde el mes siguiente a la compra; el MVP desde el lanzamiento)
    for concepto, importe, mes, vida, _ in H.INVERSIONES:
        inicio = H.MES_LANZAMIENTO_APP - 1 if "MVP" in concepto else max(mes, 0)
        for t in range(inicio, M):
            m["amortizacion"][t] += importe / (vida * 12)

    # préstamo (sistema francés)
    saldo, cuota = 0.0, cuota_prestamo(H.PRESTAMO, H.TIPO_PRESTAMO, H.PLAZO_PRESTAMO)
    amort_prestamo = [0.0] * M
    for t in range(M):
        if t + 1 == H.MES_PRESTAMO:
            saldo = H.PRESTAMO
        elif saldo > 0:
            interes = saldo * H.TIPO_PRESTAMO / 12
            m["intereses"][t] = interes
            amort_prestamo[t] = cuota - interes
            saldo -= cuota - interes
    saldo_prestamo = [0.0] * M
    s = 0.0
    for t in range(M):
        if t + 1 == H.MES_PRESTAMO:
            s = H.PRESTAMO
        s -= amort_prestamo[t]
        saldo_prestamo[t] = s

    anual = resumen_anual(m)
    anual["amort_prestamo"] = [sum(amort_prestamo[a * 12:(a + 1) * 12]) for a in range(3)]
    anual["saldo_prestamo"] = [saldo_prestamo[a * 12 + 11] for a in range(3)]
    tesoreria(m, anual)
    return {"mensual": m, "anual": anual}


def suma(lista, a):
    return sum(lista[a * 12:(a + 1) * 12])


def resumen_anual(m):
    A = {}
    for k in ["altas", "ing_setup", "ing_suscripcion", "ing_plantillas", "ing_cursos", "ing_app",
              "ingresos", "c_variables", "c_fijos", "marketing", "reta", "puesta_marcha",
              "mant_app", "amortizacion", "intereses"]:
        A[k] = [suma(m[k], a) for a in range(3)]
    A["activos_fin"] = [m["activos"][a * 12 + 11] for a in range(3)]
    A["usuarios_app_fin"] = [m["usuarios_app"][a * 12 + 11] for a in range(3)]
    A["horas_max"] = [max(m["horas"][a * 12:(a + 1) * 12]) for a in range(3)]
    A["margen_contribucion"] = [A["ingresos"][a] - A["c_variables"][a] for a in range(3)]
    A["gastos_explotacion"] = [A["c_fijos"][a] + A["marketing"][a] + A["reta"][a]
                              + A["puesta_marcha"][a] + A["mant_app"][a] for a in range(3)]
    A["ebitda"] = [A["margen_contribucion"][a] - A["gastos_explotacion"][a] for a in range(3)]
    A["ebit"] = [A["ebitda"][a] - A["amortizacion"][a] for a in range(3)]
    A["rn_previo"] = [A["ebit"][a] - A["intereses"][a] for a in range(3)]
    A["dificil_just"] = [min(max(x, 0) * H.DIFICIL_JUSTIFICACION, 2000) for x in A["rn_previo"]]
    A["rendimiento_neto"] = [A["rn_previo"][a] - A["dificil_just"][a] for a in range(3)]
    A["irpf"] = [irpf(x) for x in A["rendimiento_neto"]]
    A["resultado_neto"] = [A["rn_previo"][a] - A["irpf"][a] for a in range(3)]
    A["tipo_efectivo"] = [A["irpf"][a] / A["rn_previo"][a] if A["rn_previo"][a] > 0 else 0 for a in range(3)]
    # IRPF sin el efecto de la deuda, para el flujo de caja libre del proyecto
    A["irpf_operativo"] = [irpf(A["ebit"][a] - min(max(A["ebit"][a], 0) * H.DIFICIL_JUSTIFICACION, 2000))
                           for a in range(3)]
    return A


def tesoreria(m, A):
    """Cash-flow mensual con IVA, retenciones, modelo 130 y liquidación anual."""
    M = H.MESES
    caja = H.APORTACION_INICIAL
    inv_mes = [0.0] * (M + 1)
    for _, importe, mes, _, _ in H.INVERSIONES:
        inv_mes[mes] += importe
    caja -= inv_mes[0] * (1 + H.IVA)          # IVA de la inversión inicial
    iva_soportado_ini = inv_mes[0] * H.IVA
    iva_trim, iva_pend = -iva_soportado_ini, 0.0
    ret_acum, pagos_130, rn_acum = 0.0, 0.0, 0.0
    liquidacion_pendiente = 0.0
    cuota_p = cuota_prestamo(H.PRESTAMO, H.TIPO_PRESTAMO, H.PLAZO_PRESTAMO)
    saldo_caja, cobros_l, pagos_l = [], [], []
    iva_fin, hp_irpf_fin = [0.0] * 3, [0.0] * 3
    flujo = {k: [0.0] * 3 for k in ["cobros_explotacion", "pagos_explotacion", "iva_neto", "irpf_pagado",
                                    "inversion", "aportacion", "prestamo", "devolucion_prestamo",
                                    "retiradas"]}
    flujo["aportacion"][0] = H.APORTACION_INICIAL
    flujo["inversion"][0] = inv_mes[0] * (1 + H.IVA)
    for t in range(M):
        a = t // 12
        servicio = m["ing_setup"][t] + m["ing_suscripcion"][t]
        ret = servicio * H.CUOTA_CLIENTES_RETENCION * H.RETENCION
        cobros = m["ingresos"][t] * (1 + H.IVA) - ret
        ipc = (1 + H.INFLACION_COSTES) ** (t // 12)
        fijos_iva = sum(v for k, v in H.COSTES_FIJOS.items() if k in H.CON_IVA) * ipc
        gastos_iva = (fijos_iva
                      + m["marketing"][t] + m["mant_app"][t] + m["puesta_marcha"][t]
                      + m["c_variables"][t])
        gastos_sin_iva = m["c_fijos"][t] - fijos_iva + m["reta"][t]
        pagos = gastos_iva * (1 + H.IVA) + gastos_sin_iva + m["intereses"][t]
        inv = sum(imp for _, imp, mes, _, _ in H.INVERSIONES if mes == t + 1)
        iva_trim += (m["ingresos"][t] - gastos_iva - inv) * H.IVA
        caja += cobros - pagos - inv * (1 + H.IVA) - m_retirada(t)
        flujo["cobros_explotacion"][a] += cobros
        flujo["pagos_explotacion"][a] += pagos
        flujo["inversion"][a] += inv * (1 + H.IVA)
        flujo["retiradas"][a] += m_retirada(t)
        if t + 1 == H.MES_PRESTAMO:
            caja += H.PRESTAMO
            flujo["prestamo"][a] += H.PRESTAMO
        elif t + 1 > H.MES_PRESTAMO:
            caja -= cuota_p - m["intereses"][t]
            flujo["devolucion_prestamo"][a] += cuota_p - m["intereses"][t]
        # cierre de trimestre: liquidación de IVA del trimestre anterior y modelo 130
        if t % 3 == 0 and t > 0:
            caja -= iva_pend
            flujo["iva_neto"][a] += iva_pend
            iva_pend = 0.0
        ret_acum += ret
        rn_acum += m["ingresos"][t] - m["c_variables"][t] - m["c_fijos"][t] - m["marketing"][t] \
            - m["reta"][t] - m["mant_app"][t] - m["puesta_marcha"][t] - m["amortizacion"][t] - m["intereses"][t]
        if t % 3 == 2:
            iva_pend, iva_trim = max(iva_trim, 0), min(iva_trim, 0)  # el IVA negativo se compensa
            pago = max(H.PAGO_FRACCIONADO * rn_acum - ret_acum - pagos_130, 0)
            caja -= pago
            pagos_130 += pago
            flujo["irpf_pagado"][a] += pago
        # junio: liquidación de la renta del año anterior
        if t % 12 == 5 and a > 0:
            caja -= liquidacion_pendiente
            flujo["irpf_pagado"][a] += liquidacion_pendiente
            liquidacion_pendiente = 0.0
        if t % 12 == 11:
            liquidacion_pendiente = A["irpf"][a] - ret_acum - pagos_130
            iva_fin[a] = iva_pend
            hp_irpf_fin[a] = liquidacion_pendiente
            flujo.setdefault("retenciones", [0.0] * 3)[a] = ret_acum
            ret_acum, pagos_130, rn_acum = 0.0, 0.0, 0.0
        saldo_caja.append(caja)
    m["caja"] = saldo_caja
    A["caja_fin"] = [saldo_caja[a * 12 + 11] for a in range(3)]
    A["caja_min"] = min(saldo_caja)
    A["mes_caja_min"] = saldo_caja.index(A["caja_min"]) + 1
    A["iva_pendiente_fin"] = iva_fin
    A["hp_irpf_fin"] = hp_irpf_fin
    A["flujo"] = flujo


def m_retirada(t):
    return H.RETIRADAS[t]


def balance(res):
    A = res["anual"]
    filas = []
    inmov_bruto = {"intangible": [0.0] * 3, "material": [0.0] * 3}
    amort_acum = [0.0] * 3
    for a in range(3):
        for _, imp, mes, _, tipo in H.INVERSIONES:
            if mes <= a * 12 + 12:
                inmov_bruto[tipo][a] += imp
        amort_acum[a] = sum(A["amortizacion"][:a + 1])
    for a in range(3):
        aport = H.APORTACION_INICIAL
        resultados = sum(A["resultado_neto"][:a + 1])
        retiradas = sum(A["flujo"]["retiradas"][:a + 1])
        hp_irpf = A["hp_irpf_fin"][a]
        # IRPF: si la liquidación sale a devolver es un activo
        pasivo_irpf = max(hp_irpf, 0)
        activo_irpf = max(-hp_irpf, 0)
        inmov_neto = inmov_bruto["intangible"][a] + inmov_bruto["material"][a] - amort_acum[a]
        activo = inmov_neto + A["caja_fin"][a] + activo_irpf
        pn = aport + resultados - retiradas
        pasivo_nc = max(A["saldo_prestamo"][a] - _amort_prox(a), 0) if A["saldo_prestamo"][a] else 0
        pasivo_c_prestamo = A["saldo_prestamo"][a] - pasivo_nc
        pasivo = pasivo_nc + pasivo_c_prestamo + A["iva_pendiente_fin"][a] + pasivo_irpf
        filas.append(dict(
            intangible=inmov_bruto["intangible"][a], material=inmov_bruto["material"][a],
            amort_acum=-amort_acum[a], inmov_neto=inmov_neto, hp_deudora=activo_irpf,
            tesoreria=A["caja_fin"][a], activo=activo, aportaciones=aport,
            resultados=resultados, retiradas=-retiradas, pn=pn, deuda_lp=pasivo_nc,
            deuda_cp=pasivo_c_prestamo, hp_iva=A["iva_pendiente_fin"][a], hp_irpf=pasivo_irpf,
            pasivo=pasivo, descuadre=activo - pn - pasivo))
    return filas


def _amort_prox(a):
    """Capital del préstamo que vence en los 12 meses siguientes al cierre del año a."""
    cuota = cuota_prestamo(H.PRESTAMO, H.TIPO_PRESTAMO, H.PLAZO_PRESTAMO)
    saldo = H.PRESTAMO
    i = H.TIPO_PRESTAMO / 12
    meses_pagados = (a + 1) * 12 - H.MES_PRESTAMO
    if meses_pagados < 0:
        return 0
    for _ in range(meses_pagados):
        saldo -= cuota - saldo * i
    prox = 0
    for _ in range(12):
        c = cuota - saldo * i
        prox += c
        saldo -= c
    return prox


def valorar(res, tasa):
    A = res["anual"]
    inv0 = sum(imp for _, imp, mes, _, _ in H.INVERSIONES if mes == 0) + H.GASTOS_PUESTA_EN_MARCHA
    fondo_maniobra = H.APORTACION_INICIAL - inv0 - sum(i * H.IVA for _, i, mes, _, _ in H.INVERSIONES if mes == 0)
    inv_anual = [sum(imp for _, imp, mes, _, _ in H.INVERSIONES if a * 12 < mes <= (a + 1) * 12) for a in range(3)]
    # NOF: saldos con Hacienda que financian la actividad (IVA e IRPF pendientes)
    nof = [-(A["iva_pendiente_fin"][a] + A["hp_irpf_fin"][a]) for a in range(3)]
    var_nof = [nof[0]] + [nof[a] - nof[a - 1] for a in (1, 2)]
    ebitda_sin_pm = [A["ebitda"][a] + (H.GASTOS_PUESTA_EN_MARCHA if a == 0 else 0) for a in range(3)]
    fcl = [ebitda_sin_pm[a] - A["irpf_operativo"][a] - inv_anual[a] - var_nof[a] for a in range(3)]
    fcl[2] += fondo_maniobra   # recuperación del colchón de tesorería inicial
    flujos = [-(inv0 + fondo_maniobra)] + fcl
    van = sum(f / (1 + tasa) ** i for i, f in enumerate(flujos))
    vr = (fcl[2] - fondo_maniobra) * (1 + H.CRECIMIENTO_PERPETUO) / (tasa - H.CRECIMIENTO_PERPETUO)
    van_vr = van + vr / (1 + tasa) ** 3
    flujos_oport = [flujos[0]] + [fcl[a] - H.COSTE_OPORTUNIDAD[a] for a in range(3)]
    van_oport = sum(f / (1 + tasa) ** i for i, f in enumerate(flujos_oport))
    return dict(tasa=tasa, inversion_inicial=inv0, fondo_maniobra=fondo_maniobra,
                inversiones=inv_anual, var_nof=var_nof, fcl=fcl, flujos=flujos, van=van,
                valor_residual=vr, van_con_vr=van_vr, tir=tir(flujos),
                payback=payback(flujos), payback_desc=payback(flujos, tasa),
                flujos_oportunidad=flujos_oport, van_oportunidad=van_oport,
                tir_oportunidad=tir(flujos_oport))


def tir(flujos):
    lo, hi = -0.99, 100.0
    f = lambda r: sum(x / (1 + r) ** i for i, x in enumerate(flujos))
    if f(lo) * f(hi) > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        if f(lo) * f(mid) <= 0:
            hi = mid
        else:
            lo = mid
    return mid


def payback(flujos, tasa=0.0):
    acum = flujos[0]
    for i in range(1, len(flujos)):
        f = flujos[i] / (1 + tasa) ** i
        if acum + f >= 0:
            return i - 1 + (-acum / f)
        acum += f
    return None


def punto_equilibrio(res, a, retribucion=0.0):
    A = res["anual"]
    fijos = A["gastos_explotacion"][a] + A["amortizacion"][a] + A["intereses"][a] + retribucion
    margen_cliente = (H.PRECIO_SUSCRIPCION * (1 - H.COMISION_COBRO)) * 12
    return fijos / margen_cliente


def ratios(res, bal):
    A = res["anual"]
    out = []
    for a in range(3):
        b = bal[a]
        pc = b["deuda_cp"] + b["hp_iva"] + b["hp_irpf"]
        out.append(dict(
            margen_ebitda=A["ebitda"][a] / A["ingresos"][a],
            margen_neto=A["resultado_neto"][a] / A["ingresos"][a],
            roa=A["ebit"][a] / b["activo"],
            roe=A["resultado_neto"][a] / b["pn"] if b["pn"] > 0 else None,
            liquidez=(b["tesoreria"] + b["hp_deudora"]) / pc if pc else None,
            endeudamiento=b["pasivo"] / (b["pasivo"] + b["pn"]),
            cobertura_intereses=A["ebit"][a] / A["intereses"][a] if A["intereses"][a] else None,
            punto_equilibrio=punto_equilibrio(res, a),
            margen_seguridad=1 - (punto_equilibrio(res, a) * H.PRECIO_SUSCRIPCION * 12)
            / (A["ing_suscripcion"][a]) if A["ing_suscripcion"][a] else None,
        ))
    return out


def sensibilidad():
    filas = {}
    filas["precio"] = {p: simular("base", precio=p)["valoracion"]["van"] for p in (55, 65, 75, 85, 95)}
    filas["altas"] = {f: simular("base", factor_altas=f)["valoracion"]["van"] for f in (0.5, 0.75, 1.0, 1.25)}
    filas["tasa"] = {t: simular("base", tasa=t)["valoracion"]["van"] for t in (0.10, 0.15, 0.20, 0.30)}
    r = simular("base", indexar_precios=True)["valoracion"]
    filas["precios_indexados"] = {"van": r["van"], "van_oportunidad": r["van_oportunidad"]}
    return filas


def main():
    salida = {}
    for nombre in ESCENARIOS:
        r = simular(nombre)
        bal = balance(r)
        salida[nombre] = dict(anual=r["anual"], balance=bal, ratios=ratios(r, bal),
                              valoracion=r["valoracion"], reta=r["reta_mensual"],
                              mensual=r["mensual"],
                              pe=[punto_equilibrio(r, a) for a in range(3)],
                              pe_smi=[punto_equilibrio(r, a, H.SMI_ANUAL) for a in range(3)])
    salida["sensibilidad"] = sensibilidad()
    salida["tasa"] = H.TASA_DESCUENTO
    Path(__file__).with_name("resultados.json").write_text(json.dumps(salida, indent=1, default=str))
    b = salida["base"]
    for k in ["ingresos", "ebitda", "rendimiento_neto", "irpf", "resultado_neto", "caja_fin", "activos_fin", "horas_max"]:
        print(f"{k:18}", [round(x) for x in b["anual"][k]])
    print("caja min", round(b["anual"]["caja_min"]), "mes", b["anual"]["mes_caja_min"])
    print("reta", b["reta"], "descuadre", [round(x["descuadre"], 2) for x in b["balance"]])
    for n in ESCENARIOS:
        v = salida[n]["valoracion"]
        print(n, "VAN", round(v["van"]), "VAN+VR", round(v["van_con_vr"]), "TIR", v["tir"], "PB", v["payback"],
              "PBd", v["payback_desc"], "VANop", round(v["van_oportunidad"]), "flujos", [round(x) for x in v["flujos"]],
              "cajamin", round(salida[n]["anual"]["caja_min"]))
    print("PE", [round(x, 1) for x in b["pe"]], "PE SMI", [round(x, 1) for x in b["pe_smi"]])
    print(salida["sensibilidad"])


if __name__ == "__main__":
    main()
