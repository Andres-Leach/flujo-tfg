"""Hipótesis del plan económico-financiero de Flujo.

Todas las cifras van en euros y sin IVA salvo que se indique lo contrario.
El horizonte es de 36 meses (año 1 = 2027, año 2 = 2028, año 3 = 2029).
Cada valor lleva al lado de dónde sale, para poder revisarlo.
"""

MESES = 36

# --- Precios -----------------------------------------------------------------
PRECIO_SETUP = 200          # centro del rango 100-300 € (apartado 5.3)
PRECIO_SUSCRIPCION = 75     # centro del rango 50-100 €/mes
PRECIO_PLANTILLA = 15       # neto de IVA (≈ 18 € con IVA), rango 10-30 €
PRECIO_CURSO = 35           # neto de IVA (≈ 42 € con IVA), rango 20-60 €
PRECIO_APP = 12             # €/mes neto, rango 10-25 €

# --- Captación y retención (escenario base) ----------------------------------
# Altas de clientes de suscripción por mes. El año 1 sigue el cronograma de
# validación: los dos primeros meses son de contacto y diagnóstico gratuito.
ALTAS = (
    [0, 1, 1, 1, 1, 2, 1, 1, 2, 1, 2, 1]      # año 1: 14 altas
    + [2, 2, 3, 2, 2, 3, 2, 2, 3, 3, 3, 3]    # año 2: 30 altas
    + [3] * 12                                 # año 3: 36 altas
)
BAJA_MENSUAL = 0.04         # 4 % mensual ≈ 61 % de retención a doce meses

# Capacidad del fundador: horas de servicio disponibles al mes
HORAS_CLIENTE_MES = 2.5     # reunión mensual + revisión del sistema
HORAS_CLIENTE_ANUAL = [HORAS_CLIENTE_MES] * 3   # se reduce si se estandariza el servicio
HORAS_SETUP = 4             # diagnóstico y montaje inicial
HORAS_DISPONIBLES = [60] * 12 + [120] * 24   # año 1 a media jornada

# --- Productos digitales (desde el mes 13) -----------------------------------
def rampa(inicio, fin, n):
    return [inicio + (fin - inicio) * i / (n - 1) for i in range(n)]

UDS_PLANTILLAS = [0] * 12 + rampa(5, 25, 12) + rampa(25, 40, 12)
UDS_CURSOS = [0] * 12 + rampa(2, 10, 12) + rampa(10, 18, 12)
COMISION_PLATAFORMA = 0.10  # pasarela de pago + plataforma de cursos

# --- Palancas de mejora (desactivadas en el escenario base) -------------------
TALLERES_MES = [0] * 36     # talleres grupales para instituciones y coworkings
PRECIO_TALLER = 250         # € netos por sesión de 2 h
HORAS_TALLER = 5            # preparación + impartición + desplazamiento
TASA_DERIVACION = 0.0       # parte de las altas derivadas a una gestoría colaboradora
COMISION_DERIVACION = 60    # € por cliente derivado (pago único)

# --- Herramienta digital (SaaS), lanzamiento en el mes 31 --------------------
APP_ACTIVA = True
MES_LANZAMIENTO_APP = 31
USUARIOS_APP = [0] * 30 + [15 + 12 * i for i in range(6)]
COSTE_VARIABLE_APP = 1.5    # alojamiento y servicios por usuario y mes
MANTENIMIENTO_APP = 150     # €/mes desde el lanzamiento
DESARROLLO_APP = 12_000     # MVP con desarrollador freelance, meses 25-30

COMISION_COBRO = 0.015      # domiciliación / pasarela sobre el servicio

# --- Costes fijos mensuales --------------------------------------------------
COSTES_FIJOS = {
    "Herramientas de gestión y facturación": 40,
    "Diseño y redes sociales": 30,
    "Web, dominio y alojamiento": 10,
    "Seguro de responsabilidad civil": 25,
    "Asesor fiscal colegiado de apoyo": 40,
    "Teléfono e internet (parte afecta)": 25,
    "Formación continua": 25,
}
MARKETING = [50] * 12 + [150] * 12 + [300] * 12

# Gastos que soportan IVA (el seguro y la formación reglada están exentos)
CON_IVA = {
    "Herramientas de gestión y facturación", "Diseño y redes sociales",
    "Web, dominio y alojamiento", "Asesor fiscal colegiado de apoyo",
    "Teléfono e internet (parte afecta)",
}

# --- Inversión inicial y posteriores -----------------------------------------
# (concepto, importe, mes de compra, vida útil en años, tipo)
def inversiones():
    base = [
        ("Página web con agenda y pagos", 350, 0, 3, "intangible"),
        ("Registro de marca en la OEPM", 125, 0, 10, "intangible"),
        ("Equipo de grabación para cursos", 450, 13, 4, "material"),
    ]
    if APP_ACTIVA:
        base += [("Desarrollo de la herramienta (MVP)", DESARROLLO_APP / 6, m, 3, "intangible")
                 for m in range(25, 31)]
    return base


INVERSIONES = inversiones()
GASTOS_PUESTA_EN_MARCHA = 350   # logo y plantillas (150) + textos legales RGPD (200)

# --- Financiación ------------------------------------------------------------
APORTACION_INICIAL = 3_000      # ahorros propios
PRESTAMO = 10_000               # microcrédito para el MVP (mes 25)
MES_PRESTAMO = 25
TIPO_PRESTAMO = 0.07            # TIN anual, referencia MicroBank
PLAZO_PRESTAMO = 60             # meses, sistema francés

# Retiradas mensuales del titular para vivir (no son gasto de la actividad)
RETIRADAS = [300] * 12 + [900] * 12 + [1_200] * 12

# --- Fiscalidad --------------------------------------------------------------
IVA = 0.21
TARIFA_PLANA = 80
SMI_ANUAL = 17_094              # RD 126/2026
RETENCION = 0.07                # tipo reducido de inicio de actividad
CUOTA_CLIENTES_RETENCION = 0.6  # parte de clientes que son autónomos y retienen
PAGO_FRACCIONADO = 0.20         # modelo 130
DIFICIL_JUSTIFICACION = 0.05    # estimación directa simplificada, máx. 2.000 €
MINIMO_PERSONAL = 5_550

ESCALA_ESTATAL = [(12_450, .095), (20_200, .12), (35_200, .15), (60_000, .185),
                  (300_000, .225), (float("inf"), .245)]
# Escala autonómica de Aragón (AEAT, Manual práctico de Renta 2025)
ESCALA_ARAGON = [(13_072.5, .095), (21_210, .12), (36_960, .15), (52_500, .185),
                 (60_000, .205), (80_000, .23), (90_000, .24), (130_000, .25),
                 (float("inf"), .255)]

# Cuotas de autónomos 2026: base mínima de cada tramo (Orden PJC/297/2026, iguales
# a las de 2025) por el tipo total del 31,5 % (28,30 % comunes + 1,30 % profesionales
# + 0,90 % cese de actividad + 0,10 % formación profesional + 0,90 % MEI)
TIPO_RETA = 0.315
BASES_RETA = [(670, 653.59), (900, 718.95), (1_166.7, 849.67), (1_300, 950.98),
              (1_500, 960.78), (1_700, 960.78), (1_850, 1_143.79), (2_030, 1_209.15),
              (2_330, 1_274.51), (2_760, 1_356.21), (3_190, 1_437.91), (3_620, 1_519.61),
              (4_050, 1_601.31), (6_000, 1_732.03), (float("inf"), 1_928.10)]
TRAMOS_RETA = [(limite, round(base * TIPO_RETA, 2)) for limite, base in BASES_RETA]

# Inflación: los costes se actualizan con la previsión del Banco de España para 2027
# (IAPC 3,7 %, proyecciones de octubre de 2026). Los precios de Flujo se mantienen
# constantes en el escenario base, por prudencia.
INFLACION_COSTES = 0.037
INDEXAR_PRECIOS = False

# --- Valoración --------------------------------------------------------------
TASA_LIBRE_RIESGO = 0.034       # bono español a 10 años, julio de 2026
PRIMA_MERCADO = 0.055           # prima de riesgo del mercado español
BETA = 1.0                      # servicios profesionales, sin deuda
PRIMA_ESPECIFICA = 0.061        # tamaño, iliquidez y dependencia del fundador
TASA_DESCUENTO = TASA_LIBRE_RIESGO + BETA * PRIMA_MERCADO + PRIMA_ESPECIFICA  # 15 %
CRECIMIENTO_PERPETUO = 0.01
# Coste de oportunidad: SMI neto de cotización (6,5 %) como sueldo alternativo
COSTE_OPORTUNIDAD = [0.5 * SMI_ANUAL * 0.935, SMI_ANUAL * 0.935, SMI_ANUAL * 0.935]
