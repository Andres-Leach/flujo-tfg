// Genera TFG_Flujo_v2.docx con el formato de las directrices FECEM (junio de 2023)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, Footer, PageNumber,
  TableOfContents, LevelFormat, ImageRun, SectionType, NumberFormat,
} = require("docx");
const { R, eur, pct, num, miles } = require("./fmt");
const { FUENTE, p, h1, h2, vinetas, salto, leyenda, fuente, tabla, imagen } = require("./helpers");

const B = R.base, A = B.anual, V = B.valoracion;
const PE = R.pesimista, OP = R.optimista;
const anios = ["Año 1", "Año 2", "Año 3"];
const fila = (nombre, arr, f = eur) => [nombre, ...arr.map((x) => f(x))];
const neg = (arr) => arr.map((x) => -x);
const sum = (a) => a.reduce((x, y) => x + y, 0);
const r0 = (x) => Math.round(x);
const meses = (anios_) => Math.round(anios_ * 12);
const W4 = [4270, 1600, 1600, 1600];

// ---------------------------------------------------------------- PORTADA
const logo = fs.readFileSync(path.join(__dirname, "logo_unizar.png"));
const centro = (texto, size, extra = {}) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: 0, after: 120 },
  children: [new TextRun({ text: texto, size, font: FUENTE, ...extra })],
});
const portada = [
  new Paragraph({ alignment: AlignmentType.LEFT, children: [new ImageRun({ type: "png", data: logo, transformation: { width: 240, height: 67 } })] }),
  new Paragraph({ spacing: { before: 1200 }, children: [] }),
  centro("Trabajo Fin de Grado", 36, { bold: true }),
  new Paragraph({ spacing: { before: 300 }, children: [] }),
  centro("Viabilidad financiera de Flujo: una start-up de ayuda a jóvenes emprendedores", 32, { bold: true }),
  new Paragraph({ spacing: { before: 200 }, children: [] }),
  centro("Financial feasibility of Flujo: a start-up supporting young entrepreneurs", 26, { italics: true }),
  new Paragraph({ spacing: { before: 700 }, children: [] }),
  centro("Autor", 24, { bold: true }), centro("Andrés Leach Sagarra", 24),
  new Paragraph({ spacing: { before: 200 }, children: [] }),
  centro("Directora", 24, { bold: true }), centro("Dra. Beatrice Boumda Ijang", 24),
  new Paragraph({ spacing: { before: 700 }, children: [] }),
  centro("Grado en Administración y Dirección de Empresas", 24),
  centro("Facultad de Economía y Empresa", 24),
  centro("Curso 2026-2027", 24),
];

// ---------------------------------------------------------- INFORMACIÓN Y RESUMEN
const resumenEs = `Los jóvenes que trabajan por cuenta propia tienen que gestionar ingresos irregulares, impuestos trimestrales y gastos casi siempre sin formación financiera. La Encuesta de Competencias Financieras del Banco de España muestra que solo el 19 % de los adultos responde bien a las tres preguntas financieras básicas, y quince entrevistas exploratorias a autónomos de 20 a 35 años apuntan en la misma dirección. El objetivo de este trabajo es evaluar la viabilidad económico-financiera de Flujo, un proyecto que ofrecería organización financiera, educación práctica y, más adelante, una herramienta digital a este colectivo. Para ello se combinan fuentes oficiales, las entrevistas, el Business Model Canvas, un análisis de mercado y un plan económico-financiero a tres años con escenarios. En el escenario base los ingresos pasan de ${eur(A.ingresos[0])} a ${eur(A.ingresos[2])}, el punto de equilibrio se sitúa en ${Math.ceil(B.pe[0])} clientes de suscripción y el VAN a tres años, descontado al ${pct(R.tasa, 0)}, es de ${eur(V.van)}. Sin embargo, cuando se descuenta el coste de oportunidad del emprendedor (un sueldo equivalente al SMI), el VAN baja a ${eur(V.van_oportunidad)}. Cinco medidas de bajo coste orientadas a aumentar los ingresos por hora lo elevan a ${eur(R.palancas[R.palancas.length - 1].van_oportunidad)}, y en una simulación de Montecarlo la probabilidad de que el proyecto supere al empleo alternativo pasa del ${pct(R.montecarlo.p_van_op, 0)} al ${pct(R.montecarlo_opt.p_van_op, 0)}. El proyecto es viable, aunque depende de la captación de clientes y de la capacidad de una sola persona, y la disposición real a pagar aún debe validarse.`;
const resumenEn = `Young self-employed workers have to manage irregular income, quarterly taxes and expenses, usually without any financial training. The Bank of Spain's Financial Competences Survey shows that only 19% of adults answer the three basic financial questions correctly, and fifteen exploratory interviews with self-employed people aged 20 to 35 point in the same direction. This study assesses the economic and financial feasibility of Flujo, a project that would offer financial organisation, practical education and, later, a digital tool to this group. It combines official sources, the interviews, the Business Model Canvas, a market analysis and a three-year financial plan with scenarios. In the base scenario revenue grows from ${eur(A.ingresos[0])} to ${eur(A.ingresos[2])}, break-even is reached with ${Math.ceil(B.pe[0])} subscription clients and the three-year NPV, discounted at ${pct(R.tasa, 0)}, is ${eur(V.van)}. Once the founder's opportunity cost (a salary equal to the Spanish minimum wage) is deducted, the NPV falls to ${eur(V.van_oportunidad)}. Five low-cost measures aimed at raising revenue per hour lift it to ${eur(R.palancas[R.palancas.length - 1].van_oportunidad)}, and in a Monte Carlo simulation the probability that the project beats the alternative job rises from ${pct(R.montecarlo.p_van_op, 0)} to ${pct(R.montecarlo_opt.p_van_op, 0)}. The project is feasible, although it depends on client acquisition and on the capacity of a single person, and real willingness to pay still has to be validated.`;
const dato = (et, val) => new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: et + ": ", bold: true }), new TextRun(val)] });
const info = [
  dato("Autor del trabajo", "Andrés Leach Sagarra"),
  dato("Directora del trabajo", "Dra. Beatrice Boumda Ijang"),
  dato("Título del trabajo", "Viabilidad financiera de Flujo: una start-up de ayuda a jóvenes emprendedores"),
  dato("Title", "Financial feasibility of Flujo: a start-up supporting young entrepreneurs"),
  dato("Titulación", "Grado en Administración y Dirección de Empresas (Universidad de Zaragoza)"),
  new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: "Resumen", bold: true })] }),
  p(resumenEs, { spacing: { line: 276, lineRule: "auto", after: 80 }, run: { size: 22 } }),
  new Paragraph({ spacing: { before: 120, after: 60 }, children: [new TextRun({ text: "Abstract", bold: true })] }),
  p(resumenEn, { spacing: { line: 276, lineRule: "auto", after: 0 }, run: { size: 22 } }),
];

// ------------------------------------------------------------------ ÍNDICE
const indice = [
  new Paragraph({ pageBreakBefore: true, spacing: { after: 240 }, children: [new TextRun({ text: "Índice", bold: true, size: 28 })] }),
  new TableOfContents("Índice", { hyperlink: true, headingStyleRange: "1-2" }),
];

// ---------------------------------------------------------- 1. INTRODUCCIÓN
const cap1 = [
  h1("1. Introducción"),
  h2("1.1. Presentación y justificación del tema"),
  p(`El trabajo por cuenta propia tiene en España un peso que pocos países europeos igualan. El Régimen Especial de Trabajadores Autónomos (RETA) cerró 2025 con 3.425.767 personas afiliadas, la cifra más alta de su serie, tras sumar unas 39.000 altas netas en el año (ATA, 2026), y en septiembre de 2026 la afiliación media alcanzaba ya 3.483.645 personas, un 2,06 % más que un año antes (The Objective, 2026). Detrás de esas cifras hay muchas actividades distintas, pero casi todas comparten un mismo problema de fondo: la persona que factura es también la que tiene que llevar las cuentas, calcular el IVA y el IRPF de cada trimestre y decidir cuánto puede gastar sin poner en riesgo su actividad.`),
  p(`Esa tarea exige unos conocimientos que la mayoría de la población no tiene. Según la Encuesta de Competencias Financieras (ECF) del Banco de España, con datos recogidos entre 2021 y 2022, solo un 41 % de los adultos responde bien a una pregunta sencilla sobre interés compuesto y apenas un 19 % acierta a la vez las preguntas sobre inflación, interés compuesto y diversificación del riesgo (Banco de España, 2022). El problema está reconocido institucionalmente: el propio Banco de España ha reanalizado la encuesta como base del Plan de Educación Financiera 2026-2029 (Banco de España, 2025). Entre los ocupados, tanto por cuenta propia como ajena, entre un 38 % y un 42 % declara tener conocimientos financieros bajos o muy bajos (Banco de España, 2022), y en el plano internacional solo un 34 % de los adultos de los 39 países de la última encuesta de la OCDE alcanza la puntuación mínima de alfabetización financiera (OCDE, 2023).`),
  p(`En el caso concreto de quienes emprenden, el Observatorio del Emprendimiento de España ha analizado las respuestas de 128.598 personas encuestadas por el GEM entre 2021 y 2024 y concluye que el nivel de conocimiento financiero en España es «moderadamente bajo»: en una escala de 1 a 5, los emprendedores potenciales puntúan 2,76 y los que han emprendido recientemente 2,99 (San-Martín *et al.*, 2025). Al mismo tiempo, ocho de cada diez emprendedores consideran tener los conocimientos y habilidades necesarios para emprender (Observatorio del Emprendimiento de España, 2025), una autopercepción que encaja con lo que la literatura llama exceso de confianza financiera (Lusardi y Mitchell, 2014).`),
  p(`Este trabajo parte de esa distancia entre lo que se sabe y lo que se cree saber. En una fase exploratoria realicé quince entrevistas a autónomos y *freelancers* de entre 20 y 35 años, y en más de la mitad apareció la misma idea, formulada de distintas maneras: no sabían cuánto ganaban realmente una vez descontados impuestos y gastos. La mayoría mezclaba el dinero personal y el de la actividad, y varios reconocieron no saber cuánto debían reservar cada trimestre para Hacienda. La muestra es pequeña y no es representativa, pero sus conclusiones van en la misma dirección que los datos agregados.`),
  p(`A partir de esa necesidad detectada se plantea Flujo, un proyecto que ofrecería a este colectivo un servicio de organización financiera, formación práctica y, a medio plazo, una herramienta digital. El trabajo no pretende presentar Flujo como una idea ya validada, sino responder a una pregunta concreta: **habiendo detectado una necesidad en el mercado, ¿sería económica y financieramente viable un proyecto que se dedicara a cubrirla?**`),
  h2("1.2. Objetivos"),
  p(`El **objetivo general** es evaluar la viabilidad económico-financiera de Flujo durante sus tres primeros años de actividad. Para alcanzarlo se plantean los siguientes **objetivos específicos**:`),
  ...vinetas([
    "Describir la situación del colectivo de autónomos jóvenes en España y sus principales dificultades financieras.",
    "Definir la propuesta de valor y el modelo de negocio de Flujo a partir de esas dificultades.",
    "Analizar el mercado, la competencia y el entorno legal y fiscal en el que operaría el proyecto.",
    "Estimar la inversión inicial, la financiación necesaria, los ingresos y la estructura de costes.",
    "Elaborar la cuenta de resultados, el estado de flujos de tesorería y el balance previsionales a tres años.",
    "Calcular el VAN, la TIR, el periodo de recuperación y los principales ratios financieros.",
    "Proponer y cuantificar medidas que mejoren la rentabilidad del proyecto.",
    "Analizar los riesgos del proyecto y su comportamiento en distintos escenarios.",
  ]),
  h2("1.3. Trascendencia del trabajo"),
  p(`El interés del trabajo es triple. En el plano **empresarial**, ofrece un análisis de viabilidad que podría servir de base real para lanzar el proyecto, con hipótesis explícitas que se pueden contrastar a medida que lleguen los primeros clientes. En el plano **social**, aborda un problema, la baja educación financiera, que las instituciones consideran prioritario y que en el caso de los autónomos tiene consecuencias directas: recargos por errores fiscales, falta de colchón de liquidez o decisiones de gasto poco meditadas. En el plano **universitario**, aplica de forma conjunta herramientas de varias asignaturas del grado (dirección estratégica, contabilidad, fiscalidad y finanzas) a un caso concreto, y el modelo de cálculo se publica en abierto para que cualquier lector pueda reproducirlo.`),
  h2("1.4. Metodología y fuentes"),
  p(`El trabajo sigue un enfoque de estudio de caso con tres tipos de fuentes. Las **fuentes secundarias** proceden de organismos oficiales y entidades de referencia: Banco de España, Seguridad Social, Ministerio de Trabajo y Economía Social, Agencia Tributaria, Instituto Nacional de Estadística, Eurostat, OCDE, Observatorio del Emprendimiento de España y la Federación Nacional de Asociaciones de Trabajadores Autónomos (ATA). El Anexo IV resume qué dato se ha contrastado con cada fuente oficial. La **fuente primaria** son las quince entrevistas exploratorias semiestructuradas ya citadas (el guion se recoge en el Anexo III). Por último, el análisis se apoya en **herramientas de gestión** consolidadas: la metodología Lean Startup (Ries, 2011), el Business Model Canvas (Osterwalder y Pigneur, 2010), el análisis DAFO y los criterios clásicos de evaluación de inversiones (Brealey, Myers y Allen, 2020).`),
  p(`Las proyecciones financieras se han calculado mes a mes con un modelo propio programado en Python, que se entrega también en una hoja de cálculo. Todas las hipótesis se explican en el capítulo 6 y el código y los datos están disponibles en un repositorio público (ver Anexo II).`),
  h2("1.5. Estructura del trabajo"),
  p(`Tras esta introducción, el capítulo 2 recoge el marco teórico: el concepto de start-up, el modelo de negocio, la educación financiera y los criterios de viabilidad y financiación. El capítulo 3 presenta la necesidad detectada y la propuesta de Flujo, y el capítulo 4 analiza el mercado, la competencia y el DAFO. El capítulo 5 trata los recursos necesarios y el marco legal y fiscal. Los capítulos 6 a 9 forman el núcleo del trabajo: el plan económico-financiero, el análisis de viabilidad y rentabilidad, las medidas para mejorar esa rentabilidad y el análisis de riesgos y escenarios. El capítulo 10 recoge las conclusiones y las limitaciones. Los anexos incluyen el análisis de datos de mercado, la nota sobre el modelo de cálculo, las entrevistas y la relación de fuentes oficiales consultadas.`),
];

// ---------------------------------------------------------- 2. MARCO TEÓRICO
const cap2 = [
  h1("2. Marco teórico", false),
  h2("2.1. La start-up y el enfoque Lean Startup"),
  p(`Ries (2011) define una start-up como una institución humana diseñada para crear un producto o servicio en condiciones de incertidumbre extrema. Lo que la distingue de una pequeña empresa tradicional no es su tamaño sino que todavía no sabe con certeza quién es su cliente ni cuánto está dispuesto a pagar. Blank (2013) lo resume diciendo que una start-up busca un modelo de negocio repetible y escalable, mientras que una empresa consolidada se limita a ejecutarlo.`),
  p(`De esa incertidumbre nace el enfoque Lean Startup, que propone sustituir los planes de negocio cerrados por ciclos cortos de construir, medir y aprender. La idea central es lanzar un producto mínimo viable, observar cómo responden los clientes y ajustar la propuesta antes de comprometer recursos importantes. Para un proyecto como Flujo, este enfoque tiene una consecuencia práctica: empezar como un servicio prestado de forma manual, que apenas requiere inversión, y dejar el desarrollo tecnológico para cuando exista evidencia de que hay clientes dispuestos a pagar.`),
  h2("2.2. El modelo de negocio como herramienta de análisis"),
  p(`Osterwalder y Pigneur (2010) describen el modelo de negocio como la lógica con la que una organización crea, entrega y captura valor, y proponen representarlo en un lienzo de nueve bloques: segmentos de clientes, propuesta de valor, canales, relación con clientes, fuentes de ingresos, recursos clave, actividades clave, socios clave y estructura de costes. Su utilidad para este trabajo es doble. Por un lado, obliga a explicitar todas las hipótesis del proyecto en un formato que se puede revisar de un vistazo. Por otro, conecta directamente con el plan financiero, porque los bloques de ingresos y costes son el punto de partida de las proyecciones.`),
  h2("2.3. Educación financiera y trabajo autónomo"),
  p(`La OCDE define la educación financiera como la combinación de conocimientos, habilidades, actitudes y comportamientos necesarios para tomar decisiones financieras adecuadas (OCDE, 2020). Lusardi y Mitchell (2014) revisan la evidencia internacional y concluyen que los niveles de conocimiento financiero son bajos incluso en países desarrollados, que los jóvenes y las personas con menos estudios puntúan peor y que un mayor conocimiento se asocia con más ahorro, mejor planificación y menos errores costosos. También señalan que mucha gente sobrestima lo que sabe, lo que reduce su propensión a buscar ayuda.`),
  p(`En el caso de los autónomos el problema se agrava por tres razones. Sus ingresos son irregulares, lo que exige planificar la liquidez. No tienen un empleador que les retenga y cotice por ellos, de modo que la reserva para impuestos y Seguridad Social depende de su propia disciplina. Y la frontera entre el patrimonio personal y el de la actividad es difusa, porque responden con todos sus bienes. Para un asalariado un error financiero suele tener un coste limitado; para un autónomo puede traducirse en recargos, en tensiones de tesorería o en el cierre de la actividad.`),
  p(`La evidencia española confirma que el tamaño importa. La Encuesta de Competencias Financieras en las Pequeñas Empresas, también del Banco de España, muestra que los propietarios de empresas de menos de 20 trabajadores tienen menos conocimientos financieros que los de empresas de 20 a 49, que prácticas como comparar proveedores financieros están menos extendidas en las de menos de cinco trabajadores y que pocos propietarios han pensado en cómo financiar su jubilación. También encuentra que las empresas cuyos propietarios tenían más competencias financieras sufrieron un impacto algo menor en liquidez durante la pandemia (Banco de España, 2021). Es decir, la educación financiera no es solo una cuestión de cultura general: tiene efectos medibles sobre la resistencia de los negocios más pequeños.`),
  h2("2.4. La viabilidad económico-financiera"),
  p(`Un proyecto es **económicamente viable** cuando genera ingresos suficientes para cubrir todos sus costes y remunerar los recursos invertidos, y es **financieramente viable** cuando dispone en cada momento de la liquidez necesaria para atender sus pagos. Las dos condiciones son necesarias: un proyecto rentable puede fracasar por falta de caja, y uno con liquidez puede estar destruyendo valor. Por eso este trabajo analiza tanto la cuenta de resultados como los flujos de tesorería y el balance.`),
  p(`Para valorar la rentabilidad se emplean los criterios clásicos de evaluación de inversiones, que se resumen en la Tabla 2.1. El **valor actual neto (VAN)** descuenta los flujos de caja futuros a una tasa que refleja el riesgo del proyecto y les resta la inversión inicial; un VAN positivo indica que el proyecto crea valor por encima de esa rentabilidad exigida. La **tasa interna de rentabilidad (TIR)** es la tasa que iguala el VAN a cero. El **periodo de recuperación (payback)** mide el tiempo necesario para recuperar la inversión. Brealey, Myers y Allen (2020) recuerdan que el VAN es el criterio más sólido y que la TIR y el payback deben usarse como complementos, porque la TIR puede dar resultados engañosos cuando la inversión inicial es muy pequeña y el payback ignora lo que ocurre después de la recuperación. Esta advertencia resulta importante para Flujo, como se verá en el capítulo 7.`),
  leyenda("Tabla", "2.1", "Criterios de evaluación de inversiones utilizados"),
  tabla(["Criterio", "Qué mide", "Regla de decisión", "Limitación principal"], [
    ["VAN", "Valor creado, en euros de hoy, por encima de la rentabilidad exigida", "Aceptar si VAN > 0", "Depende de la tasa de descuento elegida"],
    ["TIR", "Rentabilidad anual implícita del proyecto", "Aceptar si TIR > tasa exigida", "Se dispara con inversiones iniciales pequeñas"],
    ["Payback", "Tiempo hasta recuperar la inversión", "Cuanto menor, menor riesgo", "No considera los flujos posteriores"],
    ["Punto de equilibrio", "Volumen mínimo para cubrir los costes fijos", "Comparar con la demanda esperada", "Supone precios y costes constantes"],
  ], [1500, 2700, 2100, 2770], { numericas: false }),
  fuente("elaboración propia a partir de Brealey, Myers y Allen (2020)."),
  p(`La tasa de descuento es el elemento más delicado. Siguiendo a Damodaran (2012), puede estimarse con el modelo CAPM, sumando a la rentabilidad sin riesgo una prima de mercado ajustada por la beta del sector. En empresas jóvenes y no cotizadas esa tasa se queda corta, porque no recoge el riesgo de fracaso, la falta de liquidez de la inversión ni la dependencia de una sola persona, y Damodaran (2009) propone añadir primas específicas para esos riesgos. El capítulo 7 aplica este procedimiento a Flujo. Finalmente, los **ratios financieros** de rentabilidad, liquidez y endeudamiento completan el análisis, aunque en una microempresa de servicios con muy pocos activos deben interpretarse con cautela.`),
  h2("2.5. La financiación de la empresa en fase inicial"),
  p(`La teoría de la jerarquía financiera (*pecking order*) de Myers y Majluf (1984) sostiene que las empresas prefieren financiarse primero con recursos propios, después con deuda y solo en último lugar con nuevos socios, porque la información asimétrica encarece la financiación externa. Berger y Udell (1998) completan esta idea con el concepto de ciclo de crecimiento financiero: en las primeras fases la empresa depende de los ahorros del fundador y de su entorno cercano; a medida que crece y genera historial accede a microcréditos, préstamos bancarios y financiación pública, y solo cuando tiene un potencial de crecimiento elevado atrae a *business angels* y fondos de capital riesgo. La Tabla 2.2 resume las fuentes disponibles en España para un proyecto como Flujo.`),
  leyenda("Tabla", "2.2", "Fuentes de financiación según la fase del proyecto"),
  tabla(["Fuente", "Fase", "Condiciones de referencia", "Encaje con Flujo"], [
    ["Ahorros propios", "Inicio", "Sin coste explícito ni garantías", "Alto: inversión inicial reducida"],
    ["Microcrédito (p. ej., MicroBank)", "Inicio y crecimiento", "Hasta 25.000 €, sin aval, interés fijo en torno al 7 %", "Alto para el desarrollo de la herramienta"],
    ["Préstamo bancario con línea ICO Empresas y Emprendedores", "Crecimiento", "De 1 a 20 años, hasta 3 de carencia; se tramita en el banco", "Medio; exige historial y garantías"],
    ["Subvención al establecimiento como autónomo (Aragón)", "Inicio", "Para desempleados inscritos; menores de 30 años, colectivo prioritario", "Incierto; no se incluye en el escenario base"],
    ["ENISA, línea Emprendedores", "Crecimiento", "Préstamo participativo de 25.000 a 300.000 €; sociedad de menos de 24 meses y fondos propios al menos iguales al préstamo", "Solo si se constituye una sociedad"],
    ["Business angels y capital riesgo", "Escalado", "Entrada en el capital; buscan crecimiento rápido", "Bajo en el horizonte analizado"],
  ], [2300, 1400, 2900, 2470], { numericas: false }),
  fuente("elaboración propia a partir de Berger y Udell (1998), Cámara de Comercio de Valencia (2023), ICO (2026), ENISA (2026) y Gobierno de Aragón (2026)."),
];

// ---------------------------------------------------------- 3. EL PROYECTO
const cap3 = [
  h1("3. El proyecto Flujo: necesidad detectada y propuesta", false),
  h2("3.1. La necesidad detectada"),
  p(`Las entrevistas exploratorias se hicieron a autónomos y *freelancers* de entre 20 y 35 años contactados a través de redes personales y comunidades profesionales: diseñadores, creadores de contenido, entrenadores personales, profesionales del marketing digital y otros perfiles creativos. Tres problemas se repitieron con frecuencia. El primero, no saber cuánto se gana realmente una vez descontados el IVA, el IRPF y la cuota de autónomos, algo que mencionaron más de la mitad de los entrevistados. El segundo, la mezcla entre las finanzas personales y las de la actividad, que impide saber qué parte de lo que hay en la cuenta puede gastarse. El tercero, la incertidumbre ante cada cierre trimestral, porque no se reserva de forma sistemática el dinero de los impuestos.`),
  p(`Los datos oficiales confirman que el problema va más allá de esta muestra. En la ECF, el índice de conocimientos básicos es del 52 % en el grupo de 18 a 34 años, frente al 56 % del grupo de 45 a 64 y una media nacional del 53 % (Banco de España, 2022). La diferencia no es muy grande, pero llega justo en la etapa en la que estos jóvenes empiezan a tomar decisiones económicas sin una nómina ni retenciones automáticas (ver Anexo I).`),
  h2("3.2. Propuesta de valor y servicio"),
  p(`La propuesta de valor de Flujo es ayudar a los autónomos jóvenes a entender y controlar su dinero con un lenguaje sencillo, sin tecnicismos y con acompañamiento humano. No pretende sustituir a la gestoría, que presenta los impuestos, sino cubrir lo que la gestoría no suele hacer: explicar al cliente su situación y darle un sistema para decidir. El nombre del proyecto viene de ahí, del flujo de dinero que un autónomo pierde de vista cuando ingresos, gastos e impuestos pasan por la misma cuenta. El servicio se organiza en tres capas que se despliegan de forma progresiva:`),
  ...vinetas([
    "**Servicio de acompañamiento (desde el año 1).** Un diagnóstico inicial y la puesta en marcha de un sistema personalizado de control de ingresos, gastos y reservas para impuestos (*setup*), seguidos de una suscripción mensual con una reunión de seguimiento y la revisión de los números.",
    "**Educación práctica (desde el año 2).** Plantillas y cursos breves basados en casos reales, como cuánto reservar para el IVA y el IRPF o cómo saber si la actividad es rentable.",
    "**Herramienta digital (desde la segunda mitad del año 3).** Un panel sencillo con previsión de tesorería y avisos, que permitiría atender a muchos más usuarios a un precio más bajo.",
  ]),
  h2("3.3. Modelo de negocio"),
  p(`La Tabla 3.1 recoge el Business Model Canvas de Flujo. Los precios del servicio se fijan dentro de los rangos que se manejaron en la fase exploratoria: un *setup* de 100 a 300 € y una suscripción de 50 a 100 € al mes, de los que el plan financiero toma el valor central.`),
  leyenda("Tabla", "3.1", "Business Model Canvas de Flujo"),
  tabla(["Bloque", "Contenido"], [
    ["Segmentos de clientes", "Autónomos y *freelancers* de 20 a 35 años, sobre todo en actividades digitales y creativas, sin formación financiera"],
    ["Propuesta de valor", "Entender y controlar el propio dinero sin tecnicismos: saber cuánto se gana, cuánto reservar y cuánto se puede gastar"],
    ["Canales", "Redes sociales (Instagram y TikTok), contenido educativo, recomendaciones de clientes y comunidades de autónomos"],
    ["Relación con clientes", "Personal y cercana en el servicio; automatizada en los productos digitales y en la herramienta"],
    ["Fuentes de ingresos", "*Setup* (200 €), suscripción mensual (75 €), plantillas (15 €), cursos (35 €) y herramienta digital (12 € al mes); precios sin IVA"],
    ["Recursos clave", "Conocimiento financiero y fiscal del fundador, sistema de plantillas propio, contenido educativo y marca"],
    ["Actividades clave", "Diagnóstico y seguimiento de clientes, creación de contenido, captación y, más adelante, desarrollo de la herramienta"],
    ["Socios clave", "Asesor fiscal colegiado para casos complejos, gestorías que deriven clientes y proveedores tecnológicos"],
    ["Estructura de costes", "Cuota de autónomos, software, seguro, marketing y, en el año 3, desarrollo y mantenimiento de la herramienta"],
  ], [2300, 6770], { numericas: false }),
  fuente("elaboración propia a partir de Osterwalder y Pigneur (2010)."),
  h2("3.4. Plan de validación"),
  p(`Siguiendo el enfoque Lean Startup, antes de invertir en producto el proyecto tendría que demostrar que alguien paga por el servicio. La Tabla 3.2 recoge las fases de validación previstas para los primeros cuatro meses. El criterio de éxito no es conseguir clientes a cualquier precio, sino conseguirlos al precio del modelo de negocio: un cliente que solo aceptaría el servicio gratis demuestra que el problema interesa, pero no que alguien pagaría por resolverlo.`),
  leyenda("Tabla", "3.2", "Cronograma de validación"),
  tabla(["Fase", "Semanas", "Objetivo", "Indicador de éxito"], [
    ["1. Contacto inicial", "1-2", "Abrir conversación con autónomos jóvenes en redes y comunidades", "15-20 conversaciones"],
    ["2. Diagnóstico gratuito", "3-5", "Sesión de diagnóstico a quien muestre interés real", "8-10 diagnósticos"],
    ["3. Conversión", "6-10", "Proponer el servicio de pago (*setup* y suscripción)", "2-3 clientes de pago"],
    ["4. Seguimiento", "11-16", "Prestar el servicio un mes completo y recoger opiniones", "Retención > 80 % y una recomendación"],
  ], [2000, 1000, 3770, 2300], { numericas: false }),
  fuente("elaboración propia."),
];

// ---------------------------------------------------------- 4. MERCADO
const tam = 0.161 * 3483645;
const cap4 = [
  h1("4. Análisis del mercado y del entorno", false),
  h2("4.1. Tamaño y evolución del mercado"),
  p(`El RETA contaba con 3.483.645 afiliados medios en septiembre de 2026 (The Objective, 2026). Según el informe de ATA sobre el perfil del autónomo, los menores de 35 años representan el 16,1 % del total (Castilla y León Económica, 2026), lo que sitúa el mercado potencial de Flujo en torno a ${miles(Math.round(tam / 1000) * 1000)} personas. No todas serían clientes posibles, porque una parte ya tiene formación financiera o no está dispuesta a pagar por este tipo de servicio, pero la cifra sirve para poner en contexto el objetivo del plan: los ${r0(A.activos_fin[2])} clientes de suscripción previstos al final del año 3 equivalen a menos del 0,01 % de ese mercado. Para Flujo el tamaño del mercado no es una restricción; el límite está en la capacidad de captar y atender clientes.`),
  p(`La tendencia, sin embargo, no es favorable. El análisis de los datos de la Seguridad Social que se recoge en el Anexo I muestra que el colectivo autónomo está envejeciendo: los menores de 30 años son solo el 7 % del RETA, mientras que los mayores de 50 ya suponen el 47,2 % (Observatorio Español de la Economía Social y del Trabajo Autónomo, 2025), y el 73,1 % del crecimiento del RETA en 2025 se explica por autónomos mayores de 56 años (Emirodgar, 2026). La estadística oficial del Ministerio de Trabajo confirma esta estructura: en marzo de 2026 el tramo de 46 a 55 años concentraba el 31,5 % de los autónomos y el de 56 a 65 años el 23,9 % (Ministerio de Trabajo y Economía Social, 2026a), mientras que los jóvenes de 16 a 29 años eran 249.456, alrededor del 7 % del RETA (Ministerio de Trabajo y Economía Social, 2026b). En la comparación europea, los autónomos de 20 a 29 años son solo el 5,9 % del total en España, frente al 7,9 % de media en la Unión Europea, uno de los porcentajes más bajos de los Veintisiete (Eurostat, 2026).`),
  p(`Hay, no obstante, un matiz reciente que conviene recoger: según el mismo informe del Ministerio, el número de jóvenes de 16 a 29 años afiliados al RETA creció un 5,1 % interanual en marzo de 2026, por encima del conjunto del régimen (Ministerio de Trabajo y Economía Social, 2026b). Incluyo los dos datos aunque el primero no favorezca al proyecto, porque es más útil dimensionar el mercado con datos reales que con un titular optimista: el segmento joven es amplio y muestra señales de recuperación, pero su peso es pequeño y Flujo tendría que ganar cuota en lugar de esperar a que el mercado se expanda.`),
  p(`A favor del proyecto juegan otras dos tendencias. La primera es regulatoria y tiene fechas concretas dentro del horizonte del plan. El Real Decreto 238/2026 desarrolla la factura electrónica obligatoria entre empresarios y profesionales y, tras la entrada en vigor de la Orden HAC/1028/2026 el 6 de octubre de 2026, los autónomos tendrán que emitir y recibir facturas electrónicas a partir del 6 de octubre de 2028 y comunicar su estado de pago desde el 6 de octubre de 2029 (Agencia Tributaria, 2026b). Además, el reglamento de sistemas de facturación (Verifactu) obliga a los autónomos a usar programas adaptados desde el 1 de julio de 2027, aunque Hacienda ha anunciado su intención de retrasar esta obligación hasta octubre de 2028 para hacerla coincidir con la factura electrónica (Agencia Tributaria, 2026c; Ministerio de Hacienda, 2026). Para muchos autónomos jóvenes, que hoy facturan con una hoja de cálculo, estos cambios obligan a ordenar sus números, y son una puerta de entrada natural para un servicio como Flujo.`),
  p(`La segunda tendencia es institucional. El Banco de España y la CNMV han preparado el Plan de Educación Financiera 2026-2029 a partir de la ECF (Banco de España, 2025), y en enero de 2026 el Ministerio de Educación renovó su colaboración con ambos organismos y con el Ministerio de Economía para impulsarlo (La Moncloa, 2026). Que la educación financiera sea una prioridad pública refuerza la legitimidad del servicio, aunque también puede aumentar la oferta gratuita de contenidos generales con la que Flujo tendría que diferenciarse.`),
  h2("4.2. Competencia"),
  p(`Ningún competidor directo cubre a la vez la organización financiera, el acompañamiento y la educación, pero hay tres grupos de soluciones parciales que compiten por el mismo presupuesto del cliente (Tabla 4.1). Las **gestorías**, tradicionales u *online*, presentan los impuestos y llevan la contabilidad; las *online* para autónomos cobran entre unos 30 y 65 € al mes sin IVA y las de barrio entre 75 y 120 € (Billeo, 2026). Las **aplicaciones de control de gastos**, como Fintonic, que cuenta con licencia del Banco de España para agregar cuentas bajo la normativa PSD2, automatizan la categorización de movimientos pero suponen que el usuario sabe interpretarlos. La **formación online** en finanzas personales es abundante y barata, pero genérica y poco aplicada al caso de cada persona.`),
  leyenda("Tabla", "4.1", "Comparativa de Flujo con las alternativas existentes"),
  tabla(["Alternativa", "Gestión", "Acompañamiento", "Educación", "Precio orientativo"], [
    ["Gestoría online", "Sí (fiscal)", "Limitado", "No", "30-65 €/mes"],
    ["Gestoría tradicional", "Sí (fiscal)", "Sí", "No", "75-120 €/mes"],
    ["Apps de control de gastos", "Sí (automática)", "No", "No", "Gratis o pocos euros"],
    ["Cursos online", "No", "No", "Sí (teórica)", "20-200 € por curso"],
    ["Flujo", "Sí (organización)", "Sí (personal)", "Sí (práctica)", "75 €/mes + 200 € de *setup*"],
  ], [2300, 1600, 1600, 1500, 2070], { numericas: false }),
  fuente("elaboración propia a partir de Billeo (2026) y de las webs de los competidores."),
  p(`La comparación deja una conclusión que conviene no esconder: el precio de Flujo es superior al de una gestoría *online*, y un cliente que ya pague una gestoría tendría que asumir los dos costes. Por eso Flujo se posiciona como complemento de la gestoría y no como sustituto, y una línea de trabajo lógica es colaborar con gestorías que deriven a sus clientes jóvenes. El efecto de un precio más bajo se analiza en el capítulo 9.`),
  h2("4.3. Análisis DAFO"),
  p(`La Tabla 4.2 resume el diagnóstico interno y externo del proyecto.`),
  leyenda("Tabla", "4.2", "Análisis DAFO de Flujo"),
  tabla(["Debilidades", "Amenazas"], [
    ["• Dependencia total del tiempo del fundador, que limita el crecimiento del servicio", "• Segmento joven que pierde peso dentro del RETA"],
    ["• Sin marca ni reputación previas, en un servicio que exige confianza", "• Competencia de gestorías y aplicaciones consolidadas"],
    ["• Precio superior al de las gestorías online", "• Posibles cambios en la cotización de autónomos a partir de 2027"],
  ], [4535, 4535], { numericas: false }),
  tabla(["Fortalezas", "Oportunidades"], [
    ["• Inversión inicial muy reducida y punto de equilibrio bajo", "• Más de 3,4 millones de autónomos, unos 560.000 menores de 35 años, y repunte reciente de los jóvenes"],
    ["• Enfoque práctico y lenguaje sencillo, que nadie ofrece de forma combinada", "• Bajo nivel de educación financiera documentado por el Banco de España"],
    ["• Contacto directo con el cliente, que permite ajustar la propuesta", "• Digitalización y facturación electrónica, que obligan a ordenar los números"],
    ["• Escalabilidad posterior mediante productos digitales", "• Factura electrónica obligatoria y Verifactu entre 2027 y 2029"],
    ["• Formación del fundador en ADE (contabilidad, fiscalidad y finanzas)", "• Posible colaboración con gestorías y apoyo público al emprendimiento (Zaragoza Activa)"],
  ], [4535, 4535], { numericas: false }),
  fuente("elaboración propia."),
];

// ------------------------------------------------- 5. RECURSOS Y MARCO LEGAL
const cap5 = [
  h1("5. Recursos, forma jurídica y aspectos legales", false),
  h2("5.1. Recursos necesarios y capacidad"),
  p(`En sus tres primeros años Flujo sería un proyecto unipersonal. El recurso principal es el tiempo del fundador, que durante el primer año compaginaría el proyecto con otra actividad (unas 60 horas de servicio al mes) y a partir del segundo se dedicaría a jornada completa (unas 120 horas de servicio, una vez descontado el tiempo de captación, contenido y administración). Cada cliente de suscripción requiere unas 2,5 horas al mes y cada alta unas 4 horas de diagnóstico y puesta en marcha. Con estas cifras, la capacidad máxima del servicio ronda los 45 clientes activos. Este límite es la razón de ser de las otras dos capas del modelo: las plantillas, los cursos y la herramienta digital permiten generar ingresos sin consumir horas por cliente.`),
  p(`Los recursos materiales son modestos: un ordenador, que se asume ya disponible; una página web con agenda y pasarela de pago; software de gestión, facturación y diseño; y, desde el año 2, un equipo básico de grabación para los cursos. El desarrollo de la herramienta digital se encargaría a un desarrollador externo en el año 3.`),
  p(`Existen además recursos públicos que encajan con la fase inicial. El Semillero de Ideas de Zaragoza Activa, del Ayuntamiento de Zaragoza, es un programa gratuito de preincubación para personas que han terminado o están terminando estudios universitarios o de formación profesional: ofrece durante seis meses un puesto en un espacio compartido, formación, tutorías individuales y la elaboración de un plan de empresa que evalúan la Universidad de Zaragoza y el Ayuntamiento (Ayuntamiento de Zaragoza, 2026). Para un recién graduado de la propia universidad sería una forma de validar Flujo con apoyo y sin coste.`),
  h2("5.2. Forma jurídica y obligaciones fiscales"),
  p(`Se ha optado por iniciar la actividad como **empresario individual (autónomo)**. Frente a una sociedad limitada, tiene menos costes de constitución y gestión, permite acogerse a la tarifa plana y encaja con un proyecto que debe validarse antes de crecer. Su inconveniente es la responsabilidad ilimitada y una tributación progresiva por IRPF que, con rendimientos elevados, supera al tipo del Impuesto sobre Sociedades, que para entidades de nueva creación es del 15 % (Agencia Tributaria, 2026a). Para las microempresas con una cifra de negocios inferior a un millón de euros, el tipo general baja además al 17 % para los primeros 50.000 € de base y al 20 % para el resto en 2027 (Agencia Tributaria, 2026a). Algunas fuentes de financiación, como ENISA, exigen también una sociedad mercantil. La Ley 28/2022, de empresas emergentes, ofrece ventajas adicionales (un Impuesto sobre Sociedades del 15 % durante cuatro años, el aplazamiento de la deuda tributaria de los dos primeros ejercicios con beneficios o la exención de pagos fraccionados), pero exige una sociedad con un modelo innovador y escalable certificado por ENISA (Ley 28/2022), algo a lo que un servicio de asesoramiento manual difícilmente podría aspirar hasta que la herramienta digital tenga tracción. La transformación en sociedad limitada, cuyo capital mínimo es hoy de un euro tras la Ley 18/2022, se plantea como un paso natural una vez superado el horizonte del plan. La Tabla 5.1 resume las principales obligaciones.`),
  leyenda("Tabla", "5.1", "Obligaciones fiscales y de Seguridad Social del proyecto"),
  tabla(["Ámbito", "Obligación", "Efecto en el plan financiero"], [
    ["Seguridad Social (RETA)", "Alta en el RETA; tarifa plana de 80 €/mes los primeros 12 meses, prorrogable otros 12 si el rendimiento neto no supera el SMI (17.094 € en 2026). Requiere no haber estado de alta en los dos años anteriores. Después, cuota según tramo: base mínima por un tipo total del 31,5 %", `Año 1 y año 2: 80 €/mes; año 3: ${eur(R.base.reta[2], 2)}/mes (tramo de 2.760 a 3.190 € de rendimiento mensual)`],
    ["IAE y censo", "Alta censal (modelo 036/037) en un epígrafe de servicios profesionales", "Sin coste"],
    ["IVA", "Facturas con IVA del 21 %; declaraciones trimestrales (modelo 303) y resumen anual (modelo 390)", "El IVA se liquida el mes siguiente a cada trimestre"],
    ["IRPF", "Estimación directa simplificada. Retención del 7 % en facturas a otros profesionales durante el año de inicio y los dos siguientes; pagos fraccionados (modelo 130) del 20 %, salvo que al menos el 70 % de los ingresos lleve retención", "Gastos de difícil justificación del 5 % (máx. 2.000 €); cuota según la escala estatal y la de Aragón"],
    ["Facturación", "Software adaptado a Verifactu desde el 1 de julio de 2027 (aplazamiento anunciado); factura electrónica obligatoria desde el 6 de octubre de 2028", "Incluido en el coste de las herramientas de facturación"],
    ["Protección de datos", "RGPD y LOPDGDD: consentimiento, política de privacidad y contrato de encargo cuando proceda", "Textos legales en la inversión inicial"],
  ], [1700, 4170, 3200], { numericas: false }),
  fuente("elaboración propia a partir del Real Decreto-ley 13/2022, la Orden PJC/297/2026, Seguridad Social (2026), la Ley 35/2006 del IRPF, la Ley 37/1992 del IVA, Agencia Tributaria (2026b, 2026c, 2026d, 2026e, 2026f) y el Real Decreto 126/2026."),
  p(`Las cuotas de autónomos para 2027 no estaban aprobadas en el momento de redactar este trabajo. La Orden PJC/297/2026 mantuvo para 2026 las mismas bases mínimas de 2025 y fijó un tipo total del 31,5 % (28,30 % por contingencias comunes, 1,30 % por contingencias profesionales, 0,90 % por cese de actividad, 0,10 % por formación profesional y 0,90 % por el mecanismo de equidad intergeneracional). El Real Decreto-ley 13/2022 prevé que, desde 2026, la cuantía de la tarifa plana la fije cada año la Ley de Presupuestos; en ausencia de cambios se mantiene en 80 € (Infoautónomos, 2026). El plan aplica, por prudencia, la tabla vigente en 2026 a los tres años.`),
  h2("5.3. Protección de datos y responsabilidad"),
  p(`Flujo trabajaría con información financiera personal de sus clientes, por lo que debe cumplir el Reglamento General de Protección de Datos y la Ley Orgánica 3/2018 desde el primer cliente. También conviene delimitar el alcance del servicio. Según la CNMV, el asesoramiento en materia de inversión consiste en recomendaciones personalizadas sobre instrumentos financieros concretos y solo pueden prestarlo entidades autorizadas e inscritas, mientras que la información general o las explicaciones con fines educativos no constituyen asesoramiento (CNMV, 2026). Flujo debe, por tanto, limitarse a la organización y la educación financiera, sin recomendar productos de inversión concretos, y derivar las cuestiones fiscales complejas a un asesor colegiado. Para la protección de datos, la AEPD ofrece herramientas gratuitas como Facilita RGPD y Facilita EMPRENDE, que generan el registro de actividades, las cláusulas informativas y unas medidas de seguridad mínimas, aunque la propia Agencia advierte de que los documentos deben revisarse y adaptarse a cada caso (AEPD, 2026). Por ese motivo el plan incluye un seguro de responsabilidad civil profesional y la colaboración de un asesor fiscal colegiado para los casos que excedan el servicio.`),
];

// ------------------------------------------------- 6. PLAN ECONÓMICO-FINANCIERO
const fijosMes = 40 + 30 + 10 + 25 + 40 + 25 + 25;
const inv0 = V.inversion_inicial;
const cap6 = [
  h1("6. Plan económico-financiero", false),
  p(`Este capítulo traduce el modelo de negocio a cifras para los tres primeros años de actividad (año 1 = 2027). Todas las cantidades se expresan sin IVA salvo en la tesorería, y se han calculado mes a mes. Las cifras son estimaciones razonadas, no datos de ventas reales, por lo que cada hipótesis se explica para que pueda revisarse cuando lleguen los primeros clientes.`, { keepNext: false }),
  h2("6.1. Hipótesis de partida"),
  leyenda("Tabla", "6.1", "Principales hipótesis del escenario base"),
  tabla(["Hipótesis", "Valor", "Justificación"], [
    ["Precio del *setup*", "200 €", "Centro del rango de 100 a 300 €"],
    ["Suscripción mensual", "75 €", "Centro del rango de 50 a 100 €"],
    ["Altas de clientes", "14 / 30 / 36 al año", "Ritmo del plan de validación; contenido y recomendaciones después"],
    ["Bajas", "4 % mensual", "Retención anual del 61 %, por debajo del objetivo de validación"],
    ["Plantillas y cursos", "15 € y 35 € por unidad", "Desde el año 2, de 5 a 40 plantillas y de 2 a 18 cursos al mes"],
    ["Herramienta digital", "12 €/mes por usuario", "Lanzamiento en el mes 31; de 15 a 75 usuarios"],
    ["Capacidad del servicio", "60 h/mes (año 1) y 120 h/mes", "2,5 h por cliente y mes; 4 h por alta"],
    ["Comisiones", "1,5 % servicio; 10 % productos", "Pasarela de pago y plataforma de cursos"],
    ["Clientes con retención de IRPF", "60 %", "Parte de los clientes son autónomos que facturan el servicio como gasto"],
    ["Inflación de costes", "3,7 % anual desde el año 2", "Previsión del IAPC del Banco de España para 2027; precios de Flujo constantes"],
  ], [2600, 2400, 4070], { numericas: false }),
  fuente("elaboración propia."),
  h2("6.2. Inversión inicial y financiación"),
  p(`La inversión inicial es reducida porque el proyecto empieza como un servicio. La Tabla 6.2 recoge las necesidades del mes 0 y del resto del horizonte. A la inversión en activos se suma un colchón de tesorería que cubre los meses de validación, en los que todavía no hay ingresos.`),
  leyenda("Tabla", "6.2", "Plan de inversiones y financiación"),
  tabla(["Concepto", "Momento", "Importe", "Financiación"], [
    ["Página web con agenda y pagos", "Mes 0", eur(350), "Ahorros propios"],
    ["Registro de marca en la OEPM", "Mes 0", eur(125), "Ahorros propios"],
    ["Gastos de puesta en marcha (logo, textos legales)", "Mes 0", eur(350), "Ahorros propios"],
    ["IVA soportado y colchón de tesorería", "Mes 0", eur(3000 - inv0), "Ahorros propios"],
    ["!Total necesidades iniciales", "", eur(3000), "Aportación de 3.000 €"],
    ["Equipo de grabación para cursos", "Mes 13", eur(450), "Autofinanciación"],
    ["Desarrollo de la herramienta (MVP)", "Meses 25-30", eur(12000), "Microcrédito de 10.000 € y autofinanciación"],
  ], [3770, 1300, 1400, 2600], { numericas: false }),
  fuente("elaboración propia."),
  p(`De acuerdo con la jerarquía financiera descrita en el capítulo 2, el arranque se financia íntegramente con 3.000 € de ahorros propios. El desarrollo de la herramienta digital, la única inversión relevante, se financiaría con un microcrédito de 10.000 € a cinco años y un tipo fijo del 7 %, en línea con las condiciones de MicroBank para emprendedores (Cámara de Comercio de Valencia, 2023), y con la caja generada por la actividad. Es un tipo superior a la TAE media de los nuevos créditos de hasta 250.000 € a sociedades no financieras, que el Banco de España situaba en el 4,58 % en enero de 2026 (Banco de España, 2026b), una diferencia lógica porque el microcrédito no exige garantías ni historial. En el escenario base el proyecto podría autofinanciar esa inversión, pero el préstamo se mantiene por prudencia: preserva un colchón de liquidez y, como se verá en el capítulo 9, sería imprescindible si la captación fuese más lenta. ENISA y los *business angels* no encajan en esta fase, porque exigen una sociedad mercantil y buscan crecimientos que un servicio unipersonal no ofrece; serían opciones a estudiar si la herramienta digital demostrara tracción.`),
  h2("6.3. Previsión de ingresos"),
  p(`Con estas hipótesis, los ingresos pasan de ${eur(A.ingresos[0])} en el año 1 a ${eur(A.ingresos[1])} en el año 2 y ${eur(A.ingresos[2])} en el año 3 (Tabla 6.3). El servicio de acompañamiento aporta la mayor parte durante todo el horizonte, mientras que los productos digitales y la herramienta ganan peso a partir del segundo año (Figura 6.1).`),
  leyenda("Tabla", "6.3", "Previsión de ingresos del escenario base (sin IVA)"),
  tabla(["Concepto", ...anios], [
    fila("Altas de clientes de suscripción", A.altas, (x) => num(x, 0)),
    fila("Clientes activos a final de año", A.activos_fin, (x) => num(x, 0)),
    fila("Ingresos por *setup*", A.ing_setup),
    fila("Ingresos por suscripción", A.ing_suscripcion),
    fila("Ingresos por plantillas y cursos", [0, 1, 2].map((a) => A.ing_plantillas[a] + A.ing_cursos[a])),
    fila("Ingresos por la herramienta digital", A.ing_app),
    fila("!Total ingresos", A.ingresos),
  ], W4),
  fuente("elaboración propia."),
  leyenda("Figura", "6.1", "Composición de los ingresos por línea de negocio"),
  imagen("fig_ingresos.png", 12),
  fuente("elaboración propia."),
  p(`La Figura 6.2 muestra un resultado que condiciona todo el plan: en el escenario base la cartera de clientes alcanza el límite de capacidad hacia el final del año 3, cuando se acerca a los 45 clientes activos. A partir de ese momento el servicio no puede crecer sin contratar o sin trasladar clientes a la herramienta digital, y el modelo frena las altas del año 3 de forma automática (se registran ${r0(A.altas[2])} en lugar de 36).`),
  leyenda("Figura", "6.2", "Clientes de suscripción activos y límite de capacidad"),
  imagen("fig_clientes.png", 12),
  fuente("elaboración propia."),
  h2("6.4. Estructura de costes y punto de equilibrio"),
  p(`Los costes fijos de explotación, sin contar la cuota de autónomos, suman ${eur(fijosMes)} al mes: herramientas de gestión y facturación (40 €), diseño y redes sociales (30 €), web y alojamiento (10 €), seguro de responsabilidad civil (25 €), asesor fiscal de apoyo (40 €), teléfono e internet (25 €) y formación (25 €). A ellos se añaden el marketing, que crece de 50 a 300 € al mes, la cuota de autónomos y, desde el lanzamiento de la herramienta, 150 € mensuales de mantenimiento. Los costes variables son las comisiones de cobro y de las plataformas, y el alojamiento de la herramienta. Todos los costes, salvo la cuota de autónomos, se actualizan un 3,7 % anual desde el segundo año, de acuerdo con la previsión de inflación del Banco de España para 2027 (Banco de España, 2026a), mientras que los precios de Flujo se mantienen constantes. Es una hipótesis prudente, porque la mayoría de servicios por suscripción actualizan sus tarifas.`),
  p(`Con un margen de contribución de ${eur(75 * 0.985, 2)} por cliente y mes, el punto de equilibrio del año 1 se alcanza con unos **${Math.ceil(B.pe[0])} clientes de suscripción** activos, de ${Math.ceil(B.pe[1])} en el año 2 y de ${Math.ceil(B.pe[2])} en el año 3, cuando sube la cuota de autónomos, crece el marketing y empiezan a amortizarse el desarrollo de la herramienta y los intereses. Esta estimación es más exigente que la de una primera versión del proyecto, que solo contaba la cuota de autónomos y el software, pero sigue siendo baja. El dato cambia mucho si se exige que el proyecto pague también un sueldo al fundador: para cubrir el equivalente al SMI harían falta unos ${Math.ceil(B.pe_smi[0])} clientes en el año 1 y ${Math.ceil(B.pe_smi[2])} en el año 3, una cifra cercana al límite de capacidad.`),
  h2("6.5. Cuenta de resultados previsional"),
  p(`La Tabla 6.4 recoge la cuenta de resultados. Al tratarse de un empresario individual, el IRPF es un impuesto personal y no un gasto de la actividad; se incluye como línea analítica, calculado con la escala estatal y la de Aragón, para conocer el resultado del que dispondría realmente el titular. El año 1 cierra con un beneficio de ${eur(A.resultado_neto[0])}, el año 2 con ${eur(A.resultado_neto[1])} y el año 3 con ${eur(A.resultado_neto[2])}, con un tipo efectivo de IRPF que pasa del 0 % al ${pct(A.tipo_efectivo[2], 0)}.`),
  leyenda("Tabla", "6.4", "Cuenta de resultados previsional del escenario base"),
  tabla(["Concepto", ...anios], [
    fila("!Importe neto de la cifra de negocios", A.ingresos),
    fila("Costes variables", neg(A.c_variables)),
    fila("!Margen de contribución", A.margen_contribucion),
    fila("Cuota de autónomos", neg(A.reta)),
    fila("Otros gastos fijos", neg(A.c_fijos)),
    fila("Marketing", neg(A.marketing)),
    fila("Mantenimiento de la herramienta", neg(A.mant_app)),
    fila("Gastos de puesta en marcha", neg(A.puesta_marcha)),
    fila("!Resultado bruto de explotación (EBITDA)", A.ebitda),
    fila("Amortización del inmovilizado", neg(A.amortizacion)),
    fila("!Resultado de explotación (EBIT)", A.ebit),
    fila("Gastos financieros", neg(A.intereses)),
    fila("!Resultado antes de impuestos", A.rn_previo),
    fila("IRPF imputable a la actividad", neg(A.irpf)),
    fila("!Resultado neto", A.resultado_neto),
  ], W4),
  fuente("elaboración propia."),
  h2("6.6. Estado de flujos de tesorería"),
  p(`La tesorería se ha calculado mes a mes incluyendo el IVA, las retenciones, los pagos fraccionados del IRPF y la liquidación anual de la renta en junio, además de las retiradas que el titular hace para vivir (300 € al mes el primer año, 900 € el segundo y 1.200 € el tercero). Con estas retiradas el saldo mínimo, de ${eur(A.caja_min)}, se produce en el mes ${A.mes_caja_min}, durante la fase de validación, y a partir del segundo año la caja crece con rapidez (Tabla 6.5). El colchón inicial resulta por tanto suficiente, aunque ajustado.`),
  leyenda("Tabla", "6.5", "Estado de flujos de tesorería del escenario base"),
  tabla(["Concepto", ...anios], (() => {
    const F = A.flujo;
    const expl = [0, 1, 2].map((a) => F.cobros_explotacion[a] - F.pagos_explotacion[a] - F.iva_neto[a] - F.irpf_pagado[a]);
    const fin = [0, 1, 2].map((a) => F.aportacion[a] + F.prestamo[a] - F.devolucion_prestamo[a] - F.retiradas[a]);
    const ini = [0, A.caja_fin[0], A.caja_fin[1]];
    return [
      fila("Cobros de explotación (con IVA)", F.cobros_explotacion),
      fila("Pagos de explotación (con IVA)", neg(F.pagos_explotacion)),
      fila("Liquidaciones de IVA", neg(F.iva_neto)),
      fila("Pagos de IRPF", neg(F.irpf_pagado)),
      fila("!Flujo de las actividades de explotación", expl),
      fila("!Flujo de las actividades de inversión", neg(F.inversion)),
      fila("Aportación del titular", F.aportacion),
      fila("Microcrédito", F.prestamo),
      fila("Devolución del principal", neg(F.devolucion_prestamo)),
      fila("Retiradas del titular", neg(F.retiradas)),
      fila("!Flujo de las actividades de financiación", fin),
      fila("Tesorería inicial", ini),
      fila("!Tesorería final", A.caja_fin),
    ];
  })(), W4),
  fuente("elaboración propia."),
  h2("6.7. Balance previsional"),
  p(`El balance (Tabla 6.6) refleja la estructura de una microempresa de servicios: un activo formado casi solo por tesorería hasta que se desarrolla la herramienta digital, un patrimonio neto que crece con los beneficios no retirados y un pasivo compuesto por el microcrédito y las deudas con Hacienda por el IVA del último trimestre y la liquidación del IRPF. El balance cuadra en los tres ejercicios, lo que confirma la coherencia entre la cuenta de resultados y la tesorería.`),
  leyenda("Tabla", "6.6", "Balance previsional a 31 de diciembre del escenario base"),
  tabla(["Concepto", ...anios], (() => {
    const b = B.balance, g = (k) => b.map((x) => x[k]);
    return [
      fila("Inmovilizado intangible", g("intangible")),
      fila("Inmovilizado material", g("material")),
      fila("Amortización acumulada", g("amort_acum")),
      fila("!Activo no corriente", g("inmov_neto")),
      fila("Hacienda Pública deudora", g("hp_deudora")),
      fila("Tesorería", g("tesoreria")),
      fila("!Total activo", g("activo")),
      fila("Aportaciones del titular", g("aportaciones")),
      fila("Resultados acumulados", g("resultados")),
      fila("Retiradas del titular", g("retiradas")),
      fila("!Patrimonio neto", g("pn")),
      fila("Deudas con entidades de crédito a largo plazo", g("deuda_lp")),
      fila("Deudas con entidades de crédito a corto plazo", g("deuda_cp")),
      fila("Hacienda Pública acreedora (IVA e IRPF)", b.map((x) => x.hp_iva + x.hp_irpf)),
      fila("!Total patrimonio neto y pasivo", b.map((x) => x.pn + x.pasivo)),
    ];
  })(), W4),
  fuente("elaboración propia."),
];

// ------------------------------------------------- 7. VIABILIDAD Y RENTABILIDAD
const RAT = B.ratios;
const cap7 = [
  h1("7. Análisis de viabilidad y rentabilidad", false),
  h2("7.1. Tasa de descuento"),
  p(`La rentabilidad exigida al proyecto se ha estimado con el procedimiento descrito en el marco teórico (Tabla 7.1). Se parte de la rentabilidad del bono español a diez años, que rondaba el 3,4 % en julio de 2026 (VT Markets, 2026), y se añade una prima de mercado del 5,5 %, coherente con la prima implícita que estima Damodaran (2026) para un mercado maduro más un pequeño diferencial por riesgo país. Se toma una beta de 1, propia de un negocio de servicios profesionales sin deuda. El resultado, un 8,9 %, sería razonable para una empresa consolidada, pero no para un proyecto que todavía no tiene clientes, que no podría venderse con facilidad y que depende por completo de una persona. Por ello se añade una prima específica de 6,1 puntos, hasta una tasa del ${pct(R.tasa, 0)}. La elección es inevitablemente subjetiva, de modo que el capítulo 9 muestra cómo cambia el VAN con tasas de entre el 10 % y el 30 %.`),
  leyenda("Tabla", "7.1", "Estimación de la tasa de descuento"),
  tabla(["Componente", "Valor"], [
    ["Rentabilidad sin riesgo (bono español a 10 años)", "3,4 %"],
    ["Prima de riesgo de mercado × beta (5,5 % × 1,0)", "5,5 %"],
    ["Prima específica (tamaño, iliquidez y dependencia del fundador)", "6,1 %"],
    ["!Tasa de descuento", pct(R.tasa, 1)],
  ], [7070, 2000]),
  fuente("elaboración propia a partir de Damodaran (2009, 2012, 2026) y VT Markets (2026)."),
  h2("7.2. VAN, TIR y periodo de recuperación"),
  p(`Los flujos de caja libres del proyecto (Tabla 7.2) se obtienen a partir del EBITDA, restando el IRPF que correspondería al resultado de explotación, las inversiones y la variación del capital circulante, y sin tener en cuenta las retiradas del titular ni la financiación. En el año 0 se incluye la inversión inicial y el colchón de tesorería, que se recupera al final del año 3.`),
  leyenda("Tabla", "7.2", "Flujos de caja libres y criterios de rentabilidad (escenario base)"),
  tabla(["Concepto", "Año 0", ...anios], [
    ["EBITDA (sin gastos de puesta en marcha)", "—", ...[0, 1, 2].map((a) => eur(A.ebitda[a] + (a === 0 ? 350 : 0)))],
    ["IRPF sobre el resultado de explotación", "—", ...A.irpf_operativo.map((x) => eur(-x))],
    ["Inversiones", eur(-(V.inversion_inicial + V.fondo_maniobra)), ...V.inversiones.map((x, a) => eur(-x + (a === 2 ? V.fondo_maniobra : 0)))],
    ["Variación del capital circulante", "—", ...V.var_nof.map((x) => eur(-x))],
    ["!Flujo de caja libre", ...V.flujos.map((x) => eur(x))],
  ], [3470, 1400, 1400, 1400, 1400]),
  fuente("elaboración propia. En el año 3 las inversiones incluyen la recuperación del colchón de tesorería inicial."),
  p(`Con una tasa del ${pct(R.tasa, 0)}, el **VAN a tres años es de ${eur(V.van)}**, sin incluir ningún valor residual. Si se supusiera que el proyecto continúa indefinidamente con los flujos del año 3 y un crecimiento del 1 %, el VAN superaría los ${eur(Math.floor(V.van_con_vr / 1000) * 1000)}, pero ese cálculo no es prudente en un negocio que depende de una sola persona y se presenta solo como referencia. El **payback** es de unos ${meses(V.payback)} meses (${meses(V.payback_desc)} en términos descontados).`),
  p(`La **TIR**, del ${miles(V.tir * 100, 0)} %, es un buen ejemplo de la limitación señalada en el capítulo 2. Una cifra tan alta no significa que el proyecto sea excepcional, sino que la inversión inicial (${eur(-V.flujos[0])}) es muy pequeña en relación con los flujos posteriores. Lo que el fundador invierte de verdad en Flujo no es dinero, sino su tiempo, y ese tiempo no aparece en los flujos de caja. Por eso el análisis se completa en el apartado siguiente.`),
  h2("7.3. El coste de oportunidad del emprendedor"),
  p(`El resultado de un empresario individual es, en realidad, su sueldo. Para que Flujo sea una opción razonable no basta con que genere caja: tiene que generar más que la alternativa de trabajar por cuenta ajena. Para medirlo se ha restado a los flujos de caja un sueldo neto equivalente al SMI de 2026 (17.094 € brutos, unos 15.980 € netos de cotización), la mitad en el año 1, en el que la dedicación es parcial. Se ha elegido el SMI por ser un umbral objetivo y conservador; un graduado en ADE podría aspirar a un sueldo algo mayor.`),
  p(`Con este ajuste los flujos pasan a ser de ${V.flujos_oportunidad.slice(1).map((x) => eur(x)).join(", ")} en los años 1 a 3, y **el VAN se reduce a ${eur(V.van_oportunidad)}**, con una TIR del ${miles(V.tir_oportunidad * 100, 0)} %. El proyecto sigue creando valor, pero el margen es mucho más estrecho: el año 1 no cubre el sueldo alternativo y el resultado depende de que los años 2 y 3 se cumplan. Este es, a mi juicio, el indicador más honesto de la viabilidad de Flujo.`),
  h2("7.4. Ratios financieros"),
  p(`La Tabla 7.3 recoge los principales ratios. Los márgenes son altos porque los costes variables apenas existen y el principal coste, el trabajo del fundador, no figura como gasto. La liquidez es holgada desde el primer año y el endeudamiento solo aumenta en el año 3 por el microcrédito, con una cobertura de intereses muy amplia. Las rentabilidades sobre activos y sobre patrimonio neto superan el 100 % porque la empresa apenas tiene activos y el titular retira buena parte del beneficio; por eso no se consideran indicadores útiles en este caso.`),
  leyenda("Tabla", "7.3", "Ratios financieros del escenario base"),
  tabla(["Ratio", ...anios], [
    fila("Margen EBITDA (EBITDA / ventas)", RAT.map((r) => r.margen_ebitda), (x) => pct(x, 1)),
    fila("Margen neto (resultado neto / ventas)", RAT.map((r) => r.margen_neto), (x) => pct(x, 1)),
    fila("Liquidez (activo corriente / pasivo corriente)", RAT.map((r) => r.liquidez), (x) => num(x, 2)),
    fila("Endeudamiento (pasivo / total patrimonio neto y pasivo)", RAT.map((r) => r.endeudamiento), (x) => pct(x, 1)),
    ["Cobertura de intereses (EBIT / gastos financieros)", "—", "—", num(RAT[2].cobertura_intereses, 1)],
    fila("Punto de equilibrio (clientes de suscripción)", B.pe, (x) => num(x, 1)),
    fila("Margen de seguridad sobre la suscripción", RAT.map((r) => r.margen_seguridad), (x) => pct(x, 1)),
  ], W4),
  fuente("elaboración propia."),
  p(`El margen de seguridad del año 1, de apenas un ${pct(RAT[0].margen_seguridad, 0)}, indica que en ese ejercicio los ingresos por suscripción superan por poco a los que serían necesarios para cubrir los costes. El primer año es, por tanto, el más vulnerable.`),
];

// ------------------------------------------------- 8. MEJORA DE LA RENTABILIDAD
const UE = R.unit_economics, PAL = R.palancas, OPT = R.optimizado, MC = R.montecarlo, MCO = R.montecarlo_opt;
const LIN = R.lineas, SL = R.sl;
const capMejora = [
  h1("8. Mejora de la rentabilidad del proyecto", false),
  p(`El capítulo anterior concluye que Flujo es viable, pero con un margen estrecho una vez se tiene en cuenta el coste de oportunidad del fundador. Este capítulo busca dónde está el margen de mejora. Primero analiza la rentabilidad de cada cliente y de cada línea de negocio, después propone y cuantifica cinco medidas concretas y, por último, estudia si la forma jurídica elegida es la más eficiente.`),
  h2("8.1. Economía unitaria del cliente"),
  p(`En los negocios por suscripción es habitual comparar el valor que aporta un cliente durante toda su relación con la empresa (*lifetime value*, LTV) con lo que cuesta captarlo (*customer acquisition cost*, CAC). Con una baja mensual del 4 %, un cliente de Flujo permanece de media ${num(UE.vida_meses, 0)} meses y deja un margen de ${eur(UE.margen_mes, 2)} al mes, a lo que se suma el *setup*. Su LTV es, por tanto, de unos ${eur(UE.ltv)} (Tabla 8.1).`),
  leyenda("Tabla", "8.1", "Valor del cliente y coste de captación (escenario base)"),
  tabla(["Indicador", ...anios], [
    ["Valor del cliente (LTV)", eur(UE.ltv), eur(UE.ltv), eur(UE.ltv)],
    fila("Coste de captación por cliente (CAC)", UE.cac),
    fila("Relación LTV / CAC", UE.ltv_cac, (x) => num(x, 1)),
    fila("Meses de margen para recuperar el CAC", UE.recuperacion_cac_meses, (x) => num(x, 1)),
  ], W4),
  fuente("elaboración propia."),
  p(`La relación LTV/CAC supera con creces el umbral de 3 que suele tomarse como referencia en negocios de suscripción, y el coste de captar un cliente se recupera en menos de dos meses. La lectura es clara: captar clientes no es caro. El cuello de botella de Flujo no es el marketing, sino las horas del fundador, que se agotan al final del año 3 (Figura 6.2). Por eso las medidas de mejora que siguen buscan, sobre todo, **aumentar los ingresos por hora trabajada** en lugar de captar más clientes.`),
  h2("8.2. Rentabilidad por línea de negocio"),
  p(`Para saber qué aporta cada capa del modelo se ha calculado el VAN incremental de cada una, es decir, la diferencia entre el plan completo y el mismo plan sin esa línea. Los productos digitales (plantillas y cursos) aportan ${eur(LIN.productos.van)} al VAN, porque generan ingresos sin consumir horas por cliente. La herramienta digital, en cambio, tiene un VAN incremental de ${eur(LIN.herramienta.van)} en el horizonte del plan: la inversión de 12.000 € se hace en el año 3 y apenas hay tiempo para recuperarla.`),
  p(`Este resultado no significa que la herramienta deba descartarse. Su valor está fuera del horizonte de tres años, ya que es la única capa que permite crecer por encima del límite de capacidad del servicio. Se comporta como una opción real: conviene conservar la posibilidad de desarrollarla, pero al menor coste posible y solo cuando el servicio haya demostrado que hay demanda. Esa idea se recoge en la última de las medidas del apartado siguiente.`),
  h2("8.3. Medidas de mejora"),
  p(`Se proponen cinco medidas de bajo coste, coherentes con el diagnóstico anterior. Todas son hipótesis que deberían validarse, igual que las del plan base:`),
  ...vinetas([
    "**Actualizar los precios cada año con el IPC.** El plan base mantiene los precios constantes mientras los costes suben un 3,7 % al año; actualizar las tarifas es una práctica habitual en los servicios por suscripción.",
    "**Estandarizar el servicio.** Con las plantillas y la rutina de revisión ya construidas, el tiempo por cliente puede bajar de 2,5 a 2 horas al mes desde el segundo año, lo que eleva la capacidad máxima de unos 45 a unos 55 clientes.",
    "**Talleres grupales para instituciones.** Sesiones de dos horas sobre organización financiera para autónomos, ofrecidas a asociaciones de autónomos, coworkings, colegios profesionales, universidades o programas públicos de emprendimiento, a 250 € por sesión: una al mes en el año 2 y dos en el año 3. Cada taller ocupa unas 5 horas y genera más ingresos por hora que el servicio individual, además de servir como canal de captación.",
    "**Acuerdo de derivación con gestorías.** Flujo no presenta impuestos, por lo que puede derivar a una gestoría colaboradora a los clientes que no tengan una, a cambio de una comisión de 60 € por cliente; se supone que el 30 % de las altas la necesita.",
    "**Herramienta digital de bajo coste.** Desarrollar la primera versión con herramientas *no-code* por unos 5.000 € en lugar de un desarrollo a medida de 12.000 €, financiada con la caja generada y sin microcrédito.",
  ]),
  p(`La Tabla 8.2 muestra el efecto de aplicar las medidas de forma acumulada.`),
  leyenda("Tabla", "8.2", "Efecto acumulado de las medidas de mejora"),
  tabla(["Medida", "VAN", "VAN con coste de oportunidad", "Mejora"], PAL.map((f, i) => [
    i === 0 ? "!Plan base" : f.palanca, eur(f.van), eur(f.van_oportunidad), i === 0 ? "—" : "+" + eur(f.delta_op),
  ]).concat([["!Plan mejorado", eur(PAL[PAL.length - 1].van), eur(PAL[PAL.length - 1].van_oportunidad),
    "+" + eur(PAL[PAL.length - 1].van_oportunidad - PAL[0].van_oportunidad)]]), [4470, 1400, 1800, 1400]),
  fuente("elaboración propia."),
  leyenda("Figura", "8.1", "Aportación de cada medida al VAN con coste de oportunidad"),
  imagen("fig_palancas.png", 12),
  fuente("elaboración propia."),
  p(`En conjunto, las cinco medidas elevan el VAN a tres años de ${eur(PAL[0].van)} a ${eur(PAL[PAL.length - 1].van)} y, lo que es más importante, multiplican por ${num(PAL[PAL.length - 1].van_oportunidad / PAL[0].van_oportunidad, 1)} el VAN con coste de oportunidad, que pasa de ${eur(PAL[0].van_oportunidad)} a ${eur(PAL[PAL.length - 1].van_oportunidad)}. Las dos que más aportan son los talleres, porque aprovechan mejor las horas del fundador, y la herramienta de bajo coste, porque reduce una inversión que no se recupera dentro del horizonte. Ninguna exige una inversión adicional.`),
  h2("8.4. El plan mejorado"),
  p(`Con las cinco medidas, los ingresos pasan de ${eur(OPT.anual.ingresos[0])} en el año 1 a ${eur(OPT.anual.ingresos[1])} en el año 2 y ${eur(OPT.anual.ingresos[2])} en el año 3, y el resultado neto del año 3 alcanza ${eur(OPT.anual.resultado_neto[2])}. La cartera llega a ${num(OPT.anual.activos_fin[2], 0)} clientes de suscripción, de nuevo cerca del límite de capacidad (que ahora incluye las horas de los talleres), y la tesorería nunca baja de ${eur(OPT.anual.caja_min)}, sin necesidad de financiación ajena. El plan sigue siendo sensible a la captación: en el escenario pesimista el VAN con coste de oportunidad sería de ${eur(R.optimizado_pesimista.van_oportunidad)}, frente a ${eur(R.pesimista.valoracion.van_oportunidad)} del plan base, y en el optimista de ${eur(R.optimizado_optimista.van_oportunidad)}.`),
  p(`Para valorar el riesgo de forma más completa que con tres escenarios cerrados se ha hecho una simulación de Montecarlo con ${miles(MC.n)} combinaciones aleatorias de las variables más inciertas: la captación de clientes (entre el 50 % y el 130 % de la prevista, con distribución triangular), la baja mensual (uniforme entre el 3 % y el 6 %) y las ventas de productos y de la herramienta (entre el 30 % y el 150 % de lo previsto). La Tabla 8.3 y la Figura 8.2 recogen los resultados.`),
  leyenda("Tabla", "8.3", "Resultados de la simulación de Montecarlo"),
  tabla(["Indicador", "Plan base", "Plan mejorado"], [
    ["Probabilidad de VAN positivo", pct(MC.p_van, 0), pct(MCO.p_van, 0)],
    ["Probabilidad de VAN con coste de oportunidad positivo", pct(MC.p_van_op, 0), pct(MCO.p_van_op, 0)],
    ["VAN con coste de oportunidad: percentil 5", eur(MC.p5_op), eur(MCO.p5_op)],
    ["VAN con coste de oportunidad: mediana", eur(MC.p50_op), eur(MCO.p50_op)],
    ["VAN con coste de oportunidad: percentil 95", eur(MC.p95_op), eur(MCO.p95_op)],
  ], [5070, 2000, 2000]),
  fuente("elaboración propia."),
  leyenda("Figura", "8.2", "Distribución del VAN con coste de oportunidad en la simulación"),
  imagen("fig_montecarlo.png", 12),
  fuente("elaboración propia."),
  p(`El VAN es positivo en todas las simulaciones, tanto en el plan base como en el mejorado. La diferencia está en el coste de oportunidad: con el plan base, Flujo supera al empleo alternativo en el ${pct(MC.p_van_op, 0)} de los casos, prácticamente una moneda al aire; con el plan mejorado, en el ${pct(MCO.p_van_op, 0)}. Las medidas propuestas no solo aumentan la rentabilidad esperada, sino que reducen de forma notable la probabilidad de que el proyecto no compense el esfuerzo del fundador.`),
  h2("8.5. Forma jurídica y fiscalidad"),
  p(`El empresario individual es la forma más sencilla para arrancar, pero su tributación progresiva aumenta con los beneficios. La Tabla 8.4 compara la carga fiscal del año 3 con la que tendría una sociedad limitada que pagara un sueldo al fundador y tributara al 15 % por el resto, el tipo de las entidades de nueva creación (Agencia Tributaria, 2026a). Se añaden unos 1.200 € anuales de gestoría y obligaciones mercantiles y se supone una cuota de autónomo similar en ambos casos.`),
  leyenda("Tabla", "8.4", "Carga fiscal anual: empresario individual frente a sociedad limitada"),
  tabla(["Rendimiento y sueldo", "IRPF como autónomo", "Impuestos con SL", "Ahorro neto con SL"], SL.map((x) => [
    `${eur(x.rendimiento)} con sueldo de ${eur(x.sueldo)}`, eur(x.irpf_autonomo), eur(x.impuestos_sl), eur(x.ahorro),
  ]), [3670, 1800, 1800, 1800]),
  fuente("elaboración propia a partir de las escalas del IRPF y de Agencia Tributaria (2026a). El ahorro descuenta 1.200 € de costes adicionales de la sociedad."),
  p(`Con el rendimiento del año 3 del plan base, en torno a ${eur(SL[0].rendimiento)}, la sociedad ahorraría unos ${eur(SL[1].ahorro)} al año, pero a cambio de que una parte del beneficio quede retenida en ella; si se repartiera como dividendos tributaría de nuevo en el IRPF del ahorro, por lo que buena parte del ahorro sería un diferimiento. Con un rendimiento de 60.000 € el ahorro supera los ${eur(Math.floor(SL[3].ahorro / 1000) * 1000)}. La conclusión es que el empresario individual es la forma adecuada durante el horizonte del plan, y que la transformación en sociedad limitada tendría sentido a partir del cuarto año si el beneficio supera los 50.000 € o si se quiere reinvertir en la herramienta digital y acceder a ENISA.`),
];

// ------------------------------------------------- 9. RIESGOS Y ESCENARIOS
const esc = ["pesimista", "base", "optimista"];
const cap8 = [
  h1("9. Análisis de riesgos y escenarios", false),
  h2("9.1. Escenarios"),
  p(`Para comprobar la solidez de los resultados se han construido dos escenarios alternativos. En el **pesimista** las altas de clientes son un 40 % inferiores a las del base, las bajas suben al 6 % mensual y las ventas de productos y usuarios de la herramienta se reducen a la mitad. En el **optimista** las altas aumentan un 30 %, las bajas bajan al 3 % y los productos y la herramienta venden un 30 % y un 50 % más, respectivamente. Los costes, la financiación y las retiradas del titular son los mismos en los tres casos, para que la comparación sea limpia.`),
  leyenda("Tabla", "9.1", "Resultados por escenario"),
  tabla(["Indicador", "Pesimista", "Base", "Optimista"], [
    ["Clientes activos al final del año 3", ...esc.map((e) => num(R[e].anual.activos_fin[2], 0))],
    ["Ingresos del año 3", ...esc.map((e) => eur(R[e].anual.ingresos[2]))],
    ["Resultado neto acumulado (3 años)", ...esc.map((e) => eur(sum(R[e].anual.resultado_neto)))],
    ["Saldo mínimo de tesorería", ...esc.map((e) => eur(R[e].anual.caja_min))],
    ["VAN a 3 años", ...esc.map((e) => eur(R[e].valoracion.van))],
    ["Payback (meses)", ...esc.map((e) => num(meses(R[e].valoracion.payback), 0))],
    ["!VAN con coste de oportunidad", ...esc.map((e) => eur(R[e].valoracion.van_oportunidad))],
  ], W4),
  fuente("elaboración propia."),
  p(`El VAN es positivo en los tres escenarios, incluso en el pesimista, lo que confirma que el proyecto, como inversión en dinero, tiene poco riesgo. La diferencia está en los otros dos indicadores. En el escenario pesimista la tesorería se vuelve negativa ya en el primer año y llega a un mínimo de ${eur(PE.anual.caja_min)} en el mes ${PE.anual.mes_caja_min} (Figura 9.1): el fundador no podría retirar lo previsto para vivir y tendría que mantener otra fuente de ingresos o retrasar el desarrollo de la herramienta. Además, el VAN con coste de oportunidad pasa a ${eur(PE.valoracion.van_oportunidad)}: con esa captación, Flujo rendiría menos que un empleo con el salario mínimo.`),
  leyenda("Figura", "9.1", "Saldo mensual de tesorería por escenario"),
  imagen("fig_tesoreria.png", 12),
  fuente("elaboración propia."),
  p(`En el escenario optimista, por el contrario, el límite de capacidad se alcanza ya en el año 2, lo que confirma que el crecimiento del servicio está acotado por las horas del fundador y no por la demanda.`),
  h2("9.2. Análisis de sensibilidad"),
  p(`La Figura 9.2 muestra cómo varía el VAN del escenario base al modificar, de una en una, las dos variables más inciertas. Una rebaja del precio de la suscripción a 55 €, cercana al de una gestoría *online*, reduciría el VAN a ${eur(R.sensibilidad.precio["55"])}, y una captación a la mitad de la prevista lo dejaría en ${eur(R.sensibilidad.altas["0.5"])}. El VAN se mantiene positivo en todos los casos. Respecto a la tasa de descuento, pasar del 10 % al 30 % lo reduce de ${eur(R.sensibilidad.tasa["0.1"])} a ${eur(R.sensibilidad.tasa["0.3"])}, una variación pequeña porque el horizonte es corto y la inversión, mínima. La captación de clientes es, por tanto, la variable crítica. Por último, si los precios de Flujo se actualizaran con la misma inflación que los costes, el VAN subiría a ${eur(R.sensibilidad.precios_indexados.van)} y el VAN con coste de oportunidad a ${eur(R.sensibilidad.precios_indexados.van_oportunidad)}.`),
  leyenda("Figura", "9.2", "Sensibilidad del VAN al precio y a la captación de clientes"),
  imagen("fig_sensibilidad.png", 12),
  fuente("elaboración propia."),
  h2("9.3. Principales riesgos y medidas de mitigación"),
  p(`Los escenarios deben leerse junto con un dato de contexto: según la Demografía Armonizada de Empresas del INE, de media el 45,3 % de las empresas nacidas en España entre 2011 y 2022 desapareció en sus tres primeros años, y solo el 41,9 % de las nacidas en 2018 seguía activa cinco años después (INE, 2025). El primer año es el de mayor mortalidad. Un VAN positivo no elimina, por tanto, el riesgo de fracaso, y la Tabla 9.2 recoge los principales riesgos identificados y las medidas para reducirlos.`),
  leyenda("Tabla", "9.2", "Matriz de riesgos del proyecto"),
  tabla(["Riesgo", "Probabilidad", "Impacto", "Medida de mitigación"], [
    ["Captación más lenta de lo previsto", "Media", "Alto", "Validar antes de invertir; mantener ingresos alternativos el primer año"],
    ["Precio percibido como alto frente a gestorías", "Media", "Medio", "Posicionarse como complemento; acuerdos con gestorías; plan reducido"],
    ["Dependencia del fundador (enfermedad, saturación)", "Media", "Alto", "Procesos y plantillas estandarizados; trasladar clientes a la herramienta"],
    ["Fracaso o retraso de la herramienta digital", "Media", "Medio", "Desarrollarla solo con tracción demostrada; MVP sencillo; préstamo acotado"],
    ["Cambios en la cotización de autónomos o en la fiscalidad", "Media", "Bajo", "Seguimiento normativo; el plan ya usa las cuotas vigentes por prudencia"],
    ["Inflación superior a la prevista", "Media", "Bajo", "Actualizar precios una vez al año; costes fijos muy bajos"],
    ["Incidente de protección de datos", "Baja", "Alto", "Cumplimiento del RGPD, herramientas con cifrado y seguro de responsabilidad civil"],
    ["Reducción del segmento joven del RETA", "Alta", "Bajo", "Ampliar el público a autónomos noveles de cualquier edad"],
  ], [3000, 1300, 1100, 3670], { numericas: false }),
  fuente("elaboración propia."),
];

// ------------------------------------------------- 9. CONCLUSIONES
const cap9 = [
  h1("10. Conclusiones", false),
  h2("10.1. Conclusiones generales y específicas"),
  p(`Este trabajo se planteaba si sería viable un proyecto dedicado a mejorar la gestión financiera de los autónomos jóvenes. La respuesta, con las hipótesis utilizadas, es afirmativa, aunque con matices importantes.`),
  p(`En primer lugar, **la necesidad existe y está documentada**. Los datos del Banco de España muestran un nivel bajo de conocimientos financieros en la población adulta, algo peor entre los jóvenes, y las quince entrevistas realizadas apuntan a problemas muy concretos: no saber cuánto se gana, mezclar las finanzas personales y las de la actividad y no reservar para los impuestos. El segmento de autónomos menores de 35 años, de unas 560.000 personas, es más que suficiente para un proyecto de este tamaño, aunque pierde peso dentro del colectivo.`),
  p(`En segundo lugar, **el proyecto es económica y financieramente viable**. La inversión inicial es de solo 3.000 €, el punto de equilibrio del primer año se sitúa en unos ${Math.ceil(B.pe[0])} clientes de suscripción y, en el escenario base, la actividad genera un beneficio acumulado de ${eur(sum(A.resultado_neto))} en tres años. El VAN es positivo en los tres escenarios analizados y el balance y la tesorería son coherentes con la cuenta de resultados.`),
  p(`En tercer lugar, **la viabilidad es menos holgada de lo que sugieren la TIR y el payback**. Cuando se incluye el coste de oportunidad del fundador, el VAN del escenario base se reduce a ${eur(V.van_oportunidad)} y pasa a ser negativo en el pesimista. Además, el servicio alcanza el límite de capacidad de una persona al final del tercer año. El éxito del proyecto depende, por tanto, de dos factores: captar clientes al ritmo previsto y conseguir que los productos digitales y la herramienta crezcan sin consumir horas por cliente.`),
  p(`En cuarto lugar, **la rentabilidad puede mejorarse sin invertir más**. El análisis por cliente muestra que captar clientes es barato (la relación LTV/CAC supera 17 en todos los años) y que la restricción es el tiempo del fundador. Cinco medidas orientadas a aumentar los ingresos por hora (actualizar precios, estandarizar el servicio, ofrecer talleres a instituciones, colaborar con gestorías y abaratar la herramienta digital) elevan el VAN con coste de oportunidad de ${eur(PAL[0].van_oportunidad)} a ${eur(PAL[PAL.length - 1].van_oportunidad)}, y la simulación de Montecarlo sitúa en el ${pct(MCO.p_van_op, 0)} la probabilidad de que el proyecto compense el esfuerzo del fundador, frente al ${pct(MC.p_van_op, 0)} del plan base.`),
  p(`En quinto lugar, **el entorno juega a favor en el corto plazo**. La factura electrónica obligatoria y Verifactu obligarán a los autónomos a revisar cómo facturan y controlan sus números entre 2027 y 2029, justo en el horizonte del plan, y la educación financiera es una prioridad institucional reconocida. En sentido contrario, la estadística del INE recuerda que casi la mitad de las empresas nuevas desaparece en sus tres primeros años, lo que obliga a leer los resultados con prudencia.`),
  p(`Estas conclusiones responden a los objetivos planteados en la introducción: se ha descrito el colectivo y sus dificultades (capítulos 3 y 4), se ha definido el modelo de negocio (capítulo 3), se han analizado el mercado, la competencia y el marco legal (capítulos 4 y 5), se han elaborado los estados financieros previsionales a tres años (capítulo 6), se han calculado los criterios de rentabilidad y los ratios (capítulo 7), se han cuantificado las medidas de mejora (capítulo 8) y se han estudiado los riesgos y escenarios (capítulo 9).`),
  h2("10.2. Trascendencia de los resultados"),
  p(`Para un emprendedor, el principal valor del trabajo es que convierte una idea en un conjunto de hipótesis medibles: cuántos clientes hacen falta, cuánto tiempo consume cada uno y qué ocurre si el ritmo de captación es menor. El modelo, publicado en abierto, puede actualizarse con los datos reales de la validación. Desde un punto de vista más general, el trabajo muestra que la falta de educación financiera de los autónomos jóvenes es también una oportunidad de negocio, y que un servicio de bajo coste puede contribuir a un objetivo que las instituciones consideran prioritario.`),
  h2("10.3. Limitaciones y líneas futuras"),
  p(`El trabajo tiene varias limitaciones. Las entrevistas son pocas y no representativas, y no miden la disposición a pagar, que solo puede comprobarse vendiendo el servicio. Los datos de competencias financieras proceden de una encuesta de 2021-2022 y no ofrecen información específica sobre autónomos jóvenes. Las cuotas de autónomos de 2027 a 2029 no están aprobadas, y la tributación por IRPF depende de la situación personal del titular, que se ha simplificado. Por último, la tasa de descuento incluye una prima específica que es, inevitablemente, una estimación.`),
  p(`Como líneas futuras se proponen ejecutar el plan de validación de la Tabla 3.2, preferiblemente dentro de un programa como el Semillero de Ideas de Zaragoza Activa, y actualizar el modelo con los datos reales, explorar acuerdos con gestorías y estudiar la transformación en sociedad limitada si la herramienta digital alcanza suficientes usuarios, lo que abriría la puerta a la financiación de ENISA o de inversores privados.`),
];

// ------------------------------------------------- BIBLIOGRAFÍA
const ref = (t) => new Paragraph({
  alignment: AlignmentType.LEFT, indent: { left: 567, hanging: 567 }, spacing: { after: 100, line: 240, lineRule: "auto" }, children: require("./helpers").runs(t, { size: 22 }),
});
const consulta = "[Consulta: 28/10/2026]";
const biblio = [
  h1("Bibliografía"),
  ...[
    `AEPD, AGENCIA ESPAÑOLA DE PROTECCIÓN DE DATOS (2026): *Orientación a pequeñas y medianas empresas (pymes)* [en línea]. Madrid. https://www.aepd.es/derechos-y-deberes/cumple-tus-deberes/directrices-de-aplicacion/pymes ${consulta}.`,
    `AGENCIA TRIBUTARIA (2026a): *Tipos de gravamen del Impuesto sobre Sociedades* [en línea]. Madrid. https://sede.agenciatributaria.gob.es/Sede/impuesto-sobre-sociedades/campana-sociedades/cuestiones-tener-cuenta-liquidacion-impuesto-2025/tipos-gravamen.html ${consulta}.`,
    `AGENCIA TRIBUTARIA (2026b): *Factura electrónica: calendario de implantación* [en línea]. Madrid. https://sede.agenciatributaria.gob.es/Sede/iva/factura-electronica/calendario-implantacion.html ${consulta}.`,
    `AGENCIA TRIBUTARIA (2026c): *Nota informativa: ampliación del plazo de adaptación de los sistemas informáticos de facturación (Verifactu)* [en línea]. Madrid. https://sede.agenciatributaria.gob.es/Sede/iva/sistemas-informaticos-facturacion-verifactu/nota-informativa-ampliacion-plazo-adaptacion-facturacion.html ${consulta}.`,
    `AGENCIA TRIBUTARIA (2026d): *Estimación directa simplificada* [en línea]. Madrid. https://sede.agenciatributaria.gob.es/Sede/irpf/empresarios-individuales-profesionales/regimenes-determinar-rendimiento-actividad/estimacion-directa-simplificada.html ${consulta}.`,
    `AGENCIA TRIBUTARIA (2026e): *Retenciones e ingresos a cuenta del IRPF en el ejercicio 2026* [en línea]. Madrid. https://sede.agenciatributaria.gob.es/static_files/Sede/Programas_ayuda/Retenciones/2026/CUADRO_TIPOS_RETENCION_IRPF_2026.pdf ${consulta}.`,
    `AGENCIA TRIBUTARIA (2026f): *Manual práctico de Renta 2025. Gravamen autonómico: Comunidad Autónoma de Aragón* [en línea]. Madrid. https://sede.agenciatributaria.gob.es/Sede/eu_es/ayuda/manuales-videos-folletos/manuales-practicos/irpf-2025/c15-calculo-impuesto-determinacion-cuotas-integras/gravamen-base-liquidable-general/gravamen-autonomico/comunidad-autonoma-aragon.html ${consulta}.`,
    `AYUNTAMIENTO DE ZARAGOZA (2026): *Semillero de Ideas de Zaragoza Activa* [en línea]. Zaragoza. https://www.zaragoza.es/semillerodeideas ${consulta}.`,
    `BANCO DE ESPAÑA (2021): *Encuesta de Competencias Financieras en las Pequeñas Empresas 2021: principales resultados*. Documentos Ocasionales, Nº 2129. Madrid. https://www.bde.es/f/webbde/SES/Secciones/Publicaciones/PublicacionesSeriadas/DocumentosOcasionales/21/Fich/do2129.pdf ${consulta}.`,
    `BANCO DE ESPAÑA (2022): *Encuesta de Competencias Financieras (ECF) 2021*. Dirección General de Economía y Estadística, Madrid. https://www.bde.es/f/webbe/SES/AnalisisEconomico/Competencis_Financieras/EncuestaCompetencias_2021.pdf ${consulta}.`,
    `BANCO DE ESPAÑA (2025): *Análisis de la Encuesta de Competencias Financieras 2021 en apoyo del Plan de Educación Financiera 2026-2029* [en línea]. Madrid. https://www.bde.es/f/webbe/INF/MenuHorizontal/AreasActuacion/educacion_financiera/analisis-ecf-2021-apoyo-pef-2026-2029.pdf ${consulta}.`,
    `BANCO DE ESPAÑA (2026a): *Proyecciones macroeconómicas e informe trimestral de la economía española. Tercer trimestre 2026* [en línea]. Madrid. https://www.bde.es/wbe/en/publicaciones/analisis-economico-investigacion/proyecciones-macro-informe-trimestral/proyecciones-e-informe-trimestral-de-la-economia-espanola-tercer-trimestre-2026.html ${consulta}.`,
    `BANCO DE ESPAÑA (2026b): *Boletín Estadístico. Cuadro 19.6: tipos de interés TAE de nuevas operaciones* [en línea]. Madrid. https://www.bde.es/webbe/es/estadisticas/compartido/datos/pdf/a1906.pdf ${consulta}.`,
    `BERGER, A. N. y UDELL, G. F. (1998): "The economics of small business finance: The roles of private equity and debt markets in the financial growth cycle". *Journal of Banking & Finance*, Vol. 22, Nº 6-8, pp. 613-673.`,
    `BILLEO (2026): *Nota de datos: precios de gestorías online y software 2026* [en línea]. https://www.billeo.es/data/nota-datos-precios-gestorias-2026.md ${consulta}.`,
    `BLANK, S. (2013): "Why the lean start-up changes everything". *Harvard Business Review*, Vol. 91, Nº 5, pp. 63-72.`,
    `BREALEY, R. A., MYERS, S. C. y ALLEN, F. (2020): *Principles of Corporate Finance* (13ª ed.). McGraw-Hill Education, Nueva York.`,
    `CÁMARA DE COMERCIO DE VALENCIA (2023): *Microcréditos para el emprendimiento MicroBank 2023* [en línea]. Valencia. https://www.camaravalencia.com/wp-content/uploads/2023/05/EMP3_Microcreditos-Emprendimiento-Microbank-2023.pdf ${consulta}.`,
    `CASTILLA Y LEÓN ECONÓMICA (2026): *6 de cada 10 autónomos llevan más de 5 años en su negocio* [en línea]. 9 de febrero de 2026. https://www.castillayleoneconomica.es/6-10-autonomos-llevan-5-anos-frente-negocio/ ${consulta}.`,
    `CNMV, COMISIÓN NACIONAL DEL MERCADO DE VALORES (2026): *Asesoramiento en materia de inversión* [en línea]. Madrid. https://www.cnmv.es/portal/inversor/asesoramiento ${consulta}.`,
    `DAMODARAN, A. (2009): *Valuing Young, Start-up and Growth Companies: Estimation Issues and Valuation Challenges*. Working paper, Stern School of Business, New York University.`,
    `DAMODARAN, A. (2012): *Investment Valuation: Tools and Techniques for Determining the Value of Any Asset* (3ª ed.). John Wiley & Sons, Hoboken.`,
    `DAMODARAN, A. (2026): *Country Default Spreads and Risk Premiums* [en línea]. Actualización de enero de 2026. https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html ${consulta}.`,
    `EMIRODGAR (2026): *Análisis y evolución de los autónomos en España* [en línea]. https://emirodgar.es/analisis-autonomos ${consulta}.`,
    `ENISA (2026): *Línea Emprendedores* [en línea]. Madrid. https://www.enisa.es/es/financiacion/info/lineas-enisa/emprendedores ${consulta}.`,
    `EUROSTAT (2026): *EU sees 2.06 million young entrepreneurs in 2025* [en línea]. Luxemburgo, 21 de abril de 2026. https://ec.europa.eu/eurostat/web/products-eurostat-news/w/edn-20260421-1 ${consulta}.`,
    `FEDERACIÓN NACIONAL DE ASOCIACIONES DE TRABAJADORES AUTÓNOMOS, ATA (2026): "El RETA ganó 39.000 autónomos en 2025, pero ATA prevé peores perspectivas para 2026". *Autónomos y Emprendedor*, 5 de enero de 2026. https://www.autonomosyemprendedor.es/articulo/autonomos/reta-gano-39000-autonomos-2025-menor-crecimiento-peores-perspectivas-2026/20260105114615047592.html ${consulta}.`,
    `GOBIERNO DE ARAGÓN (2026): *Subvenciones para el fomento y consolidación del empleo autónomo. Convocatoria 2026 (Orden EMC/1877/2025)* [en línea]. Zaragoza. https://www.aragon.es/tramitador/-/tramite/programa-emprendedores-autonomos ${consulta}.`,
    `ICO, INSTITUTO DE CRÉDITO OFICIAL (2026): *ICO Empresas y Emprendedores* [en línea]. Madrid. https://www.ico.es/en/ico-empresas-y-emprendedores ${consulta}.`,
    `INE, INSTITUTO NACIONAL DE ESTADÍSTICA (2025): *Demografía Armonizada de Empresas. Año 2023*. Nota de prensa. Madrid. https://www.ine.es/dyngs/Prensa/DAE2023.htm ${consulta}.`,
    `INFOAUTÓNOMOS (2026): *La tarifa plana de 80 euros para autónomos en 2026: requisitos, cuantía y renovación* [en línea]. https://www.infoautonomos.com/seguridad-social/tarifa-plana-autonomos/ ${consulta}.`,
    `LA MONCLOA (2026): *Educación renueva su colaboración con el Banco de España, la CNMV y el Ministerio de Economía para fomentar la educación financiera* [en línea]. 16 de enero de 2026. https://www.lamoncloa.gob.es/serviciosdeprensa/notasprensa/educacion-fp-deportes/paginas/2026/160126-tolon-fomento-educacion-financiera.aspx ${consulta}.`,
    `LEY 35/2006, de 28 de noviembre, del Impuesto sobre la Renta de las Personas Físicas. *Boletín Oficial del Estado*, Nº 285, de 29 de noviembre de 2006.`,
    `LEY 37/1992, de 28 de diciembre, del Impuesto sobre el Valor Añadido. *Boletín Oficial del Estado*, Nº 312, de 29 de diciembre de 1992.`,
    `LEY 18/2022, de 28 de septiembre, de creación y crecimiento de empresas. *Boletín Oficial del Estado*, Nº 234, de 29 de septiembre de 2022.`,
    `LEY 28/2022, de 21 de diciembre, de fomento del ecosistema de las empresas emergentes. *Boletín Oficial del Estado*, Nº 306, de 22 de diciembre de 2022. https://www.boe.es/buscar/act.php?id=BOE-A-2022-21739 ${consulta}.`,
    `LEY ORGÁNICA 3/2018, de 5 de diciembre, de Protección de Datos Personales y garantía de los derechos digitales. *Boletín Oficial del Estado*, Nº 294, de 6 de diciembre de 2018.`,
    `LUSARDI, A. y MITCHELL, O. S. (2014): "The economic importance of financial literacy: Theory and evidence". *Journal of Economic Literature*, Vol. 52, Nº 1, pp. 5-44.`,
    `MINISTERIO DE HACIENDA (2026): *Nota informativa: previsión de modificación de los plazos de aplicación del Real Decreto 1007/2023 (Verifactu)* [en línea]. Madrid, 5 de octubre de 2026. https://www.hacienda.gob.es/sgt/gabsehacienda/nota-informativa-verifactu.pdf ${consulta}.`,
    `MINISTERIO DE TRABAJO Y ECONOMÍA SOCIAL (2026a): *Estadística de personas trabajadoras por cuenta propia afiliadas a la Seguridad Social. Marzo 2026* [en línea]. Madrid. https://www.mites.gob.es/estadisticas/AUT/AUT_03_2026.pdf ${consulta}.`,
    `MINISTERIO DE TRABAJO Y ECONOMÍA SOCIAL (2026b): *Informe trimestral de jóvenes y mercado de trabajo*, Nº 49, mayo de 2026 [en línea]. Madrid. https://www.mites.gob.es/ficheros/ministerio/sec_trabajo/analisis_mercado_trabajo/jovenes/2026/Informe_Jovenes_Num49_mayo_2026.pdf ${consulta}.`,
    `MYERS, S. C. y MAJLUF, N. S. (1984): "Corporate financing and investment decisions when firms have information that investors do not have". *Journal of Financial Economics*, Vol. 13, Nº 2, pp. 187-221.`,
    `OBSERVATORIO DEL EMPRENDIMIENTO DE ESPAÑA (2025): *Global Entrepreneurship Monitor. Informe GEM España 2024-2025*. Editorial de la Universidad de Cantabria, Santander. https://www.editorial.unican.es/libro/global-entrepreneurship-monitor-informe-gem-espana-2024-2025 ${consulta}.`,
    `OBSERVATORIO ESPAÑOL DE LA ECONOMÍA SOCIAL Y DEL TRABAJO AUTÓNOMO (2025): *Aumenta considerablemente el número y porcentaje de trabajadores autónomos con más de 50 años de edad* [en línea]. https://observatorioeconomiasocial.es/aumenta-considerablemente-el-numero-y-porcentaje-de-trabajadores-autonomos-con-mas-de-50-anos-de-edad/ ${consulta}.`,
    `OCDE (2020): *OECD/INFE 2020 International Survey of Adult Financial Literacy*. OECD Publishing, París. https://www.oecd.org/en/publications/2020/06/oecd-infe-2020-international-survey-of-adult-financial-literacy_bbad9b27.html ${consulta}.`,
    `OCDE (2023): *OECD/INFE 2023 International Survey of Adult Financial Literacy*. OECD Business and Finance Policy Papers, Nº 39. OECD Publishing, París. https://www.oecd.org/en/publications/oecd-infe-2023-international-survey-of-adult-financial-literacy_56003a32-en.html ${consulta}.`,
    `ORDEN HAC/1028/2026, de 2 de octubre, por la que se regula la solución pública de facturación electrónica. *Boletín Oficial del Estado*, Nº 247, de 5 de octubre de 2026. https://www.boe.es/diario_boe/txt.php?id=BOE-A-2026-20587 ${consulta}.`,
    `ORDEN PJC/297/2026, de 30 de marzo, por la que se desarrollan las normas legales de cotización a la Seguridad Social, desempleo, protección por cese de actividad, Fondo de Garantía Salarial y formación profesional para el ejercicio 2026. *Boletín Oficial del Estado*, Nº 79, de 31 de marzo de 2026. https://www.boe.es/diario_boe/txt.php?id=BOE-A-2026-7296 ${consulta}.`,
    `OSTERWALDER, A. y PIGNEUR, Y. (2010): *Business Model Generation*. John Wiley & Sons, Hoboken.`,
    `REAL DECRETO 126/2026, de 17 de febrero, por el que se fija el salario mínimo interprofesional para 2026. *Boletín Oficial del Estado*, de 19 de febrero de 2026.`,
    `REAL DECRETO 238/2026, de 25 de marzo, por el que se desarrolla el sistema de facturación electrónica obligatoria entre empresarios y profesionales. *Boletín Oficial del Estado*, Nº 79, de 31 de marzo de 2026. https://boe.es/eli/es/rd/2026/03/25/238 ${consulta}.`,
    `REAL DECRETO-LEY 13/2022, de 26 de julio, por el que se establece un nuevo sistema de cotización para los trabajadores por cuenta propia o autónomos. *Boletín Oficial del Estado*, Nº 179, de 27 de julio de 2022.`,
    `REGLAMENTO (UE) 2016/679 del Parlamento Europeo y del Consejo, de 27 de abril de 2016, relativo a la protección de las personas físicas en lo que respecta al tratamiento de datos personales (Reglamento General de Protección de Datos). *Diario Oficial de la Unión Europea*, L 119, de 4 de mayo de 2016.`,
    `RIES, E. (2011): *The Lean Startup*. Crown Business, Nueva York.`,
    `SAN-MARTÍN, [[iniciales]], TORRE-OLMO, [[iniciales]], SAMSAMI, [[iniciales]], SCHØTT, T. y TRUEBA-CASTAÑEDA, [[iniciales]] (2025): *El conocimiento financiero en España: brechas y desafíos*. Cuadernos de divulgación científica, Nº 5/2025. Observatorio del Emprendimiento de España, Madrid. https://observatoriodelemprendimiento.es/wp-content/uploads/El-conocimiento-financiero-en-Espana.pdf ${consulta}.`,
    `SEGURIDAD SOCIAL (2026): *Cotización de trabajadores autónomos: bases y tramos de rendimientos 2026* [en línea]. Madrid. https://www.seg-social.es/wps/portal/wss/internet/Trabajadores/CotizacionRecaudacionTrabajadores/10721/10724/1320/1322/boni9 ${consulta}.`,
    `THE OBJECTIVE (2026): *La Seguridad Social sumó 92.327 afiliados en septiembre, el mayor aumento en este mes* [en línea]. 2 de octubre de 2026. https://theobjective.com/economia/2026-10-02/empleo-incremento-afiliados/ ${consulta}.`,
    `VT MARKETS (2026): *La rentabilidad del bono español a 10 años repunta ligeramente en la subasta* [en línea]. 2 de julio de 2026. https://www.vtmarkets.com/es-eu/live-updates/la-rentabilidad-del-bono-espanol-a-10-anos-repunta-ligeramente-en-la-subasta-mientras-la-inflacion-de-la-eurozona-mantiene-un-panorama-de-tipos-restrictivo/ ${consulta}.`,
  ].map(ref),
];

// ------------------------------------------------- ANEXOS
const anexos = [
  h1("Anexo I. Análisis de datos de mercado"),
  p(`Los capítulos 3 y 4 se apoyan en cifras publicadas por el Banco de España, la Seguridad Social y ATA. Para no quedarme con la cita puntual de cada informe, he reunido esos datos en un pequeño análisis propio con el objetivo de responder a tres preguntas: qué peso tienen los autónomos jóvenes en el total, si tienen peor comprensión financiera que el resto y cómo está evolucionando el colectivo.`),
  p(`**¿Qué peso tienen los autónomos jóvenes?** Los menores de 30 años representan solo el 7 % de los afiliados al RETA, mientras que los mayores de 50 suponen el 47,2 % (Figura I.1). El tramo de 30 a 49 años no se publica de forma directa y se ha obtenido por diferencia. Si se toma el corte de 35 años, el informe de ATA sitúa a los menores de esa edad en el 16,1 % del total, unas ${miles(Math.round(tam / 1000) * 1000)} personas con la afiliación de septiembre de 2026.`),
  leyenda("Figura", "I.1", "Distribución de los autónomos por edad en España (2025)"),
  imagen("fig_anexo_reta_edad.png", 13),
  fuente("elaboración propia a partir de Observatorio Español de la Economía Social y del Trabajo Autónomo (2025)."),
  p(`**¿Tienen los jóvenes peor comprensión financiera?** El índice que resume los aciertos en las tres preguntas básicas de la ECF (inflación, interés compuesto y diversificación del riesgo) es del 52 % entre los 18 y los 34 años, frente al 56 % del grupo de 45 a 64 años, el que mejor puntúa, y a una media nacional del 53 % (Figura I.2). La diferencia, de cuatro puntos, no es abismal, pero confirma la hipótesis de partida y coincide con lo que se escuchó en las entrevistas: los autónomos jóvenes no parten de cero, pero su base es algo más débil justo cuando empiezan a decidir solos.`),
  leyenda("Figura", "I.2", "Aciertos medios en las preguntas financieras básicas por edad (2021)"),
  imagen("fig_anexo_big3.png", 13),
  fuente("elaboración propia a partir de Banco de España (2022)."),
  p(`**¿Cómo evoluciona el colectivo?** Este fue el hallazgo que no esperaba. Según el análisis de las series de la Seguridad Social de Emirodgar (2026), España perdió un 26,5 % de sus autónomos menores de 35 años entre 2009 y 2019, y el tramo de 26 a 35 años ha perdido más de 136.000 afiliados desde 2013. En 2025, el 73,1 % del crecimiento neto del RETA se explica por autónomos mayores de 56 años, y un 35,2 % corresponde solo a mayores de 65 (Figura I.3). El segmento al que se dirige Flujo es amplio, pero no crece por sí solo.`),
  p(`**Contraste con la estadística oficial.** Los datos del Ministerio de Trabajo y Economía Social corroboran el envejecimiento, aunque matizan la tendencia reciente. En marzo de 2026 el RETA tenía 3.437.099 afiliados, de los que el 31,5 % tenía entre 46 y 55 años y el 23,9 % entre 56 y 65 (Ministerio de Trabajo y Economía Social, 2026a). Los jóvenes de 16 a 29 años sumaban 249.456, un 7,3 % del total, y crecían un 5,1 % interanual (Ministerio de Trabajo y Economía Social, 2026b). Eurostat (2026) sitúa a España entre los países de la Unión Europea con menor proporción de autónomos jóvenes: los de 20 a 29 años son el 5,9 % de los autónomos españoles, frente al 7,9 % de la media europea.`),
  leyenda("Figura", "I.3", "Reparto del crecimiento neto del RETA en 2025 por edad"),
  imagen("fig_anexo_crecimiento.png", 13),
  fuente("elaboración propia a partir de Emirodgar (2026)."),

  h1("Anexo II. Modelo de cálculo y reproducibilidad"),
  p(`El plan económico-financiero se ha calculado con un modelo mensual de 36 periodos programado en Python. El modelo parte de un único fichero de hipótesis y genera, para cada escenario, la cuenta de resultados, la tesorería con el calendario real del IVA y del IRPF, el balance, los criterios de rentabilidad y el análisis de sensibilidad. Como comprobación de coherencia, el programa verifica que el activo y la suma del patrimonio neto y el pasivo coinciden en cada cierre. Los resultados se vuelcan también en una hoja de cálculo (Modelo_financiero_Flujo.xlsx) con el detalle mes a mes.`),
  p(`El código, los datos de mercado con la fuente de cada cifra y los gráficos de este trabajo están disponibles en el repositorio https://github.com/Andres-Leach/flujo-tfg. Para reproducir los resultados basta con ejecutar, en este orden, modelo.py, graficos.py y exportar_excel.py, y cualquier hipótesis puede cambiarse en hipotesis.py para recalcular todo el plan.`),
  p(`Todos los porcentajes del Anexo I proceden de fuentes secundarias ya citadas en la bibliografía. La aportación propia consiste en reunir, cruzar y representar esos datos, y en construir el modelo financiero; los únicos datos primarios del trabajo son las entrevistas.`),

  h1("Anexo III. Entrevistas exploratorias"),
  p(`**Método.** Se realizaron quince entrevistas semiestructuradas a autónomos y *freelancers* de entre 20 y 35 años, contactados a través de redes personales y comunidades profesionales. Las entrevistas se hicieron entre [[mes y año]] y [[mes y año]], por [[videollamada / en persona]], con una duración aproximada de [[XX]] minutos. Se garantizó el anonimato de los participantes, que se identifican con un número.`),
  p(`**Guion.** Las preguntas siguieron este orden, aunque se adaptaron a cada conversación:`),
  ...vinetas([
    "¿A qué te dedicas y cuánto tiempo llevas como autónomo?",
    "¿Cómo organizas tus ingresos y gastos? ¿Usas alguna herramienta?",
    "¿Separas el dinero de tu actividad del personal?",
    "Si tuvieras que decirme cuánto ganaste el mes pasado después de impuestos y gastos, ¿sabrías hacerlo?",
    "¿Cómo calculas lo que tienes que reservar para el IVA y el IRPF cada trimestre?",
    "¿Has tenido alguna sorpresa negativa con Hacienda o con la Seguridad Social?",
    "¿Trabajas con una gestoría? ¿Qué te aporta y qué echas en falta?",
    "¿Pagarías por un servicio que te ayudara a entender y organizar tu dinero? ¿Cuánto te parecería razonable?",
  ]),
  leyenda("Tabla", "III.1", "Perfil de los entrevistados y respuestas principales"),
  tabla(["Nº", "Edad", "Actividad", "Años como autónomo", "¿Sabe cuánto gana neto?", "¿Separa cuentas?", "¿Gestoría?"],
    Array.from({ length: 15 }, (_, i) => [String(i + 1), "[[ ]]", "[[ ]]", "[[ ]]", "[[ ]]", "[[ ]]", "[[ ]]"]),
    [600, 800, 2270, 1300, 1500, 1300, 1300], { numericas: false }),
  fuente("elaboración propia a partir de las entrevistas realizadas."),
  p(`**Resumen de resultados.** De los quince entrevistados, [[X]] no sabían calcular cuánto ganaban una vez descontados impuestos y gastos, [[X]] mezclaban la cuenta personal y la de la actividad y [[X]] trabajaban con una gestoría. [[X]] afirmaron que pagarían por un servicio como el descrito, con un precio razonable de entre [[X]] y [[X]] € al mes. Entre las respuestas más repetidas destacan: [[dos o tres frases textuales y anónimas de los entrevistados]].`),
];

anexos.push(
  h1("Anexo IV. Fuentes oficiales consultadas"),
  p(`La tabla siguiente relaciona cada fuente oficial con el dato que se ha contrastado y el apartado del trabajo en que se utiliza. Las direcciones web completas figuran en la bibliografía; todas se consultaron el 28 de octubre de 2026.`),
  leyenda("Tabla", "IV.1", "Fuentes oficiales y datos contrastados"),
  tabla(["Organismo", "Fuente", "Dato contrastado", "Apartado"], [
    ["BOE / Seguridad Social", "Orden PJC/297/2026; tablas de cotización 2026", "Bases mínimas por tramo y tipo total del 31,5 %", "5.2, 6.4"],
    ["Seguridad Social / BOE", "Real Decreto-ley 13/2022", "Tarifa plana de 80 € y prórroga si el rendimiento no supera el SMI", "5.2"],
    ["BOE", "Real Decreto 126/2026", "SMI de 2026: 17.094 € anuales", "5.2, 7.3"],
    ["Agencia Tributaria", "Estimación directa simplificada; cuadro de retenciones 2026", "5 % de gastos de difícil justificación (máx. 2.000 €); retención del 7 %; regla del 70 % del modelo 130", "5.2, 6.5"],
    ["Agencia Tributaria", "Manual de Renta 2025, escala de Aragón", "Tramos autonómicos del IRPF", "6.5"],
    ["Agencia Tributaria", "Tipos del Impuesto sobre Sociedades", "15 % nueva creación; 17 % y 20 % microempresas en 2027", "5.2"],
    ["BOE / Agencia Tributaria", "RD 238/2026, Orden HAC/1028/2026 y calendario", "Factura electrónica obligatoria: 6/10/2028 para autónomos", "4.1, 5.2"],
    ["Agencia Tributaria / Hacienda", "Nota Verifactu (RDL 15/2025) y nota de 5/10/2026", "Verifactu el 1/7/2027; aplazamiento anunciado a octubre de 2028", "4.1, 5.2"],
    ["BOE / ENISA", "Ley 28/2022 de empresas emergentes; certificación ENISA", "Requisitos y ventajas fiscales de las startups", "5.2"],
    ["Ministerio de Trabajo", "Estadística AUT, marzo de 2026", "Distribución de los autónomos por edad", "4.1, Anexo I"],
    ["Ministerio de Trabajo", "Informe de jóvenes y mercado de trabajo nº 49", "249.456 autónomos de 16 a 29 años; +5,1 % interanual", "4.1, Anexo I"],
    ["Eurostat", "Nota de 21 de abril de 2026", "Autónomos de 20 a 29 años: 5,9 % en España y 7,9 % en la UE", "4.1, Anexo I"],
    ["INE", "Demografía Armonizada de Empresas 2023", "45,3 % de las empresas desaparecen en tres años", "8.3"],
    ["Banco de España", "ECF 2021 y análisis para el PEF 2026-2029", "Conocimientos financieros de la población y de los ocupados", "1.1, 3.1"],
    ["Banco de España", "ECF en las Pequeñas Empresas 2021", "Menores competencias en empresas de menos de 20 trabajadores", "2.3"],
    ["Banco de España", "Proyecciones del tercer trimestre de 2026", "Inflación prevista del 3,7 % en 2027", "6.4"],
    ["Banco de España", "Boletín Estadístico, cuadro 19.6", "TAE del 4,58 % en créditos de hasta 250.000 €", "6.2"],
    ["OCDE", "OECD/INFE 2023 Survey of Adult Financial Literacy", "34 % de adultos alcanza el mínimo de alfabetización financiera", "1.1"],
    ["Observatorio del Emprendimiento", "El conocimiento financiero en España (2025)", "Conocimiento financiero de los emprendedores (escala 1-5)", "1.1"],
    ["CNMV", "Asesoramiento en materia de inversión", "Límite entre educación financiera y asesoramiento regulado", "5.3"],
    ["AEPD", "Facilita RGPD y Facilita EMPRENDE", "Herramientas gratuitas de cumplimiento", "5.3"],
    ["ICO / ENISA / Gobierno de Aragón", "Líneas de financiación y ayudas 2026", "Condiciones de financiación pública", "2.5"],
    ["Ayuntamiento de Zaragoza", "Semillero de Ideas (Zaragoza Activa)", "Preincubación gratuita para graduados", "5.1, 9.3"],
    ["La Moncloa", "Nota de 16 de enero de 2026", "Renovación del Plan de Educación Financiera", "4.1"]
  ], [1900, 2700, 3270, 1200], { numericas: false }),
  fuente("elaboración propia."),
);

// ------------------------------------------------- DOCUMENTO
const pie = new Footer({
  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 20 })] })],
});
const pagina = {
  page: { size: { width: 11906, height: 16838 }, margin: { top: 1418, bottom: 1418, left: 1701, right: 1701, header: 709, footer: 709 } },
};

const doc = new Document({
  creator: "Andrés Leach Sagarra",
  title: "Viabilidad financiera de Flujo",
  features: { updateFields: true },
  styles: {
    default: { document: { run: { font: FUENTE, size: 24 }, paragraph: { spacing: { before: 0, after: 120, line: 360, lineRule: "auto", lineRule: "auto" } } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FUENTE, size: 28, bold: true }, paragraph: { spacing: { before: 0, after: 240 }, outlineLevel: 0, keepNext: true } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FUENTE, size: 24, bold: true }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1, keepNext: true } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FUENTE, size: 24, bold: true, italics: true }, paragraph: { spacing: { before: 120, after: 120 }, outlineLevel: 2, keepNext: true } },
    ],
  },
  numbering: { config: [{ reference: "vinetas", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 284 } } } }] }] },
  sections: [
    { properties: { ...pagina }, children: portada },
    { properties: { ...pagina, type: SectionType.NEXT_PAGE }, children: info },
    { properties: { ...pagina, type: SectionType.NEXT_PAGE, page: { ...pagina.page, pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL } } },
      footers: { default: pie },
      children: [...indice, ...cap1, ...cap2, ...cap3, ...cap4, ...cap5, ...cap6, ...cap7, ...capMejora, ...cap8, ...cap9, ...biblio, ...anexos] },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(__dirname, "..", "TFG_Flujo_v2.docx");
  fs.writeFileSync(out, buf);
  console.log("guardado", out);
});
