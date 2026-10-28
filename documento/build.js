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
const resumenEs = `Los jóvenes que trabajan por cuenta propia tienen que gestionar ingresos irregulares, impuestos trimestrales y gastos sin haber recibido, en la mayoría de los casos, formación financiera. La Encuesta de Competencias Financieras del Banco de España muestra que solo el 19 % de los adultos responde bien a las tres preguntas financieras básicas, y quince entrevistas exploratorias a autónomos de 20 a 35 años apuntan en la misma dirección. El objetivo de este trabajo es evaluar la viabilidad económico-financiera de Flujo, un proyecto que ofrecería organización financiera, educación práctica y, más adelante, una herramienta digital a este colectivo. Para ello se combinan fuentes secundarias, las entrevistas, el Business Model Canvas, un análisis de mercado y competencia, y un plan económico-financiero a tres años con tres escenarios. En el escenario base los ingresos pasan de ${eur(A.ingresos[0])} a ${eur(A.ingresos[2])}, el punto de equilibrio se sitúa en ${Math.ceil(B.pe[0])} clientes de suscripción y el VAN a tres años, descontado al ${pct(R.tasa, 0)}, es de ${eur(V.van)}. Sin embargo, cuando se descuenta el coste de oportunidad del emprendedor (un sueldo equivalente al SMI), el VAN baja a ${eur(V.van_oportunidad)} y pasa a ser negativo en el escenario pesimista. El proyecto es viable, pero su margen depende de la captación de clientes y de la capacidad de una sola persona, y la disposición real a pagar aún debe validarse con clientes.`;
const resumenEn = `Young self-employed workers have to manage irregular income, quarterly taxes and expenses, usually without any financial training. The Bank of Spain's Financial Competences Survey shows that only 19% of adults answer the three basic financial questions correctly, and fifteen exploratory interviews with self-employed people aged 20 to 35 point in the same direction. This study assesses the economic and financial feasibility of Flujo, a project that would offer financial organisation, practical education and, later, a digital tool to this group. It combines secondary sources, the interviews, the Business Model Canvas, a market and competitor analysis and a three-year financial plan with three scenarios. In the base scenario revenue grows from ${eur(A.ingresos[0])} to ${eur(A.ingresos[2])}, break-even is reached with ${Math.ceil(B.pe[0])} subscription clients and the three-year NPV, discounted at ${pct(R.tasa, 0)}, is ${eur(V.van)}. Once the founder's opportunity cost (a salary equal to the Spanish minimum wage) is deducted, the NPV falls to ${eur(V.van_oportunidad)} and turns negative in the pessimistic scenario. The project is feasible, but its margin depends on client acquisition and on the capacity of a single person, and real willingness to pay still has to be validated with clients.`;
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
  p(`Esa tarea exige unos conocimientos que la mayoría de la población no tiene. Según la Encuesta de Competencias Financieras (ECF) del Banco de España, con datos recogidos entre 2021 y 2022, solo un 41 % de los adultos responde bien a una pregunta sencilla sobre interés compuesto y apenas un 19 % acierta a la vez las preguntas sobre inflación, interés compuesto y diversificación del riesgo (Banco de España, 2022). El problema está reconocido institucionalmente: el propio Banco de España ha reanalizado la encuesta como base del Plan de Educación Financiera 2026-2029 (Banco de España, 2025). En el caso de quienes emprenden, el Informe GEM España 2024-2025 recoge que ocho de cada diez emprendedores consideran tener los conocimientos y habilidades necesarios para emprender (Observatorio del Emprendimiento de España, 2025), una autopercepción que contrasta con los resultados objetivos de la ECF y que encaja con lo que la literatura llama exceso de confianza financiera (Lusardi y Mitchell, 2014).`),
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
    "Analizar los riesgos del proyecto y su comportamiento en distintos escenarios.",
  ]),
  h2("1.3. Trascendencia del trabajo"),
  p(`El interés del trabajo es triple. En el plano **empresarial**, ofrece un análisis de viabilidad que podría servir de base real para lanzar el proyecto, con hipótesis explícitas que se pueden contrastar a medida que lleguen los primeros clientes. En el plano **social**, aborda un problema, la baja educación financiera, que las instituciones consideran prioritario y que en el caso de los autónomos tiene consecuencias directas: recargos por errores fiscales, falta de colchón de liquidez o decisiones de gasto poco meditadas. En el plano **universitario**, aplica de forma conjunta herramientas de varias asignaturas del grado (dirección estratégica, contabilidad, fiscalidad y finanzas) a un caso concreto, y el modelo de cálculo se publica en abierto para que cualquier lector pueda reproducirlo.`),
  h2("1.4. Metodología y fuentes"),
  p(`El trabajo sigue un enfoque de estudio de caso con tres tipos de fuentes. Las **fuentes secundarias** proceden de organismos oficiales y entidades de referencia: Banco de España, Seguridad Social, Agencia Tributaria, Observatorio del Emprendimiento de España y la Federación Nacional de Asociaciones de Trabajadores Autónomos (ATA). La **fuente primaria** son las quince entrevistas exploratorias semiestructuradas ya citadas (el guion se recoge en el Anexo III). Por último, el análisis se apoya en **herramientas de gestión** consolidadas: la metodología Lean Startup (Ries, 2011), el Business Model Canvas (Osterwalder y Pigneur, 2010), el análisis DAFO y los criterios clásicos de evaluación de inversiones (Brealey, Myers y Allen, 2020).`),
  p(`Las proyecciones financieras se han calculado mes a mes con un modelo propio programado en Python, que se entrega también en una hoja de cálculo. Todas las hipótesis se explican en el capítulo 6 y el código y los datos están disponibles en un repositorio público (ver Anexo II).`),
  h2("1.5. Estructura del trabajo"),
  p(`Tras esta introducción, el capítulo 2 recoge el marco teórico: el concepto de start-up, el modelo de negocio, la educación financiera y los criterios de viabilidad y financiación. El capítulo 3 presenta la necesidad detectada y la propuesta de Flujo, y el capítulo 4 analiza el mercado, la competencia y el DAFO. El capítulo 5 trata los recursos necesarios y el marco legal y fiscal. Los capítulos 6, 7 y 8 forman el núcleo del trabajo: el plan económico-financiero, el análisis de viabilidad y rentabilidad, y el análisis de riesgos y escenarios. El capítulo 9 recoge las conclusiones y las limitaciones. Los anexos incluyen el análisis de datos de mercado, la nota sobre el modelo de cálculo y el guion de las entrevistas.`),
];

// ---------------------------------------------------------- 2. MARCO TEÓRICO
const cap2 = [
  h1("2. Marco teórico"),
  h2("2.1. La start-up y el enfoque Lean Startup"),
  p(`Ries (2011) define una start-up como una institución humana diseñada para crear un producto o servicio en condiciones de incertidumbre extrema. Lo que la distingue de una pequeña empresa tradicional no es su tamaño sino que todavía no sabe con certeza quién es su cliente ni cuánto está dispuesto a pagar. Blank (2013) lo resume diciendo que una start-up busca un modelo de negocio repetible y escalable, mientras que una empresa consolidada se limita a ejecutarlo.`),
  p(`De esa incertidumbre nace el enfoque Lean Startup, que propone sustituir los planes de negocio cerrados por ciclos cortos de construir, medir y aprender. La idea central es lanzar un producto mínimo viable, observar cómo responden los clientes y ajustar la propuesta antes de comprometer recursos importantes. Para un proyecto como Flujo, este enfoque tiene una consecuencia práctica: empezar como un servicio prestado de forma manual, que apenas requiere inversión, y dejar el desarrollo tecnológico para cuando exista evidencia de que hay clientes dispuestos a pagar.`),
  h2("2.2. El modelo de negocio como herramienta de análisis"),
  p(`Osterwalder y Pigneur (2010) describen el modelo de negocio como la lógica con la que una organización crea, entrega y captura valor, y proponen representarlo en un lienzo de nueve bloques: segmentos de clientes, propuesta de valor, canales, relación con clientes, fuentes de ingresos, recursos clave, actividades clave, socios clave y estructura de costes. Su utilidad para este trabajo es doble. Por un lado, obliga a explicitar todas las hipótesis del proyecto en un formato que se puede revisar de un vistazo. Por otro, conecta directamente con el plan financiero, porque los bloques de ingresos y costes son el punto de partida de las proyecciones.`),
  h2("2.3. Educación financiera y trabajo autónomo"),
  p(`La OCDE define la educación financiera como la combinación de conocimientos, habilidades, actitudes y comportamientos necesarios para tomar decisiones financieras adecuadas (OCDE, 2020). Lusardi y Mitchell (2014) revisan la evidencia internacional y concluyen que los niveles de conocimiento financiero son bajos incluso en países desarrollados, que los jóvenes y las personas con menos estudios puntúan peor y que un mayor conocimiento se asocia con más ahorro, mejor planificación y menos errores costosos. También señalan que mucha gente sobrestima lo que sabe, lo que reduce su propensión a buscar ayuda.`),
  p(`En el caso de los autónomos el problema se agrava por tres razones. Sus ingresos son irregulares, lo que exige planificar la liquidez. No tienen un empleador que les retenga y cotice por ellos, de modo que la reserva para impuestos y Seguridad Social depende de su propia disciplina. Y la frontera entre el patrimonio personal y el de la actividad es difusa, porque responden con todos sus bienes. Para un asalariado un error financiero suele tener un coste limitado; para un autónomo puede traducirse en recargos, en tensiones de tesorería o en el cierre de la actividad.`),
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
    ["Subvenciones autonómicas al autoempleo (Aragón)", "Inicio", "Requieren estar desempleado e inscrito; cuantías variables", "Incierto; no se incluye en el escenario base"],
    ["ENISA Jóvenes Emprendedores", "Crecimiento", "Préstamo participativo de 25.000 a 75.000 €; exige sociedad mercantil", "Solo si se constituye una sociedad"],
    ["Business angels y capital riesgo", "Escalado", "Entrada en el capital; buscan crecimiento rápido", "Bajo en el horizonte analizado"],
  ], [2300, 1400, 2900, 2470], { numericas: false }),
  fuente("elaboración propia a partir de Berger y Udell (1998), Cámara de Comercio de Valencia (2023), ENISA (2026) y Gobierno de Aragón (2026)."),
];

// ---------------------------------------------------------- 3. EL PROYECTO
const cap3 = [
  h1("3. El proyecto Flujo: necesidad detectada y propuesta"),
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
  h1("4. Análisis del mercado y del entorno"),
  h2("4.1. Tamaño y evolución del mercado"),
  p(`El RETA contaba con 3.483.645 afiliados medios en septiembre de 2026 (The Objective, 2026). Según el informe de ATA sobre el perfil del autónomo, los menores de 35 años representan el 16,1 % del total (Castilla y León Económica, 2026), lo que sitúa el mercado potencial de Flujo en torno a ${miles(Math.round(tam / 1000) * 1000)} personas. No todas serían clientes posibles, porque una parte ya tiene formación financiera o no está dispuesta a pagar por este tipo de servicio, pero la cifra sirve para poner en contexto el objetivo del plan: los ${r0(A.activos_fin[2])} clientes de suscripción previstos al final del año 3 equivalen a menos del 0,01 % de ese mercado. Para Flujo el tamaño del mercado no es una restricción; el límite está en la capacidad de captar y atender clientes.`),
  p(`La tendencia, sin embargo, no es favorable. El análisis de los datos de la Seguridad Social que se recoge en el Anexo I muestra que el colectivo autónomo está envejeciendo: los menores de 30 años son solo el 7 % del RETA, mientras que los mayores de 50 ya suponen el 47,2 % (Observatorio Español de la Economía Social y del Trabajo Autónomo, 2025), y el 73,1 % del crecimiento del RETA en 2025 se explica por autónomos mayores de 56 años (Emirodgar, 2026). Incluyo este dato aunque no favorezca al proyecto porque es más útil dimensionar el mercado con datos reales que con un titular optimista: el segmento joven es amplio, pero no está creciendo, y Flujo tendría que ganar cuota en lugar de esperar a que el mercado se expanda.`),
  p(`A favor del proyecto juegan otras dos tendencias. Por un lado, la digitalización de los servicios financieros y la normalización de la facturación electrónica obligan a los autónomos a ordenar sus números. Por otro, las instituciones reconocen cada vez más la necesidad de mejorar la educación financiera de la población adulta (Banco de España, 2025).`),
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
  p(`La comparación deja una conclusión que conviene no esconder: el precio de Flujo es superior al de una gestoría *online*, y un cliente que ya pague una gestoría tendría que asumir los dos costes. Por eso Flujo se posiciona como complemento de la gestoría y no como sustituto, y una línea de trabajo lógica es colaborar con gestorías que deriven a sus clientes jóvenes. El efecto de un precio más bajo se analiza en el capítulo 8.`),
  h2("4.3. Análisis DAFO"),
  p(`La Tabla 4.2 resume el diagnóstico interno y externo del proyecto.`),
  leyenda("Tabla", "4.2", "Análisis DAFO de Flujo"),
  tabla(["Debilidades", "Amenazas"], [
    ["• Dependencia total del tiempo del fundador, que limita el crecimiento del servicio", "• Segmento joven que pierde peso dentro del RETA"],
    ["• Sin marca ni reputación previas, en un servicio que exige confianza", "• Competencia de gestorías y aplicaciones consolidadas"],
    ["• Precio superior al de las gestorías online", "• Posibles cambios en la cotización de autónomos a partir de 2027"],
  ], [4535, 4535], { numericas: false }),
  tabla(["Fortalezas", "Oportunidades"], [
    ["• Inversión inicial muy reducida y punto de equilibrio bajo", "• Más de 3,4 millones de autónomos, unos 560.000 menores de 35 años"],
    ["• Enfoque práctico y lenguaje sencillo, que nadie ofrece de forma combinada", "• Bajo nivel de educación financiera documentado por el Banco de España"],
    ["• Contacto directo con el cliente, que permite ajustar la propuesta", "• Digitalización y facturación electrónica, que obligan a ordenar los números"],
    ["• Escalabilidad posterior mediante productos digitales", "• Posible colaboración con gestorías que deriven clientes"],
  ], [4535, 4535], { numericas: false }),
  fuente("elaboración propia."),
];

// ------------------------------------------------- 5. RECURSOS Y MARCO LEGAL
const cap5 = [
  h1("5. Recursos, forma jurídica y aspectos legales"),
  h2("5.1. Recursos necesarios y capacidad"),
  p(`En sus tres primeros años Flujo sería un proyecto unipersonal. El recurso principal es el tiempo del fundador, que durante el primer año compaginaría el proyecto con otra actividad (unas 60 horas de servicio al mes) y a partir del segundo se dedicaría a jornada completa (unas 120 horas de servicio, una vez descontado el tiempo de captación, contenido y administración). Cada cliente de suscripción requiere unas 2,5 horas al mes y cada alta unas 4 horas de diagnóstico y puesta en marcha. Con estas cifras, la capacidad máxima del servicio ronda los 45 clientes activos. Este límite es la razón de ser de las otras dos capas del modelo: las plantillas, los cursos y la herramienta digital permiten generar ingresos sin consumir horas por cliente.`),
  p(`Los recursos materiales son modestos: un ordenador, que se asume ya disponible; una página web con agenda y pasarela de pago; software de gestión, facturación y diseño; y, desde el año 2, un equipo básico de grabación para los cursos. El desarrollo de la herramienta digital se encargaría a un desarrollador externo en el año 3.`),
  h2("5.2. Forma jurídica y obligaciones fiscales"),
  p(`Se ha optado por iniciar la actividad como **empresario individual (autónomo)**. Frente a una sociedad limitada, tiene menos costes de constitución y gestión, permite acogerse a la tarifa plana y encaja con un proyecto que debe validarse antes de crecer. Su inconveniente es la responsabilidad ilimitada y una tributación progresiva por IRPF que, con rendimientos elevados, supera al tipo del Impuesto sobre Sociedades, que para entidades de nueva creación es del 15 % (Agencia Tributaria, 2026). Además, algunas fuentes de financiación, como ENISA, exigen una sociedad mercantil. La transformación en sociedad limitada, cuyo capital mínimo es hoy de un euro tras la Ley 18/2022, se plantea como un paso natural una vez superado el horizonte del plan. La Tabla 5.1 resume las principales obligaciones.`),
  leyenda("Tabla", "5.1", "Obligaciones fiscales y de Seguridad Social del proyecto"),
  tabla(["Ámbito", "Obligación", "Efecto en el plan financiero"], [
    ["Seguridad Social (RETA)", "Alta en el RETA; tarifa plana de 80 €/mes los primeros 12 meses, prorrogable otros 12 si el rendimiento neto no supera el SMI (17.094 € en 2026)", `Año 1 y año 2: 80 €/mes; año 3: ${R.base.reta[2]} €/mes según tramo de rendimientos`],
    ["IAE y censo", "Alta censal (modelo 036/037) en un epígrafe de servicios profesionales", "Sin coste"],
    ["IVA", "Facturas con IVA del 21 %; declaraciones trimestrales (modelo 303) y resumen anual (modelo 390)", "El IVA se liquida el mes siguiente a cada trimestre"],
    ["IRPF", "Estimación directa simplificada. Retención del 7 % en facturas a otros autónomos durante el año de inicio y los dos siguientes; pagos fraccionados (modelo 130) del 20 %", "Gastos de difícil justificación del 5 % (máx. 2.000 €); cuota según la escala estatal y la de Aragón"],
    ["Protección de datos", "RGPD y LOPDGDD: consentimiento, política de privacidad y contrato de encargo cuando proceda", "Textos legales en la inversión inicial"],
  ], [1700, 4170, 3200], { numericas: false }),
  fuente("elaboración propia a partir del Real Decreto-ley 13/2022, la Ley 35/2006 del IRPF, la Ley 37/1992 del IVA, Infoautónomos (2026a, 2026b) y el Real Decreto 126/2026."),
  p(`Las cuotas de autónomos para 2027 no estaban aprobadas en el momento de redactar este trabajo; la tarifa plana se mantiene en 80 € y los tramos de 2026 son los mismos de 2025 (Infoautónomos, 2026a). El plan aplica, por prudencia, la tabla vigente en 2026 a los tres años.`),
  h2("5.3. Protección de datos y responsabilidad"),
  p(`Flujo trabajaría con información financiera personal de sus clientes, por lo que debe cumplir el Reglamento General de Protección de Datos y la Ley Orgánica 3/2018 desde el primer cliente. También conviene delimitar el alcance del servicio: Flujo ofrece organización y educación financiera, no asesoramiento fiscal oficial ni asesoramiento en materia de inversión, que es una actividad regulada y supervisada por la CNMV. Por ese motivo el plan incluye un seguro de responsabilidad civil profesional y la colaboración de un asesor fiscal colegiado para los casos que excedan el servicio.`),
];

// ------------------------------------------------- 6. PLAN ECONÓMICO-FINANCIERO
const fijosMes = 40 + 30 + 10 + 25 + 40 + 25 + 25;
const inv0 = V.inversion_inicial;
const cap6 = [
  h1("6. Plan económico-financiero"),
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
  p(`De acuerdo con la jerarquía financiera descrita en el capítulo 2, el arranque se financia íntegramente con 3.000 € de ahorros propios. El desarrollo de la herramienta digital, la única inversión relevante, se financiaría con un microcrédito de 10.000 € a cinco años y un tipo fijo del 7 %, en línea con las condiciones de MicroBank para emprendedores (Cámara de Comercio de Valencia, 2023), y con la caja generada por la actividad. En el escenario base el proyecto podría autofinanciar esa inversión, pero el préstamo se mantiene por prudencia: preserva un colchón de liquidez y, como se verá en el capítulo 8, sería imprescindible si la captación fuese más lenta. ENISA y los *business angels* no encajan en esta fase, porque exigen una sociedad mercantil y buscan crecimientos que un servicio unipersonal no ofrece; serían opciones a estudiar si la herramienta digital demostrara tracción.`),
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
  imagen("fig_ingresos.png", 14),
  fuente("elaboración propia."),
  p(`La Figura 6.2 muestra un resultado que condiciona todo el plan: en el escenario base la cartera de clientes alcanza el límite de capacidad hacia el final del año 3, cuando se acerca a los 45 clientes activos. A partir de ese momento el servicio no puede crecer sin contratar o sin trasladar clientes a la herramienta digital, y el modelo frena las altas del año 3 de forma automática (se registran ${r0(A.altas[2])} en lugar de 36).`),
  leyenda("Figura", "6.2", "Clientes de suscripción activos y límite de capacidad"),
  imagen("fig_clientes.png", 14),
  fuente("elaboración propia."),
  h2("6.4. Estructura de costes y punto de equilibrio"),
  p(`Los costes fijos de explotación, sin contar la cuota de autónomos, suman ${eur(fijosMes)} al mes: herramientas de gestión y facturación (40 €), diseño y redes sociales (30 €), web y alojamiento (10 €), seguro de responsabilidad civil (25 €), asesor fiscal de apoyo (40 €), teléfono e internet (25 €) y formación (25 €). A ellos se añaden el marketing, que crece de 50 a 300 € al mes, la cuota de autónomos y, desde el lanzamiento de la herramienta, 150 € mensuales de mantenimiento. Los costes variables son las comisiones de cobro y de las plataformas, y el alojamiento de la herramienta.`),
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
  h1("7. Análisis de viabilidad y rentabilidad"),
  h2("7.1. Tasa de descuento"),
  p(`La rentabilidad exigida al proyecto se ha estimado con el procedimiento descrito en el marco teórico (Tabla 7.1). Se parte de la rentabilidad del bono español a diez años, que rondaba el 3,4 % en julio de 2026 (VT Markets, 2026), y se añade una prima de mercado del 5,5 %, coherente con la prima implícita que estima Damodaran (2026) para un mercado maduro más un pequeño diferencial por riesgo país. Se toma una beta de 1, propia de un negocio de servicios profesionales sin deuda. El resultado, un 8,9 %, sería razonable para una empresa consolidada, pero no para un proyecto que todavía no tiene clientes, que no podría venderse con facilidad y que depende por completo de una persona. Por ello se añade una prima específica de 6,1 puntos, hasta una tasa del ${pct(R.tasa, 0)}. La elección es inevitablemente subjetiva, de modo que el capítulo 8 muestra cómo cambia el VAN con tasas de entre el 10 % y el 30 %.`),
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

// ------------------------------------------------- 8. RIESGOS Y ESCENARIOS
const esc = ["pesimista", "base", "optimista"];
const cap8 = [
  h1("8. Análisis de riesgos y escenarios"),
  h2("8.1. Escenarios"),
  p(`Para comprobar la solidez de los resultados se han construido dos escenarios alternativos. En el **pesimista** las altas de clientes son un 40 % inferiores a las del base, las bajas suben al 6 % mensual y las ventas de productos y usuarios de la herramienta se reducen a la mitad. En el **optimista** las altas aumentan un 30 %, las bajas bajan al 3 % y los productos y la herramienta venden un 30 % y un 50 % más, respectivamente. Los costes, la financiación y las retiradas del titular son los mismos en los tres casos, para que la comparación sea limpia.`),
  leyenda("Tabla", "8.1", "Resultados por escenario"),
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
  p(`El VAN es positivo en los tres escenarios, incluso en el pesimista, lo que confirma que el proyecto, como inversión en dinero, tiene poco riesgo. La diferencia está en los otros dos indicadores. En el escenario pesimista la tesorería se vuelve negativa ya en el primer año y llega a un mínimo de ${eur(PE.anual.caja_min)} en el mes ${PE.anual.mes_caja_min} (Figura 8.1): el fundador no podría retirar lo previsto para vivir y tendría que mantener otra fuente de ingresos o retrasar el desarrollo de la herramienta. Además, el VAN con coste de oportunidad pasa a ${eur(PE.valoracion.van_oportunidad)}: con esa captación, Flujo rendiría menos que un empleo con el salario mínimo.`),
  leyenda("Figura", "8.1", "Saldo mensual de tesorería por escenario"),
  imagen("fig_tesoreria.png", 14),
  fuente("elaboración propia."),
  p(`En el escenario optimista, por el contrario, el límite de capacidad se alcanza ya en el año 2, lo que confirma que el crecimiento del servicio está acotado por las horas del fundador y no por la demanda.`),
  h2("8.2. Análisis de sensibilidad"),
  p(`La Figura 8.2 muestra cómo varía el VAN del escenario base al modificar, de una en una, las dos variables más inciertas. Una rebaja del precio de la suscripción a 55 €, cercana al de una gestoría *online*, reduciría el VAN a ${eur(R.sensibilidad.precio["55"])}, y una captación a la mitad de la prevista lo dejaría en ${eur(R.sensibilidad.altas["0.5"])}. El VAN se mantiene positivo en todos los casos. Respecto a la tasa de descuento, pasar del 10 % al 30 % lo reduce de ${eur(R.sensibilidad.tasa["0.1"])} a ${eur(R.sensibilidad.tasa["0.3"])}, una variación pequeña porque el horizonte es corto y la inversión, mínima. La captación de clientes es, por tanto, la variable crítica.`),
  leyenda("Figura", "8.2", "Sensibilidad del VAN al precio y a la captación de clientes"),
  imagen("fig_sensibilidad.png", 14.5),
  fuente("elaboración propia."),
  h2("8.3. Principales riesgos y medidas de mitigación"),
  leyenda("Tabla", "8.2", "Matriz de riesgos del proyecto"),
  tabla(["Riesgo", "Probabilidad", "Impacto", "Medida de mitigación"], [
    ["Captación más lenta de lo previsto", "Media", "Alto", "Validar antes de invertir; mantener ingresos alternativos el primer año"],
    ["Precio percibido como alto frente a gestorías", "Media", "Medio", "Posicionarse como complemento; acuerdos con gestorías; plan reducido"],
    ["Dependencia del fundador (enfermedad, saturación)", "Media", "Alto", "Procesos y plantillas estandarizados; trasladar clientes a la herramienta"],
    ["Fracaso o retraso de la herramienta digital", "Media", "Medio", "Desarrollarla solo con tracción demostrada; MVP sencillo; préstamo acotado"],
    ["Cambios en la cotización de autónomos o en la fiscalidad", "Media", "Bajo", "Seguimiento normativo; el plan ya usa las cuotas vigentes por prudencia"],
    ["Incidente de protección de datos", "Baja", "Alto", "Cumplimiento del RGPD, herramientas con cifrado y seguro de responsabilidad civil"],
    ["Reducción del segmento joven del RETA", "Alta", "Bajo", "Ampliar el público a autónomos noveles de cualquier edad"],
  ], [3000, 1300, 1100, 3670], { numericas: false }),
  fuente("elaboración propia."),
];

// ------------------------------------------------- 9. CONCLUSIONES
const cap9 = [
  h1("9. Conclusiones"),
  h2("9.1. Conclusiones generales y específicas"),
  p(`Este trabajo se planteaba si sería viable un proyecto dedicado a mejorar la gestión financiera de los autónomos jóvenes. La respuesta, con las hipótesis utilizadas, es afirmativa, aunque con matices importantes.`),
  p(`En primer lugar, **la necesidad existe y está documentada**. Los datos del Banco de España muestran un nivel bajo de conocimientos financieros en la población adulta, algo peor entre los jóvenes, y las quince entrevistas realizadas apuntan a problemas muy concretos: no saber cuánto se gana, mezclar las finanzas personales y las de la actividad y no reservar para los impuestos. El segmento de autónomos menores de 35 años, de unas 560.000 personas, es más que suficiente para un proyecto de este tamaño, aunque pierde peso dentro del colectivo.`),
  p(`En segundo lugar, **el proyecto es económica y financieramente viable**. La inversión inicial es de solo 3.000 €, el punto de equilibrio del primer año se sitúa en unos ${Math.ceil(B.pe[0])} clientes de suscripción y, en el escenario base, la actividad genera un beneficio acumulado de ${eur(sum(A.resultado_neto))} en tres años. El VAN es positivo en los tres escenarios analizados y el balance y la tesorería son coherentes con la cuenta de resultados.`),
  p(`En tercer lugar, **la viabilidad es menos holgada de lo que sugieren la TIR y el payback**. Cuando se incluye el coste de oportunidad del fundador, el VAN del escenario base se reduce a ${eur(V.van_oportunidad)} y pasa a ser negativo en el pesimista. Además, el servicio alcanza el límite de capacidad de una persona al final del tercer año. El éxito del proyecto depende, por tanto, de dos factores: captar clientes al ritmo previsto y conseguir que los productos digitales y la herramienta crezcan sin consumir horas por cliente.`),
  p(`Estas conclusiones responden a los objetivos planteados en la introducción: se ha descrito el colectivo y sus dificultades (capítulos 3 y 4), se ha definido el modelo de negocio (capítulo 3), se han analizado el mercado, la competencia y el marco legal (capítulos 4 y 5), se han elaborado los estados financieros previsionales a tres años (capítulo 6), se han calculado los criterios de rentabilidad y los ratios (capítulo 7) y se han estudiado los riesgos y escenarios (capítulo 8).`),
  h2("9.2. Trascendencia de los resultados"),
  p(`Para un emprendedor, el principal valor del trabajo es que convierte una idea en un conjunto de hipótesis medibles: cuántos clientes hacen falta, cuánto tiempo consume cada uno y qué ocurre si el ritmo de captación es menor. El modelo, publicado en abierto, puede actualizarse con los datos reales de la validación. Desde un punto de vista más general, el trabajo muestra que la falta de educación financiera de los autónomos jóvenes es también una oportunidad de negocio, y que un servicio de bajo coste puede contribuir a un objetivo que las instituciones consideran prioritario.`),
  h2("9.3. Limitaciones y líneas futuras"),
  p(`El trabajo tiene varias limitaciones. Las entrevistas son pocas y no representativas, y no miden la disposición a pagar, que solo puede comprobarse vendiendo el servicio. Los datos de competencias financieras proceden de una encuesta de 2021-2022 y no ofrecen información específica sobre autónomos jóvenes. Las cuotas de autónomos de 2027 a 2029 no están aprobadas, y la tributación por IRPF depende de la situación personal del titular, que se ha simplificado. Por último, la tasa de descuento incluye una prima específica que es, inevitablemente, una estimación.`),
  p(`Como líneas futuras se proponen ejecutar el plan de validación de la Tabla 3.2 y actualizar el modelo con los datos reales, explorar acuerdos con gestorías y estudiar la transformación en sociedad limitada si la herramienta digital alcanza suficientes usuarios, lo que abriría la puerta a la financiación de ENISA o de inversores privados.`),
];

// ------------------------------------------------- BIBLIOGRAFÍA
const ref = (t) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED, indent: { left: 567, hanging: 567 }, spacing: { after: 120 }, children: require("./helpers").runs(t),
});
const consulta = "[Consulta: 28/10/2026]";
const biblio = [
  h1("Bibliografía"),
  ...[
    `AGENCIA TRIBUTARIA (2026): *Tipos de gravamen del Impuesto sobre Sociedades* [en línea]. Madrid. https://sede.agenciatributaria.gob.es/Sede/impuesto-sobre-sociedades/campana-sociedades/cuestiones-tener-cuenta-liquidacion-impuesto-2025/tipos-gravamen.html ${consulta}.`,
    `BANCO DE ESPAÑA (2022): *Encuesta de Competencias Financieras (ECF) 2021*. Dirección General de Economía y Estadística, Madrid. https://www.bde.es/f/webbe/SES/AnalisisEconomico/Competencis_Financieras/EncuestaCompetencias_2021.pdf ${consulta}.`,
    `BANCO DE ESPAÑA (2025): *Análisis de la Encuesta de Competencias Financieras 2021 en apoyo del Plan de Educación Financiera 2026-2029* [en línea]. Madrid. https://www.bde.es/f/webbe/INF/MenuHorizontal/AreasActuacion/educacion_financiera/analisis-ecf-2021-apoyo-pef-2026-2029.pdf ${consulta}.`,
    `BERGER, A. N. y UDELL, G. F. (1998): "The economics of small business finance: The roles of private equity and debt markets in the financial growth cycle". *Journal of Banking & Finance*, Vol. 22, Nº 6-8, pp. 613-673.`,
    `BILLEO (2026): *Nota de datos: precios de gestorías online y software 2026* [en línea]. https://www.billeo.es/data/nota-datos-precios-gestorias-2026.md ${consulta}.`,
    `BLANK, S. (2013): "Why the lean start-up changes everything". *Harvard Business Review*, Vol. 91, Nº 5, pp. 63-72.`,
    `BREALEY, R. A., MYERS, S. C. y ALLEN, F. (2020): *Principles of Corporate Finance* (13ª ed.). McGraw-Hill Education, Nueva York.`,
    `CÁMARA DE COMERCIO DE VALENCIA (2023): *Microcréditos para el emprendimiento MicroBank 2023* [en línea]. Valencia. https://www.camaravalencia.com/wp-content/uploads/2023/05/EMP3_Microcreditos-Emprendimiento-Microbank-2023.pdf ${consulta}.`,
    `CASTILLA Y LEÓN ECONÓMICA (2026): *6 de cada 10 autónomos llevan más de 5 años en su negocio* [en línea]. 9 de febrero de 2026. https://www.castillayleoneconomica.es/6-10-autonomos-llevan-5-anos-frente-negocio/ ${consulta}.`,
    `DAMODARAN, A. (2009): *Valuing Young, Start-up and Growth Companies: Estimation Issues and Valuation Challenges*. Working paper, Stern School of Business, New York University.`,
    `DAMODARAN, A. (2012): *Investment Valuation: Tools and Techniques for Determining the Value of Any Asset* (3ª ed.). John Wiley & Sons, Hoboken.`,
    `DAMODARAN, A. (2026): *Country Default Spreads and Risk Premiums* [en línea]. Actualización de enero de 2026. https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html ${consulta}.`,
    `EMIRODGAR (2026): *Análisis y evolución de los autónomos en España* [en línea]. https://emirodgar.es/analisis-autonomos ${consulta}.`,
    `ENISA (2026): *Línea Jóvenes Emprendedores* [en línea]. Madrid. https://www.enisa.es ${consulta}.`,
    `FEDERACIÓN NACIONAL DE ASOCIACIONES DE TRABAJADORES AUTÓNOMOS, ATA (2026): "El RETA ganó 39.000 autónomos en 2025, pero ATA prevé peores perspectivas para 2026". *Autónomos y Emprendedor*, 5 de enero de 2026. https://www.autonomosyemprendedor.es/articulo/autonomos/reta-gano-39000-autonomos-2025-menor-crecimiento-peores-perspectivas-2026/20260105114615047592.html ${consulta}.`,
    `GOBIERNO DE ARAGÓN (2026): *Programa Emprendedores Autónomos* [en línea]. Zaragoza. https://www.aragon.es/tramitador/-/tramite/programa-emprendedores-autonomos ${consulta}.`,
    `INFOAUTÓNOMOS (2026a): *La tarifa plana de 80 euros para autónomos en 2026: requisitos, cuantía y renovación* [en línea]. https://www.infoautonomos.com/seguridad-social/tarifa-plana-autonomos/ ${consulta}.`,
    `INFOAUTÓNOMOS (2026b): *Las retenciones del IRPF de los profesionales autónomos* [en línea]. https://www.infoautonomos.com/fiscalidad/retenciones-irpf-profesionales-autonomos/ ${consulta}.`,
    `LEY 35/2006, de 28 de noviembre, del Impuesto sobre la Renta de las Personas Físicas. *Boletín Oficial del Estado*, Nº 285, de 29 de noviembre de 2006.`,
    `LEY 37/1992, de 28 de diciembre, del Impuesto sobre el Valor Añadido. *Boletín Oficial del Estado*, Nº 312, de 29 de diciembre de 1992.`,
    `LEY 18/2022, de 28 de septiembre, de creación y crecimiento de empresas. *Boletín Oficial del Estado*, Nº 234, de 29 de septiembre de 2022.`,
    `LEY ORGÁNICA 3/2018, de 5 de diciembre, de Protección de Datos Personales y garantía de los derechos digitales. *Boletín Oficial del Estado*, Nº 294, de 6 de diciembre de 2018.`,
    `LUSARDI, A. y MITCHELL, O. S. (2014): "The economic importance of financial literacy: Theory and evidence". *Journal of Economic Literature*, Vol. 52, Nº 1, pp. 5-44.`,
    `MYERS, S. C. y MAJLUF, N. S. (1984): "Corporate financing and investment decisions when firms have information that investors do not have". *Journal of Financial Economics*, Vol. 13, Nº 2, pp. 187-221.`,
    `OBSERVATORIO DEL EMPRENDIMIENTO DE ESPAÑA (2025): *Global Entrepreneurship Monitor. Informe GEM España 2024-2025*. Editorial de la Universidad de Cantabria, Santander. https://www.editorial.unican.es/libro/global-entrepreneurship-monitor-informe-gem-espana-2024-2025 ${consulta}.`,
    `OBSERVATORIO ESPAÑOL DE LA ECONOMÍA SOCIAL Y DEL TRABAJO AUTÓNOMO (2025): *Aumenta considerablemente el número y porcentaje de trabajadores autónomos con más de 50 años de edad* [en línea]. https://observatorioeconomiasocial.es/aumenta-considerablemente-el-numero-y-porcentaje-de-trabajadores-autonomos-con-mas-de-50-anos-de-edad/ ${consulta}.`,
    `OCDE (2020): *OECD/INFE 2020 International Survey of Adult Financial Literacy*. OECD Publishing, París. https://www.oecd.org/en/publications/2020/06/oecd-infe-2020-international-survey-of-adult-financial-literacy_bbad9b27.html ${consulta}.`,
    `OSTERWALDER, A. y PIGNEUR, Y. (2010): *Business Model Generation*. John Wiley & Sons, Hoboken.`,
    `REAL DECRETO 126/2026, de 17 de febrero, por el que se fija el salario mínimo interprofesional para 2026. *Boletín Oficial del Estado*, de 19 de febrero de 2026.`,
    `REAL DECRETO-LEY 13/2022, de 26 de julio, por el que se establece un nuevo sistema de cotización para los trabajadores por cuenta propia o autónomos. *Boletín Oficial del Estado*, Nº 179, de 27 de julio de 2022.`,
    `REGLAMENTO (UE) 2016/679 del Parlamento Europeo y del Consejo, de 27 de abril de 2016, relativo a la protección de las personas físicas en lo que respecta al tratamiento de datos personales (Reglamento General de Protección de Datos). *Diario Oficial de la Unión Europea*, L 119, de 4 de mayo de 2016.`,
    `RIES, E. (2011): *The Lean Startup*. Crown Business, Nueva York.`,
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
  leyenda("Figura", "I.3", "Reparto del crecimiento neto del RETA en 2025 por edad"),
  imagen("fig_anexo_crecimiento.png", 13),
  fuente("elaboración propia a partir de Emirodgar (2026)."),

  h1("Anexo II. Modelo de cálculo y reproducibilidad"),
  p(`El plan económico-financiero se ha calculado con un modelo mensual de 36 periodos programado en Python. El modelo parte de un único fichero de hipótesis y genera, para cada escenario, la cuenta de resultados, la tesorería con el calendario real del IVA y del IRPF, el balance, los criterios de rentabilidad y el análisis de sensibilidad. Como comprobación de coherencia, el programa verifica que el activo y la suma del patrimonio neto y el pasivo coinciden en cada cierre. Los resultados se vuelcan también en una hoja de cálculo (Modelo_financiero_Flujo.xlsx) con el detalle mes a mes.`),
  p(`El código, los datos de mercado con la fuente de cada cifra y los gráficos de este trabajo están disponibles en el repositorio [[https://github.com/USUARIO/flujo-tfg]]. Para reproducir los resultados basta con ejecutar, en este orden, modelo.py, graficos.py y exportar_excel.py, y cualquier hipótesis puede cambiarse en hipotesis.py para recalcular todo el plan.`),
  p(`Todos los porcentajes del Anexo I proceden de fuentes secundarias ya citadas en la bibliografía. La aportación propia consiste en reunir, cruzar y representar esos datos, y en construir el modelo financiero; los únicos datos primarios del trabajo son las entrevistas.`),

  h1("Anexo III. Guion de las entrevistas exploratorias"),
  p(`[[PENDIENTE DE COMPLETAR POR EL AUTOR: fechas en que se hicieron las entrevistas, medio (presencial, videollamada) y tabla anónima con el perfil de los quince entrevistados (edad, actividad, años como autónomo). A continuación figura una propuesta de guion que debe ajustarse a las preguntas que se plantearon realmente.]]`),
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
];

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
      children: [...indice, ...cap1, ...cap2, ...cap3, ...cap4, ...cap5, ...cap6, ...cap7, ...cap8, ...cap9, ...biblio, ...anexos] },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(__dirname, "..", "TFG_Flujo_v2.docx");
  fs.writeFileSync(out, buf);
  console.log("guardado", out);
});
