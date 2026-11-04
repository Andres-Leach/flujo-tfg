// Genera TFG_Flujo_v2.docx con el formato de las directrices FECEM (junio de 2023)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, Footer, PageNumber,
  TableOfContents, LevelFormat, ImageRun, SectionType, NumberFormat, StyleLevel,
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
const resumenEs = `Los autónomos jóvenes gestionan ingresos irregulares e impuestos trimestrales casi siempre sin formación financiera: solo el 19 % de los adultos acierta las tres preguntas financieras básicas del Banco de España. Este trabajo evalúa la viabilidad económico-financiera de Flujo, un proyecto de organización y educación financiera para autónomos de 20 a 35 años. Combina fuentes oficiales, una encuesta propia y un plan financiero a tres años con escenarios. En el escenario base los ingresos pasan de ${eur(A.ingresos[0])} a ${eur(A.ingresos[2])}, el punto de equilibrio es de ${Math.ceil(B.pe[0])} clientes y el VAN, de ${eur(V.van)}; descontado el coste de oportunidad del fundador, baja a ${eur(V.van_oportunidad)}. Cinco medidas de bajo coste lo elevan a ${eur(R.palancas[R.palancas.length - 1].van_oportunidad)} y sitúan en el ${pct(R.montecarlo_opt.p_van_op, 0)} la probabilidad de superar al empleo alternativo. El proyecto presenta una viabilidad potencial, condicionada a la captación de clientes y al precio, y la disposición a pagar debe validarse.`;
const resumenEn = `Young self-employed workers manage irregular income and quarterly taxes, usually without financial training: only 19% of adults answer the Bank of Spain's three basic financial questions correctly. This study assesses the economic and financial feasibility of Flujo, a financial organisation and education project for self-employed people aged 20 to 35. It combines official sources, an original survey and a three-year financial plan with scenarios. In the base scenario revenue grows from ${eur(A.ingresos[0])} to ${eur(A.ingresos[2])}, break-even is reached with ${Math.ceil(B.pe[0])} clients and the NPV is ${eur(V.van)}; after deducting the founder's opportunity cost it falls to ${eur(V.van_oportunidad)}. Five low-cost measures raise it to ${eur(R.palancas[R.palancas.length - 1].van_oportunidad)} and put at ${pct(R.montecarlo_opt.p_van_op, 0)} the probability of beating the alternative job. The project shows potential feasibility, conditional on client acquisition and pricing, and willingness to pay must be validated.`;
const dato = (et, val) => new Paragraph({ spacing: { before: 0, after: 120 }, children: [new TextRun({ text: et + ": ", bold: true }), new TextRun(val)] });
const info = [
  dato("Autor del trabajo", "Andrés Leach Sagarra"),
  dato("Directora del trabajo", "Dra. Beatrice Boumda Ijang"),
  dato("Título del trabajo", "Viabilidad financiera de Flujo: una start-up de ayuda a jóvenes emprendedores"),
  dato("Title", "Financial feasibility of Flujo: a start-up supporting young entrepreneurs"),
  dato("Titulación", "Grado en Administración y Dirección de Empresas (Universidad de Zaragoza)"),
  new Paragraph({ spacing: { before: 0, after: 120 }, children: [new TextRun({ text: "Resumen", bold: true })] }),
  p(resumenEs),
  new Paragraph({ spacing: { before: 0, after: 120 }, children: [new TextRun({ text: "Abstract", bold: true })] }),
  p(resumenEn),
];

// ------------------------------------------------------------------ ÍNDICE
const indice = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 120 }, children: [new TextRun({ text: "Índice", bold: true, size: 28 })] }),
  new TableOfContents("Índice", { hyperlink: true, headingStyleRange: "1-2" }),
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 120 }, children: [new TextRun({ text: "Índice de tablas", bold: true, size: 28 })] }),
  new TableOfContents("Índice de tablas", { hyperlink: true, stylesWithLevels: [new StyleLevel("LeyendaTabla", 1)] }),
  new Paragraph({ keepNext: true, spacing: { before: 240, after: 120 }, children: [new TextRun({ text: "Índice de figuras", bold: true, size: 28 })] }),
  new TableOfContents("Índice de figuras", { hyperlink: true, stylesWithLevels: [new StyleLevel("LeyendaFigura", 1)] }),
];

// ---------------------------------------------------------- 1. INTRODUCCIÓN
const cap1 = [
  h1("1. Introducción", false),
  h2("1.1. Presentación y justificación del tema"),
  p(`El trabajo por cuenta propia tiene en España un peso económico y social considerable. El Régimen Especial de Trabajadores Autónomos (RETA) cerró 2025 con 3.425.767 personas afiliadas, la cifra más alta de su serie, tras sumar unas 39.000 altas netas en el año (ATA, 2026), y en septiembre de 2026 la afiliación media alcanzaba ya 3.483.645 personas, un 2,06 % más que un año antes (Ministerio de Inclusión, Seguridad Social y Migraciones, 2026). Detrás de esas cifras hay muchas actividades distintas, pero casi todas comparten un mismo problema de fondo: la persona que factura es también la que tiene que llevar las cuentas, calcular el IVA y el IRPF de cada trimestre y decidir cuánto puede gastar sin poner en riesgo su actividad.`),
  p(`Esa tarea exige unos conocimientos que la mayoría de la población no tiene. Según la Encuesta de Competencias Financieras (ECF) del Banco de España, con datos recogidos entre 2021 y 2022, solo un 41 % de los adultos responde bien a una pregunta sencilla sobre interés compuesto y apenas un 19 % acierta a la vez las preguntas sobre inflación, interés compuesto y diversificación del riesgo (Banco de España, 2022). El problema está reconocido institucionalmente: el propio Banco de España ha reanalizado la encuesta como base del Plan de Educación Financiera 2026-2029 (Banco de España, 2025). Entre los ocupados, tanto por cuenta propia como ajena, entre un 38 % y un 42 % declara tener conocimientos financieros bajos o muy bajos (Banco de España, 2022), y en el plano internacional solo un 34 % de los adultos de los 39 países de la última encuesta de la OCDE alcanza la puntuación mínima de alfabetización financiera (OCDE, 2023).`),
  p(`En el caso concreto de quienes emprenden, el Observatorio del Emprendimiento de España ha analizado las respuestas de 128.598 personas encuestadas por el GEM entre 2021 y 2024 y concluye que el nivel de conocimiento financiero en España es «moderadamente bajo»: en una escala de 1 a 5, los emprendedores potenciales puntúan 2,76 y los que han emprendido recientemente 2,99 (San-Martín *et al.*, 2025). Al mismo tiempo, ocho de cada diez emprendedores consideran tener los conocimientos y habilidades necesarios para emprender (Observatorio del Emprendimiento de España, 2025), una autopercepción que encaja con lo que la literatura llama exceso de confianza financiera (Lusardi y Mitchell, 2014).`),
  p(`Este trabajo parte de esa distancia entre lo que se sabe y lo que se cree saber. En el día a día de un autónomo, esa distancia se traduce en problemas muy concretos: saber cuánto se gana realmente una vez descontados impuestos y gastos, separar el dinero de la actividad del personal y reservar cada trimestre lo que corresponde a Hacienda. Para comprobar hasta qué punto se dan en el colectivo al que se dirige el proyecto, he realizado una encuesta propia a autónomos y *freelancers* de 20 a 35 años, cuyos resultados se recogen en el capítulo 3 y en el Anexo III.`),
  p(`Elegí este tema porque siempre me ha llamado la atención la idea de crear un negocio propio, y quería aprovechar el TFG para desarrollar una idea desde cero y comprobar si realmente podría funcionar. Me interesaba, además, un problema que afecta a muchos autónomos: tener ingresos, gastos e impuestos que gestionar sin saber siempre con claridad cuánto dinero están ganando realmente ni cómo organizarse.`),
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
  p(`Para mí, el trabajo era también una oportunidad de aplicar a un proyecto concreto conocimientos de distintas áreas del grado, como las finanzas, el marketing o la estrategia empresarial. Más allá de eso, su interés es triple. En el plano **empresarial**, ofrece un análisis de viabilidad que podría servir de base real para lanzar el proyecto, con hipótesis explícitas que se pueden contrastar a medida que lleguen los primeros clientes. En el plano **social**, aborda un problema, la baja educación financiera, que las instituciones consideran prioritario y que en el caso de los autónomos tiene consecuencias directas: recargos por errores fiscales, falta de colchón de liquidez o decisiones de gasto poco meditadas. En el plano **universitario**, integra herramientas de contabilidad, fiscalidad, finanzas y estrategia en un mismo caso, y el modelo de cálculo se publica en abierto para que cualquier lector pueda reproducirlo.`),
  h2("1.4. Metodología y fuentes"),
  p(`El trabajo sigue un enfoque de estudio de caso con tres tipos de fuentes. Las **fuentes secundarias** proceden de organismos oficiales y entidades de referencia: Banco de España, Seguridad Social, Ministerio de Trabajo y Economía Social, Agencia Tributaria, Instituto Nacional de Estadística, Eurostat, OCDE, Observatorio del Emprendimiento de España y la Federación Nacional de Asociaciones de Trabajadores Autónomos (ATA). El Anexo IV resume qué dato se ha contrastado con cada fuente oficial. La **fuente primaria** es una encuesta propia en línea, anónima, a autónomos y *freelancers* de 20 a 35 años, que pregunta por su gestión financiera y por su disposición a pagar, medida con el método de sensibilidad al precio de Van Westendorp (1976); el cuestionario y los resultados se recogen en el Anexo III. Por último, el análisis se apoya en **herramientas de gestión** consolidadas: la metodología Lean Startup (Ries, 2011), el Business Model Canvas (Osterwalder y Pigneur, 2010), el análisis DAFO y los criterios clásicos de evaluación de inversiones (Brealey, Myers y Allen, 2020).`),
  p(`Las proyecciones financieras se han calculado mes a mes con un modelo propio programado en Python, que se entrega también en una hoja de cálculo. Todas las hipótesis se explican en el capítulo 6 y el código y los datos están disponibles en un repositorio público (ver Anexo II).`),
  h2("1.5. Estructura del trabajo"),
  p(`Tras esta introducción, el capítulo 2 recoge el marco teórico: el concepto de start-up, el modelo de negocio, la educación financiera y los criterios de viabilidad y financiación. El capítulo 3 presenta la necesidad detectada y la propuesta de Flujo, y el capítulo 4 analiza el mercado, la competencia y el DAFO. El capítulo 5 trata los recursos necesarios y el marco legal y fiscal. Los capítulos 6 a 9 forman el núcleo del trabajo: el plan económico-financiero, el análisis de viabilidad y rentabilidad, las medidas para mejorar esa rentabilidad y el análisis de riesgos y escenarios. El capítulo 10 recoge las conclusiones y las limitaciones. Los anexos incluyen el análisis de datos de mercado, la nota sobre el modelo de cálculo, la encuesta y la relación de fuentes oficiales consultadas.`),
];

// ---------------------------------------------------------- 2. MARCO TEÓRICO
const cap2 = [
  h1("2. Marco teórico", false),
  h2("2.1. La start-up y el enfoque Lean Startup"),
  p(`Ries (2011) define una start-up como una institución humana diseñada para crear un producto o servicio en condiciones de incertidumbre extrema. Lo que la distingue de una pequeña empresa tradicional no es su tamaño sino que todavía no sabe con certeza quién es su cliente ni cuánto está dispuesto a pagar. Blank (2013) lo resume diciendo que una start-up busca un modelo de negocio repetible y escalable, mientras que una empresa consolidada se limita a ejecutarlo. Shane y Venkataraman (2000) sitúan el origen de cualquier iniciativa emprendedora en el descubrimiento y la evaluación de una oportunidad, es decir, en detectar una necesidad que el mercado no cubre y decidir si merece la pena explotarla, que es exactamente la pregunta que se plantea este trabajo.`),
  p(`De esa incertidumbre nace el enfoque Lean Startup, que propone sustituir los planes de negocio cerrados por ciclos cortos de construir, medir y aprender. La idea central es lanzar un producto mínimo viable, observar cómo responden los clientes y ajustar la propuesta antes de comprometer recursos importantes. Para un proyecto como Flujo, este enfoque tiene una consecuencia práctica: empezar como un servicio prestado de forma manual, que apenas requiere inversión, y dejar el desarrollo tecnológico para cuando exista evidencia de que hay clientes dispuestos a pagar.`),
  h2("2.2. El modelo de negocio como herramienta de análisis"),
  p(`Osterwalder y Pigneur (2010) describen el modelo de negocio como la lógica con la que una organización crea, entrega y captura valor, y proponen representarlo en un lienzo de nueve bloques: segmentos de clientes, propuesta de valor, canales, relación con clientes, fuentes de ingresos, recursos clave, actividades clave, socios clave y estructura de costes. Su utilidad para este trabajo es doble. Por un lado, obliga a explicitar todas las hipótesis del proyecto en un formato que se puede revisar de un vistazo. Por otro, conecta directamente con el plan financiero, porque los bloques de ingresos y costes son el punto de partida de las proyecciones. En el caso de Flujo, el bloque más incierto es el de las fuentes de ingresos: los costes son bajos y conocidos, pero el precio que el cliente aceptará y el tiempo que seguirá suscrito no se pueden conocer sin venderle el servicio. Por eso el plan los trata como hipótesis de respaldo bajo (Tabla 6.1) y la encuesta se centra en ellos.`),
  h2("2.3. Educación financiera y trabajo autónomo"),
  p(`La OCDE define la educación financiera como la combinación de conocimientos, habilidades, actitudes y comportamientos necesarios para tomar decisiones financieras adecuadas (OCDE, 2020). Lusardi y Mitchell (2014) revisan la evidencia internacional y concluyen que los niveles de conocimiento financiero son bajos incluso en países desarrollados, que los jóvenes y las personas con menos estudios puntúan peor y que un mayor conocimiento se asocia con más ahorro, mejor planificación y menos errores costosos. También señalan que mucha gente sobrestima lo que sabe, lo que reduce su propensión a buscar ayuda. Para Flujo esta idea tiene una doble lectura: confirma que el problema existe, pero anticipa también una dificultad comercial, porque quien cree que no necesita ayuda difícilmente paga por ella. Es una de las razones por las que la demanda del servicio debe validarse en lugar de darse por supuesta.`),
  p(`En el caso de los autónomos el problema se agrava por tres razones. Sus ingresos son irregulares, lo que exige planificar la liquidez. No tienen un empleador que les retenga y cotice por ellos, de modo que la reserva para impuestos y Seguridad Social depende de su propia disciplina. Y la frontera entre el patrimonio personal y el de la actividad es difusa, porque responden con todos sus bienes. Para un asalariado un error financiero suele tener un coste limitado; para un autónomo puede traducirse en recargos, en tensiones de tesorería o en el cierre de la actividad.`),
  p(`La evidencia española confirma que el tamaño importa. La Encuesta de Competencias Financieras en las Pequeñas Empresas, también del Banco de España, muestra que los propietarios de empresas de menos de 20 trabajadores tienen menos conocimientos financieros que los de empresas de 20 a 49, que prácticas como comparar proveedores financieros están menos extendidas en las de menos de cinco trabajadores y que pocos propietarios han pensado en cómo financiar su jubilación. También encuentra que las empresas cuyos propietarios tenían más competencias financieras sufrieron un impacto algo menor en liquidez durante la pandemia (Banco de España, 2021). Es decir, la educación financiera no es solo una cuestión de cultura general: tiene efectos medibles sobre la resistencia de los negocios más pequeños.`),
  p(`La evidencia también indica qué tipo de formación funciona. Drexler, Fischer y Schoar (2014) compararon, en dos experimentos con microempresarios de la República Dominicana, un curso de contabilidad convencional con una formación basada en reglas prácticas sencillas. La formación sencilla mejoró de forma significativa las prácticas financieras, la calidad de la información y los ingresos de los negocios, y su efecto fue mayor entre quienes partían con menos formación. Es un resultado relevante para Flujo, cuya propuesta consiste precisamente en sustituir los tecnicismos por reglas y sistemas fáciles de aplicar.`),
  h2("2.4. La viabilidad económico-financiera"),
  p(`Un proyecto es **económicamente viable** cuando genera ingresos suficientes para cubrir todos sus costes y remunerar los recursos invertidos, y es **financieramente viable** cuando dispone en cada momento de la liquidez necesaria para atender sus pagos. Las dos condiciones son necesarias: un proyecto rentable puede fracasar por falta de caja, y uno con liquidez puede estar destruyendo valor. Por eso este trabajo analiza tanto la cuenta de resultados como los flujos de tesorería y el balance.`),
  p(`Para valorar la rentabilidad se emplean los criterios clásicos de evaluación de inversiones. El **valor actual neto (VAN)** descuenta los flujos de caja futuros a una tasa que refleja el riesgo del proyecto y les resta la inversión inicial; un VAN positivo indica que el proyecto crea valor por encima de esa rentabilidad exigida. La **tasa interna de rentabilidad (TIR)** es la tasa que iguala el VAN a cero. El **periodo de recuperación (payback)** mide el tiempo necesario para recuperar la inversión. Brealey, Myers y Allen (2020) recuerdan que el VAN es el criterio más sólido y que la TIR y el payback deben usarse como complementos, porque la TIR puede dar resultados engañosos cuando la inversión inicial es muy pequeña y el payback ignora lo que ocurre después de la recuperación. Esta advertencia resulta importante para Flujo, como se verá en el capítulo 7. Hay además un recurso que estos criterios no recogen y que en un negocio unipersonal es el principal: el tiempo del fundador. Si su trabajo no se valora, un proyecto puede parecer muy rentable solo porque quien lo dirige trabaja sin cobrar, y por eso el análisis calcula también el VAN después de descontar un sueldo alternativo.`),
  p(`La tasa de descuento es el elemento más delicado. Siguiendo a Damodaran (2012), puede estimarse con el modelo CAPM, sumando a la rentabilidad sin riesgo una prima de mercado ajustada por la beta del sector. En empresas jóvenes y no cotizadas esa tasa se queda corta, porque no recoge el riesgo de fracaso, la falta de liquidez de la inversión ni la dependencia de una sola persona, y Damodaran (2009) propone añadir primas específicas para esos riesgos. El capítulo 7 aplica este procedimiento a Flujo.`),
  p(`Cuando la incertidumbre es alta, dos herramientas complementan estos criterios. La simulación de Montecarlo, propuesta para la evaluación de inversiones por Hertz (1964), sustituye las estimaciones puntuales por distribuciones de probabilidad y permite estimar con qué probabilidad el VAN será positivo. La teoría de las opciones reales (Dixit y Pindyck, 1994) recuerda que poder esperar, ampliar o abandonar una inversión tiene valor en sí mismo, algo útil para decidir cuándo desarrollar la herramienta digital. En los negocios por suscripción, además, es habitual valorar a los clientes como activos a partir de los márgenes que generarán durante su relación con la empresa (Gupta y Lehmann, 2003). Finalmente, los **ratios financieros** de rentabilidad, liquidez y endeudamiento completan el análisis, aunque en una microempresa de servicios con muy pocos activos deben interpretarse con cautela.`),
  h2("2.5. La financiación de la empresa en fase inicial"),
  p(`La teoría de la jerarquía financiera (*pecking order*) de Myers y Majluf (1984) sostiene que las empresas prefieren financiarse primero con recursos propios, después con deuda y solo en último lugar con nuevos socios, porque la información asimétrica encarece la financiación externa. Berger y Udell (1998) completan esta idea con el concepto de ciclo de crecimiento financiero: en las primeras fases la empresa depende de los ahorros del fundador y de su entorno cercano; a medida que crece y genera historial accede a microcréditos, préstamos bancarios y financiación pública, y solo cuando tiene un potencial de crecimiento elevado atrae a *business angels* y fondos de capital riesgo. La evidencia empírica matiza este esquema: Robb y Robinson (2014), con datos de miles de empresas nuevas estadounidenses, muestran que recurren a la deuda bancaria mucho más de lo que se suponía, y Cassar (2004) encuentra que la financiación de las start-ups depende sobre todo de su tamaño, de la proporción de activos tangibles y de su orientación al crecimiento. Para un servicio casi sin activos como Flujo, que no puede ofrecer garantías, esto refuerza el papel de los recursos propios y del microcrédito. La Tabla 2.1 resume las fuentes disponibles en España para un proyecto como Flujo.`),
  leyenda("Tabla", "2.1", "Fuentes de financiación según la fase del proyecto"),
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
  p(`El marco teórico y los datos oficiales permiten concretar tres problemas habituales en quien trabaja por cuenta propia. El primero, no saber cuánto se gana realmente una vez descontados el IVA, el IRPF y la cuota de autónomos. El segundo, la mezcla entre las finanzas personales y las de la actividad, que impide saber qué parte de lo que hay en la cuenta puede gastarse. El tercero, la incertidumbre ante cada cierre trimestral cuando no se reserva de forma sistemática el dinero de los impuestos.`),
  p(`La encuesta propia contrasta estas hipótesis en tres niveles: el problema (si existen dificultades reales de gestión), el interés (si el servicio se percibe útil) y la demanda (si se pagaría su precio), con umbrales fijados antes de conocer los resultados (Anexo III, Tabla III.2). De las [[N]] personas de 20 a 35 años que respondieron, el [[X]] % no sabría decir con exactitud cuánto ganó el mes anterior después de impuestos y gastos, el [[X]] % no separa, o solo a medias, el dinero personal del de la actividad y el [[X]] % no aparta cada mes lo que tendrá que pagar de IVA e IRPF (ver Anexo III). [[Completar cuando estén los resultados.]]`),
  p(`Los datos oficiales apuntan en la misma dirección: en la ECF, los jóvenes de 18 a 34 años puntúan algo por debajo de la media en conocimientos financieros básicos, justo en la etapa en la que empiezan a decidir sin una nómina ni retenciones automáticas (Banco de España, 2022; ver Anexo I).`),
  h2("3.2. Propuesta de valor y servicio"),
  p(`El servicio se dirige a autónomos de 20 a 35 años porque es la etapa en la que muchas personas empiezan a trabajar por cuenta propia o a desarrollar sus primeros proyectos: suelen dominar su actividad, pero tienen menos experiencia gestionando sus finanzas. Es también un público acostumbrado a las herramientas digitales y, previsiblemente, más abierto a este tipo de acompañamiento. El servicio podría funcionar con personas de más edad, pero he preferido empezar con un público concreto y bien definido.`),
  p(`La propuesta de valor de Flujo es ayudar a los autónomos jóvenes a entender y controlar su dinero con un lenguaje sencillo, sin tecnicismos y con acompañamiento humano. Decidí que Flujo no presentara impuestos porque no quería convertirlo en una gestoría ni competir con los servicios que estas ya ofrecen. Su función es otra: ayudar al autónomo a entender sus números, organizar sus ingresos y gastos, prever sus obligaciones y saber cuánto dinero puede utilizar realmente. Ambos servicios se complementan: la gestoría se ocupa de las obligaciones fiscales y administrativas, y Flujo de que el cliente tenga una visión clara de su situación financiera. El nombre del proyecto viene de ahí, del flujo de dinero que un autónomo pierde de vista cuando ingresos, gastos e impuestos pasan por la misma cuenta. El servicio se organiza en tres capas que se despliegan de forma progresiva:`),
  ...vinetas([
    "**Servicio de acompañamiento (desde el año 1).** Un diagnóstico inicial y la puesta en marcha de un sistema personalizado de control de ingresos, gastos y reservas para impuestos (*setup*), seguidos de una suscripción mensual con una reunión de seguimiento y la revisión de los números.",
    "**Educación práctica (desde el año 2).** Plantillas y cursos breves basados en casos reales, como cuánto reservar para el IVA y el IRPF o cómo saber si la actividad es rentable.",
    "**Herramienta digital (desde la segunda mitad del año 3).** Un panel sencillo con previsión de tesorería y avisos, que permitiría atender a muchos más usuarios a un precio más bajo.",
  ]),
  h2("3.3. Modelo de negocio"),
  p(`La Tabla 3.1 recoge el Business Model Canvas de Flujo. Los precios del servicio parten de unos rangos iniciales, un *setup* de 100 a 300 € y una suscripción de 50 a 100 € al mes, situados entre lo que cobra una gestoría *online* y una tradicional (Tabla 4.1). El plan financiero toma el valor central. Los 200 € iniciales retribuyen el trabajo de analizar la situación económica del cliente y dejar organizada su estructura financiera desde el principio; los 75 € mensuales, un servicio continuado en el que no se trata de entregar una plantilla, sino de acompañar al cliente y ayudarle a tomar mejores decisiones. Aun así, para alguien que empieza como autónomo puede ser un gasto importante, y por eso es uno de los aspectos que más hay que comprobar: la encuesta lo contrasta con la disposición a pagar del público objetivo (Anexo III).`),
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
  p(`Siguiendo el enfoque Lean Startup, antes de invertir en producto el proyecto tendría que demostrar que alguien paga por el servicio. La Tabla 3.2 recoge las fases de validación previstas para los primeros cuatro meses. El criterio de éxito no es conseguir clientes a cualquier precio, sino conseguirlos al precio del modelo de negocio: un cliente que solo aceptaría el servicio gratis demuestra que el problema interesa, pero no que alguien pagaría por resolverlo. Me replantearía el proyecto si comprobara que no hay suficiente demanda real o que los clientes no están dispuestos a pagar, pero también si captar clientes resultara demasiado caro, si cada cliente exigiera más tiempo del previsto o si los ingresos no compensaran el trabajo. Una cosa es que el modelo financiero muestre resultados positivos y otra que el negocio funcione en la práctica.`),
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
  p(`El RETA contaba con 3.483.645 afiliados medios en septiembre de 2026 (Ministerio de Inclusión, Seguridad Social y Migraciones, 2026). El segmento al que se dirige Flujo puede acotarse con dos cifras. La estadística oficial del Ministerio de Trabajo cifra en 249.456 los autónomos de 16 a 29 años (Ministerio de Trabajo y Economía Social, 2026b), y el informe de ATA sobre el perfil del autónomo eleva el peso de los menores de 35 años al 16,1 % del total (Castilla y León Económica, 2026), unas ${miles(Math.round(tam / 10000) * 10000)} personas. Con estas cifras pueden distinguirse tres niveles. El mercado total lo forman entre 250.000 y 560.000 autónomos jóvenes, según el corte de edad. El mercado alcanzable se reduce a quienes tienen el problema que Flujo resuelve: aplicando el 38-42 % de ocupados que declara conocimientos financieros bajos o muy bajos (Banco de España, 2022), serían entre unas 100.000 y 230.000 personas. El mercado objetivo son los ${r0(A.activos_fin[2])} clientes de suscripción que el plan prevé al final del año 3, menos del 0,05 % incluso del mercado alcanzable más pequeño. Que la cifra sea pequeña no significa que captarlos sea fácil: la dificultad está en llegar a esos clientes y convencerlos de que paguen. Para Flujo el tamaño del mercado no es una restricción; el límite está en la capacidad de captar y atender clientes.`),
  p(`La estructura del colectivo, en cambio, no favorece al proyecto. Los autónomos de 46 a 65 años suman el 55,4 % del RETA y los de 16 a 29 apenas el 7,2 % (ver Anexo I), y la Estrategia Nacional de Impulso del Trabajo Autónomo identifica la franja de 25 a 39 años como la que más peso ha perdido en los últimos años (Ministerio de Trabajo y Economía Social, 2022). En la comparación europea, los autónomos de 20 a 29 años son solo el 5,9 % del total en España, frente al 7,9 % de media en la Unión Europea (Eurostat, 2026).`),
  p(`Hay, no obstante, un matiz reciente: el número de jóvenes de 16 a 29 años afiliados al RETA creció un 5,1 % interanual en marzo de 2026, por encima del conjunto del régimen (Ministerio de Trabajo y Economía Social, 2026b). Incluyo los dos datos aunque el primero no favorezca al proyecto, porque es más útil dimensionar el mercado con datos reales que con un titular optimista: el segmento joven es amplio y muestra señales de recuperación, pero su peso es pequeño y Flujo tendría que ganar cuota en lugar de esperar a que el mercado se expanda.`),
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
  ], [1900, 1500, 1800, 1500, 2370], { numericas: false }),
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
    ["• Inversión inicial muy reducida y punto de equilibrio bajo", "• Entre 250.000 y 560.000 autónomos jóvenes, con un repunte reciente"],
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
  p(`Se ha optado por iniciar la actividad como **empresario individual (autónomo)**. Frente a una sociedad limitada, tiene menos costes de constitución y gestión, permite acogerse a la tarifa plana y encaja con un proyecto que debe validarse antes de crecer. Sus inconvenientes son la responsabilidad ilimitada y una tributación progresiva por IRPF que, con beneficios elevados, supera al Impuesto sobre Sociedades. Además, la financiación de ENISA y las ventajas de la Ley 28/2022 de empresas emergentes exigen una sociedad mercantil. El apartado 8.5 compara la carga fiscal de las dos formas y concluye a partir de qué momento convendría constituir una sociedad limitada, cuyo capital mínimo es hoy de un euro tras la Ley 18/2022. En síntesis, el proyecto debe darse de alta en el RETA y en el censo de la Agencia Tributaria, repercutir el IVA del 21 % y presentarlo cada trimestre, tributar por IRPF en estimación directa simplificada con una retención reducida del 7 % durante el año de inicio y los dos siguientes, adaptar su facturación a Verifactu y a la factura electrónica, y cumplir la normativa de protección de datos. El Anexo V (Tabla V.1) detalla cada obligación y su efecto en el plan financiero.`),
  p(`Las cuotas de autónomos para 2027 no estaban aprobadas en el momento de redactar este trabajo. La Orden PJC/297/2026 mantuvo para 2026 las mismas bases mínimas de 2025 y fijó un tipo total del 31,5 % (28,30 % por contingencias comunes, 1,30 % por contingencias profesionales, 0,90 % por cese de actividad, 0,10 % por formación profesional y 0,90 % por el mecanismo de equidad intergeneracional). El Real Decreto-ley 13/2022 prevé que, desde 2026, la cuantía de la tarifa plana la fije cada año la Ley de Presupuestos; en ausencia de cambios se mantiene en 80 € (Infoautónomos, 2026). El rendimiento que determina el tramo es el rendimiento neto de la actividad más las propias cuotas pagadas, menos un 7 % de gastos genéricos (art. 308 de la Ley General de la Seguridad Social, en la redacción del Real Decreto-ley 13/2022). La prórroga de la tarifa plana exige prever que el rendimiento del segundo año será inferior al SMI; en el escenario base ese rendimiento supera los ${miles(Math.floor((R.base.anual.rendimiento_neto[1] + 12 * R.base.reta[1]) / 1000) * 1000)} €, por lo que el modelo solo aplica la tarifa plana el primer año. El plan aplica, por prudencia, la tabla vigente en 2026 a los tres años.`),
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
  p(`La Tabla 6.1 resume las hipótesis principales del escenario base y las clasifica según su respaldo: alto cuando proceden de la normativa o de fuentes oficiales, medio cuando son estimaciones razonadas y bajo cuando dependen de la validación comercial. Las cifras más sólidas son las fiscales y de costes; las más inciertas, precisamente las que más pesan en el resultado: precio, captación y bajas. La captación es, a mi juicio, la más difícil de estimar, sobre todo en los primeros meses: un negocio nuevo y sin marca conocida no sabe cuánto le costará conseguir sus primeros clientes ni cuánto tardarán en llegar. Tampoco el interés que alguien muestre en una encuesta garantiza que después contrate. Por eso los resultados del modelo deben entenderse como una estimación basada en estas hipótesis, no como una garantía. La retención de IRPF es la menos relevante: solo la practican los clientes que pagan el servicio como empresarios o profesionales, y solo si la actividad se da de alta en un epígrafe profesional; en cualquier caso, se limita a adelantar el pago del impuesto, y sin ninguna retención el VAN no cambia y el saldo mínimo de tesorería pasa de ${eur(A.caja_min)} a ${eur(R.sensibilidad.sin_retencion.caja_min)}. El resto de supuestos figura en el modelo de cálculo (ver Anexo II).`),
  p(`La hipótesis más delicada es el ritmo de altas, que se ha construido con un embudo de captación sencillo. En el año 1 se suponen unas 10 conversaciones al mes con autónomos interesados, procedentes de redes personales, comunidades y contenido; un 40 % acepta el diagnóstico gratuito y un 30 % de estos contrata, lo que da algo más de una alta al mes (14 en el año), en línea con el plan de validación de la Tabla 3.2. En los años 2 y 3, el contenido en redes, las recomendaciones de clientes y el presupuesto de marketing elevan los contactos a unos 20 y 25 al mes con las mismas tasas de conversión, lo que equivale a unas 2,5 y 3 altas mensuales. Las tasas de conversión son supuestos que la encuesta y la validación permitirán afinar, y el capítulo 9 calcula qué captación mínima necesita el proyecto para seguir compensando.`),
  leyenda("Tabla", "6.1", "Principales hipótesis del escenario base y nivel de respaldo"),
  tabla(["Hipótesis", "Valor", "Justificación", "Respaldo"], [
    ["Cuotas, IVA, IRPF e inflación", "Normativa 2026; IPC 3,7 %", "BOE, Seguridad Social, Agencia Tributaria y Banco de España", "Alto"],
    ["Comisiones", "1,5 % servicio; 10 % productos", "Tarifas habituales de pasarelas y plataformas", "Alto"],
    ["Costes fijos", "195 €/mes más marketing", "Precios de mercado del software y servicios", "Medio"],
    ["Capacidad del servicio", "60 h/mes (año 1) y 120 h/mes", "2,5 h por cliente y mes; 4 h por alta", "Medio"],
    ["Clientes con retención de IRPF", "60 %", "Clientes que pagan como profesionales; solo adelanta el IRPF", "Medio"],
    ["Precio del *setup* y de la suscripción", "200 € y 75 €/mes", "Centro de los rangos de partida; entre gestoría *online* y tradicional", "Bajo: pendiente de la encuesta"],
    ["Altas de clientes", "14 / 30 / 36 al año", "Embudo de captación descrito en el texto", "Bajo: pendiente de validación"],
    ["Bajas", "4 % mensual", "Retención anual del 61 %, por debajo del objetivo de validación", "Bajo: pendiente de validación"],
    ["Plantillas, cursos y herramienta", "15 €, 35 € y 12 €/mes", "Desde el año 2; herramienta desde el mes 31", "Bajo: pendiente de validación"],
  ], [2300, 1900, 3270, 1600], { numericas: false }),
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
  p(`De acuerdo con la jerarquía financiera descrita en el capítulo 2, el arranque se financia íntegramente con 3.000 € de ahorros propios. El desarrollo de la herramienta digital, la única inversión relevante, se ha presupuestado en 12.000 €, el coste aproximado de unas 300 horas de un desarrollador freelance a 40 €/h para un producto mínimo con panel de ingresos y gastos, previsión de tesorería y avisos; es una estimación propia, y el capítulo 8 analiza una alternativa *no-code* de 5.000 €. Se financiaría con un microcrédito de 10.000 € a cinco años y un tipo fijo del 7 %, en línea con las condiciones de MicroBank para emprendedores (Cámara de Comercio de Valencia, 2023), y con la caja generada por la actividad. Es un tipo superior a la TAE media de los nuevos créditos de hasta 250.000 € a sociedades no financieras, que el Banco de España situaba en el 4,58 % en enero de 2026 (Banco de España, 2026b), una diferencia lógica porque el microcrédito no exige garantías ni historial. En el escenario base el proyecto podría autofinanciar esa inversión, pero el préstamo se mantiene por prudencia: preserva un colchón de liquidez y, como se verá en el capítulo 9, sería imprescindible si la captación fuese más lenta. En el plan base la herramienta se incluye en el año 3 para medir su efecto, pero su desarrollo debería quedar condicionado a que el servicio y los productos digitales hayan demostrado demanda, como se analiza en el apartado 8.2. ENISA y los *business angels* no encajan en esta fase, porque exigen una sociedad mercantil y buscan crecimientos que un servicio unipersonal no ofrece; serían opciones a estudiar si la herramienta digital demostrara tracción.`),
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
  p(`Con un margen de contribución de ${eur(75 * 0.985, 2)} por cliente y mes, el punto de equilibrio del año 1 se alcanza con unos **${Math.ceil(B.pe[0])} clientes de suscripción** activos, de ${Math.ceil(B.pe[1])} en el año 2, cuando termina la tarifa plana, y de ${Math.ceil(B.pe[2])} en el año 3, cuando además crece el marketing, empieza a amortizarse la herramienta y aparecen los intereses del microcrédito. Son umbrales bajos en comparación con la capacidad del servicio. El dato cambia mucho si se exige que el proyecto pague también un sueldo al fundador: para cubrir el equivalente al SMI harían falta unos ${Math.ceil(B.pe_smi[0])} clientes en el año 1 y ${Math.ceil(B.pe_smi[2])} en el año 3, una cifra cercana al límite de capacidad.`),
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
  p(`La rentabilidad exigida al proyecto se ha estimado con el procedimiento descrito en el marco teórico (Tabla 7.1). Se parte de la rentabilidad de las Obligaciones del Estado a diez años, cuyo tipo medio fue del 4,176 % en la subasta del 1 de octubre de 2026 (Tesoro Público, 2026), y se añade una prima de mercado del 5,5 %, coherente con la prima implícita que estima Damodaran (2026) para un mercado maduro más un pequeño diferencial por riesgo país. Con una beta de 1, propia de un negocio de servicios profesionales sin deuda, el resultado es del 9,7 %. Esa tasa sería razonable para una empresa consolidada, pero no para un proyecto que todavía no tiene clientes, que no podría venderse con facilidad y que depende por completo de una persona.`),
  p(`En el otro extremo, los fondos de capital riesgo exigen a las empresas jóvenes tasas de entre el 30 % y el 70 % según su fase (Damodaran, 2009), pero esas tasas incorporan la probabilidad de fracaso, que en este trabajo se analiza aparte con escenarios y con una simulación de Montecarlo. Entre ambos extremos se añade una prima específica de algo más de 5 puntos por tamaño, iliquidez y dependencia del fundador, y la tasa se redondea al ${pct(R.tasa, 0)}. Es un supuesto razonado, no un cálculo exacto, y por eso el capítulo 9 muestra cómo cambia el VAN con tasas de entre el 10 % y el 30 %.`),
  leyenda("Tabla", "7.1", "Estimación de la tasa de descuento"),
  tabla(["Componente", "Valor"], [
    ["Rentabilidad sin riesgo (Obligaciones del Estado a 10 años)", "4,2 %"],
    ["Prima de riesgo de mercado × beta (5,5 % × 1,0)", "5,5 %"],
    ["Prima específica (tamaño, iliquidez y dependencia del fundador)", "5,3 %"],
    ["!Tasa de descuento", pct(R.tasa, 1)],
  ], [7070, 2000]),
  fuente("elaboración propia a partir de Damodaran (2009, 2012, 2026) y Tesoro Público (2026)."),
  h2("7.2. VAN, TIR y periodo de recuperación"),
  p(`Los flujos de caja libres del proyecto (Tabla 7.2) se obtienen a partir del EBITDA, restando el IRPF que correspondería al resultado de explotación, las inversiones y la variación del capital circulante, y sin tener en cuenta las retiradas del titular ni la financiación. En el año 0 se incluye la inversión inicial y el colchón de tesorería, que se recupera al final del año 3. El flujo del año 0 (${eur(-V.flujos[0])}) es algo inferior a la aportación de 3.000 € de la Tabla 6.2 porque no incluye los ${eur(3000 + V.flujos[0], 2)} de IVA soportado en la inversión inicial, que se recuperan en la primera liquidación del impuesto y quedan recogidos en la variación del capital circulante.`),
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
  p(`El resultado de un empresario individual es, en realidad, su sueldo. Para que Flujo sea una opción razonable no basta con que genere caja: tiene que generar más que la alternativa de trabajar por cuenta ajena. Para medirlo se ha restado a los flujos de caja un sueldo neto equivalente al SMI de 2026 (17.094 € brutos, unos 15.980 € netos de cotización), la mitad en el año 1, en el que la dedicación es parcial. Se ha elegido el SMI por ser un umbral objetivo; más abajo se usa también un listón más exigente.`),
  p(`Con este ajuste los flujos pasan a ser de ${V.flujos_oportunidad.slice(1).map((x) => eur(x)).join(", ")} en los años 1 a 3, y **el VAN se reduce a ${eur(V.van_oportunidad)}**, con una TIR del ${miles(V.tir_oportunidad * 100, 0)} %. El proyecto sigue creando valor, pero el margen es mucho más estrecho: el año 1 no cubre el sueldo alternativo y el resultado depende de que los años 2 y 3 se cumplan. Es probablemente lo que más me ha sorprendido del análisis: los resultados parecen muy positivos hasta que se valora el tiempo que exige el negocio, y entonces la rentabilidad se reduce de forma considerable. No basta con que una empresa genere beneficios; también tiene que compensar el esfuerzo de quien la dirige. Este es, a mi juicio, el indicador más honesto de la viabilidad de Flujo.`),
  p(`El SMI es, sin embargo, un listón bajo para un graduado en ADE. Como alternativa más exigente se ha usado la ganancia media anual de los asalariados de 25 a 29 años, 20.460 € brutos según la Encuesta de Estructura Salarial del INE (2024), que equivalen a unos ${eur(V.neto_ine)} netos de cotización e IRPF. Con ese sueldo como referencia, el VAN con coste de oportunidad baja a ${eur(V.van_ine)}. El dato de la encuesta es de 2022, por lo que el listón real sería algo mayor: con el plan base, Flujo apenas compensa frente a un empleo de un joven con estudios, lo que obliga a buscar mejoras (capítulo 8).`),
  p(`La Tabla 7.3 muestra la misma idea desde el punto de vista del fundador: lo que el negocio deja por cada hora trabajada, contando media jornada el primer año y jornada completa (unas 1.760 horas anuales) después, frente a lo que ganaría por hora en un empleo por cuenta ajena.`),
  leyenda("Tabla", "7.3", "Resultado del negocio y ganancia del fundador por hora trabajada"),
  tabla(["Concepto", ...anios], [
    fila("Resultado neto del negocio (plan base)", A.resultado_neto),
    fila("Retiradas previstas para vivir", A.flujo.retiradas),
    fila("Horas de trabajo al año", [880, 1760, 1760], (x) => miles(x)),
    fila("!Ganancia por hora: plan base", R.ganancia_hora.base, (x) => eur(x, 1)),
    fila("Ganancia por hora: plan mejorado", R.ganancia_hora.optimizado, (x) => eur(x, 1)),
    fila("Ganancia por hora: escenario pesimista", R.ganancia_hora.pesimista, (x) => eur(x, 1)),
    ["Referencia: asalariado con el SMI (neto)", eur(R.ganancia_hora.smi, 1), eur(R.ganancia_hora.smi, 1), eur(R.ganancia_hora.smi, 1)],
    ["Referencia: asalariado de 25 a 29 años (neto, INE)", eur(R.ganancia_hora.ine, 1), eur(R.ganancia_hora.ine, 1), eur(R.ganancia_hora.ine, 1)],
  ], W4),
  fuente("elaboración propia a partir del modelo, el Real Decreto 126/2026 e INE (2024)."),
  p(`El primer año el fundador ganaría unos ${eur(R.ganancia_hora.base[0], 0)} por hora, menos de la mitad que un asalariado; a partir del segundo año la relación se invierte y en el tercero llegaría a unos ${eur(R.ganancia_hora.base[2], 0)} por hora, casi el doble. En el escenario pesimista, en cambio, no alcanzaría el salario mínimo en ningún año. Flujo solo resulta atractivo para su fundador si se supera el primer año y la captación se acerca a la prevista, y a partir del segundo año las retiradas previstas son inferiores al resultado, lo que permite ir formando un colchón de liquidez.`),
  h2("7.4. Ratios financieros"),
  p(`La Tabla 7.4 recoge los principales ratios. Los márgenes son altos porque los costes variables apenas existen y el principal coste, el trabajo del fundador, no figura como gasto. La liquidez es holgada desde el primer año y el endeudamiento solo aumenta en el año 3 por el microcrédito, con una cobertura de intereses muy amplia. Las rentabilidades sobre activos y sobre patrimonio neto superan el 100 % porque la empresa apenas tiene activos y el titular retira buena parte del beneficio; por eso no se consideran indicadores útiles en este caso.`),
  leyenda("Tabla", "7.4", "Ratios financieros del escenario base"),
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
    ["!LTV ajustado por el tiempo de atención", eur(UE.ltv_aj), eur(UE.ltv_aj), eur(UE.ltv_aj)],
    fila("!CAC ajustado por el tiempo de captación", UE.cac_aj),
    fila("!Relación LTV / CAC ajustada", UE.ltv_cac_aj, (x) => num(x, 1)),
  ], W4),
  fuente("elaboración propia."),
  p(`Estas dos cifras no incluyen el tiempo del fundador: el LTV no descuenta las horas de atención de cada cliente y el CAC solo recoge el gasto en marketing. Por eso la Tabla 8.1 añade una versión ajustada, que valora cada hora a ${eur(UE.valor_hora, 2)}, el salario neto por hora de un asalariado de 25 a 29 años, y suma al CAC las ${num(7.5, 1)} horas de conversaciones y diagnósticos que, según el embudo del apartado 6.1, exige cada alta. Incluso así, la relación LTV/CAC se mantiene entre ${num(UE.ltv_cac_aj[2], 0)} y ${num(UE.ltv_cac_aj[0], 0)}, por encima del umbral de 3 que suele tomarse como referencia en los negocios de suscripción. La lectura es clara: captar clientes no es caro. El cuello de botella de Flujo no es el marketing, sino las horas del fundador, que se agotan al final del año 3 (Figura 6.2). Por eso las medidas de mejora que siguen buscan, sobre todo, **aumentar los ingresos por hora trabajada** en lugar de captar más clientes.`),
  h2("8.2. Rentabilidad por línea de negocio"),
  p(`Para saber qué aporta cada capa del modelo se ha calculado el VAN incremental de cada una, es decir, la diferencia entre el plan completo y el mismo plan sin esa línea. Los productos digitales (plantillas y cursos) aportan ${eur(LIN.productos.van)} al VAN, porque las horas que exigen no dependen del número de clientes. La herramienta digital, en cambio, tiene un VAN incremental de ${eur(LIN.herramienta.van)} en el horizonte del plan: la inversión de 12.000 € se hace en el año 3 y apenas hay tiempo para recuperarla.`),
  p(`Fue un resultado que no esperaba, pero no significa que la herramienta deba descartarse. Su valor está fuera del horizonte de tres años, ya que es la única capa que permite crecer por encima del límite de capacidad del servicio. Se comporta como una opción real: conviene conservar la posibilidad de desarrollarla, pero al menor coste posible y solo cuando el servicio haya demostrado que hay demanda. Esa idea se recoge en la última de las medidas del apartado siguiente.`),
  h2("8.3. Medidas de mejora"),
  p(`Se proponen cinco medidas de bajo coste, coherentes con el diagnóstico anterior. Todas son hipótesis que deberían validarse, igual que las del plan base:`),
  ...vinetas([
    "**Actualizar los precios cada año con el IPC.** El plan base mantiene los precios constantes mientras los costes suben un 3,7 % al año; actualizar las tarifas es una práctica habitual en los servicios por suscripción.",
    "**Estandarizar el servicio.** Con las plantillas y la rutina de revisión ya construidas, el tiempo por cliente puede bajar de 2,5 a 2 horas al mes desde el segundo año, lo que eleva la capacidad máxima de unos 45 a unos 55 clientes.",
    "**Talleres grupales para instituciones.** Sesiones de dos horas sobre organización financiera para autónomos, ofrecidas a asociaciones de autónomos, coworkings, colegios profesionales, universidades o programas públicos de emprendimiento, a 250 € por sesión: una al mes en el año 2 y dos en el año 3. Cada taller ocupa unas 5 horas y genera más ingresos por hora que el servicio individual, además de servir como canal de captación.",
    "**Acuerdo de derivación con gestorías.** Flujo no presenta impuestos, por lo que puede derivar a una gestoría colaboradora a los clientes que no tengan una, a cambio de una comisión de 60 € por cliente; se supone que el 30 % de las altas la necesita.",
    "**Herramienta digital de bajo coste.** Desarrollar la primera versión con herramientas *no-code* por unos 5.000 € en lugar de un desarrollo a medida de 12.000 €, financiada con la caja generada y sin microcrédito.",
  ]),
  p(`La Tabla 8.2 y la Figura 8.1 muestran el efecto de aplicar las medidas de forma acumulada.`),
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
  p(`El VAN es positivo en todas las simulaciones, tanto en el plan base como en el mejorado. La diferencia está en el coste de oportunidad: con el plan base, Flujo supera al empleo alternativo en el ${pct(MC.p_van_op, 0)} de los casos, prácticamente una moneda al aire; con el plan mejorado, en el ${pct(MCO.p_van_op, 0)}. Con el listón más exigente del salario medio de 25 a 29 años, las probabilidades son del ${pct(MC.p_van_ine, 0)} y del ${pct(MCO.p_van_ine, 0)}, respectivamente. Estas probabilidades deben leerse con cautela: están condicionadas a los rangos y distribuciones elegidos, que son a su vez hipótesis, y no recogen riesgos como un precio aceptable inferior a 75 € o un coste de captación mayor. Su utilidad no es predecir el éxito, sino comparar el plan base y el mejorado bajo los mismos supuestos, y en esa comparación las medidas reducen claramente el riesgo de que el proyecto no compense el esfuerzo del fundador.`),
  h2("8.5. Forma jurídica y fiscalidad"),
  p(`El empresario individual es la forma más sencilla para arrancar, pero su tributación progresiva aumenta con los beneficios. En una sociedad limitada, en cambio, los beneficios tributan al 15 % durante los primeros ejercicios con base positiva si es de nueva creación, y las microempresas con una cifra de negocios inferior a un millón de euros tributan en 2027 al 17 % por los primeros 50.000 € y al 20 % por el resto (Agencia Tributaria, 2026a). La Ley 28/2022 añade ventajas para las empresas emergentes (un tipo del 15 % durante cuatro años, el aplazamiento de la deuda tributaria de los dos primeros ejercicios con beneficios o la exención de pagos fraccionados), pero exige un modelo innovador y escalable certificado por ENISA, al que un servicio manual difícilmente podría aspirar hasta que la herramienta digital tenga tracción. La Tabla 8.4 compara la carga fiscal del año 3 con la que tendría una sociedad limitada que pagara un sueldo al fundador y tributara al 15 % por el resto, el tipo de las entidades de nueva creación (Agencia Tributaria, 2026a). Se añaden unos 1.200 € anuales de gestoría y obligaciones mercantiles y se supone una cuota de autónomo similar en ambos casos.`),
  leyenda("Tabla", "8.4", "Carga fiscal anual: empresario individual frente a sociedad limitada"),
  tabla(["Rendimiento y sueldo", "IRPF como autónomo", "Impuestos con SL", "Ahorro neto con SL"], [SL[1], SL[2], SL[3]].map((x) => [
    `${eur(x.rendimiento)} con sueldo de ${eur(x.sueldo)}`, eur(x.irpf_autonomo), eur(x.impuestos_sl), eur(x.ahorro),
  ]), [3670, 1800, 1800, 1800]),
  fuente("elaboración propia a partir de las escalas del IRPF y de Agencia Tributaria (2026a). El ahorro descuenta 1.200 € de costes adicionales de la sociedad."),
  p(`Con el rendimiento del año 3 del plan base, de unos ${eur(Math.round(SL[0].rendimiento / 1000) * 1000)}, la sociedad ahorraría unos ${eur(SL[1].ahorro)} al año, pero a cambio de que una parte del beneficio quede retenida en ella; si se repartiera como dividendos tributaría de nuevo en el IRPF del ahorro, por lo que buena parte del ahorro sería un diferimiento. Con un rendimiento de 60.000 € el ahorro supera los ${eur(Math.floor(SL[3].ahorro / 1000) * 1000)}. La conclusión es que el empresario individual es la forma adecuada durante el horizonte del plan, y que la transformación en sociedad limitada tendría sentido a partir del cuarto año si el beneficio supera los 50.000 € o si se quiere reinvertir en la herramienta digital y acceder a ENISA.`),
];

// ------------------------------------------------- 9. RIESGOS Y ESCENARIOS
const esc = ["pesimista", "base", "optimista"];
const cap8 = [
  h1("9. Análisis de riesgos y escenarios", false),
  h2("9.1. Escenarios"),
  p(`Para comprobar la solidez de los resultados se han construido dos escenarios alternativos. En el **pesimista** las altas de clientes son un 40 % inferiores a las del base, las bajas suben al 6 % mensual y las ventas de productos y usuarios de la herramienta se reducen a la mitad. En el **optimista** las altas aumentan un 30 %, las bajas bajan al 3 % y los productos y la herramienta venden un 30 % y un 50 % más, respectivamente. Los costes, la financiación y las retiradas del titular son los mismos en los tres casos, para que la comparación sea limpia. La Tabla 9.1 resume los resultados.`),
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
  p(`El VAN es positivo en los tres escenarios, pero solo porque la inversión en dinero es muy pequeña; no significa que el proyecto tenga poco riesgo. Los otros dos indicadores lo muestran. En el escenario pesimista la tesorería se vuelve negativa ya en el primer año y llega a un mínimo de ${eur(PE.anual.caja_min)} en el mes ${PE.anual.mes_caja_min} (Figura 9.1): el fundador no podría retirar lo previsto para vivir y tendría que mantener otra fuente de ingresos o retrasar el desarrollo de la herramienta. Además, el VAN con coste de oportunidad pasa a ${eur(PE.valoracion.van_oportunidad)}: con esa captación, Flujo rendiría menos que un empleo con el salario mínimo.`),
  leyenda("Figura", "9.1", "Saldo mensual de tesorería por escenario"),
  imagen("fig_tesoreria.png", 12),
  fuente("elaboración propia."),
  p(`En el escenario optimista, por el contrario, el límite de capacidad se alcanza ya en el año 2, lo que confirma que el crecimiento del servicio está acotado por las horas del fundador y no por la demanda.`),
  h2("9.2. Análisis de sensibilidad"),
  p(`Al modificar, de una en una, las dos variables más inciertas, el VAN del escenario base se comporta así: una rebaja del precio de la suscripción a 55 €, cercana al de una gestoría *online*, reduciría el VAN a ${eur(R.sensibilidad.precio["55"])}, y una captación a la mitad de la prevista lo dejaría en ${eur(R.sensibilidad.altas["0.5"])}. El VAN se mantiene positivo en todos los casos. Respecto a la tasa de descuento, pasar del 10 % al 30 % lo reduce de ${eur(R.sensibilidad.tasa["0.1"])} a ${eur(R.sensibilidad.tasa["0.3"])}, una variación pequeña porque el horizonte es corto y la inversión, mínima. La captación de clientes es, por tanto, la variable crítica. Por último, si los precios de Flujo se actualizaran con la misma inflación que los costes (la primera medida del capítulo 8), el VAN subiría a ${eur(R.sensibilidad.precios_indexados.van)} y el VAN con coste de oportunidad a ${eur(R.sensibilidad.precios_indexados.van_oportunidad)}.`),
  p(`Más útil que mover una variable cada vez es preguntarse a partir de qué valor el proyecto deja de compensar frente al empleo alternativo. Con el plan base, el VAN con coste de oportunidad se anula con una suscripción de unos ${num(R.criticos.precio_van_oportunidad, 0)} €/mes (${num(R.criticos.precio_van_ine, 0)} €/mes con el salario medio de 25 a 29 años) o con una captación del ${pct(R.criticos.altas_van_oportunidad, 0)} de la prevista. Con el plan mejorado, esos umbrales bajan a unos ${num(R.criticos_opt.precio_van_oportunidad, 0)} y ${num(R.criticos_opt.precio_van_ine, 0)} €/mes y a una captación del ${pct(R.criticos_opt.altas_van_oportunidad, 0)}: Flujo seguiría compensando aunque tuviera que cobrar lo mismo que una gestoría *online*. Estos umbrales son los valores que la encuesta y la validación deben contrastar.`),
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
  p(`Este trabajo se planteaba si sería viable un proyecto dedicado a mejorar la gestión financiera de los autónomos jóvenes. La respuesta es que Flujo presenta una viabilidad económico-financiera potencial bajo las hipótesis del escenario base, pero condicionada a tres factores: alcanzar el ritmo de captación previsto, mantener un precio cercano al planteado y gestionar la capacidad del fundador. La validación comercial es, por tanto, una condición previa a su puesta en marcha.`),
  p(`De todo el análisis se desprenden tres ideas que condicionan la decisión sobre el proyecto.`),
  p(`**La primera es que el problema existe, pero la demanda está por demostrar.** Los datos del Banco de España muestran un nivel bajo de conocimientos financieros en la población adulta, algo peor entre los jóvenes, y la encuesta propia [[confirma / matiza]] que en el público objetivo son frecuentes problemas muy concretos: [[resumen de los principales resultados de la encuesta]]. El segmento, de entre 250.000 y 560.000 autónomos jóvenes, es más que suficiente para un proyecto de este tamaño. Pero que alguien tenga un problema no significa que pague por resolverlo, y más cuando, como señala la literatura, muchos creen saber más de lo que saben. El precio de 75 €/mes, superior al de una gestoría *online*, es la hipótesis más expuesta del plan.`),
  p(`**La segunda es que la rentabilidad de Flujo depende sobre todo del tiempo de su fundador, no del dinero invertido.** Con solo 3.000 € de inversión y un punto de equilibrio de unos ${Math.ceil(B.pe[0])} clientes el primer año, el VAN del escenario base es de ${eur(V.van)} y la TIR, del ${miles(V.tir * 100, 0)} %, pero esas cifras dicen poco. Cuando se descuenta un sueldo alternativo, el VAN baja a ${eur(V.van_oportunidad)} con el SMI y a ${eur(V.van_ine)} con el salario medio de un joven de 25 a 29 años; el primer año el fundador ganaría unos ${eur(R.ganancia_hora.base[0], 0)} por hora, y en el escenario pesimista no llegaría al salario mínimo en ningún año, con una tesorería que cae hasta ${eur(PE.anual.caja_min)}. Además, el servicio se queda sin horas hacia el final del tercer año. Que el proyecto sea rentable no significa, por tanto, que convenga emprenderlo tal como está planteado: con el plan base, apenas compensa frente a un empleo por cuenta ajena.`),
  p(`**La tercera es que el modelo puede mejorarse sin invertir más, y que la decisión razonable es avanzar por etapas.** Incluso imputando el tiempo del fundador, captar un cliente cuesta mucho menos de lo que aporta (la relación LTV/CAC supera ${num(UE.ltv_cac_aj[2], 0)}), así que el margen de mejora no está en captar más, sino en ganar más por cada hora. Cinco medidas de bajo coste (actualizar precios, estandarizar el servicio, ofrecer talleres a instituciones, colaborar con gestorías y abaratar la herramienta digital) elevan el VAN con coste de oportunidad de ${eur(PAL[0].van_oportunidad)} a ${eur(PAL[PAL.length - 1].van_oportunidad)} y bajan el precio mínimo necesario a unos ${num(R.criticos_opt.precio_van_oportunidad, 0)} €/mes, y bajo los supuestos de la simulación de Montecarlo elevan del ${pct(MC.p_van_op, 0)} al ${pct(MCO.p_van_op, 0)} la probabilidad de que el proyecto compense. La herramienta digital, en cambio, resta valor dentro del horizonte del plan y debería tratarse como una opción que solo se ejerce cuando haya demanda probada. La secuencia que se deduce del análisis es validar primero el servicio con clientes reales, añadir después las plantillas y los cursos, probar la herramienta con una versión *no-code* y estudiar la transformación en sociedad limitada solo si el negocio crece.`),
  p(`El contexto refuerza esta prudencia en los dos sentidos: la factura electrónica y Verifactu obligarán a muchos autónomos a ordenar sus números entre 2027 y 2029, lo que favorece al proyecto, pero casi la mitad de las empresas nuevas desaparece en sus tres primeros años (INE, 2025). Con estas conclusiones se cubren los objetivos planteados en la introducción, desde la descripción del colectivo y sus dificultades (capítulos 3 y 4) hasta la viabilidad, la mejora de la rentabilidad y los riesgos del proyecto (capítulos 6 a 9).`),
  h2("10.2. Trascendencia de los resultados"),
  p(`Para un emprendedor, el principal valor del trabajo es que convierte una idea en un conjunto de hipótesis medibles: cuántos clientes hacen falta, cuánto tiempo consume cada uno y qué ocurre si el ritmo de captación es menor. El modelo, publicado en abierto, puede actualizarse con los datos reales de la validación. Desde un punto de vista más general, el trabajo muestra que la falta de educación financiera de los autónomos jóvenes es también una oportunidad de negocio, y que un servicio de bajo coste puede contribuir a un objetivo que las instituciones consideran prioritario.`),
  h2("10.3. Limitaciones y líneas futuras"),
  p(`El trabajo tiene varias limitaciones. La encuesta propia es una muestra de conveniencia, difundida por redes y comunidades de autónomos, y no es representativa; además, la disposición a pagar declarada en una encuesta suele ser mayor que la real, que solo puede comprobarse vendiendo el servicio. Los datos de competencias financieras proceden de una encuesta de 2021-2022 y no ofrecen información específica sobre autónomos jóvenes, y el peso de los menores de 35 años procede de un informe de ATA y no de la estadística oficial, que solo publica el tramo de 16 a 29 años. Las cuotas de autónomos de 2027 a 2029 no están aprobadas, y la tributación por IRPF depende de la situación personal del titular, que se ha simplificado. Por último, la tasa de descuento incluye una prima específica que es, inevitablemente, una estimación.`),
  p(`Después de este análisis, sí intentaría poner en marcha Flujo, pero de forma prudente: sin una gran inversión inicial y sin desarrollar directamente una herramienta digital. Empezaría ofreciendo el servicio de forma manual a un grupo reducido de clientes para conocer cuánto tiempo necesita cada uno, qué partes del servicio valoran más y si mantendrían una suscripción mensual, e iría ajustando a partir de ahí los precios y la forma de trabajar. Solo si comprobara que existe demanda suficiente tendría sentido invertir en automatizar procesos o en una herramienta propia.`),
  p(`Como líneas futuras se proponen, por tanto, ejecutar el plan de validación de la Tabla 3.2, preferiblemente dentro de un programa como el Semillero de Ideas de Zaragoza Activa, y actualizar el modelo con los datos reales, explorar acuerdos con gestorías y estudiar la transformación en sociedad limitada si la herramienta digital alcanza suficientes usuarios, lo que abriría la puerta a la financiación de ENISA o de inversores privados.`),
];

// ------------------------------------------------- BIBLIOGRAFÍA
const ref = (t) => new Paragraph({
  alignment: AlignmentType.LEFT, indent: { left: 567, hanging: 567 }, spacing: { before: 0, after: 120, line: 360, lineRule: "auto" }, children: require("./helpers").runs(t),
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
    `CASSAR, G. (2004): "The financing of business start-ups". *Journal of Business Venturing*, Vol. 19, Nº 2, pp. 261-283.`,
    `CASTILLA Y LEÓN ECONÓMICA (2026): *6 de cada 10 autónomos llevan más de 5 años en su negocio* [en línea]. 9 de febrero de 2026. https://www.castillayleoneconomica.es/6-10-autonomos-llevan-5-anos-frente-negocio/ ${consulta}.`,
    `CNMV, COMISIÓN NACIONAL DEL MERCADO DE VALORES (2026): *Asesoramiento en materia de inversión* [en línea]. Madrid. https://www.cnmv.es/portal/inversor/asesoramiento ${consulta}.`,
    `DAMODARAN, A. (2009): *Valuing Young, Start-up and Growth Companies: Estimation Issues and Valuation Challenges*. Working paper, Stern School of Business, New York University.`,
    `DAMODARAN, A. (2012): *Investment Valuation: Tools and Techniques for Determining the Value of Any Asset* (3ª ed.). John Wiley & Sons, Hoboken.`,
    `DAMODARAN, A. (2026): *Country Default Spreads and Risk Premiums* [en línea]. Actualización de enero de 2026. https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html ${consulta}.`,
    `DIXIT, A. K. y PINDYCK, R. S. (1994): *Investment under Uncertainty*. Princeton University Press, Princeton.`,
    `DREXLER, A., FISCHER, G. y SCHOAR, A. (2014): "Keeping it simple: Financial literacy and rules of thumb". *American Economic Journal: Applied Economics*, Vol. 6, Nº 2, pp. 1-31.`,
        `ENISA (2026): *Línea Emprendedores* [en línea]. Madrid. https://www.enisa.es/es/financiacion/info/lineas-enisa/emprendedores ${consulta}.`,
    `EUROSTAT (2026): *EU sees 2.06 million young entrepreneurs in 2025* [en línea]. Luxemburgo, 21 de abril de 2026. https://ec.europa.eu/eurostat/web/products-eurostat-news/w/edn-20260421-1 ${consulta}.`,
    `FEDERACIÓN NACIONAL DE ASOCIACIONES DE TRABAJADORES AUTÓNOMOS, ATA (2026): "El RETA ganó 39.000 autónomos en 2025, pero ATA prevé peores perspectivas para 2026". *Autónomos y Emprendedor*, 5 de enero de 2026. https://www.autonomosyemprendedor.es/articulo/autonomos/reta-gano-39000-autonomos-2025-menor-crecimiento-peores-perspectivas-2026/20260105114615047592.html ${consulta}.`,
    `GOBIERNO DE ARAGÓN (2026): *Subvenciones para el fomento y consolidación del empleo autónomo. Convocatoria 2026 (Orden EMC/1877/2025)* [en línea]. Zaragoza. https://www.aragon.es/tramitador/-/tramite/programa-emprendedores-autonomos ${consulta}.`,
    `GUPTA, S. y LEHMANN, D. R. (2003): "Customers as assets". *Journal of Interactive Marketing*, Vol. 17, Nº 1, pp. 9-24.`,
    `HERTZ, D. B. (1964): "Risk analysis in capital investment". *Harvard Business Review*, Vol. 42, Nº 1, pp. 95-106.`,
    `ICO, INSTITUTO DE CRÉDITO OFICIAL (2026): *ICO Empresas y Emprendedores* [en línea]. Madrid. https://www.ico.es/en/ico-empresas-y-emprendedores ${consulta}.`,
    `INE, INSTITUTO NACIONAL DE ESTADÍSTICA (2024): *Encuesta de Estructura Salarial 2022. Principales resultados*. Madrid. https://www.ine.es/metodologia/t22/ees_prinre22.pdf ${consulta}.`,
    `INE, INSTITUTO NACIONAL DE ESTADÍSTICA (2025): *Demografía Armonizada de Empresas. Año 2023*. Nota de prensa. Madrid. https://www.ine.es/dyngs/Prensa/DAE2023.htm ${consulta}.`,
    `INFOAUTÓNOMOS (2026): *La tarifa plana de 80 euros para autónomos en 2026: requisitos, cuantía y renovación* [en línea]. https://www.infoautonomos.com/seguridad-social/tarifa-plana-autonomos/ ${consulta}.`,
    `LA MONCLOA (2026): *Educación renueva su colaboración con el Banco de España, la CNMV y el Ministerio de Economía para fomentar la educación financiera* [en línea]. 16 de enero de 2026. https://www.lamoncloa.gob.es/serviciosdeprensa/notasprensa/educacion-fp-deportes/paginas/2026/160126-tolon-fomento-educacion-financiera.aspx ${consulta}.`,
    `LEY 18/2022, de 28 de septiembre, de creación y crecimiento de empresas. *Boletín Oficial del Estado*, Nº 234, de 29 de septiembre de 2022.`,
    `LEY 28/2022, de 21 de diciembre, de fomento del ecosistema de las empresas emergentes. *Boletín Oficial del Estado*, Nº 306, de 22 de diciembre de 2022. https://www.boe.es/buscar/act.php?id=BOE-A-2022-21739 ${consulta}.`,
    `LEY 35/2006, de 28 de noviembre, del Impuesto sobre la Renta de las Personas Físicas. *Boletín Oficial del Estado*, Nº 285, de 29 de noviembre de 2006.`,
    `LEY 37/1992, de 28 de diciembre, del Impuesto sobre el Valor Añadido. *Boletín Oficial del Estado*, Nº 312, de 29 de diciembre de 1992.`,
    `LEY ORGÁNICA 3/2018, de 5 de diciembre, de Protección de Datos Personales y garantía de los derechos digitales. *Boletín Oficial del Estado*, Nº 294, de 6 de diciembre de 2018.`,
    `LUSARDI, A. y MITCHELL, O. S. (2014): "The economic importance of financial literacy: Theory and evidence". *Journal of Economic Literature*, Vol. 52, Nº 1, pp. 5-44.`,
    `MINISTERIO DE HACIENDA (2026): *Nota informativa: previsión de modificación de los plazos de aplicación del Real Decreto 1007/2023 (Verifactu)* [en línea]. Madrid, 5 de octubre de 2026. https://www.hacienda.gob.es/sgt/gabsehacienda/nota-informativa-verifactu.pdf ${consulta}.`,
    `MINISTERIO DE INCLUSIÓN, SEGURIDAD SOCIAL Y MIGRACIONES (2026): *El empleo logra el mejor septiembre de la historia después de sumar 92.327 afiliados en el último mes y 739.888 respecto a hace un año* [en línea]. Madrid, 2 de octubre de 2026. https://www.inclusion.gob.es/en/w/el-empleo-logra-el-mejor-septiembre-de-la-historia-despues-de-sumar-92.327-afiliados-en-el-ultimo-mes-y-739.888-respecto-a-hace-un-ano ${consulta}.`,
    `MINISTERIO DE TRABAJO Y ECONOMÍA SOCIAL (2022): *Estrategia Nacional de Impulso del Trabajo Autónomo (ENDITA) 2022-2027*. Madrid. https://www.mites.gob.es/ficheros/ministerio/inicio_destacados/ENDITA-2022-2027.pdf ${consulta}.`,
    `MINISTERIO DE TRABAJO Y ECONOMÍA SOCIAL (2026a): *Estadística de personas trabajadoras por cuenta propia afiliadas a la Seguridad Social. Marzo 2026* [en línea]. Madrid. https://www.mites.gob.es/estadisticas/AUT/AUT_03_2026.pdf ${consulta}.`,
    `MINISTERIO DE TRABAJO Y ECONOMÍA SOCIAL (2026b): *Informe trimestral de jóvenes y mercado de trabajo*, Nº 49, mayo de 2026 [en línea]. Madrid. https://www.mites.gob.es/ficheros/ministerio/sec_trabajo/analisis_mercado_trabajo/jovenes/2026/Informe_Jovenes_Num49_mayo_2026.pdf ${consulta}.`,
    `MYERS, S. C. y MAJLUF, N. S. (1984): "Corporate financing and investment decisions when firms have information that investors do not have". *Journal of Financial Economics*, Vol. 13, Nº 2, pp. 187-221.`,
    `OBSERVATORIO DEL EMPRENDIMIENTO DE ESPAÑA (2025): *Global Entrepreneurship Monitor. Informe GEM España 2024-2025*. Editorial de la Universidad de Cantabria, Santander. https://www.editorial.unican.es/libro/global-entrepreneurship-monitor-informe-gem-espana-2024-2025 ${consulta}.`,
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
    `ROBB, A. M. y ROBINSON, D. T. (2014): "The capital structure decisions of new firms". *The Review of Financial Studies*, Vol. 27, Nº 1, pp. 153-179.`,
    `SAN-MARTÍN, P., TORRE-OLMO, B., SAMSAMI, M., SCHØTT, T. y TRUEBA-CASTAÑEDA, L. (2025): *El conocimiento financiero en España: brechas y desafíos de una competencia clave para la sociedad*. Cuadernos de divulgación científica, Nº 5/2025. Observatorio del Emprendimiento de España, Madrid. https://observatoriodelemprendimiento.es/wp-content/uploads/El-conocimiento-financiero-en-Espana.pdf ${consulta}.`,
    `SEGURIDAD SOCIAL (2026): *Cotización de trabajadores autónomos: bases y tramos de rendimientos 2026* [en línea]. Madrid. https://www.seg-social.es/wps/portal/wss/internet/Trabajadores/CotizacionRecaudacionTrabajadores/10721/10724/1320/1322/boni9 ${consulta}.`,
    `SHANE, S. y VENKATARAMAN, S. (2000): "The promise of entrepreneurship as a field of research". *Academy of Management Review*, Vol. 25, Nº 1, pp. 217-226.`,
        `TESORO PÚBLICO (2026): *Resultado de la subasta de Obligaciones del Estado a 10 años de 1 de octubre de 2026* [en línea]. Madrid. https://www.tesoro.es/deuda-publica/subastas/resultados-subastas-anteriores/obligaciones-del-estado/subasta-de-394 ${consulta}.`,
    `VAN WESTENDORP, P. (1976): "NSS Price Sensitivity Meter (PSM): A new approach to study consumer perception of prices". *Proceedings of the 29th ESOMAR Congress*, Venecia, pp. 139-167.`,
  ].map(ref),
];

// ------------------------------------------------- ANEXOS
const anexos = [
  h1("Anexo I. Análisis de datos de mercado"),
  p(`Los capítulos 3 y 4 se apoyan en cifras publicadas por el Banco de España, el Ministerio de Trabajo y Economía Social, Eurostat y ATA. Este anexo las reúne para responder a tres preguntas: qué peso tienen los autónomos jóvenes, si tienen peor comprensión financiera que el resto y cómo está evolucionando el colectivo.`),
  p(`**¿Qué peso tienen los autónomos jóvenes?** En marzo de 2026 había 3.448.635 autónomos. Los de 46 a 55 años eran el 31,5 % y los de 56 a 65 el 23,9 % (Ministerio de Trabajo y Economía Social, 2026a), mientras que los de 16 a 29 años, 249.456 personas, solo representaban el 7,2 % (Ministerio de Trabajo y Economía Social, 2026b). El resto, de 30 a 45 años y de 66 o más, se obtiene por diferencia (Figura I.1). Si se toma el corte de 35 años, el informe de ATA sitúa a los menores de esa edad en el 16,1 % del total, unas ${miles(Math.round(tam / 10000) * 10000)} personas.`),
  leyenda("Figura", "I.1", "Distribución de los autónomos por edad en España (marzo de 2026)"),
  imagen("fig_anexo_reta_edad.png", 13),
  fuente("elaboración propia a partir de Ministerio de Trabajo y Economía Social (2026a, 2026b)."),
  p(`**¿Tienen los jóvenes peor comprensión financiera?** El índice que resume los aciertos en las tres preguntas básicas de la ECF (inflación, interés compuesto y diversificación del riesgo) es del 52 % entre los 18 y los 34 años, frente al 56 % del grupo de 45 a 64 años, el que mejor puntúa, y a una media nacional del 53 % (Figura I.2). Las tres preguntas miden conceptos básicos: si se entiende que la inflación reduce el poder de compra del dinero ahorrado, si se sabe calcular de forma aproximada el efecto del interés compuesto sobre un depósito a cinco años y si se reconoce que diversificar entre muchas acciones reduce el riesgo de una inversión. La diferencia, de cuatro puntos, no es abismal, pero es coherente con la hipótesis de partida: los autónomos jóvenes no parten de cero, pero su base es algo más débil justo cuando empiezan a decidir solos.`),
  leyenda("Figura", "I.2", "Aciertos medios en las preguntas financieras básicas por edad (2021)"),
  imagen("fig_anexo_big3.png", 13),
  fuente("elaboración propia a partir de Banco de España (2022)."),
  p(`**¿Cómo evoluciona el colectivo?** El envejecimiento es claro. Entre marzo de 2025 y marzo de 2026, el tramo de 56 a 65 años fue el que más creció en términos absolutos (14.222 autónomos más) y el de 66 o más años el que más lo hizo en términos relativos (+8,9 %), mientras que el de 36 a 45 años perdió 6.889 afiliados (Ministerio de Trabajo y Economía Social, 2026a). La Estrategia Nacional de Impulso del Trabajo Autónomo ya señalaba la franja de 25 a 39 años como la que más peso había perdido (Ministerio de Trabajo y Economía Social, 2022), y España es uno de los países de la Unión Europea con menor proporción de autónomos de 20 a 29 años: el 5,9 %, frente al 7,9 % de media (Eurostat, 2026). La única señal en sentido contrario es el aumento del 5,1 % de los autónomos de 16 a 29 años en el último año (Ministerio de Trabajo y Economía Social, 2026b).`),

  h1("Anexo II. Modelo de cálculo y reproducibilidad"),
  p(`El plan económico-financiero se ha calculado con un modelo mensual de 36 periodos programado en Python. El modelo parte de un único fichero de hipótesis y genera, para cada escenario, la cuenta de resultados, la tesorería con el calendario real del IVA y del IRPF, el balance, los criterios de rentabilidad y el análisis de sensibilidad. Como comprobación de coherencia, el programa verifica que el activo y la suma del patrimonio neto y el pasivo coinciden en cada cierre. El modelo se entrega también como hoja de cálculo (Modelo_financiero_Flujo.xlsx) con fórmulas vivas: desde un panel de control se puede elegir el escenario y activar cada una de las medidas del capítulo 8, y cualquier cambio en las hipótesis recalcula la cuenta de resultados, la tesorería, el balance, el VAN y la TIR. La hoja reproduce los resultados del código hasta el céntimo.`),
  p(`El código, los datos de mercado con la fuente de cada cifra y los gráficos de este trabajo están disponibles en el repositorio https://github.com/Andres-Leach/flujo-tfg. Para reproducir los resultados basta con ejecutar, en este orden, modelo.py, graficos.py y excel_formulas.py (las respuestas de la encuesta se procesan con analizar_encuesta.py), y cualquier hipótesis puede cambiarse en hipotesis.py para recalcular todo el plan.`),
  p(`Todos los porcentajes del Anexo I proceden de fuentes oficiales o institucionales ya citadas en la bibliografía. La aportación propia consiste en reunir, cruzar y representar esos datos, y en construir el modelo financiero; los únicos datos primarios del trabajo proceden de la encuesta del Anexo III.`),

  h1("Anexo III. Encuesta a autónomos jóvenes"),
  p(`**Método.** La encuesta se realizó con Google Forms entre el [[fecha]] y el [[fecha]] y se difundió por [[canales: redes sociales, grupos de autónomos, coworkings, Zaragoza Activa...]]. Es anónima: no recoge nombres, correos ni datos que permitan identificar a nadie, y empieza pidiendo el consentimiento para participar con fines académicos. Se trata de una muestra de conveniencia, por lo que sus resultados describen a quienes respondieron y no pueden extrapolarse al conjunto del colectivo. El cuestionario puede consultarse en https://docs.google.com/forms/d/e/1FAIpQLSecTHUjiV4dEWECRdpd1_hL_OelT9GEsyCB3qt3yq8BkEEn9A/viewform.`),
  p(`**Selección de la muestra.** El análisis principal incluye a las personas de 20 a 35 años que trabajan por cuenta propia como autónomos dados de alta o como *freelancers* sin alta, que son el público de Flujo. Quienes tienen una sociedad se analizan aparte, porque su gestión financiera y fiscal es distinta, y se excluye a quien no trabaja por cuenta propia o está fuera del rango de edad. La Tabla III.1 recoge el proceso.`),
  leyenda("Tabla", "III.1", "Selección de la muestra"),
  tabla(["Paso", "Respuestas"], [
    ["Respuestas recibidas", "[[ ]]"],
    ["Excluidas por no dar el consentimiento", "[[ ]]"],
    ["Excluidas por edad (menos de 20 o más de 35 años)", "[[ ]]"],
    ["Excluidas por no trabajar por cuenta propia", "[[ ]]"],
    ["Socios de sociedades (analizados aparte)", "[[ ]]"],
    ["!Muestra final válida", "[[ ]]"],
    ["  de ellos, autónomos dados de alta / *freelancers* sin alta", "[[ ]] / [[ ]]"],
  ], [6570, 2500], { numericas: false }),
  fuente("elaboración propia a partir de la encuesta."),
  p(`**Cuestionario.** Consta de 22 preguntas en cuatro bloques: perfil; gestión financiera (si sabe cuánto gana neto, si separa cuentas, cómo las lleva, si tiene gestoría, si reserva para impuestos, sorpresas con Hacienda y escalas de capacidad y estrés); interés por un servicio como Flujo y formatos preferidos; y precio. Para el precio se usa el método de Van Westendorp (1976), que pide cuatro precios (tan barato que haría dudar de la calidad, barato, caro pero aceptable y tan caro que no se pagaría) cuyas curvas acumuladas delimitan el rango aceptable, junto con una pregunta directa sobre la intención de contratar a 200 € de puesta en marcha y 75 € al mes.`),
  p(`**Resultados y decisiones.** Los resultados se organizan en los tres niveles de validación del capítulo 3: problema, interés y demanda. Para evitar interpretarlos a conveniencia, los umbrales que confirman cada hipótesis y la decisión que se tomaría si no se alcanzan se fijaron antes de conocer los resultados (Tabla III.2).`),
  leyenda("Tabla", "III.2", "Resultados de la encuesta y su efecto en el plan"),
  tabla(["Nivel e indicador", "Umbral", "Resultado", "Decisión si no se alcanza"], [
    ["!Problema", "", "", ""],
    ["No sabe con exactitud cuánto gana neto", "≥ 50 %", "[[ ]] %", "Reorientar la propuesta hacia la fiscalidad"],
    ["No separa, o solo a medias, las cuentas", "≥ 40 %", "[[ ]] %", "Reducir el peso del *setup*"],
    ["No aparta siempre el dinero de IVA e IRPF", "≥ 40 %", "[[ ]] %", "Revisar la propuesta de valor"],
    ["!Interés", "", "", ""],
    ["Interés alto en el servicio (4 o 5 sobre 5)", "≥ 40 %", "[[ ]] %", "Replantear el servicio antes de lanzarlo"],
    ["Prefiere plantillas o cursos", "≥ 30 %", "[[ ]] %", "Retrasar los productos digitales"],
    ["Prefiere una aplicación", "≥ 30 %", "[[ ]] %", "Mantener la herramienta como opción a largo plazo"],
    ["Le interesan los talleres en grupo", "≥ 20 %", "[[ ]] %", "Revisar la medida 3 del capítulo 8"],
    ["!Demanda", "", "", ""],
    ["Contrataría a 75 €/mes (seguro o probablemente sí)", "≥ 25 %", "[[ ]] %", "Bajar el precio y recalcular el plan"],
    ["Rango de precio aceptable (Van Westendorp)", "Incluye 75 €", "[[ ]] - [[ ]] €", "Fijar el precio en el óptimo y recalcular"],
    ["Precio óptimo (Van Westendorp)", `> ${num(R.criticos_opt.precio_van_oportunidad, 0)} €`, "[[ ]] €", "Por debajo del umbral crítico: el proyecto no compensa"],
  ], [3270, 1300, 1500, 3000], { numericas: false }),
  fuente(`elaboración propia. El umbral de ${num(R.criticos_opt.precio_van_oportunidad, 0)} € es el precio mínimo del plan mejorado (apartado 9.2).`),
  p(`[[Comentario de los resultados: qué hipótesis se confirman, cuáles no y cómo cambia el plan. Si el precio aceptable queda por debajo de 75 €/mes, recalcular el plan con el nuevo precio.]]`),
];

anexos.push(
  h1("Anexo IV. Fuentes oficiales consultadas"),
  p(`La Tabla IV.1 relaciona cada fuente oficial con el dato que se ha contrastado y el apartado del trabajo en que se utiliza. Las direcciones web completas figuran en la bibliografía; todas se consultaron el 28 de octubre de 2026.`),
  leyenda("Tabla", "IV.1", "Fuentes oficiales y datos contrastados"),
  tabla(["Organismo", "Fuente", "Dato contrastado", "Apartado"], [
    ["BOE / Seguridad Social", "Orden PJC/297/2026; tablas de cotización 2026", "Bases mínimas por tramo y tipo total del 31,5 %", "5.2, 6.4"],
    ["Seguridad Social / BOE", "Real Decreto-ley 13/2022", "Tarifa plana de 80 € y prórroga si el rendimiento no supera el SMI", "5.2"],
    ["BOE", "Real Decreto 126/2026", "SMI de 2026: 17.094 € anuales", "5.2, 7.3"],
    ["Tesoro Público", "Subasta de Obligaciones a 10 años de 1/10/2026", "Tipo medio del 4,176 % (rentabilidad sin riesgo)", "7.1"],
    ["INE", "Encuesta de Estructura Salarial 2022", "Ganancia media de 25 a 29 años: 20.460 €", "7.3"],
    ["Ministerio de Inclusión", "Balance de afiliación de septiembre de 2026", "3.483.645 autónomos en el RETA", "1.1, 4.1"],
    ["Ministerio de Trabajo", "Estrategia Nacional de Impulso del Trabajo Autónomo 2022-2027", "Pérdida de peso de la franja de 25 a 39 años", "4.1, Anexo I"],
    ["Agencia Tributaria", "Estimación directa simplificada; cuadro de retenciones 2026", "5 % de gastos de difícil justificación (máx. 2.000 €); retención del 7 %; regla del 70 % del modelo 130", "5.2, 6.5"],
    ["Agencia Tributaria", "Manual de Renta 2025, escala de Aragón", "Tramos autonómicos del IRPF", "6.5"],
    ["Agencia Tributaria", "Tipos del Impuesto sobre Sociedades", "15 % nueva creación; 17 % y 20 % microempresas en 2027", "5.2"],
    ["BOE / Agencia Tributaria", "RD 238/2026, Orden HAC/1028/2026 y calendario", "Factura electrónica obligatoria: 6/10/2028 para autónomos", "4.1, 5.2"],
    ["Agencia Tributaria / Hacienda", "Nota de la AEAT sobre Verifactu y nota de Hacienda de 5/10/2026", "Verifactu el 1/7/2027; aplazamiento anunciado a octubre de 2028", "4.1, 5.2"],
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

anexos.push(
  h1("Anexo V. Obligaciones fiscales y de Seguridad Social"),
  p(`La Tabla V.1 recoge las obligaciones legales del proyecto como empresario individual y su efecto en el plan económico-financiero (ver apartado 5.2).`),
  leyenda("Tabla", "V.1", "Obligaciones fiscales y de Seguridad Social del proyecto"),
  tabla(["Ámbito", "Obligación", "Efecto en el plan financiero"], [
    ["Seguridad Social (RETA)", "Alta en el RETA; tarifa plana de 80 €/mes los primeros 12 meses, prorrogable otros 12 si el rendimiento neto no supera el SMI (17.094 € en 2026). Requiere no haber estado de alta en los dos años anteriores. Después, cuota según tramo: base mínima por un tipo total del 31,5 %", `Año 1: 80 €/mes; año 2: ${eur(R.base.reta[1], 2)}/mes y año 3: ${eur(R.base.reta[2], 2)}/mes según el tramo de rendimientos (en el escenario base no se cumple el requisito de la prórroga)`],
    ["IAE y censo", "Alta censal (modelo 036/037) en un epígrafe de servicios profesionales", "Sin coste"],
    ["IVA", "Facturas con IVA del 21 %; declaraciones trimestrales (modelo 303) y resumen anual (modelo 390)", "El IVA se liquida el mes siguiente a cada trimestre"],
    ["IRPF", "Estimación directa simplificada. Retención del 7 % en facturas a otros profesionales durante el año de inicio y los dos siguientes; pagos fraccionados (modelo 130) del 20 %, salvo que al menos el 70 % de los ingresos lleve retención", "Gastos de difícil justificación del 5 % (máx. 2.000 €); cuota según la escala estatal y la de Aragón"],
    ["Facturación", "Software adaptado a Verifactu desde el 1 de julio de 2027 (aplazamiento anunciado); factura electrónica obligatoria desde el 6 de octubre de 2028", "Incluido en el coste de las herramientas de facturación"],
    ["Protección de datos", "RGPD y LOPDGDD: consentimiento, política de privacidad y contrato de encargo cuando proceda", "Textos legales en la inversión inicial"],
  ], [1700, 4170, 3200], { numericas: false }),
  fuente("elaboración propia a partir del Real Decreto-ley 13/2022, la Orden PJC/297/2026, Seguridad Social (2026), la Ley 35/2006 del IRPF, la Ley 37/1992 del IVA, Agencia Tributaria (2026b, 2026c, 2026d, 2026e, 2026f) y el Real Decreto 126/2026."),
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
        run: { font: FUENTE, size: 28, bold: true }, paragraph: { spacing: { before: 0, after: 120 }, outlineLevel: 0, keepNext: true } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FUENTE, size: 24, bold: true }, paragraph: { spacing: { before: 0, after: 120 }, outlineLevel: 1, keepNext: true } },
      { id: "TOC1", name: "toc 1", basedOn: "Normal", next: "Normal", paragraph: { spacing: { before: 0, after: 120, line: 360, lineRule: "auto" } } },
      { id: "TOC2", name: "toc 2", basedOn: "Normal", next: "Normal", paragraph: { spacing: { before: 0, after: 120, line: 360, lineRule: "auto" }, indent: { left: 284 } } },
      { id: "LeyendaTabla", name: "LeyendaTabla", basedOn: "Normal", next: "Normal", run: { font: FUENTE, size: 24 } },
      { id: "LeyendaFigura", name: "LeyendaFigura", basedOn: "Normal", next: "Normal", run: { font: FUENTE, size: 24 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FUENTE, size: 24, bold: true, italics: true }, paragraph: { spacing: { before: 0, after: 120 }, outlineLevel: 2, keepNext: true } },
    ],
  },
  numbering: { config: [{ reference: "vinetas", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 284 } } } }] }] },
  sections: [
    { properties: { ...pagina }, children: portada },
    { properties: { ...pagina, type: SectionType.NEXT_PAGE }, children: info },
    { properties: { ...pagina, type: SectionType.NEXT_PAGE, page: { ...pagina.page, pageNumbers: { start: 1, formatType: NumberFormat.LOWER_ROMAN } } },
      footers: { default: pie },
      children: indice },
    { properties: { ...pagina, type: SectionType.NEXT_PAGE, page: { ...pagina.page, pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL } } },
      footers: { default: pie },
      children: [...cap1, ...cap2, ...cap3, ...cap4, ...cap5, ...cap6, ...cap7, ...capMejora, ...cap8, ...cap9, ...biblio, ...anexos] },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(__dirname, "..", "TFG_Flujo_v2.docx");
  fs.writeFileSync(out, buf);
  console.log("guardado", out);
});
