/**
 * Crea en tu Google Drive el formulario de la encuesta del TFG "Flujo".
 *
 * Cómo usarlo:
 *   1. Entra en https://script.google.com con tu cuenta de Google y pulsa "Nuevo proyecto".
 *   2. Borra lo que aparece, pega todo este archivo y guarda (icono del disquete).
 *   3. Elige la función "crearFormulario" arriba y pulsa "Ejecutar". La primera vez
 *      Google te pedirá permiso para crear formularios en tu cuenta: acéptalo.
 *   4. En "Registro de ejecución" aparecen dos enlaces: el de edición y el que
 *      tienes que compartir. El formulario también queda guardado en tu Drive.
 *
 * Las preguntas llevan un prefijo (P1, P2...) para que el script de análisis
 * (analizar_encuesta.py) las reconozca en el CSV de respuestas.
 */
function crearFormulario() {
  const form = FormApp.create('Encuesta: los autónomos jóvenes y sus finanzas (TFG Universidad de Zaragoza)');
  form.setDescription(
    'Hola. Soy Andrés Leach, estudiante de ADE en la Universidad de Zaragoza, y estoy haciendo mi ' +
    'Trabajo Fin de Grado sobre cómo gestionan su dinero los autónomos y freelancers jóvenes.\n\n' +
    'La encuesta dura unos 5 minutos y es anónima: no se piden nombres, correos ni datos que te ' +
    'identifiquen. Las respuestas solo se usarán de forma agregada en el TFG.\n\n' +
    'Va dirigida a personas de 20 a 35 años que trabajan por cuenta propia, con o sin alta como autónomo.');
  form.setCollectEmail(false);
  form.setProgressBar(true);
  form.setShowLinkToRespondAgain(false);
  form.setConfirmationMessage('¡Gracias por tu tiempo! Tus respuestas me ayudan mucho con el TFG.');

  const NUM = (min, max) => FormApp.createTextValidation()
    .setHelpText('Escribe solo un número (en euros, sin decimales).')
    .requireNumberBetween(min, max).build();
  const SEGUIR = FormApp.PageNavigationType.CONTINUE;
  const FIN = FormApp.PageNavigationType.SUBMIT;

  // --- Sección 1: consentimiento -------------------------------------------
  const consentimiento = form.addMultipleChoiceItem()
    .setTitle('P0. ¿Aceptas participar en esta encuesta anónima con fines académicos?')
    .setRequired(true);
  consentimiento.setChoices([
    consentimiento.createChoice('Sí, acepto', SEGUIR),
    consentimiento.createChoice('No', FIN),
  ]);

  // --- Sección 2: edad -------------------------------------------------------
  form.addPageBreakItem().setTitle('Sobre ti');
  const edad = form.addMultipleChoiceItem().setTitle('P1. ¿Qué edad tienes?').setRequired(true);
  edad.setChoices([
    edad.createChoice('Menos de 20 años', FIN),
    edad.createChoice('De 20 a 24 años', SEGUIR),
    edad.createChoice('De 25 a 29 años', SEGUIR),
    edad.createChoice('De 30 a 35 años', SEGUIR),
    edad.createChoice('Más de 35 años', FIN),
  ]);

  // --- Sección 3: situación laboral -----------------------------------------
  form.addPageBreakItem().setTitle('Tu actividad');
  const situacion = form.addMultipleChoiceItem()
    .setTitle('P2. ¿Cuál es tu situación?').setRequired(true);
  situacion.setChoices([
    situacion.createChoice('Estoy dado/a de alta como autónomo/a', SEGUIR),
    situacion.createChoice('Trabajo como freelance pero sin alta como autónomo/a', SEGUIR),
    situacion.createChoice('Tengo una sociedad (SL u otra)', SEGUIR),
    situacion.createChoice('No trabajo por cuenta propia', FIN),
  ]);

  // --- Sección 4: perfil y finanzas -----------------------------------------
  form.addPageBreakItem().setTitle('Tu actividad y tus finanzas');
  form.addMultipleChoiceItem().setTitle('P3. ¿A qué te dedicas principalmente?').setRequired(true)
    .setChoiceValues(['Diseño, ilustración, fotografía o vídeo', 'Marketing, redes sociales o creación de contenido',
      'Programación, desarrollo web o informática', 'Deporte, formación o coaching',
      'Consultoría o servicios profesionales', 'Estética, belleza o salud', 'Comercio u hostelería'])
    .showOtherOption(true);
  form.addMultipleChoiceItem().setTitle('P4. ¿Cuánto tiempo llevas trabajando por cuenta propia?').setRequired(true)
    .setChoiceValues(['Menos de 1 año', 'De 1 a 2 años', 'De 3 a 5 años', 'Más de 5 años']);
  form.addMultipleChoiceItem().setTitle('P5. ¿Cuánto facturas al mes, aproximadamente, sin IVA?').setRequired(false)
    .setChoiceValues(['Menos de 1.000 €', 'De 1.000 a 2.000 €', 'De 2.000 a 3.000 €', 'Más de 3.000 €',
      'Prefiero no contestar']);
  form.addMultipleChoiceItem().setTitle('P6. ¿Dónde trabajas?').setRequired(true)
    .setChoiceValues(['Zaragoza', 'Otra provincia de Aragón', 'Otra comunidad autónoma']);
  form.addMultipleChoiceItem()
    .setTitle('P7. Si tuvieras que decir cuánto ganaste el mes pasado después de impuestos y gastos, ¿sabrías hacerlo?')
    .setRequired(true)
    .setChoiceValues(['Sí, con exactitud', 'Más o menos', 'No']);
  form.addMultipleChoiceItem().setTitle('P8. ¿Separas el dinero de tu actividad del dinero personal (cuentas distintas)?')
    .setRequired(true).setChoiceValues(['Sí', 'A medias', 'No']);
  form.addCheckboxItem().setTitle('P9. ¿Cómo llevas tus cuentas? (puedes marcar varias)').setRequired(true)
    .setChoiceValues(['No llevo ningún control', 'Hoja de cálculo (Excel, Google Sheets)',
      'Aplicación de gastos o del banco', 'Programa de facturación', 'Me lo lleva una gestoría o asesoría']);
  form.addMultipleChoiceItem().setTitle('P10. ¿Trabajas con una gestoría o asesoría? ¿Cuánto pagas al mes?')
    .setRequired(true)
    .setChoiceValues(['No tengo gestoría', 'Sí, menos de 40 € al mes', 'Sí, entre 40 y 80 € al mes',
      'Sí, más de 80 € al mes', 'Sí, pero no sé cuánto pago']);
  form.addMultipleChoiceItem().setTitle('P11. ¿Apartas cada mes el dinero que tendrás que pagar de IVA e IRPF?')
    .setRequired(true).setChoiceValues(['Sí, siempre', 'A veces', 'No']);
  form.addMultipleChoiceItem()
    .setTitle('P12. En el último año, ¿te has llevado alguna sorpresa negativa con Hacienda o la Seguridad Social (un pago mayor del esperado, un recargo...)?')
    .setRequired(true).setChoiceValues(['Sí', 'No', 'No lo sé']);
  form.addScaleItem().setTitle('P13. ¿Hasta qué punto te sientes capaz de gestionar las finanzas de tu actividad?')
    .setBounds(1, 5).setLabels('Nada capaz', 'Totalmente capaz').setRequired(true);
  form.addScaleItem().setTitle('P14. ¿Cuánto estrés te generan las finanzas de tu actividad?')
    .setBounds(1, 5).setLabels('Ninguno', 'Muchísimo').setRequired(true);

  // --- Sección 5: el servicio y el precio -----------------------------------
  form.addPageBreakItem().setTitle('Un servicio para ordenar tus finanzas')
    .setHelpText('Imagina un servicio que, en lenguaje sencillo y sin tecnicismos, te monta un sistema para ' +
      'controlar tus ingresos, gastos y lo que debes reservar para impuestos, y después revisa tus números ' +
      'contigo una vez al mes. No sustituye a la gestoría: te ayuda a entender tu dinero y a decidir.');
  form.addScaleItem().setTitle('P15. ¿Cuánto te interesaría un servicio así?')
    .setBounds(1, 5).setLabels('Nada', 'Muchísimo').setRequired(true);
  form.addCheckboxItem().setTitle('P16. ¿Qué formato te resultaría más útil? (puedes marcar varios)').setRequired(true)
    .setChoiceValues(['Una sesión mensual personalizada', 'Plantillas para llevar las cuentas',
      'Cursos cortos con ejemplos prácticos', 'Una aplicación con avisos y previsiones',
      'Talleres en grupo con otros autónomos', 'Ninguno me interesa']);
  form.addTextItem().setTitle('P17. ¿A partir de qué precio al mes te parecería tan barato que dudarías de su calidad?')
    .setRequired(true).setValidation(NUM(0, 1000));
  form.addTextItem().setTitle('P18. ¿Qué precio al mes te parecería barato, una buena compra?')
    .setRequired(true).setValidation(NUM(0, 1000));
  form.addTextItem().setTitle('P19. ¿Qué precio al mes te parecería caro, pero aún te plantearías pagarlo?')
    .setRequired(true).setValidation(NUM(0, 1000));
  form.addTextItem().setTitle('P20. ¿A partir de qué precio al mes te parecería tan caro que no lo contratarías?')
    .setRequired(true).setValidation(NUM(0, 1000));
  form.addMultipleChoiceItem()
    .setTitle('P21. Si el servicio costara 200 € de puesta en marcha y 75 € al mes (sin IVA), ¿lo contratarías?')
    .setRequired(true)
    .setChoiceValues(['Seguro que sí', 'Probablemente sí', 'No lo sé', 'Probablemente no', 'Seguro que no']);
  form.addParagraphTextItem().setTitle('P22. ¿Cuál es tu mayor problema o duda con el dinero de tu actividad? (opcional)')
    .setRequired(false);

  Logger.log('Enlace para editar: ' + form.getEditUrl());
  Logger.log('Enlace para compartir: ' + form.getPublishedUrl());
}
