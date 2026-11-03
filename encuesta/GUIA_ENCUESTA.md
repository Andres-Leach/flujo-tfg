# Guía de la encuesta del TFG

## 1. Crear el formulario (5 minutos)
1. Entra en https://script.google.com con tu cuenta de Google y pulsa **Nuevo proyecto**.
2. Borra el contenido, pega todo `crear_formulario.gs` y guarda.
3. Arriba, elige la función `crearFormulario` y pulsa **Ejecutar**. Acepta los permisos (Google avisa de que la app "no está verificada": es tu propio script; pulsa *Configuración avanzada* > *Ir a…*).
4. En el **Registro de ejecución** salen dos enlaces: el de edición y el de compartir. El formulario queda en tu Drive.
5. Ábrelo, revísalo en vista previa (icono del ojo) y haz tú una respuesta de prueba. Después bórrala en *Respuestas*.

## 2. Objetivo de muestra
- **Mínimo: 30 respuestas válidas** (20-35 años y por cuenta propia). **Ideal: 50 o más.**
- Cuenta con que una parte se filtra (edad o situación), así que necesitarás que la abran unas 60-80 personas.

## 3. Difusión (2-3 semanas)
- Tus contactos: compañeros de clase que ya trabajen como freelance, amigos y conocidos, con la petición de que la reenvíen.
- Instagram y LinkedIn: publicación y stories con el enlace. Explica que es para tu TFG en Unizar y que tarda 5 minutos.
- Grupos y comunidades: grupos de WhatsApp o Telegram de freelancers, comunidades de diseñadores, programadores y creadores de contenido, foros de autónomos.
- Zaragoza: Zaragoza Activa (Semillero de Ideas), coworkings, CIPAJ, asociaciones de autónomos (UPTA, ATA Aragón) y el servicio de emprendimiento de Unizar. Pide permiso antes de difundir en sus canales.
- Recuerda cada 4-5 días a quien te dijo que la compartiría.

Mensaje de ejemplo:
> Hola, soy Andrés y estoy haciendo mi TFG de ADE en la Universidad de Zaragoza sobre cómo gestionan su dinero los autónomos y freelancers jóvenes. Si tienes entre 20 y 35 años y trabajas por cuenta propia, me ayudaría mucho que respondieras esta encuesta anónima (5 minutos): [enlace]. ¡Y si puedes compartirla, mejor todavía! Gracias.

## 4. Reglas para que los datos valgan
- No respondas tú ni pidas a nadie que responda "por ayudar" si no cumple el perfil: los filtros están para eso.
- No modifiques respuestas ni añadas filas a mano.
- Guarda el CSV original tal cual lo descargas.

## 5. Analizar y meterlo en el TFG
1. En el formulario: *Respuestas* > *Vincular con Hojas de cálculo*. En la hoja: *Archivo* > *Descargar* > *CSV*.
2. Pásame el CSV (o ejecuta `python analizar_encuesta.py respuestas.csv`).
3. Con los resultados relleno los huecos en amarillo de la introducción, el capítulo 3, las conclusiones y el Anexo III, y añado la figura del rango de precios.
4. Si el rango de precio aceptable sale por debajo de 75 €/mes, revisaremos el precio del plan y su efecto en el VAN.

## Calendario propuesto
| Semana | Tarea |
|---|---|
| 3-9 noviembre | Crear el formulario, probarlo y lanzarlo |
| 10-23 noviembre | Difusión y recordatorios |
| 24-30 noviembre | Cerrar la encuesta, analizar y actualizar el TFG |
| Diciembre | Enviar a Beatrice junto al entregable 3 |
