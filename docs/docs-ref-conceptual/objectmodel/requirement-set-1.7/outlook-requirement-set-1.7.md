# <a name="outlook-add-in-api-requirement-set-17"></a>Requisito de Outlook en API establece 1.7

El subconjunto de la API de complementos de Outlook de la API de JavaScript para Office incluye objetos, métodos, propiedades y eventos que se pueden usar en un complemento de Outlook.

## <a name="whats-new-in-17"></a>¿Qué es una novedad de 1.7?

Conjunto de requisitos 1.7 incluye todas las características del [conjunto de requisito 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md). Agregan las siguientes características.

- Se agregaron nuevas API relacionadas con el patrón de periodicidad en los mensajes que son las convocatorias de reunión y citas.
- Puede modificar la propiedad item.from para que también esté disponible en modo de redacción.
- Se agregó compatibilidad para eventos RecurrenceChanged, RecipientsChanged y AppointmentTimeChanged.

### <a name="change-log"></a>Registro de cambios

- Se agregó [desde](/javascript/api/outlook_1_7/office.from): agrega un nuevo objeto que proporciona un método para obtener el valor de.
- Se agregó el [Organizador](/javascript/api/outlook_1_7/office.organizer): agrega un nuevo objeto que proporciona un método para obtener el valor del organizador.
- Se agregó la [Periodicidad](/javascript/api/outlook_1_7/office.recurrence): agrega un nuevo objeto que proporciona métodos para obtener y establecer el patrón de periodicidad de citas, pero sólo obtener el patrón de periodicidad de los mensajes que son las convocatorias de reunión.
- Agregado [RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone): agrega un nuevo objeto que representa la configuración de zona horaria del patrón de periodicidad.
- Agregado [SeriesTime](/javascript/api/outlook_1_7/office.seriestime): agrega un nuevo objeto que proporciona métodos para obtener y establecer las fechas y horas de citas en una serie periódica y para obtener las fechas y horas de convocatorias de reunión en una serie periódica.
- Agregado [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback): agrega un nuevo método que agrega un controlador de eventos para un evento compatible.
- Modificar [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom): modifica para obtener el valor en modo de redacción.
- Modificado [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) - modifica para obtener el valor de organizador en modo de redacción.
- Agregado [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence): agrega una nueva propiedad que obtiene o establece un objeto que proporciona métodos para administrar el patrón de periodicidad de un elemento de cita. Esta propiedad también puede utilizarse para obtener el patrón de periodicidad de una reunión de elemento de solicitud.
- Agregado [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback): agrega un nuevo método que quita un controlador de eventos.
- Agregado [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string): agrega una nueva propiedad que obtiene el identificador de la serie huérfana pertenece a.
- Agregado [Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days): agrega una nueva enumeración que especifica el día de la semana o tipo de día.
- Agregado [Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month): agrega una nueva enumeración que especifica el mes.
- Agregado [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone): agrega una nueva enumeración que especifica la zona horaria que se aplica a la periodicidad.
- Agregado [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype): agrega una nueva enumeración que especifica el tipo de periodicidad.
- Agregado [Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber): agrega una nueva enumeración que especifica la semana del mes.
- Modificar [Office.EventType](/javascript/api/office/office.eventtype): modifica para admitir eventos RecurrenceChanged, RecipientsChanged y AppointmentTimeChanged a través de adición de `RecurrenceChanged`, `RecipientsChanged`, y `AppointmentTimeChanged` entradas respectivamente.

## <a name="see-also"></a>Vea también

- [Complementos de Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Ejemplos de código de complementos de Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introducción](https://docs.microsoft.com/outlook/add-ins/quick-start)