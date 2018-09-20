# <a name="outlook-add-in-api-preview-requirement-set"></a>Conjunto de requisitos de la versión preliminar de la API de complementos de Outlook

El subconjunto de la API de complementos de Outlook de la API de JavaScript para Office incluye objetos, métodos, propiedades y eventos que se pueden usar en un complemento de Outlook.

> [!NOTE]
> Esta documentación es para una **vista previa de** [requisito establecido](/javascript/office/requirement-sets/outlook-api-requirement-sets). Este conjunto de requisitos no está completamente implementado todavía, y los clientes no notificarán compatibilidad para este de manera precisa. No debe especificar este conjunto de requisitos en su manifiesto de complemento. Los métodos y las propiedades que se presentan en este conjunto de requisitos deben probarse individualmente para su disponibilidad antes de usarlos.

El conjunto de requisitos de vista previa incluye todas las características del [conjunto de requisito 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md).

## <a name="features-in-preview"></a>Características de la versión preliminar

Las características siguientes están en versión preliminar.

- [De](/javascript/api/outlook/office.from) - agrega un nuevo objeto que proporciona un método para obtener el valor de.
- [Organizador](/javascript/api/outlook/office.organizer) - agrega un nuevo objeto que proporciona un método para obtener el valor del organizador.
- [Periodicidad](/javascript/api/outlook/office.recurrence) - agrega un nuevo objeto que proporciona métodos para obtener y establecer el patrón de periodicidad de citas, pero sólo obtener el patrón de periodicidad de los mensajes que son las convocatorias de reunión.
- [SeriesTime](/javascript/api/outlook/office.seriestime) - agrega un nuevo objeto que proporciona métodos para obtener y establecer las fechas y horas de citas en una serie periódica y para obtener las fechas y horas de convocatorias de reunión en una serie periódica.
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-): un nuevo parámetro `options` opcional, que es un diccionario con un valor `allowEvent` aceptado. Este valor se usa para cancelar la ejecución de un evento.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - agrega un nuevo método que se adjunta un archivo de la base64 codificación a un mensaje o cita.
- [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback) - agrega un nuevo método que agrega un controlador de eventos para un evento compatible.
- [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) - modificado para obtener el valor en modo de redacción.
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback): se ha agregado una nueva función que devuelve los datos de inicialización que se pasan cuando [un mensaje accionable activa](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message) el complemento.
- [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) - modificada para obtener el valor del organizador en modo de redacción
- [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) - agrega una nueva propiedad que obtiene o establece un objeto que proporciona métodos para administrar el patrón de periodicidad de un elemento de cita. Esta propiedad también puede utilizarse para obtener el patrón de periodicidad de una reunión de elemento de solicitud.
- [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback) - agrega un método nuevo que quita un controlador de eventos.
- [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string) - agrega una nueva propiedad que obtiene el identificador de la serie huérfana pertenece a.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference): se ha agregado el acceso a `getAccessTokenAsync`, que permite que los complementos [accedan al token](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) de la API de Microsoft Graph.
- [Office.MailboxEnums.Days](/javascript/api/outlook/office.mailboxenums.days) - agrega una nueva enumeración que especifica el día de la semana o tipo de día.
- [Office.MailboxEnums.Month](/javascript/api/outlook/office.mailboxenums.month) - agrega una nueva enumeración que especifica el mes.
- [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook/office.mailboxenums.recurrencetimezone) - agrega una nueva enumeración que especifica la zona horaria que se aplica a la periodicidad.
- [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook/office.mailboxenums.recurrencetype) - agrega una nueva enumeración que especifica el tipo de periodicidad.
- [Office.MailboxEnums.WeekNumber](/javascript/api/outlook/office.mailboxenums.weeknumber) - agrega una nueva enumeración que especifica la semana del mes.
- [Office.EventType](/javascript/api/office/office.eventtype) - modificado para admitir eventos RecurrenceChanged, RecipientsChanged, AppointmentTimeChanged y OfficeThemeChanged a través de adición de `RecurrencePatternChanged`, `RecipientsChanged`, `AppointmentTimeChanged`, y `OfficeThemeChanged` entradas respectivamente.

## <a name="see-also"></a>Vea también

- [Complementos de Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Ejemplos de código de complementos de Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introducción](https://docs.microsoft.com/outlook/add-ins/quick-start)