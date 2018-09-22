# <a name="outlook-add-in-api-requirement-set-13"></a>Conjunto de requisitos de la API de complementos de Outlook 1.3

El Outlook complemento subconjunto de API de la API de JavaScript para Office incluye objetos, métodos, propiedades y eventos que puede usar en un Outlook complemento.

> [!NOTE]
> Esta documentación es para un [requisito de establece](/javascript/office/requirement-sets/outlook-api-requirement-sets) que no sea el último conjunto de requisitos. 

## <a name="whats-new-in-13"></a>¿Cuáles son las novedades de la versión 1.3?

El conjunto de requisitos 1.3 incluye todas las características del [conjunto de requisitos 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md). Se han agregado las siguientes características.

- Compatibilidad agregada para los [comandos de complemento](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook).
- Capacidad agregada para guardar o cerrar un elemento que se está redactando.
- Objeto [Body](/javascript/api/outlook_1_3/office.body) mejorado para permitir que los complementos obtengan o establezcan el cuerpo completo.
- Métodos de conversión agregados para convertir identificadores entre formatos EWS y REST.
- Capacidad agregada para agregar mensajes de notificación a la barra de información de los elementos.

### <a name="change-log"></a>Registro de cambios

- [Body.getAsync](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) agregado: Devuelve el cuerpo actual en un formato especificado.
- [Body.setAsync](/javascript/api/outlook_1_3/office.body#setasync-data--options--callback-) agregado: Reemplaza todo el cuerpo por el texto especificado.
- [Office.context.officeTheme](office.context.md#officetheme-object) agregado: Proporciona acceso a los colores del tema de Office.
- Objeto [Event](/javascript/api/office/office.addincommands.event) agregado: Se pasa como un parámetro a las funciones de comandos de forma directa en un complemento de Outlook. Se usa para indicar la finalización del proceso.
- [Office.context.mailbox.item.close](office.context.mailbox.item.md#close) agregado: Cierra el elemento actual que se está redactando.
- [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#saveasyncoptions-callback) agregado: Guarda un elemento de manera asincrónica.
- [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages) agregado: Obtiene los mensajes de notificación de un elemento.
- [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#converttoewsiditemid-restversion--string) agregado: Convierte un identificador de elemento con formato para REST al formato EWS.
- [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) agregado: Convierte un identificador de elemento con formato para EWS al formato REST.
- [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook_1_3/office.mailboxenums.itemnotificationmessagetype) agregado: Especifica el tipo de mensaje de notificación de una cita o un mensaje.
- [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_3/office.mailboxenums.restversion) agregado: Especifica la versión de la API de REST que corresponde a un identificador de elemento con formato REST.
- Objeto [NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages) agregado: Proporciona métodos para tener acceso a mensajes de notificación en un complemento de Outlook.
- Tipo [NotificationMessageDetails](/javascript/api/outlook_1_3/office.notificationmessagedetails) agregado: Devuelto por el método `NotificationMessages.getAllAsync`.

## <a name="see-also"></a>Vea también

- [Complementos de Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Ejemplos de código de complementos de Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introducción](https://docs.microsoft.com/outlook/add-ins/quick-start)