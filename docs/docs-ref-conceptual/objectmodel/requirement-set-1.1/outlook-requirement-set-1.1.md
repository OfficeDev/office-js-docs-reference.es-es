# <a name="outlook-add-in-api-requirement-set-11"></a>Conjunto de requisitos de la API de complementos de Outlook 1.1

El subconjunto de la API de complementos de Outlook de la API de JavaScript para Office incluye objetos, métodos, propiedades y eventos que se pueden usar en un complemento de Outlook.

> [!NOTE]
> Esta documentación es para un [requisito de establece](/javascript/office/requirement-sets/outlook-api-requirement-sets) que no sea el último conjunto de requisitos. 

## <a name="whats-new-in-11"></a>¿Cuáles son las novedades de la versión 1.1?

El conjunto de requisitos 1.1 incluye todas las características del conjunto de requisitos 1.0. Se ha agregado la capacidad de que los complementos tengan acceso al cuerpo de los mensajes y a las citas, y la capacidad de modificar el elemento actual.

### <a name="change-log"></a>Registro de cambios

- Objeto [Body](/javascript/api/outlook_1_1/office.body) agregado: Proporciona métodos para agregar y actualizar el contenido de un elemento en un complemento de Outlook.
- Objeto [Location](/javascript/api/outlook_1_1/office.location) agregado: Proporciona métodos para obtener y establecer la ubicación de una reunión en un complemento de Outlook.
- Objeto [Recipients](/javascript/api/outlook_1_1/office.recipients) agregado: Proporciona métodos para obtener y establecer los destinatarios de una cita o un mensaje en un complemento de Outlook.
- Objeto [Subject](/javascript/api/outlook_1_1/office.subject) agregado: Proporciona métodos para obtener y establecer el asunto de una cita o un mensaje en un complemento de Outlook.
- Objeto [Time](/javascript/api/outlook_1_1/office.time) agregado: Proporciona métodos para obtener y establecer la hora de inicio y de finalización de una reunión en un complemento de Outlook.
- [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback) agregado: Agrega un archivo a un mensaje o cita como datos adjuntos.
- [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback) agregado: Agrega un elemento de Exchange (por ejemplo, un mensaje) como datos adjuntos al mensaje o a la cita.
- [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) agregado: Quita los datos adjuntos de un mensaje o cita.
- [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-bodyjavascriptapioutlook11officebody) agregado: Obtiene un objeto que proporciona métodos para manipular el cuerpo de un elemento.
- [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#bcc-recipientsjavascriptapioutlook11officerecipients) agregado: Obtiene o establece los destinatarios en la línea CCO (copia carbón oculta) de un mensaje.
- [Office.MailboxEnums.RecipientType](/javascript/api/outlook_1_1/office.mailboxenums.recipienttype) agregado: Especifica el tipo de destinatario de una cita.

## <a name="see-also"></a>Vea también

- [Complementos de Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Ejemplos de código de complementos de Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introducción](https://docs.microsoft.com/outlook/add-ins/quick-start)