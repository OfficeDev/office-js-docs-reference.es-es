# <a name="outlook-add-in-api-requirement-set-15"></a>Conjunto de requisitos de la API de complementos de Outlook 1.5

El subconjunto de la API de complementos de Outlook de la API de JavaScript para Office incluye objetos, métodos, propiedades y eventos que se pueden usar en un complemento de Outlook.

> [!NOTE]
> Esta documentación es para un [requisito de establece](/javascript/office/requirement-sets/outlook-api-requirement-sets) que no sea el último conjunto de requisitos.

## <a name="whats-new-in-15"></a>¿Cuáles son las novedades de la versión 1.5?

El conjunto de requisitos 1.5 incluye todas las características del [conjunto de requisitos 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). Se han agregado las siguientes características.

- Se ha agregado compatibilidad para los [paneles de tareas anclables](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).
- Se ha agregado compatibilidad para llamar a las [API de REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api).
- Se ha agregado la capacidad de marcar datos adjuntos como insertados.
- Se ha agregado la capacidad de cerrar un panel de tareas o un cuadro de diálogo.

### <a name="change-log"></a>Registro de cambios

- [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback) agregado: Agrega un controlador de eventos para un evento admitido.
- Agregado [Office.EventType](office.md#eventtype-string): especifica el evento asociado a un controlador de eventos e incluye compatibilidad para el evento ItemChanged.
- [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string) agregado: Obtiene la URL del punto de conexión REST para esta cuenta de correo electrónico.
- [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback) modificado: Se ha agregado una nueva versión de este método con una firma nueva (`getCallbackTokenAsync([options], callback)`). La versión original todavía está disponible y es invariable.
- Se ha agregado [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).
- [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback) modificado: Un valor nuevo en el diccionario `options` denominado `isInline`, usado para especificar que una imagen se usa en línea en el cuerpo del mensaje.
- [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata) modificado: Un valor nuevo en el diccionario `formData.attachments` denominado `isInline`, usado para especificar que una imagen se usa en línea en el cuerpo del mensaje.
- [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata) modificado: Un valor nuevo en el diccionario `formData.attachments` denominado `isInline`, usado para especificar que una imagen se usa en línea en el cuerpo del mensaje.

## <a name="see-also"></a>Vea también

- [Complementos de Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Ejemplos de código de complementos de Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introducción](https://docs.microsoft.com/outlook/add-ins/quick-start)