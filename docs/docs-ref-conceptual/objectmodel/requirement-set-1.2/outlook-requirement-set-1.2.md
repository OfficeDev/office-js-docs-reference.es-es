# <a name="outlook-add-in-api-requirement-set-12"></a>Conjunto de requisitos de la API de complementos de Outlook 1.2

El subconjunto de la API de complementos de Outlook de la API de JavaScript para Office incluye objetos, métodos, propiedades y eventos que se pueden usar en un complemento de Outlook.

> [!NOTE]
> Esta documentación es para un [requisito de establece](/javascript/office/requirement-sets/outlook-api-requirement-sets) que no sea el último conjunto de requisitos. 

## <a name="whats-new-in-12"></a>¿Cuáles son las novedades de la versión 1.2?

El conjunto de requisitos 1.2 incluye todas las características del [conjunto de requisitos 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Se ha agregado la capacidad de que los complementos inserten texto en el cursor del usuario, en el asunto o en el cuerpo del mensaje.

### <a name="change-log"></a>Registro de cambios

- Agregado [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): forma asincrónica devuelve datos seleccionados desde el asunto o el cuerpo de un mensaje.
- [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback) agregado: Inserta asincrónicamente datos en el cuerpo o el asunto de un mensaje.
- [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata) modificado: Propiedad `attachments` agregada al parámetro `formData`.
- [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata) modificado: Propiedad `attachments` agregada al parámetro `formData`.

## <a name="see-also"></a>Vea también

- [Complementos de Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Ejemplos de código de complementos de Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introducción](https://docs.microsoft.com/outlook/add-ins/quick-start)