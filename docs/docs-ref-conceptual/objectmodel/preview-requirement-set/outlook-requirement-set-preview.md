# <a name="outlook-add-in-api-preview-requirement-set"></a>Conjunto de requisitos de la versión preliminar de la API de complementos de Outlook

El Outlook complemento subconjunto de API de la API de JavaScript para Office incluye objetos, métodos, propiedades y eventos que puede usar en un Outlook complemento.

> [!NOTE]
> Esta documentación es para una **vista previa de** [requisito establecido](/javascript/office/requirement-sets/outlook-api-requirement-sets). Este conjunto de requisitos no está completamente implementado todavía, y los clientes no notificarán compatibilidad para este de manera precisa. No debe especificar este conjunto de requisitos en su manifiesto de complemento. Los métodos y las propiedades que se presentan en este conjunto de requisitos deben probarse individualmente para su disponibilidad antes de usarlos.

El conjunto de requisitos de vista previa incluye todas las características del [conjunto de requisito 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## <a name="features-in-preview"></a>Características de la versión preliminar

Las características siguientes están en versión preliminar.

- [SharedProperties](/javascript/api/outlook/office.sharedproperties) - agrega un nuevo objeto que representa las propiedades de un elemento de cita o un mensaje en una carpeta compartida, calendario o buzón de correo.
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-): un nuevo parámetro `options` opcional, que es un diccionario con un valor `allowEvent` aceptado. Este valor se usa para cancelar la ejecución de un evento.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - agrega un nuevo método que se adjunta un archivo de la base64 codificación a un mensaje o cita.
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback): se ha agregado una nueva función que devuelve los datos de inicialización que se pasan cuando [un mensaje accionable activa](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message) el complemento.
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback) - agrega un nuevo método que obtiene un objeto que representa la sharedProperties de una cita o un elemento de mensaje.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference): se ha agregado el acceso a `getAccessTokenAsync`, que permite que los complementos [accedan al token](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) de la API de Microsoft Graph.
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) - agrega una nueva enumeración de indicador de bits que especifica los permisos de delegado.
- [Office.EventType](/javascript/api/office/office.eventtype) - modificado para admitir eventos OfficeThemeChanged a través de adición de `OfficeThemeChanged` entrada.

## <a name="see-also"></a>Vea también

- [Complementos de Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Ejemplos de código de complementos de Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introducción](https://docs.microsoft.com/outlook/add-ins/quick-start)