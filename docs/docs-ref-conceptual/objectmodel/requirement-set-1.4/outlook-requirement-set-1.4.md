# <a name="outlook-add-in-api-requirement-set-14"></a>Conjunto de requisitos de la API de complementos de Outlook 1.4

El Outlook complemento subconjunto de API de la API de JavaScript para Office incluye objetos, métodos, propiedades y eventos que puede usar en un Outlook complemento.

> [!NOTE]
> Esta documentación es para un [requisito de establece](/javascript/office/requirement-sets/outlook-api-requirement-sets) que no sea el último conjunto de requisitos.

## <a name="whats-new-in-14"></a>¿Cuáles son las novedades de la versión 1.4?

El conjunto de requisitos 1.4 incluye todas las características del [conjunto de requisitos 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). Se ha agregado acceso al espacio de nombres `Office.ui`.

### <a name="change-log"></a>Registro de cambios

- [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) agregado: Muestra un cuadro de diálogo en un host de Office.
- [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-) agregado: Entrega un mensaje desde el cuadro de diálogo a su página primaria o de apertura.
- Se ha agregado un objeto de [cuadro de diálogo](/javascript/api/office/office.dialog) : el objeto que se devuelve cuando el [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) se llama al método.

## <a name="see-also"></a>Vea también

- [Complementos de Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Ejemplos de código de complementos de Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introducción](https://docs.microsoft.com/outlook/add-ins/quick-start)