# <a name="whats-changed-in-the-javascript-api-for-office"></a>Novedades en la API de JavaScript para Office

Con el fin de ampliar la funcionalidad de sus Complementos de Office, la API de JavaScript para Office se actualiza periódicamente con objetos, métodos, propiedades, eventos y enumeraciones nuevos y actualizados. Siga los vínculos siguientes para ver los miembros de la API nuevos y actualizados.

Para desarrollar complementos con nuevos miembros de API, necesita [actualizar los archivos de la API de JavaScript para Office en el proyecto](https://docs.microsoft.com/office/dev/add-ins/develop/update-your-javascript-api-for-office-and-manifest-schema-version).

Para ver todos los miembros de la API incluidos los que no han cambiado desde actualizaciones anteriores, consulte [API de JavaScript para Office](javascript-api-for-office.md).

## <a name="new-and-updated-apis"></a>API nuevas y actualizadas

### <a name="new-and-updated-objects"></a>Objetos nuevos y actualizados

|**Objeto**|**Descripción**|**Versión agregada o actualizada**|
|:-----|:-----|:-----|
|`Item`|Actualizaciones y adiciones para:<br><ul><li><p>El `getSelectedDataAsync` y `setSelectedDataAsync` métodos para admitir la obtención de la selección del usuario y sobrescribir en el asunto y el cuerpo de un mensaje o cita.</p></li><li><p>El `displayReplyAllForm` y `displayReplyForm` métodos para admitir la adición de datos adjuntos en el formulario de respuesta de una cita.</p></li></ul>|Mailbox 1.2|
|`Item`|Se actualizó para incluir métodos y campos para la creación de complementos de Outlook en modo de redacción. |1.1|
|`Binding`|Se actualizó para admitir enlaces de tablas en complementos de contenido para Access.|1.1|
|`Bindings`|Se actualizó para admitir enlaces de tablas en complementos de contenido para Access.|1.1|
|`Body`|Se agregó para habilitar la creación y edición del cuerpo de un mensaje o una cita en complementos de Outlook en modo de redacción.|1.1|
|`Document`|Actualizaciones y adiciones para: <ul><li><p>Admitir las propiedades <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js" target="_blank">mode</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#settings" target="_blank">settings</a> y <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js" target="_blank">url</a> en complementos de contenido para Access.</p></li><li><p>Obtener el documento como PDF con el método <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfileasync-filetype--options--callback-" target="_blank">getFileAsync</a> en complementos para PowerPoint y Word.</p></li><li><p>Obtener propiedades de archivo con el método <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-" target="_blank">getFileProperties</a> en complementos para Excel, PowerPoint y Word.</p></li><li><p>Navegar a ubicaciones y objetos dentro del documento con el método <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-" target="_blank">goToByIdAsync</a> en complementos para Excel y PowerPoint.</p></li><li><p>Obtener el identificador, el título y el índice de las diapositivas seleccionadas con el método <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-" target="_blank">getSelectedDataAsync</a> (cuando especifique la nueva enumeración <span class="keyword">Office.CoercionType.SlideRange</span><a href="https://docs.microsoft.com/javascript/api/office/office.coerciontype?view=office-js" target="_blank">coercionType</a>) en complementos para PowerPoint.</p></li></ul>|1.1|
|`Location`|Se agregó para habilitar la configuración de ubicación de una cita en complementos de Outlook en modo de redacción.|1.1|
|`Office`|Se actualizó el método de selección para admitir la obtención de enlaces en complementos de contenido para Access.|1.1|
|`Recipients`|Se agregó para habilitar la obtención y configuración de destinatarios de un mensaje o una cita en modo de redacción.|1.1|
|`Settings`|Se actualizó para admitir la creación de configuraciones personalizadas en complementos de contenido para Access.|1.1|
|`Subject`|Se agregó para habilitar la obtención o configuración del asunto de un mensaje o una cita en complementos de Outlook en modo de redacción.|1.1|
|`Time`|Se agregó para habilitar la obtención y configuración de la hora de inicio y finalización de una cita en complementos de Outlook en modo de redacción.|1.1|

### <a name="new-and-updated-enumerations"></a>Enumeraciones nuevas y actualizadas

|**Objeto**|**Descripción**|**Version**|
|:-----|:-----|:-----|
|`ActiveView`|Especifica el estado de la vista activa del documento (por ejemplo, si el usuario puede editar o no el documento).Se agregó para que los complementos para PowerPoint puedan determinar si los usuarios están viendo la presentación ( **Presentación con diapositivas**) o modificando diapositivas. |1.1|
|`CoercionType`|Se actualizó con **Office.CoercionType.SlideRange** para admitir la obtención del intervalo de diapositivas seleccionado con el método **getSelectedDataAsync** en complementos para PowerPoint.|1.1|
|`EventType`|Se actualizó para incluir el nuevo evento ActiveViewChanged.|1.1|
|`FileType`|Se actualizó para especificar el resultado en formato PDF.|1.1|
|`GoToType`|Se agregó para especificar el lugar u objeto del documento al que se debe ir.|1.1|

