
# <a name="item"></a>elemento

### [Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item

El espacio de nombres `item` se usa para acceder al mensaje seleccionado actualmente, a una convocatoria de reunión o a una cita. Puede determinar el tipo del `item` mediante la propiedad [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype).

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restringido|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

### <a name="example"></a>Ejemplo

En el siguiente ejemplo de código de JavaScript, se muestra cómo tener acceso a la propiedad `subject` del elemento actual en Outlook.

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a>Miembros

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a>attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)>

Obtiene una matriz de datos adjuntos para el elemento. Solo modo Lectura.

> [!NOTE]
> Ciertos tipos de archivos están bloqueados por Outlook debido a problemas de seguridad y, por tanto, no se devuelven. Para obtener más información, vea [datos adjuntos bloqueados en Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).

##### <a name="type"></a>Tipo:

*   Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)>

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="example"></a>Ejemplo

El siguiente código crea una cadena HTML con los detalles de todos los datos adjuntos en el elemento actual.

```JavaScript
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a>bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)

Obtiene un objeto que proporciona métodos para obtener o actualizar a los destinatarios de la línea de CCO (con copia oculta) de un mensaje. Solo modo Redacción.

##### <a name="type"></a>Tipo:

*   [Destinatarios](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción|

##### <a name="example"></a>Ejemplo

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a>body :[Body](/javascript/api/outlook_1_1/office.body)

Obtiene un objeto que proporciona métodos para manipular el cuerpo de un elemento.

##### <a name="type"></a>Tipo:

*   [Cuerpo](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a>cc: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook_1_1/office.recipients)

Proporciona acceso a los destinatarios de Cc (con copia) de un mensaje. El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `cc` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **CC** del mensaje. La colección está limitada a un máximo de 100 miembros.

##### <a name="compose-mode"></a>Modo de redacción

El `cc` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **Cc** del mensaje.

##### <a name="type"></a>Tipo:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a>(nullable) conversationId :String

Obtiene un identificador de la conversación de correo electrónico que contiene un mensaje determinado.

Puede obtener un entero para esta propiedad si la aplicación de correo se activa en respuestas o formularios de lectura o en formularios de redacción. Si, después, el usuario cambia el asunto del mensaje de respuesta, al enviar dicha respuesta, el Id. de conversación de ese mensaje cambiará y el valor que se obtuvo anteriormente ya no será de aplicación.

Obtiene un valor NULL para esta propiedad para un nuevo elemento de un formulario de redacción. Si el usuario establece un asunto y guarda el elemento, la propiedad `conversationId` devolverá un valor.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

#### <a name="datetimecreated-date"></a>dateTimeCreated :Date

Obtiene la fecha y hora de creación de un elemento. Solo modo Lectura.

##### <a name="type"></a>Tipo:

*   Fecha

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="example"></a>Ejemplo

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a>dateTimeModified :Date

Obtiene la fecha y hora en que se modificó por última vez un elemento. Solo modo Lectura.

> [!NOTE]
> Este miembro no se admite en Outlook para iOS o Outlook para Android.

##### <a name="type"></a>Tipo:

*   Fecha

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="example"></a>Ejemplo

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a>end :Date|[Time](/javascript/api/outlook_1_1/office.time)

Obtiene o establece la fecha y la hora de finalización de la cita.

La propiedad `end` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) para convertir el valor de la propiedad final a la fecha y hora local del cliente.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `end` devuelve un objeto `Date`.

##### <a name="compose-mode"></a>Modo de redacción

La propiedad `end` devuelve un objeto `Time`.

Si usa el método [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) para establecer la hora de finalización, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.

##### <a name="type"></a>Tipo:

*   Date | [Time](/javascript/api/outlook_1_1/office.time)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="example"></a>Ejemplo

En el ejemplo siguiente, se establece la hora de finalización de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) del objeto `Time`.

```JavaScript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a>from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)

Obtiene la dirección de correo electrónico del remitente de un mensaje. Solo modo Lectura.

Las propiedades `from` y [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.

> [!NOTE]
> El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `from` propiedad es `undefined`.

##### <a name="type"></a>Tipo:

*   [EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

#### <a name="internetmessageid-string"></a>internetMessageId :String

Obtiene el identificador de mensaje de Internet de un mensaje de correo electrónico. Solo modo Lectura.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="example"></a>Ejemplo

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a>itemClass :String

Obtiene la clase de elemento Servicios Web Exchange del elemento seleccionado. Solo modo Lectura.

La propiedad `itemClass` especifica la clase de mensaje del elemento seleccionado. Las siguientes son las clases de mensaje predeterminadas para el elemento de mensaje o cita.

| Tipo | Descripción | Clase de elemento |
| --- | --- | --- |
| Elementos de cita | Estos son los elementos de calendario de la clase de elemento `IPM.Appointment` o `IPM.Appointment.Occurence`. | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| Elementos de mensaje | Entre estos se encuentran los mensajes de correo electrónico que tienen la clase de mensaje predeterminada `IPM.Note` y las convocatorias de reunión, respuestas y cancelaciones que usan `IPM.Schedule.Meeting` como clase de mensaje base. | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

Puede crear clases de mensaje personalizadas para ampliar una clase de mensaje predeterminada. Por ejemplo, puede crear una clase de mensaje de cita personalizada `IPM.Appointment.Contoso`.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="example"></a>Ejemplo

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a>(nullable) itemId :String

Obtiene el identificador de elemento de Servicios Web Exchange para el elemento actual. Solo modo Lectura.

> [!NOTE]
> El identificador que devuelve la propiedad `itemId` es el mismo que el identificador de elemento de Servicios Web Exchange. El `itemId` (propiedad) no es idéntico al identificador de entrada de Outlook o el identificador de la API de REST de Outlook. Antes de realizar llamadas a API de REST utilizando este valor, se debe convertir mediante `Office.context.mailbox.convertToRestId`, que está disponible a partir de requisito establece 1.3. Para obtener más información, vea [usar las API de REST de Outlook desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="example"></a>Ejemplo

El siguiente código comprueba la presencia de un identificador de elemento. Si la propiedad `itemId` devuelve `null` o `undefined`, guarda el elemento en el servidor y obtiene el identificador de elemento desde el resultado asincrónico.

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a>itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

Obtiene el tipo de elemento que representa una instancia.

La propiedad `itemType` devuelve uno de los valores de enumeración de `ItemType`, lo que indica si la instancia del objeto `item` es un mensaje o una cita.

##### <a name="type"></a>Tipo:

*   [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a>location :String|[Location](/javascript/api/outlook_1_1/office.location)

Obtiene o establece la ubicación de una cita.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `location` devuelve una cadena que contiene la ubicación de la cita.

##### <a name="compose-mode"></a>Modo de redacción

La propiedad `location` devuelve un objeto `Location` que proporciona métodos para obtener y establecer la ubicación de la cita.

##### <a name="type"></a>Tipo:

*   String | [Location](/javascript/api/outlook_1_1/office.location)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a>normalizedSubject :String

Obtiene el asunto de un elemento, con todos los prefijos quitados (incluidos `RE:` y `FWD:`). Solo modo Lectura.

La propiedad normalizedSubject obtiene el asunto del elemento, con los prefijos estándar (como `RE:` y `FW:`) agregados por los programas de correo electrónico. Para obtener el asunto del elemento con los prefijos intactos, use la propiedad [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject).

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="example"></a>Ejemplo

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a>optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)

Proporciona acceso a los asistentes opcionales de un evento. El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `optionalAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente opcional a la reunión.

##### <a name="compose-mode"></a>Modo de redacción

El `optionalAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes opcionales de una reunión.

##### <a name="type"></a>Tipo:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a>organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)

Obtiene la dirección de correo electrónico del organizador de una reunión especificada. Solo modo Lectura.

##### <a name="type"></a>Tipo:

*   [EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="example"></a>Ejemplo

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a>requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)

Proporciona acceso a los asistentes necesarios de un evento. El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `requiredAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente requerido a la reunión.

##### <a name="compose-mode"></a>Modo de redacción

El `requiredAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes necesarios de una reunión.

##### <a name="type"></a>Tipo:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a>sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)

Obtiene la dirección de correo electrónico del remitente de un mensaje de correo electrónico. Solo modo Lectura.

Las propiedades [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) y `sender` representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.

> [!NOTE]
> El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `from` propiedad es `undefined`.

##### <a name="type"></a>Tipo:

*   [EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="example"></a>Ejemplo

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a>start :Date|[Time](/javascript/api/outlook_1_1/office.time)

Obtiene o establece la fecha y la hora de inicio de la cita.

La propiedad `start` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) para convertir el valor a la fecha y hora local del cliente.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `start` devuelve un objeto `Date`.

##### <a name="compose-mode"></a>Modo de redacción

La propiedad `start` devuelve un objeto `Time`.

Si usa el método [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) para establecer la hora de inicio, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.

##### <a name="type"></a>Tipo:

*   Date | [Time](/javascript/api/outlook_1_1/office.time)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="example"></a>Ejemplo

En el ejemplo siguiente, se establece la hora de inicio de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) del objeto `Time`.

```JavaScript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a>subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)

Obtiene o establece la descripción que se muestra en el campo de asunto de un elemento.

La propiedad `subject` obtiene o establece el asunto completo del elemento, como lo ha enviado el servidor de correo electrónico.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `subject` devuelve una cadena. Use la propiedad [`normalizedSubject`](#normalizedsubject-string) para obtener el asunto menos cualquier prefijo inicial como `RE:` y `FW:`.

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a>Modo de redacción

La propiedad `subject` devuelve un objeto `Subject` que proporciona métodos para obtener y establecer el asunto.

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a>Tipo:

*   String | [Subject](/javascript/api/outlook_1_1/office.subject)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a>para: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook_1_1/office.recipients)

Proporciona acceso a los destinatarios en la línea **para** de un mensaje. El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.

##### <a name="read-mode"></a>Modo de lectura

La propiedad `to` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **Para** del mensaje. La colección está limitada a un máximo de 100 miembros.

##### <a name="compose-mode"></a>Modo de redacción

El `to` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **para** del mensaje.

##### <a name="type"></a>Tipo:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a>Métodos

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

Agrega un archivo a un mensaje o cita como datos adjuntos.

El método `addFileAttachmentAsync` carga el archivo en el URI especificado y lo asocia al elemento en el formulario de redacción.

Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`uri`| String||El URI que proporciona la ubicación del archivo que se va a adjuntar al mensaje o a la cita. La longitud máxima es de 2048 caracteres.|
|`attachmentName`| String||El nombre de los datos adjuntos que se muestra mientras estos se cargan. La longitud máxima es de 255 caracteres.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.|
|`options.asyncContext`| Objeto| &lt;optional&gt;|Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.<br/>Si se produce un error en la carga de los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.|

##### <a name="errors"></a>Errores

| Código de error | Descripción |
|------------|-------------|
| `AttachmentSizeExceeded` | Los datos adjuntos son más grandes de lo permitido. |
| `FileTypeNotSupported` | Los datos adjuntos tienen una extensión que no está permitida. |
| `NumberOfAttachmentsExceeded` | El mensaje o cita tiene demasiados datos adjuntos. |

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción|

##### <a name="example"></a>Ejemplo

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

Agrega un elemento de Exchange (por ejemplo, un mensaje) como datos adjuntos al mensaje o a la cita.

El método `addItemAttachmentAsync` asocia el elemento con el identificador especificado de Exchange al elemento en el formulario de redacción. Si se especifica un método de devolución de llamada, se llama al método con un parámetro, `asyncResult`, que contiene el identificador de datos adjuntos o un código que indique cualquier error que se ha producido al adjuntar el elemento. Si es necesario, puede usar el parámetro `options` para pasar información de estado al método de devolución de llamada.

Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.

Si el complemento de Office se está ejecutando en Outlook Web App, el `addItemAttachmentAsync` método puede adjuntar elementos a los elementos que no sea el elemento que se está editando; Sin embargo, esto no se admite y no se recomienda.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`itemId`| String||El identificador de Exchange del elemento que debe adjuntarse. La longitud máxima es de 100 caracteres.|
|`attachmentName`| String||El asunto del elemento que debe adjuntarse. La longitud máxima es de 255 caracteres.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.|
|`options.asyncContext`| Objeto| &lt;optional&gt;|Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.<br/>Si se produce un error al agregar los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.|

##### <a name="errors"></a>Errores

| Código de error | Descripción |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | El mensaje o cita tiene demasiados datos adjuntos. |

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción|

##### <a name="example"></a>Ejemplo

En el ejemplo siguiente, se agrega un elemento existente de Outlook como un archivo adjunto con el nombre `My Attachment`.

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

#### <a name="displayreplyallformformdata"></a>displayReplyAllForm(formData)

Muestra un formulario de respuesta que incluye el remitente y todos los destinatarios del mensaje seleccionado o el organizador y todos los asistentes de la cita seleccionada.

> [!NOTE]
> Este método no se admite en Outlook para iOS o Outlook para Android.

En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.

Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyAllForm` produce una excepción.

> [!NOTE]
> La capacidad de incluir datos adjuntos en la llamada a `displayReplyAllForm` no se admite en el conjunto de requisitos 1.1. Se ha agregado la compatibilidad con datos adjuntos a `displayReplyAllForm` en el conjunto de requisitos 1.2 y versiones superiores.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`formData`| String &#124; Object| |Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.<br/>**O**<br/>Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente. |
| `formData.htmlBody` | String | &lt;optional&gt; | Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.
| `callback` | función | &lt;optional&gt; | Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult). |

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="examples"></a>Ejemplos

El código siguiente pasa una cadena a la función `displayReplyAllForm`.

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

Responder con un cuerpo vacío.

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

Responder solo con un cuerpo.

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

Responder con un cuerpo y una devolución de llamada.

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a>displayReplyForm(formData)

Muestra un formulario de respuesta que incluye solo el remitente del mensaje seleccionado o el organizador de la cita seleccionada.

> [!NOTE]
> Este método no se admite en Outlook para iOS o Outlook para Android.

En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.

Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyForm` produce una excepción.

> [!NOTE]
> La capacidad de incluir datos adjuntos en la llamada a `displayReplyForm` no se admite en el conjunto de requisitos 1.1. Se ha agregado la compatibilidad con datos adjuntos a `displayReplyForm` en el conjunto de requisitos 1.2 y versiones superiores.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`formData`| String &#124; Object| | Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.<br/>**O**<br/>Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente. |
| `formData.htmlBody` | String | &lt;optional&gt; | Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.
| `callback` | función | &lt;optional&gt; | Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult). |

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="examples"></a>Ejemplos

El código siguiente pasa una cadena a la función `displayReplyForm`.

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

Responder con un cuerpo vacío.

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

Responder solo con un cuerpo.

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

Responder con un cuerpo y una devolución de llamada.

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a>getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}

Obtiene las entidades que se encuentran en el cuerpo del elemento seleccionado.

> [!NOTE]
> Este método no se admite en Outlook para iOS o Outlook para Android.

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="returns"></a>Valores devueltos:

Tipo: [Entidades](/javascript/api/outlook_1_1/office.entities)

##### <a name="example"></a>Ejemplo

En el ejemplo siguiente se obtiene acceso a las entidades de los contactos en el cuerpo del elemento actual.

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}

Obtiene una matriz de todas las entidades del tipo de entidad especificado que se encuentran en el cuerpo del elemento seleccionado.

> [!NOTE]
> Este método no se admite en Outlook para iOS o Outlook para Android.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|Uno de los valores de enumeración de EntityType.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restringido|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="returns"></a>Valores devueltos:

Si el valor que se pasa a `entityType` no es un miembro válido de la enumeración de `EntityType`, el método devuelve un valor NULL. Si no existen entidades del tipo especificado están presentes en el cuerpo del elemento, el método devuelve una matriz vacía. De otro modo, el tipo de los objetos en la matriz devuelta depende del tipo de entidad solicitada en el parámetro `entityType`.

Aunque el nivel de permiso mínimo para usar este método es **Restricted**, algunos tipos de entidad requieren **ReadItem** para tener acceso, como se especifica en la tabla siguiente.

| Valor de `entityType` | Tipo de objetos de la matriz devuelta | Nivel de permiso requerido |
| --- | --- | --- |
| `Address` | Cadena | **Restringido** |
| `Contact` | Contacto | **ReadItem** |
| `EmailAddress` | Cadena | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **Restringido** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | Cadena | **Restringido** |

Tipo:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>


##### <a name="example"></a>Ejemplo

En el ejemplo siguiente se muestra cómo obtener acceso a una matriz de cadenas que representan direcciones postales en el cuerpo del elemento actual.

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}

Devuelve entidades conocidas en el elemento seleccionado que pasan el filtro con nombre definido en el archivo XML de manifiesto.

> [!NOTE]
> Este método no se admite en Outlook para iOS o Outlook para Android.

El método `getFilteredEntitiesByName` devuelve las entidades que coinciden con la expresión regular definida en el elemento de regla [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) del archivo XML de manifiesto con el valor de elemento `FilterName` especificado.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`name`| String|El nombre del elemento de regla `ItemHasKnownEntity` que define el filtro para que coincidan.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="returns"></a>Valores devueltos:

Si no existe ningún elemento `ItemHasKnownEntity` en el manifiesto con un valor de elemento `FilterName` que coincida con el parámetro `name`, el método devuelve `null`. Si el parámetro `name` no coincide con un elemento `ItemHasKnownEntity` del manifiesto, pero no existen entidades en el elemento actual que coincidan, el método devuelve una matriz vacía.


Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>


#### <a name="getregexmatches--object"></a>getRegExMatches() → {Object}

Devuelve valores de cadena en el elemento seleccionado que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto.

> [!NOTE]
> Este método no se admite en Outlook para iOS o Outlook para Android.

El método `getRegExMatches` devuelve las cadenas que coinciden con la expresión regular definida en cada elemento de regla `ItemHasRegularExpressionMatch` o `ItemHasKnownEntity` del archivo XML de manifiesto. Para una regla `ItemHasRegularExpressionMatch`, tiene que darse una cadena coincidente en la propiedad del elemento especificado por la regla. El tipo simple `PropertyName` define las propiedades admitidas.

Por ejemplo, considere un manifiesto de complemento que tenga el siguiente elemento `Rule`:

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

El objeto devuelto desde `getRegExMatches` tendría dos propiedades: `fruits` y `veggies`.

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados.

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="returns"></a>Valores devueltos:

Un objeto que contiene matrices de cadenas que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. El nombre de cada matriz es igual al valor correspondiente del atributo `RegExName` de la regla `ItemHasRegularExpressionMatch` coincidente o al atributo `FilterName` de la regla `ItemHasKnownEntity` coincidente.

<dl class="param-type">

<dt>Tipo</dt>

<dd>Object</dd>

</dl>

##### <a name="example"></a>Ejemplo

En el siguiente ejemplo, se muestra cómo tener acceso a la matriz de coincidencias de los <rule>elementos de expresión regular `fruits`y `veggies`, que se especifican en el manifiesto.</rule>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a>→ de getregexmatchesbyname (Name) (nullable) {matriz. < cadena >}

Devuelve valores de cadena en el elemento seleccionado que coinciden con la expresión regular de nombre definida en el archivo XML de manifiesto.

> [!NOTE]
> Este método no se admite en Outlook para iOS o Outlook para Android.

El método `getRegExMatchesByName` devuelve las cadenas que coinciden con la expresión regular definida en el elemento de regla `ItemHasRegularExpressionMatch` del archivo XML de manifiesto con el valor de elemento `RegExName` especificado.

Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Descripción|
|---|---|---|
|`name`| String|El nombre del elemento de regla `ItemHasRegularExpressionMatch` que define el filtro para que coincidan.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lectura|

##### <a name="returns"></a>Valores devueltos:

Una matriz que contiene las cadenas que coinciden con la expresión regular definida en el archivo XML de manifiesto.

<dl class="param-type">

<dt>Tipo</dt>

<dd>< Cadena > de matriz.</dd>

</dl>

##### <a name="example"></a>Ejemplo

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

Carga de forma asincrónica las propiedades personalizadas de este complemento en el elemento seleccionado.

Las propiedades personalizadas se almacenan como pares de clave/valor según la aplicación y el elemento. Este método devuelve un objeto `CustomProperties` en la devolución de llamada, que proporciona métodos para tener acceso a las propiedades personalizadas específicas del elemento y el complemento actual. Las propiedades personalizadas no están cifradas en el elemento, por lo que no debería usarse como almacenamiento seguro.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`callback`| función||Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Las propiedades personalizadas se proporcionan como un objeto [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) en la propiedad `asyncResult.value`. Este objeto se puede utilizar para obtener, establecer y elimina las propiedades personalizadas del elemento y guarde los cambios en la propiedad personalizada que se vuelve a establecer en el servidor.|
|`userContext`| Objeto| &lt;optional&gt;|Los desarrolladores pueden proporcionar cualquier objeto que desean tener acceso en la función de devolución de llamada. Se puede tener acceso a este objeto por el `asyncResult.asyncContext` (propiedad) en la función de devolución de llamada.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="example"></a>Ejemplo

En el siguiente ejemplo de código se muestra cómo usar el método `loadCustomPropertiesAsync` para cargar de forma asincrónica propiedades personalizadas que son específicas del elemento actual. En el ejemplo también se muestra cómo usar el método `CustomProperties.saveAsync` para guardar estas propiedades en el servidor. Después de cargar las propiedades personalizadas, el ejemplo de código usa el método `CustomProperties.get` para leer la propiedad personalizada `myProp`, el método `CustomProperties.set` para escribir en la propiedad personalizada `otherProp` y finalmente se llama al método `saveAsync` para guardar las propiedades personalizadas.

```JavaScript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId, [options], [callback])

Quita los datos adjuntos de un mensaje o cita.

El método `removeAttachmentAsync` quita del elemento los datos adjuntos con el identificador especificado. Como práctica recomendada, debe usar el identificador de datos adjuntos para quitar datos adjuntos solo si la misma aplicación de correo ha agregado los datos adjuntos en la misma sesión. En Outlook Web App y OWA para los dispositivos, el identificador de datos adjuntos es válido solo en la misma sesión. Una sesión finaliza cuando el usuario cierra la aplicación, o si el usuario comienza a redactar en un formulario en línea y posteriormente se extrae el formulario en línea para continuar en una ventana independiente.

##### <a name="parameters"></a>Parámetros:

|Nombre| Tipo| Atributos| Descripción|
|---|---|---|---|
|`attachmentId`| String||El identificador de los datos adjuntos que se deben quitar. La longitud máxima de la cadena es de 100 caracteres.|
|`options`| Object| &lt;optional&gt;|Un objeto literal que contiene una o más de las siguientes propiedades.|
|`options.asyncContext`| Objeto| &lt;optional&gt;|Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.|
|`callback`| función| &lt;optional&gt;|Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>Si se produce un error en la eliminación de los datos adjuntos, la propiedad `asyncResult.error` contendrá un código de error con el motivo del error.|

##### <a name="errors"></a>Errores

| Código de error | Descripción |
|------------|-------------|
| `InvalidAttachmentId` | El identificador de datos adjuntos no existe. |

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción|

##### <a name="example"></a>Ejemplo

El código siguiente quita un archivo adjunto con un identificador de "0".

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```