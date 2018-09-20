# <a name="office"></a>Office

El espacio de nombres de Office proporciona interfaces compartidas que los complementos usan en todas las aplicaciones de Office. Este listado documenta solo aquellas interfaces que usan los complementos de Outlook. Para obtener un listado completo del espacio de nombres de Office, vea [API compartida](/javascript/api/office).

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="members-and-methods"></a>Miembros y métodos

| Miembro	 | Tipo |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | Miembro	 |
| [CoercionType](#coerciontype-string) | Miembro	 |
| [EventType](#eventtype-string) | Miembro	 |
| [SourceProperty](#sourceproperty-string) | Miembro	 |

### <a name="namespaces"></a>Espacios de nombres

[contexto](office.context.md): proporciona interfaces compartidas del espacio de nombres de contexto de la API de complementos de Office para su uso en la API de complemento de Outlook.

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Incluye las enumeraciones ItemType, EntityType, AttachmentType, RecipientType, ResponseType y ItemNotificationMessageType.

### <a name="members"></a>Miembros

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

Especifica el resultado de una llamada asíncrona.

##### <a name="type"></a>Tipo:

*   String

##### <a name="properties"></a>Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`Succeeded`| String|La llamada ha sido correcta.|
|`Failed`| String|La llamada ha fallado.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

---

####  <a name="coerciontype-string"></a>CoercionType :String

Especifica cómo convertir los datos que el método invocado ha devuelto o definido.

##### <a name="type"></a>Tipo:

*   String

##### <a name="properties"></a>Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`Html`| String|Solicita que los datos se devuelvan en formato HTML.|
|`Text`| String|Solicita que los datos se devuelvan en formato de texto.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

---

####  <a name="eventtype-string"></a>EventType :String

Especifica el evento asociado con un controlador de eventos.

##### <a name="type"></a>Tipo:

*   String

##### <a name="properties"></a>Propiedades:

| Nombre | Tipo | Descripción |
|---|---|---|
|`ItemChanged`| String | El elemento seleccionado ha cambiado. |

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1,5 |
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura |

---

####  <a name="sourceproperty-string"></a>SourceProperty :String

Especifica el origen de los datos devueltos por el método invocado.

##### <a name="type"></a>Tipo:

*   String

##### <a name="properties"></a>Propiedades:

|Nombre| Tipo| Descripción|
|---|---|---|
|`Body`| String|El origen de los datos proviene del cuerpo de un mensaje.|
|`Subject`| String|El origen de los datos proviene del asunto de un mensaje.|

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|