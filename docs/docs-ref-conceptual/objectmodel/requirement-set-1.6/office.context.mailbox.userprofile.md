
# <a name="userprofile"></a>userProfile

### [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="members-and-methods"></a>Miembros y métodos

| Miembro	 | Tipo |
|--------|------|
| [accountType](#accounttype-string) | Member |
| [displayName](#displayname-string) | Miembro	 |
| [emailAddress](#emailaddress-string) | Miembro	 |
| [timeZone](#timezone-string) | Miembro	 |

### <a name="members"></a>Members

####  <a name="accounttype-string"></a>accountType: cadena

> [!NOTE]
> Este miembro es actualmente sólo admitidos en 2016 de Outlook para Mac, compilación 16.9.1212 y mayor.

Obtiene el tipo de cuenta del usuario asociado con el buzón de correo. Los valores posibles se enumeran en la siguiente tabla.

| Valor | Descripción |
|-------|-------------|
| `enterprise` | El buzón se encuentra en un servidor de Exchange local. |
| `gmail` | El buzón de correo está asociado con una cuenta de Gmail. |
| `office365` | El buzón de correo está asociada con un Office 365 funciona o escuela cuenta. |
| `outlookCom` | El buzón de correo está asociado con una cuenta de Outlook.com personal. |

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a>displayName :String

Obtiene el nombre para mostrar del usuario.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a>emailAddress :String

Obtiene la dirección de correo electrónico SMTP del usuario.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a>timeZone :String

Obtiene la zona horaria predeterminada del usuario.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="example"></a>Ejemplo

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```