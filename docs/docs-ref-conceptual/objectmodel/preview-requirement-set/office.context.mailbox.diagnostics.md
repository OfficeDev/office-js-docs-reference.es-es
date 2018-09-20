
# <a name="diagnostics"></a>diagnostics

### [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics

Proporciona información de diagnóstico a un complemento de Outlook.

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

##### <a name="members-and-methods"></a>Miembros y métodos

| Miembro	 | Tipo |
|--------|------|
| [hostName](#hostname-string) | Miembro	 |
| [hostVersion](#hostversion-string) | Miembro	 |
| [OWAView](#owaview-string) | Miembro	 |

### <a name="members"></a>Miembros

####  <a name="hostname-string"></a>hostName :String

Obtiene una cadena que representa el nombre de la aplicación host.

Una cadena que puede ser uno de los siguientes valores: `Outlook`, `Mac Outlook`, `OutlookIOS` o `OutlookWebApp`.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

####  <a name="hostversion-string"></a>hostVersion :String

Obtiene una cadena que representa la versión de la aplicación host o de Exchange Server.

Si el complemento de correo se ejecuta en el cliente de escritorio de Outlook o en Outlook para iOS, la propiedad `hostVersion` devuelve la versión de la aplicación host, Outlook. En Outlook Web App, la propiedad devuelve la versión de Exchange Server. Un ejemplo es la cadena `15.0.468.0`.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|

####  <a name="owaview-string"></a>OWAView :String

Obtiene una cadena que representa la vista actual de Outlook Web App.

La cadena devuelta puede ser uno de los valores siguientes: `OneColumn`, `TwoColumns` o `ThreeColumns`.

Si la aplicación host no es Outlook Web App, obtener acceso a esta propiedad será `undefined`.

Outlook Web App tiene tres vistas que se corresponden con el ancho de la pantalla, con la ventana y con el número de columnas que pueden mostrarse:

*   `OneColumn` se muestra cuando la pantalla es estrecha. Outlook Web App usa este diseño de una sola columna en la pantalla completa de un smartphone.
*   `TwoColumns` se muestra cuando la pantalla es más ancha. Outlook Web App usa esta vista en la mayor parte de las tabletas.
*   `ThreeColumns` se muestra cuando la pantalla es ancha. Por ejemplo, Outlook Web App usa esta vista en una ventana de pantalla completa en un equipo de escritorio.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requirement| Valor|
|---|---|
|[Versión del conjunto de requisitos mínimos del buzón](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nivel de permisos mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo de Outlook aplicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redacción o lectura|