
# <a name="mailbox"></a><span data-ttu-id="8be8b-101">buzón de correo</span><span class="sxs-lookup"><span data-stu-id="8be8b-101">mailbox</span></span>

### <span data-ttu-id="8be8b-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="8be8b-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="8be8b-104">Proporciona acceso al modelo de objeto de complemento de Outlook para Microsoft Outlook y Microsoft Outlook en el web.</span><span class="sxs-lookup"><span data-stu-id="8be8b-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8be8b-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8be8b-105">Requirements</span></span>

|<span data-ttu-id="8be8b-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="8be8b-106">Requirement</span></span>| <span data-ttu-id="8be8b-107">Valor</span><span class="sxs-lookup"><span data-stu-id="8be8b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be8b-108">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8be8b-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be8b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8be8b-109">1.0</span></span>|
|[<span data-ttu-id="8be8b-110">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8be8b-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be8b-111">Restringido</span><span class="sxs-lookup"><span data-stu-id="8be8b-111">Restricted</span></span>|
|[<span data-ttu-id="8be8b-112">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8be8b-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be8b-113">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8be8b-113">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="8be8b-114">Espacios de nombres</span><span class="sxs-lookup"><span data-stu-id="8be8b-114">Namespaces</span></span>

<span data-ttu-id="8be8b-115">[diagnostics](Office.context.mailbox.diagnostics.md): Proporciona información de diagnóstico a un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="8be8b-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="8be8b-116">[item](Office.context.mailbox.item.md): Proporciona métodos y propiedades para tener acceso a un mensaje o cita en un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="8be8b-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="8be8b-117">[userProfile](Office.context.mailbox.userProfile.md): Proporciona información sobre el usuario en un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="8be8b-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="8be8b-118">Miembros</span><span class="sxs-lookup"><span data-stu-id="8be8b-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="8be8b-119">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="8be8b-119">ewsUrl :String</span></span>

<span data-ttu-id="8be8b-p102">Obtiene la dirección URL del punto de conexión de Servicios Web Exchange (EWS) para esta cuenta de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8be8b-122">Este miembro no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8be8b-122">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8be8b-p103">El valor `ewsUrl` puede usarse por un servicio remoto para realizar llamadas EWS al buzón del usuario. Por ejemplo, puede crear un servicio remoto para [obtener datos adjuntos desde el elemento seleccionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="8be8b-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="8be8b-125">La aplicación debe tener el permiso **ReadItem** especificado en su manifiesto para poder llamar al miembro `ewsUrl` en el modo lectura.</span><span class="sxs-lookup"><span data-stu-id="8be8b-125">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="8be8b-p104">En el modo de redacción debe llamar al método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar el miembro `ewsUrl`. Su aplicación debe tener permisos **ReadWriteItem** para llamar al método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="8be8b-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8be8b-128">Type:</span></span>

*   <span data-ttu-id="8be8b-129">String</span><span class="sxs-lookup"><span data-stu-id="8be8b-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8be8b-130">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8be8b-130">Requirements</span></span>

|<span data-ttu-id="8be8b-131">Requirement</span><span class="sxs-lookup"><span data-stu-id="8be8b-131">Requirement</span></span>| <span data-ttu-id="8be8b-132">Valor</span><span class="sxs-lookup"><span data-stu-id="8be8b-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be8b-133">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8be8b-133">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be8b-134">1.0</span><span class="sxs-lookup"><span data-stu-id="8be8b-134">1.0</span></span>|
|[<span data-ttu-id="8be8b-135">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8be8b-135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be8b-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8be8b-136">ReadItem</span></span>|
|[<span data-ttu-id="8be8b-137">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8be8b-137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be8b-138">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8be8b-138">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="8be8b-139">Métodos</span><span class="sxs-lookup"><span data-stu-id="8be8b-139">Methods</span></span>

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="8be8b-140">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="8be8b-140">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="8be8b-141">Convierte un identificador de elemento con formato para REST al formato EWS.</span><span class="sxs-lookup"><span data-stu-id="8be8b-141">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="8be8b-142">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8be8b-142">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8be8b-p105">Los identificadores de elemento obtenidos a través de una API de REST (como la [API de correo de Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) o [Microsoft Graph](http://graph.microsoft.io/)) usan un formato diferente al formato que usa Exchange Web Services (EWS). El método `convertToEwsId` convierte un identificador con formato REST al formato adecuado para EWS.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8be8b-145">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8be8b-145">Parameters:</span></span>

|<span data-ttu-id="8be8b-146">Nombre</span><span class="sxs-lookup"><span data-stu-id="8be8b-146">Name</span></span>| <span data-ttu-id="8be8b-147">Tipo</span><span class="sxs-lookup"><span data-stu-id="8be8b-147">Type</span></span>| <span data-ttu-id="8be8b-148">Descripción</span><span class="sxs-lookup"><span data-stu-id="8be8b-148">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="8be8b-149">String</span><span class="sxs-lookup"><span data-stu-id="8be8b-149">String</span></span>|<span data-ttu-id="8be8b-150">Un identificador de elemento con formato para las API de REST de Outlook</span><span class="sxs-lookup"><span data-stu-id="8be8b-150">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="8be8b-151">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="8be8b-151">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.restversion)|<span data-ttu-id="8be8b-152">Un valor que indica la versión de la API de REST de Outlook que se usa para recuperar el identificador de elemento.</span><span class="sxs-lookup"><span data-stu-id="8be8b-152">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8be8b-153">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8be8b-153">Requirements</span></span>

|<span data-ttu-id="8be8b-154">Requirement</span><span class="sxs-lookup"><span data-stu-id="8be8b-154">Requirement</span></span>| <span data-ttu-id="8be8b-155">Valor</span><span class="sxs-lookup"><span data-stu-id="8be8b-155">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be8b-156">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8be8b-156">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be8b-157">1.3</span><span class="sxs-lookup"><span data-stu-id="8be8b-157">1.3</span></span>|
|[<span data-ttu-id="8be8b-158">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8be8b-158">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be8b-159">Restringido</span><span class="sxs-lookup"><span data-stu-id="8be8b-159">Restricted</span></span>|
|[<span data-ttu-id="8be8b-160">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8be8b-160">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be8b-161">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8be8b-161">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8be8b-162">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="8be8b-162">Returns:</span></span>

<span data-ttu-id="8be8b-163">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="8be8b-163">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="8be8b-164">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8be8b-164">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime"></a><span data-ttu-id="8be8b-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_4/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="8be8b-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_4/office.LocalClientTime)}</span></span>

<span data-ttu-id="8be8b-166">Obtiene un diccionario con información de tiempo en el tiempo del cliente local.</span><span class="sxs-lookup"><span data-stu-id="8be8b-166">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="8be8b-p106">Las fechas y horas usadas por una aplicación de correo para Outlook o Outlook Web App pueden usar distintas zonas horarias. Outlook usa la zona horaria del equipo cliente; Outlook Web App usa la zona horaria definida en el Centro de administración de Exchange (EAC). Debería tratar los valores de fecha y hora de modo que los valores que aparezcan en la interfaz de usuario sean siempre coherentes con la zona horaria que el usuario espera.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p106">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="8be8b-p107">Si se está ejecutando la aplicación de correo en Outlook, el método `convertToLocalClientTime` devolverá un objeto de diccionario con los valores establecidos para la zona horaria del equipo cliente. Si se está ejecutando la aplicación de correo en Outlook Web App, el método `convertToLocalClientTime` devolverá un objeto de diccionario con los valores establecidos para la zona horaria especificada en el CEF.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p107">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8be8b-172">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8be8b-172">Parameters:</span></span>

|<span data-ttu-id="8be8b-173">Nombre</span><span class="sxs-lookup"><span data-stu-id="8be8b-173">Name</span></span>| <span data-ttu-id="8be8b-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="8be8b-174">Type</span></span>| <span data-ttu-id="8be8b-175">Descripción</span><span class="sxs-lookup"><span data-stu-id="8be8b-175">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="8be8b-176">Fecha</span><span class="sxs-lookup"><span data-stu-id="8be8b-176">Date</span></span>|<span data-ttu-id="8be8b-177">Un objeto Date</span><span class="sxs-lookup"><span data-stu-id="8be8b-177">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8be8b-178">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8be8b-178">Requirements</span></span>

|<span data-ttu-id="8be8b-179">Requirement</span><span class="sxs-lookup"><span data-stu-id="8be8b-179">Requirement</span></span>| <span data-ttu-id="8be8b-180">Valor</span><span class="sxs-lookup"><span data-stu-id="8be8b-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be8b-181">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8be8b-181">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be8b-182">1.0</span><span class="sxs-lookup"><span data-stu-id="8be8b-182">1.0</span></span>|
|[<span data-ttu-id="8be8b-183">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8be8b-183">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be8b-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8be8b-184">ReadItem</span></span>|
|[<span data-ttu-id="8be8b-185">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8be8b-185">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be8b-186">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8be8b-186">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8be8b-187">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="8be8b-187">Returns:</span></span>

<span data-ttu-id="8be8b-188">Tipo: [LocalClientTime](/javascript/api/outlook_1_4/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="8be8b-188">Type: [LocalClientTime](/javascript/api/outlook_1_4/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="8be8b-189">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="8be8b-189">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="8be8b-190">Convierte un identificador de elemento con formato para EWS al formato REST.</span><span class="sxs-lookup"><span data-stu-id="8be8b-190">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="8be8b-191">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8be8b-191">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8be8b-p108">Los identificadores de elemento obtenidos a través de EWS o de la propiedad `itemId` usan un formato diferente al formato que usan las API de REST (como la [API de correo de Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) o [Microsoft Graph](http://graph.microsoft.io/)). El método `convertToRestId` convierte un identificador con formato EWS al formato adecuado para REST.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8be8b-194">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8be8b-194">Parameters:</span></span>

|<span data-ttu-id="8be8b-195">Nombre</span><span class="sxs-lookup"><span data-stu-id="8be8b-195">Name</span></span>| <span data-ttu-id="8be8b-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="8be8b-196">Type</span></span>| <span data-ttu-id="8be8b-197">Descripción</span><span class="sxs-lookup"><span data-stu-id="8be8b-197">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="8be8b-198">String</span><span class="sxs-lookup"><span data-stu-id="8be8b-198">String</span></span>|<span data-ttu-id="8be8b-199">Un identificador de elemento con formato para Exchange Web Services (EWS)</span><span class="sxs-lookup"><span data-stu-id="8be8b-199">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="8be8b-200">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="8be8b-200">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.restversion)|<span data-ttu-id="8be8b-201">Un valor que indica la versión de la API de REST de Outlook con la que se usará el identificador convertido.</span><span class="sxs-lookup"><span data-stu-id="8be8b-201">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8be8b-202">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8be8b-202">Requirements</span></span>

|<span data-ttu-id="8be8b-203">Requirement</span><span class="sxs-lookup"><span data-stu-id="8be8b-203">Requirement</span></span>| <span data-ttu-id="8be8b-204">Valor</span><span class="sxs-lookup"><span data-stu-id="8be8b-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be8b-205">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8be8b-205">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be8b-206">1.3</span><span class="sxs-lookup"><span data-stu-id="8be8b-206">1.3</span></span>|
|[<span data-ttu-id="8be8b-207">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8be8b-207">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be8b-208">Restringido</span><span class="sxs-lookup"><span data-stu-id="8be8b-208">Restricted</span></span>|
|[<span data-ttu-id="8be8b-209">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8be8b-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be8b-210">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8be8b-210">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8be8b-211">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="8be8b-211">Returns:</span></span>

<span data-ttu-id="8be8b-212">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="8be8b-212">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="8be8b-213">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8be8b-213">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="8be8b-214">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="8be8b-214">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="8be8b-215">Obtiene un objeto Date del diccionario que contiene información de tiempo.</span><span class="sxs-lookup"><span data-stu-id="8be8b-215">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="8be8b-216">El método `convertToUtcClientTime` convierte un diccionario que contiene la fecha y la hora locales en un objeto Date con los valores correctos para la fecha y la hora locales.</span><span class="sxs-lookup"><span data-stu-id="8be8b-216">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8be8b-217">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8be8b-217">Parameters:</span></span>

|<span data-ttu-id="8be8b-218">Nombre</span><span class="sxs-lookup"><span data-stu-id="8be8b-218">Name</span></span>| <span data-ttu-id="8be8b-219">Tipo</span><span class="sxs-lookup"><span data-stu-id="8be8b-219">Type</span></span>| <span data-ttu-id="8be8b-220">Descripción</span><span class="sxs-lookup"><span data-stu-id="8be8b-220">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="8be8b-221">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="8be8b-221">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="8be8b-222">El valor de la hora local para convertir.</span><span class="sxs-lookup"><span data-stu-id="8be8b-222">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8be8b-223">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8be8b-223">Requirements</span></span>

|<span data-ttu-id="8be8b-224">Requirement</span><span class="sxs-lookup"><span data-stu-id="8be8b-224">Requirement</span></span>| <span data-ttu-id="8be8b-225">Valor</span><span class="sxs-lookup"><span data-stu-id="8be8b-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be8b-226">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8be8b-226">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be8b-227">1.0</span><span class="sxs-lookup"><span data-stu-id="8be8b-227">1.0</span></span>|
|[<span data-ttu-id="8be8b-228">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8be8b-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be8b-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8be8b-229">ReadItem</span></span>|
|[<span data-ttu-id="8be8b-230">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8be8b-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be8b-231">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8be8b-231">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8be8b-232">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="8be8b-232">Returns:</span></span>

<span data-ttu-id="8be8b-233">Objeto Date con el tiempo expresado en UTC.</span><span class="sxs-lookup"><span data-stu-id="8be8b-233">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="8be8b-234">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="8be8b-234">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8be8b-235">Fecha</span><span class="sxs-lookup"><span data-stu-id="8be8b-235">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="8be8b-236">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="8be8b-236">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="8be8b-237">Muestra una cita de calendario existente.</span><span class="sxs-lookup"><span data-stu-id="8be8b-237">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8be8b-238">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8be8b-238">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8be8b-239">El método `displayAppointmentForm` abre una cita de calendario existente en una nueva ventana del escritorio o en un cuadro de diálogo en los dispositivos móviles.</span><span class="sxs-lookup"><span data-stu-id="8be8b-239">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="8be8b-p109">En Outlook para Mac, puede usar este método para mostrar una cita que no forme parte de una serie periódica o la cita principal de una serie periódica, pero no puede mostrar una instancia de la serie. Esto es porque en Outlook para Mac no se puede tener acceso a las propiedades (incluido el identificador de elemento) de las instancias de una serie periódica.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p109">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="8be8b-242">En Outlook Web App, este método abre el formato especificado solo si el cuerpo del formulario es inferior o igual a 32 KB en el número de caracteres.</span><span class="sxs-lookup"><span data-stu-id="8be8b-242">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="8be8b-243">Si el identificador de elemento especificado no identifica una cita existente, se abrirá una página en blanco en el dispositivo o equipo cliente y no se generará ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="8be8b-243">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8be8b-244">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8be8b-244">Parameters:</span></span>

|<span data-ttu-id="8be8b-245">Nombre</span><span class="sxs-lookup"><span data-stu-id="8be8b-245">Name</span></span>| <span data-ttu-id="8be8b-246">Tipo</span><span class="sxs-lookup"><span data-stu-id="8be8b-246">Type</span></span>| <span data-ttu-id="8be8b-247">Descripción</span><span class="sxs-lookup"><span data-stu-id="8be8b-247">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="8be8b-248">String</span><span class="sxs-lookup"><span data-stu-id="8be8b-248">String</span></span>|<span data-ttu-id="8be8b-249">Identificador de los servicios web de Exchange (EWS) para una cita de calendario existente.</span><span class="sxs-lookup"><span data-stu-id="8be8b-249">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8be8b-250">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8be8b-250">Requirements</span></span>

|<span data-ttu-id="8be8b-251">Requirement</span><span class="sxs-lookup"><span data-stu-id="8be8b-251">Requirement</span></span>| <span data-ttu-id="8be8b-252">Valor</span><span class="sxs-lookup"><span data-stu-id="8be8b-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be8b-253">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8be8b-253">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be8b-254">1.0</span><span class="sxs-lookup"><span data-stu-id="8be8b-254">1.0</span></span>|
|[<span data-ttu-id="8be8b-255">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8be8b-255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be8b-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8be8b-256">ReadItem</span></span>|
|[<span data-ttu-id="8be8b-257">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8be8b-257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be8b-258">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8be8b-258">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8be8b-259">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8be8b-259">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="8be8b-260">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="8be8b-260">displayMessageForm(itemId)</span></span>

<span data-ttu-id="8be8b-261">Muestra un mensaje existente.</span><span class="sxs-lookup"><span data-stu-id="8be8b-261">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="8be8b-262">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8be8b-262">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8be8b-263">El método `displayMessageForm` abre un mensaje existente en una nueva ventana del escritorio o en un cuadro de diálogo en los dispositivos móviles.</span><span class="sxs-lookup"><span data-stu-id="8be8b-263">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="8be8b-264">En Outlook Web App, este método abre el formato especificado solo si el cuerpo del formulario es inferior o igual a 32 KB en el número de caracteres.</span><span class="sxs-lookup"><span data-stu-id="8be8b-264">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="8be8b-265">Si el identificador de elemento especificado no identifica un mensaje existente, no se mostrará ningún mensaje en el equipo cliente ni se generará ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="8be8b-265">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="8be8b-p110">No use el valor `displayMessageForm` con un `itemId` que represente una cita. Use el método `displayAppointmentForm` para mostrar una cita existente y `displayNewAppointmentForm` para mostrar un formulario para crear una cita nueva.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8be8b-268">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8be8b-268">Parameters:</span></span>

|<span data-ttu-id="8be8b-269">Nombre</span><span class="sxs-lookup"><span data-stu-id="8be8b-269">Name</span></span>| <span data-ttu-id="8be8b-270">Tipo</span><span class="sxs-lookup"><span data-stu-id="8be8b-270">Type</span></span>| <span data-ttu-id="8be8b-271">Descripción</span><span class="sxs-lookup"><span data-stu-id="8be8b-271">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="8be8b-272">String</span><span class="sxs-lookup"><span data-stu-id="8be8b-272">String</span></span>|<span data-ttu-id="8be8b-273">Identificador de los servicios web de Exchange (EWS) para un mensaje existente.</span><span class="sxs-lookup"><span data-stu-id="8be8b-273">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8be8b-274">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8be8b-274">Requirements</span></span>

|<span data-ttu-id="8be8b-275">Requirement</span><span class="sxs-lookup"><span data-stu-id="8be8b-275">Requirement</span></span>| <span data-ttu-id="8be8b-276">Valor</span><span class="sxs-lookup"><span data-stu-id="8be8b-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be8b-277">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8be8b-277">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be8b-278">1.0</span><span class="sxs-lookup"><span data-stu-id="8be8b-278">1.0</span></span>|
|[<span data-ttu-id="8be8b-279">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8be8b-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be8b-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8be8b-280">ReadItem</span></span>|
|[<span data-ttu-id="8be8b-281">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8be8b-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be8b-282">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8be8b-282">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8be8b-283">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8be8b-283">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="8be8b-284">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="8be8b-284">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="8be8b-285">Muestra un formulario para crear una nueva cita de calendario.</span><span class="sxs-lookup"><span data-stu-id="8be8b-285">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8be8b-286">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8be8b-286">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8be8b-p111">El método `displayNewAppointmentForm` abre un formulario que permite al usuario crear una nueva cita o reunión. Si se especifican parámetros, los campos de formulario de cita se rellenan automáticamente con el contenido de los parámetros.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="8be8b-p112">En Outlook Web App y OWA para dispositivos, este método muestra siempre un formulario con un campo de asistentes. Si no especifica ningún asistente como argumento de entrada, el método muestra un formulario con un botón **Guardar**. Si ha especificado asistentes, el formulario incluirá a los asistentes y un botón **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p112">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="8be8b-p113">En el cliente enriquecido de Outlook y Outlook RT, si se especifica cualquier asistente o recurso en los parámetros `requiredAttendees`, `optionalAttendees` o `resources`, este método muestra un formulario de reunión con un botón **Enviar**. Si no se especifica ningún destinatario, este método muestra un formulario de cita con un botón **Guardar y cerrar**.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="8be8b-294">Si cualquiera de los parámetros supera los límites de tamaño especificados o si se especifica un nombre de parámetro desconocido, se genera una excepción.</span><span class="sxs-lookup"><span data-stu-id="8be8b-294">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8be8b-295">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8be8b-295">Parameters:</span></span>

|<span data-ttu-id="8be8b-296">Nombre</span><span class="sxs-lookup"><span data-stu-id="8be8b-296">Name</span></span>| <span data-ttu-id="8be8b-297">Tipo</span><span class="sxs-lookup"><span data-stu-id="8be8b-297">Type</span></span>| <span data-ttu-id="8be8b-298">Descripción</span><span class="sxs-lookup"><span data-stu-id="8be8b-298">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="8be8b-299">Object</span><span class="sxs-lookup"><span data-stu-id="8be8b-299">Object</span></span> | <span data-ttu-id="8be8b-300">Un diccionario de parámetros que describen la nueva cita.</span><span class="sxs-lookup"><span data-stu-id="8be8b-300">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="8be8b-301">Matriz.&lt;Cadena&gt; &#124; Matriz.&lt;[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="8be8b-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="8be8b-p114">Una matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los asistentes necesarios de la cita. La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="8be8b-304">Matriz.&lt;Cadena&gt; &#124; Matriz.&lt;[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="8be8b-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="8be8b-p115">Una matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los asistentes opcionales de la cita. La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="8be8b-307">Fecha</span><span class="sxs-lookup"><span data-stu-id="8be8b-307">Date</span></span> | <span data-ttu-id="8be8b-308">Un objeto `Date` que especifica la fecha y hora de inicio de la cita.</span><span class="sxs-lookup"><span data-stu-id="8be8b-308">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="8be8b-309">Fecha</span><span class="sxs-lookup"><span data-stu-id="8be8b-309">Date</span></span> | <span data-ttu-id="8be8b-310">Un objeto `Date` que especifica la fecha y hora de finalización de la cita.</span><span class="sxs-lookup"><span data-stu-id="8be8b-310">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="8be8b-311">String</span><span class="sxs-lookup"><span data-stu-id="8be8b-311">String</span></span> | <span data-ttu-id="8be8b-p116">Una cadena que contiene la ubicación de la cita. La cadena está limitada a un máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="8be8b-314">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="8be8b-314">Array.&lt;String&gt;</span></span> | <span data-ttu-id="8be8b-p117">Una matriz de cadenas que contiene los recursos necesarios para la cita. La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="8be8b-317">String</span><span class="sxs-lookup"><span data-stu-id="8be8b-317">String</span></span> | <span data-ttu-id="8be8b-p118">Una cadena que contiene al asunto de la cita. La cadena está limitada a un máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="8be8b-320">String</span><span class="sxs-lookup"><span data-stu-id="8be8b-320">String</span></span> | <span data-ttu-id="8be8b-p119">El cuerpo de la cita. El contenido del cuerpo está limitado a un tamaño máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8be8b-323">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8be8b-323">Requirements</span></span>

|<span data-ttu-id="8be8b-324">Requirement</span><span class="sxs-lookup"><span data-stu-id="8be8b-324">Requirement</span></span>| <span data-ttu-id="8be8b-325">Valor</span><span class="sxs-lookup"><span data-stu-id="8be8b-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be8b-326">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8be8b-326">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be8b-327">1.0</span><span class="sxs-lookup"><span data-stu-id="8be8b-327">1.0</span></span>|
|[<span data-ttu-id="8be8b-328">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8be8b-328">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be8b-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8be8b-329">ReadItem</span></span>|
|[<span data-ttu-id="8be8b-330">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8be8b-330">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be8b-331">Lectura</span><span class="sxs-lookup"><span data-stu-id="8be8b-331">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8be8b-332">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8be8b-332">Example</span></span>

```
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="8be8b-333">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8be8b-333">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="8be8b-334">Obtiene una cadena que contiene un token usado para obtener datos adjuntos o un elemento de Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="8be8b-334">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="8be8b-p120">El método `getCallbackTokenAsync` realiza una llamada asincrónica para obtener un token opaco desde Exchange Server que hospeda el buzón del usuario. La duración del token de devolución de llamada es de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p120">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="8be8b-p121">Puede pasar el token y un identificador de archivo adjunto o el identificador del elemento a un sistema de terceros. El sistema de terceros usa el token como token de autorización del portador para llamar a los Servicios web de Exchange (EWS) o a las operaciones [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) o [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) para devolver datos adjuntos o un elemento. Por ejemplo, puede crear un servicio remoto para [obtener datos adjuntos desde el elemento seleccionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="8be8b-p121">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="8be8b-340">La aplicación debe tener el permiso **ReadItem** especificado en su manifiesto para poder llamar al método `getCallbackTokenAsync` en el modo lectura.</span><span class="sxs-lookup"><span data-stu-id="8be8b-340">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="8be8b-p122">En el modo de redacción debe llamar al método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obtener un identificador de elemento para pasar al método `getCallbackTokenAsync`. Su aplicación debe tener permisos **ReadWriteItem** para llamar al método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p122">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8be8b-343">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8be8b-343">Parameters:</span></span>

|<span data-ttu-id="8be8b-344">Nombre</span><span class="sxs-lookup"><span data-stu-id="8be8b-344">Name</span></span>| <span data-ttu-id="8be8b-345">Tipo</span><span class="sxs-lookup"><span data-stu-id="8be8b-345">Type</span></span>| <span data-ttu-id="8be8b-346">Atributos</span><span class="sxs-lookup"><span data-stu-id="8be8b-346">Attributes</span></span>| <span data-ttu-id="8be8b-347">Descripción</span><span class="sxs-lookup"><span data-stu-id="8be8b-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8be8b-348">función</span><span class="sxs-lookup"><span data-stu-id="8be8b-348">function</span></span>||<span data-ttu-id="8be8b-p123">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). El token se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p123">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="8be8b-351">Objeto</span><span class="sxs-lookup"><span data-stu-id="8be8b-351">Object</span></span>| <span data-ttu-id="8be8b-352">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8be8b-352">&lt;optional&gt;</span></span>|<span data-ttu-id="8be8b-353">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="8be8b-353">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8be8b-354">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8be8b-354">Requirements</span></span>

|<span data-ttu-id="8be8b-355">Requirement</span><span class="sxs-lookup"><span data-stu-id="8be8b-355">Requirement</span></span>| <span data-ttu-id="8be8b-356">Valor</span><span class="sxs-lookup"><span data-stu-id="8be8b-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be8b-357">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8be8b-357">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be8b-358">1.3</span><span class="sxs-lookup"><span data-stu-id="8be8b-358">1.3</span></span>|
|[<span data-ttu-id="8be8b-359">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8be8b-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be8b-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8be8b-360">ReadItem</span></span>|
|[<span data-ttu-id="8be8b-361">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8be8b-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be8b-362">Redacción y lectura</span><span class="sxs-lookup"><span data-stu-id="8be8b-362">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="8be8b-363">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8be8b-363">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="8be8b-364">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8be8b-364">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="8be8b-365">Obtiene un token que identifica al usuario y al complemento de Office.</span><span class="sxs-lookup"><span data-stu-id="8be8b-365">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="8be8b-366">El método `getUserIdentityTokenAsync` devuelve un token que puede usar para identificar y [autenticar el complemento y el usuario mediante un sistema de terceros](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="8be8b-366">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="8be8b-367">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8be8b-367">Parameters:</span></span>

|<span data-ttu-id="8be8b-368">Nombre</span><span class="sxs-lookup"><span data-stu-id="8be8b-368">Name</span></span>| <span data-ttu-id="8be8b-369">Tipo</span><span class="sxs-lookup"><span data-stu-id="8be8b-369">Type</span></span>| <span data-ttu-id="8be8b-370">Atributos</span><span class="sxs-lookup"><span data-stu-id="8be8b-370">Attributes</span></span>| <span data-ttu-id="8be8b-371">Descripción</span><span class="sxs-lookup"><span data-stu-id="8be8b-371">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8be8b-372">función</span><span class="sxs-lookup"><span data-stu-id="8be8b-372">function</span></span>||<span data-ttu-id="8be8b-373">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8be8b-373">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8be8b-374">El token se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8be8b-374">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="8be8b-375">Object</span><span class="sxs-lookup"><span data-stu-id="8be8b-375">Object</span></span>| <span data-ttu-id="8be8b-376">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8be8b-376">&lt;optional&gt;</span></span>|<span data-ttu-id="8be8b-377">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="8be8b-377">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8be8b-378">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8be8b-378">Requirements</span></span>

|<span data-ttu-id="8be8b-379">Requirement</span><span class="sxs-lookup"><span data-stu-id="8be8b-379">Requirement</span></span>| <span data-ttu-id="8be8b-380">Valor</span><span class="sxs-lookup"><span data-stu-id="8be8b-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be8b-381">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8be8b-381">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be8b-382">1.0</span><span class="sxs-lookup"><span data-stu-id="8be8b-382">1.0</span></span>|
|[<span data-ttu-id="8be8b-383">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8be8b-383">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be8b-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8be8b-384">ReadItem</span></span>|
|[<span data-ttu-id="8be8b-385">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8be8b-385">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be8b-386">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8be8b-386">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8be8b-387">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8be8b-387">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="8be8b-388">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8be8b-388">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="8be8b-389">Realiza una solicitud asincrónica a un servicio de Servicios Web Exchange (EWS) en el servidor Exchange que hospeda el buzón del usuario.</span><span class="sxs-lookup"><span data-stu-id="8be8b-389">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="8be8b-390">Este método no se admite en los siguientes escenarios.</span><span class="sxs-lookup"><span data-stu-id="8be8b-390">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="8be8b-391">En Outlook para iOS o Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="8be8b-391">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="8be8b-392">Cuando el complemento se carga en un buzón de Gmail</span><span class="sxs-lookup"><span data-stu-id="8be8b-392">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="8be8b-393">En estos casos, complementos deben [usar API de REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) para tener acceso al buzón del usuario en su lugar.</span><span class="sxs-lookup"><span data-stu-id="8be8b-393">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="8be8b-394">El método `makeEwsRequestAsync` envía una solicitud de EWS en nombre del complemento a Exchange.</span><span class="sxs-lookup"><span data-stu-id="8be8b-394">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="8be8b-395">Para obtener una lista de las operaciones de EWS compatibles, vea [llamar a servicios web desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="8be8b-395">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="8be8b-396">No puede solicitar elementos asociados de las carpetas con el método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="8be8b-396">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="8be8b-397">La solicitud XML debe especificar codificación UTF-8.</span><span class="sxs-lookup"><span data-stu-id="8be8b-397">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="8be8b-p125">Su complemento debe tener el permiso **ReadWriteMailbox** para usar el método `makeEwsRequestAsync`. Para obtener información sobre cómo usar el permiso **ReadWriteMailbox** y sobre las operaciones EWS que puede llamar con el método `makeEwsRequestAsync`, consulte [Especificar permisos para el acceso del complemento de correo al buzón del usuario](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="8be8b-p125">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="8be8b-400">El administrador del servidor debe establecer `OAuthAuthentication` en true en el directorio EWS del servidor de acceso de cliente para habilitar la `makeEwsRequestAsync` método para realizar EWS solicita.</span><span class="sxs-lookup"><span data-stu-id="8be8b-400">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="8be8b-401">Diferencias de versión</span><span class="sxs-lookup"><span data-stu-id="8be8b-401">Version differences</span></span>

<span data-ttu-id="8be8b-402">Si usa el método `makeEwsRequestAsync` en aplicaciones de correo que se ejecutan en versiones de Outlook anteriores a 15.0.4535.1004, debe establecer el valor de codificación a `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="8be8b-402">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="8be8b-p126">No es necesario establecer el valor de codificación si la aplicación de correo se ejecuta en Outlook en la web. Puede averiguar si su aplicación de correo se ejecuta en Outlook o en Outlook en la web usando la propiedad mailbox.diagnostics.hostName. Para averiguar qué versión de Outlook se está ejecutando, use la propiedad mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="8be8b-p126">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8be8b-406">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8be8b-406">Parameters:</span></span>

|<span data-ttu-id="8be8b-407">Nombre</span><span class="sxs-lookup"><span data-stu-id="8be8b-407">Name</span></span>| <span data-ttu-id="8be8b-408">Tipo</span><span class="sxs-lookup"><span data-stu-id="8be8b-408">Type</span></span>| <span data-ttu-id="8be8b-409">Atributos</span><span class="sxs-lookup"><span data-stu-id="8be8b-409">Attributes</span></span>| <span data-ttu-id="8be8b-410">Descripción</span><span class="sxs-lookup"><span data-stu-id="8be8b-410">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="8be8b-411">String</span><span class="sxs-lookup"><span data-stu-id="8be8b-411">String</span></span>||<span data-ttu-id="8be8b-412">La solicitud de EWS.</span><span class="sxs-lookup"><span data-stu-id="8be8b-412">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="8be8b-413">función</span><span class="sxs-lookup"><span data-stu-id="8be8b-413">function</span></span>||<span data-ttu-id="8be8b-414">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8be8b-414">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8be8b-415">El resultado XML de la llamada EWS se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8be8b-415">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="8be8b-416">Si el resultado supera 1 MB de tamaño, se devuelve un mensaje de error en su lugar.</span><span class="sxs-lookup"><span data-stu-id="8be8b-416">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="8be8b-417">Objeto</span><span class="sxs-lookup"><span data-stu-id="8be8b-417">Object</span></span>| <span data-ttu-id="8be8b-418">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8be8b-418">&lt;optional&gt;</span></span>|<span data-ttu-id="8be8b-419">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="8be8b-419">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8be8b-420">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8be8b-420">Requirements</span></span>

|<span data-ttu-id="8be8b-421">Requirement</span><span class="sxs-lookup"><span data-stu-id="8be8b-421">Requirement</span></span>| <span data-ttu-id="8be8b-422">Valor</span><span class="sxs-lookup"><span data-stu-id="8be8b-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be8b-423">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8be8b-423">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be8b-424">1.0</span><span class="sxs-lookup"><span data-stu-id="8be8b-424">1.0</span></span>|
|[<span data-ttu-id="8be8b-425">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8be8b-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be8b-426">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="8be8b-426">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="8be8b-427">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8be8b-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be8b-428">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8be8b-428">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8be8b-429">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8be8b-429">Example</span></span>

<span data-ttu-id="8be8b-430">En el siguiente ejemplo, se llama a `makeEwsRequestAsync` para usar la operación `GetItem` para obtener el asunto de un elemento.</span><span class="sxs-lookup"><span data-stu-id="8be8b-430">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item.
   var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
}
```