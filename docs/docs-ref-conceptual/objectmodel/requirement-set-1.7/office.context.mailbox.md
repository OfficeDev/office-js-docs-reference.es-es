
# <a name="mailbox"></a><span data-ttu-id="8e5ad-101">buzón de correo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-101">mailbox</span></span>

### <span data-ttu-id="8e5ad-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="8e5ad-104">Proporciona acceso al modelo de objeto de complemento de Outlook para Microsoft Outlook y Microsoft Outlook en el web.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e5ad-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-105">Requirements</span></span>

|<span data-ttu-id="8e5ad-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-106">Requirement</span></span>| <span data-ttu-id="8e5ad-107">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-108">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8e5ad-109">1.0</span></span>|
|[<span data-ttu-id="8e5ad-110">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-111">Restringido</span><span class="sxs-lookup"><span data-stu-id="8e5ad-111">Restricted</span></span>|
|[<span data-ttu-id="8e5ad-112">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-113">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8e5ad-114">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-114">Members and methods</span></span>

| <span data-ttu-id="8e5ad-115">Miembro	</span><span class="sxs-lookup"><span data-stu-id="8e5ad-115">Member</span></span> | <span data-ttu-id="8e5ad-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8e5ad-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="8e5ad-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="8e5ad-118">Miembro	</span><span class="sxs-lookup"><span data-stu-id="8e5ad-118">Member</span></span> |
| [<span data-ttu-id="8e5ad-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="8e5ad-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="8e5ad-120">Miembro	</span><span class="sxs-lookup"><span data-stu-id="8e5ad-120">Member</span></span> |
| [<span data-ttu-id="8e5ad-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="8e5ad-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="8e5ad-122">Método</span><span class="sxs-lookup"><span data-stu-id="8e5ad-122">Method</span></span> |
| [<span data-ttu-id="8e5ad-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="8e5ad-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="8e5ad-124">Método</span><span class="sxs-lookup"><span data-stu-id="8e5ad-124">Method</span></span> |
| [<span data-ttu-id="8e5ad-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="8e5ad-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) | <span data-ttu-id="8e5ad-126">Método</span><span class="sxs-lookup"><span data-stu-id="8e5ad-126">Method</span></span> |
| [<span data-ttu-id="8e5ad-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="8e5ad-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="8e5ad-128">Método</span><span class="sxs-lookup"><span data-stu-id="8e5ad-128">Method</span></span> |
| [<span data-ttu-id="8e5ad-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="8e5ad-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="8e5ad-130">Método</span><span class="sxs-lookup"><span data-stu-id="8e5ad-130">Method</span></span> |
| [<span data-ttu-id="8e5ad-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="8e5ad-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="8e5ad-132">Método</span><span class="sxs-lookup"><span data-stu-id="8e5ad-132">Method</span></span> |
| [<span data-ttu-id="8e5ad-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="8e5ad-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="8e5ad-134">Método</span><span class="sxs-lookup"><span data-stu-id="8e5ad-134">Method</span></span> |
| [<span data-ttu-id="8e5ad-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="8e5ad-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="8e5ad-136">Método</span><span class="sxs-lookup"><span data-stu-id="8e5ad-136">Method</span></span> |
| [<span data-ttu-id="8e5ad-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="8e5ad-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="8e5ad-138">Método</span><span class="sxs-lookup"><span data-stu-id="8e5ad-138">Method</span></span> |
| [<span data-ttu-id="8e5ad-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="8e5ad-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="8e5ad-140">Método</span><span class="sxs-lookup"><span data-stu-id="8e5ad-140">Method</span></span> |
| [<span data-ttu-id="8e5ad-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="8e5ad-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="8e5ad-142">Método</span><span class="sxs-lookup"><span data-stu-id="8e5ad-142">Method</span></span> |
| [<span data-ttu-id="8e5ad-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="8e5ad-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="8e5ad-144">Método</span><span class="sxs-lookup"><span data-stu-id="8e5ad-144">Method</span></span> |
| [<span data-ttu-id="8e5ad-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="8e5ad-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="8e5ad-146">Método</span><span class="sxs-lookup"><span data-stu-id="8e5ad-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="8e5ad-147">Espacios de nombres</span><span class="sxs-lookup"><span data-stu-id="8e5ad-147">Namespaces</span></span>

<span data-ttu-id="8e5ad-148">[diagnostics](Office.context.mailbox.diagnostics.md): Proporciona información de diagnóstico a un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="8e5ad-149">[item](Office.context.mailbox.item.md): Proporciona métodos y propiedades para tener acceso a un mensaje o cita en un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="8e5ad-150">[userProfile](Office.context.mailbox.userProfile.md): Proporciona información sobre el usuario en un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="8e5ad-151">Miembros</span><span class="sxs-lookup"><span data-stu-id="8e5ad-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="8e5ad-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-152">ewsUrl :String</span></span>

<span data-ttu-id="8e5ad-p102">Obtiene la dirección URL del punto de conexión de Servicios Web Exchange (EWS) para esta cuenta de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8e5ad-155">Este miembro no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e5ad-p103">El valor `ewsUrl` puede usarse por un servicio remoto para realizar llamadas EWS al buzón del usuario. Por ejemplo, puede crear un servicio remoto para [obtener datos adjuntos desde el elemento seleccionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="8e5ad-158">La aplicación debe tener el permiso **ReadItem** especificado en su manifiesto para poder llamar al miembro `ewsUrl` en el modo lectura.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="8e5ad-p104">En el modo de redacción debe llamar al método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar el miembro `ewsUrl`. Su aplicación debe tener permisos **ReadWriteItem** para llamar al método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="8e5ad-161">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-161">Type:</span></span>

*   <span data-ttu-id="8e5ad-162">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e5ad-163">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-163">Requirements</span></span>

|<span data-ttu-id="8e5ad-164">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-164">Requirement</span></span>| <span data-ttu-id="8e5ad-165">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-166">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-166">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-167">1.0</span><span class="sxs-lookup"><span data-stu-id="8e5ad-167">1.0</span></span>|
|[<span data-ttu-id="8e5ad-168">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e5ad-169">ReadItem</span></span>|
|[<span data-ttu-id="8e5ad-170">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-171">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="8e5ad-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-172">restUrl :String</span></span>

<span data-ttu-id="8e5ad-173">Obtiene la URL del punto de conexión REST para esta cuenta de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="8e5ad-174">El valor `restUrl` puede usarse para realizar llamadas [API de REST](https://docs.microsoft.com/outlook/rest/) al buzón del usuario.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="8e5ad-175">La aplicación debe tener el permiso **ReadItem** especificado en su manifiesto para poder llamar al miembro `restUrl` en el modo lectura.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="8e5ad-p105">En el modo de redacción debe llamar al método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar el miembro `restUrl`. Su aplicación debe tener permisos **ReadWriteItem** para llamar al método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="8e5ad-178">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-178">Type:</span></span>

*   <span data-ttu-id="8e5ad-179">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e5ad-180">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-180">Requirements</span></span>

|<span data-ttu-id="8e5ad-181">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-181">Requirement</span></span>| <span data-ttu-id="8e5ad-182">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-183">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-183">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-184">1,5</span><span class="sxs-lookup"><span data-stu-id="8e5ad-184">1.5</span></span> |
|[<span data-ttu-id="8e5ad-185">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e5ad-186">ReadItem</span></span>|
|[<span data-ttu-id="8e5ad-187">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-188">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="8e5ad-189">Métodos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="8e5ad-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8e5ad-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="8e5ad-191">Agrega un controlador de eventos para un evento admitido.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="8e5ad-192">Actualmente, los tipos de eventos admitidos son `Office.EventType.ItemChanged` y `Office.EventType.OfficeThemeChanged`.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-192">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e5ad-193">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-193">Parameters:</span></span>

| <span data-ttu-id="8e5ad-194">Nombre</span><span class="sxs-lookup"><span data-stu-id="8e5ad-194">Name</span></span> | <span data-ttu-id="8e5ad-195">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-195">Type</span></span> | <span data-ttu-id="8e5ad-196">Atributos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-196">Attributes</span></span> | <span data-ttu-id="8e5ad-197">Descripción</span><span class="sxs-lookup"><span data-stu-id="8e5ad-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="8e5ad-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="8e5ad-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="8e5ad-199">El evento que debe invocar el controlador.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="8e5ad-200">Función</span><span class="sxs-lookup"><span data-stu-id="8e5ad-200">Function</span></span> || <span data-ttu-id="8e5ad-p106">La función que va a controlar el evento. La función debe aceptar un único parámetro, que es un literal de objeto. La propiedad `type` del parámetro coincidirá con el parámetro `eventType` que se ha pasado a `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="8e5ad-204">Objeto</span><span class="sxs-lookup"><span data-stu-id="8e5ad-204">Object</span></span> | <span data-ttu-id="8e5ad-205">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-205">&lt;optional&gt;</span></span> | <span data-ttu-id="8e5ad-206">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="8e5ad-207">Objeto</span><span class="sxs-lookup"><span data-stu-id="8e5ad-207">Object</span></span> | <span data-ttu-id="8e5ad-208">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-208">&lt;optional&gt;</span></span> | <span data-ttu-id="8e5ad-209">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="8e5ad-210">función</span><span class="sxs-lookup"><span data-stu-id="8e5ad-210">function</span></span>| <span data-ttu-id="8e5ad-211">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-211">&lt;optional&gt;</span></span>|<span data-ttu-id="8e5ad-212">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e5ad-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e5ad-213">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-213">Requirements</span></span>

|<span data-ttu-id="8e5ad-214">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-214">Requirement</span></span>| <span data-ttu-id="8e5ad-215">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-216">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-216">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-217">1,5</span><span class="sxs-lookup"><span data-stu-id="8e5ad-217">1.5</span></span> |
|[<span data-ttu-id="8e5ad-218">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e5ad-219">ReadItem</span></span> |
|[<span data-ttu-id="8e5ad-220">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-221">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-221">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e5ad-222">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-222">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="8e5ad-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="8e5ad-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="8e5ad-224">Convierte un identificador de elemento con formato para REST al formato EWS.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="8e5ad-225">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-225">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e5ad-p107">Los identificadores de elemento obtenidos a través de una API de REST (como la [API de correo de Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) o [Microsoft Graph](http://graph.microsoft.io/)) usan un formato diferente al formato que usa Exchange Web Services (EWS). El método `convertToEwsId` convierte un identificador con formato REST al formato adecuado para EWS.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e5ad-228">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-228">Parameters:</span></span>

|<span data-ttu-id="8e5ad-229">Nombre</span><span class="sxs-lookup"><span data-stu-id="8e5ad-229">Name</span></span>| <span data-ttu-id="8e5ad-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-230">Type</span></span>| <span data-ttu-id="8e5ad-231">Descripción</span><span class="sxs-lookup"><span data-stu-id="8e5ad-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="8e5ad-232">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-232">String</span></span>|<span data-ttu-id="8e5ad-233">Un identificador de elemento con formato para las API de REST de Outlook</span><span class="sxs-lookup"><span data-stu-id="8e5ad-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="8e5ad-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="8e5ad-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="8e5ad-235">Un valor que indica la versión de la API de REST de Outlook que se usa para recuperar el identificador de elemento.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e5ad-236">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-236">Requirements</span></span>

|<span data-ttu-id="8e5ad-237">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-237">Requirement</span></span>| <span data-ttu-id="8e5ad-238">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-239">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-239">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-240">1.3</span><span class="sxs-lookup"><span data-stu-id="8e5ad-240">1.3</span></span>|
|[<span data-ttu-id="8e5ad-241">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-241">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-242">Restringido</span><span class="sxs-lookup"><span data-stu-id="8e5ad-242">Restricted</span></span>|
|[<span data-ttu-id="8e5ad-243">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-243">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-244">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-244">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e5ad-245">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-245">Returns:</span></span>

<span data-ttu-id="8e5ad-246">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="8e5ad-247">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-247">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime"></a><span data-ttu-id="8e5ad-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="8e5ad-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span></span>

<span data-ttu-id="8e5ad-249">Obtiene un diccionario con información de tiempo en el tiempo del cliente local.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="8e5ad-p108">Las fechas y horas usadas por una aplicación de correo para Outlook o Outlook Web App pueden usar distintas zonas horarias. Outlook usa la zona horaria del equipo cliente; Outlook Web App usa la zona horaria definida en el Centro de administración de Exchange (EAC). Debería tratar los valores de fecha y hora de modo que los valores que aparezcan en la interfaz de usuario sean siempre coherentes con la zona horaria que el usuario espera.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="8e5ad-p109">Si se está ejecutando la aplicación de correo en Outlook, el método `convertToLocalClientTime` devolverá un objeto de diccionario con los valores establecidos para la zona horaria del equipo cliente. Si se está ejecutando la aplicación de correo en Outlook Web App, el método `convertToLocalClientTime` devolverá un objeto de diccionario con los valores establecidos para la zona horaria especificada en el CEF.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e5ad-255">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-255">Parameters:</span></span>

|<span data-ttu-id="8e5ad-256">Nombre</span><span class="sxs-lookup"><span data-stu-id="8e5ad-256">Name</span></span>| <span data-ttu-id="8e5ad-257">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-257">Type</span></span>| <span data-ttu-id="8e5ad-258">Descripción</span><span class="sxs-lookup"><span data-stu-id="8e5ad-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="8e5ad-259">Fecha</span><span class="sxs-lookup"><span data-stu-id="8e5ad-259">Date</span></span>|<span data-ttu-id="8e5ad-260">Un objeto Date</span><span class="sxs-lookup"><span data-stu-id="8e5ad-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e5ad-261">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-261">Requirements</span></span>

|<span data-ttu-id="8e5ad-262">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-262">Requirement</span></span>| <span data-ttu-id="8e5ad-263">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-264">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-265">1.0</span><span class="sxs-lookup"><span data-stu-id="8e5ad-265">1.0</span></span>|
|[<span data-ttu-id="8e5ad-266">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e5ad-267">ReadItem</span></span>|
|[<span data-ttu-id="8e5ad-268">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-269">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-269">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e5ad-270">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-270">Returns:</span></span>

<span data-ttu-id="8e5ad-271">Tipo: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="8e5ad-271">Type: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="8e5ad-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="8e5ad-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="8e5ad-273">Convierte un identificador de elemento con formato para EWS al formato REST.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="8e5ad-274">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-274">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e5ad-p110">Los identificadores de elemento obtenidos a través de EWS o de la propiedad `itemId` usan un formato diferente al formato que usan las API de REST (como la [API de correo de Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) o [Microsoft Graph](http://graph.microsoft.io/)). El método `convertToRestId` convierte un identificador con formato EWS al formato adecuado para REST.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e5ad-277">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-277">Parameters:</span></span>

|<span data-ttu-id="8e5ad-278">Nombre</span><span class="sxs-lookup"><span data-stu-id="8e5ad-278">Name</span></span>| <span data-ttu-id="8e5ad-279">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-279">Type</span></span>| <span data-ttu-id="8e5ad-280">Descripción</span><span class="sxs-lookup"><span data-stu-id="8e5ad-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="8e5ad-281">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-281">String</span></span>|<span data-ttu-id="8e5ad-282">Un identificador de elemento con formato para Exchange Web Services (EWS)</span><span class="sxs-lookup"><span data-stu-id="8e5ad-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="8e5ad-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="8e5ad-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="8e5ad-284">Un valor que indica la versión de la API de REST de Outlook con la que se usará el identificador convertido.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e5ad-285">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-285">Requirements</span></span>

|<span data-ttu-id="8e5ad-286">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-286">Requirement</span></span>| <span data-ttu-id="8e5ad-287">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-288">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-288">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-289">1.3</span><span class="sxs-lookup"><span data-stu-id="8e5ad-289">1.3</span></span>|
|[<span data-ttu-id="8e5ad-290">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-291">Restringido</span><span class="sxs-lookup"><span data-stu-id="8e5ad-291">Restricted</span></span>|
|[<span data-ttu-id="8e5ad-292">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-293">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-293">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e5ad-294">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-294">Returns:</span></span>

<span data-ttu-id="8e5ad-295">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="8e5ad-296">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-296">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="8e5ad-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="8e5ad-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="8e5ad-298">Obtiene un objeto Date del diccionario que contiene información de tiempo.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="8e5ad-299">El método `convertToUtcClientTime` convierte un diccionario que contiene la fecha y la hora locales en un objeto Date con los valores correctos para la fecha y la hora locales.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e5ad-300">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-300">Parameters:</span></span>

|<span data-ttu-id="8e5ad-301">Nombre</span><span class="sxs-lookup"><span data-stu-id="8e5ad-301">Name</span></span>| <span data-ttu-id="8e5ad-302">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-302">Type</span></span>| <span data-ttu-id="8e5ad-303">Descripción</span><span class="sxs-lookup"><span data-stu-id="8e5ad-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="8e5ad-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="8e5ad-304">LocalClientTime</span></span>](/javascript/api/outlook_1_7/office.LocalClientTime)|<span data-ttu-id="8e5ad-305">El valor de la hora local para convertir.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e5ad-306">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-306">Requirements</span></span>

|<span data-ttu-id="8e5ad-307">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-307">Requirement</span></span>| <span data-ttu-id="8e5ad-308">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-309">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-309">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-310">1.0</span><span class="sxs-lookup"><span data-stu-id="8e5ad-310">1.0</span></span>|
|[<span data-ttu-id="8e5ad-311">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-311">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e5ad-312">ReadItem</span></span>|
|[<span data-ttu-id="8e5ad-313">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-313">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-314">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-314">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e5ad-315">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-315">Returns:</span></span>

<span data-ttu-id="8e5ad-316">Objeto Date con el tiempo expresado en UTC.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-316">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="8e5ad-317">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="8e5ad-317">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8e5ad-318">Fecha</span><span class="sxs-lookup"><span data-stu-id="8e5ad-318">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="8e5ad-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="8e5ad-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="8e5ad-320">Muestra una cita de calendario existente.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8e5ad-321">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-321">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e5ad-322">El método `displayAppointmentForm` abre una cita de calendario existente en una nueva ventana del escritorio o en un cuadro de diálogo en los dispositivos móviles.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="8e5ad-p111">En Outlook para Mac, puede usar este método para mostrar una cita que no forme parte de una serie periódica o la cita principal de una serie periódica, pero no puede mostrar una instancia de la serie. Esto es porque en Outlook para Mac no se puede tener acceso a las propiedades (incluido el identificador de elemento) de las instancias de una serie periódica.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="8e5ad-325">En Outlook Web App, este método abre el formato especificado solo si el cuerpo del formulario es inferior o igual a 32 KB en el número de caracteres.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-325">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="8e5ad-326">Si el identificador de elemento especificado no identifica una cita existente, se abrirá una página en blanco en el dispositivo o equipo cliente y no se generará ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e5ad-327">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-327">Parameters:</span></span>

|<span data-ttu-id="8e5ad-328">Nombre</span><span class="sxs-lookup"><span data-stu-id="8e5ad-328">Name</span></span>| <span data-ttu-id="8e5ad-329">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-329">Type</span></span>| <span data-ttu-id="8e5ad-330">Descripción</span><span class="sxs-lookup"><span data-stu-id="8e5ad-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="8e5ad-331">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-331">String</span></span>|<span data-ttu-id="8e5ad-332">Identificador de los servicios web de Exchange (EWS) para una cita de calendario existente.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e5ad-333">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-333">Requirements</span></span>

|<span data-ttu-id="8e5ad-334">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-334">Requirement</span></span>| <span data-ttu-id="8e5ad-335">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-336">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-337">1.0</span><span class="sxs-lookup"><span data-stu-id="8e5ad-337">1.0</span></span>|
|[<span data-ttu-id="8e5ad-338">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e5ad-339">ReadItem</span></span>|
|[<span data-ttu-id="8e5ad-340">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-341">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e5ad-342">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-342">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="8e5ad-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="8e5ad-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="8e5ad-344">Muestra un mensaje existente.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="8e5ad-345">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-345">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e5ad-346">El método `displayMessageForm` abre un mensaje existente en una nueva ventana del escritorio o en un cuadro de diálogo en los dispositivos móviles.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="8e5ad-347">En Outlook Web App, este método abre el formato especificado solo si el cuerpo del formulario es inferior o igual a 32 KB en el número de caracteres.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-347">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="8e5ad-348">Si el identificador de elemento especificado no identifica un mensaje existente, no se mostrará ningún mensaje en el equipo cliente ni se generará ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="8e5ad-p112">No use el valor `displayMessageForm` con un `itemId` que represente una cita. Use el método `displayAppointmentForm` para mostrar una cita existente y `displayNewAppointmentForm` para mostrar un formulario para crear una cita nueva.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e5ad-351">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-351">Parameters:</span></span>

|<span data-ttu-id="8e5ad-352">Nombre</span><span class="sxs-lookup"><span data-stu-id="8e5ad-352">Name</span></span>| <span data-ttu-id="8e5ad-353">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-353">Type</span></span>| <span data-ttu-id="8e5ad-354">Descripción</span><span class="sxs-lookup"><span data-stu-id="8e5ad-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="8e5ad-355">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-355">String</span></span>|<span data-ttu-id="8e5ad-356">Identificador de los servicios web de Exchange (EWS) para un mensaje existente.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e5ad-357">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-357">Requirements</span></span>

|<span data-ttu-id="8e5ad-358">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-358">Requirement</span></span>| <span data-ttu-id="8e5ad-359">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-360">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-360">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-361">1.0</span><span class="sxs-lookup"><span data-stu-id="8e5ad-361">1.0</span></span>|
|[<span data-ttu-id="8e5ad-362">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e5ad-363">ReadItem</span></span>|
|[<span data-ttu-id="8e5ad-364">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-365">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e5ad-366">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-366">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="8e5ad-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="8e5ad-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="8e5ad-368">Muestra un formulario para crear una nueva cita de calendario.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8e5ad-369">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-369">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e5ad-p113">El método `displayNewAppointmentForm` abre un formulario que permite al usuario crear una nueva cita o reunión. Si se especifican parámetros, los campos de formulario de cita se rellenan automáticamente con el contenido de los parámetros.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="8e5ad-p114">En Outlook Web App y OWA para dispositivos, este método muestra siempre un formulario con un campo de asistentes. Si no especifica ningún asistente como argumento de entrada, el método muestra un formulario con un botón **Guardar**. Si ha especificado asistentes, el formulario incluirá a los asistentes y un botón **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="8e5ad-p115">En el cliente enriquecido de Outlook y Outlook RT, si se especifica cualquier asistente o recurso en los parámetros `requiredAttendees`, `optionalAttendees` o `resources`, este método muestra un formulario de reunión con un botón **Enviar**. Si no se especifica ningún destinatario, este método muestra un formulario de cita con un botón **Guardar y cerrar**.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="8e5ad-377">Si cualquiera de los parámetros supera los límites de tamaño especificados o si se especifica un nombre de parámetro desconocido, se genera una excepción.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e5ad-378">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-378">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="8e5ad-379">Todos los parámetros son opcionales.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-379">All parameters are optional.</span></span>

|<span data-ttu-id="8e5ad-380">Nombre</span><span class="sxs-lookup"><span data-stu-id="8e5ad-380">Name</span></span>| <span data-ttu-id="8e5ad-381">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-381">Type</span></span>| <span data-ttu-id="8e5ad-382">Descripción</span><span class="sxs-lookup"><span data-stu-id="8e5ad-382">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="8e5ad-383">Object</span><span class="sxs-lookup"><span data-stu-id="8e5ad-383">Object</span></span> | <span data-ttu-id="8e5ad-384">Un diccionario de parámetros que describen la nueva cita.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-384">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="8e5ad-385">Matriz.&lt;Cadena&gt; &#124; Matriz.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="8e5ad-p116">Una matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los asistentes necesarios de la cita. La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="8e5ad-388">Matriz.&lt;Cadena&gt; &#124; Matriz.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="8e5ad-p117">Una matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los asistentes opcionales de la cita. La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="8e5ad-391">Fecha</span><span class="sxs-lookup"><span data-stu-id="8e5ad-391">Date</span></span> | <span data-ttu-id="8e5ad-392">Un objeto `Date` que especifica la fecha y hora de inicio de la cita.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-392">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="8e5ad-393">Fecha</span><span class="sxs-lookup"><span data-stu-id="8e5ad-393">Date</span></span> | <span data-ttu-id="8e5ad-394">Un objeto `Date` que especifica la fecha y hora de finalización de la cita.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-394">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="8e5ad-395">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-395">String</span></span> | <span data-ttu-id="8e5ad-p118">Una cadena que contiene la ubicación de la cita. La cadena está limitada a un máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="8e5ad-398">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-398">Array.&lt;String&gt;</span></span> | <span data-ttu-id="8e5ad-p119">Una matriz de cadenas que contiene los recursos necesarios para la cita. La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="8e5ad-401">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-401">String</span></span> | <span data-ttu-id="8e5ad-p120">Una cadena que contiene al asunto de la cita. La cadena está limitada a un máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="8e5ad-404">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-404">String</span></span> | <span data-ttu-id="8e5ad-p121">El cuerpo de la cita. El contenido del cuerpo está limitado a un tamaño máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8e5ad-407">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-407">Requirements</span></span>

|<span data-ttu-id="8e5ad-408">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-408">Requirement</span></span>| <span data-ttu-id="8e5ad-409">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-410">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-410">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-411">1.0</span><span class="sxs-lookup"><span data-stu-id="8e5ad-411">1.0</span></span>|
|[<span data-ttu-id="8e5ad-412">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e5ad-413">ReadItem</span></span>|
|[<span data-ttu-id="8e5ad-414">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-415">Lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e5ad-416">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-416">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="8e5ad-417">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="8e5ad-417">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="8e5ad-418">Muestra un formulario para crear un mensaje nuevo.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-418">Displays a form for creating a new message.</span></span>

<span data-ttu-id="8e5ad-419">El método `displayNewMessageForm` abre un formulario que permite al usuario crear un mensaje nuevo.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-419">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="8e5ad-420">Si se especifican parámetros, los campos de formulario de mensaje se rellenan automáticamente con el contenido de los parámetros.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-420">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="8e5ad-421">Si cualquiera de los parámetros supera los límites de tamaño especificados o si se indica un nombre de parámetro desconocido, se genera una excepción.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-421">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e5ad-422">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-422">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="8e5ad-423">Todos los parámetros son opcionales.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-423">All parameters are optional.</span></span>

|<span data-ttu-id="8e5ad-424">Nombre</span><span class="sxs-lookup"><span data-stu-id="8e5ad-424">Name</span></span>| <span data-ttu-id="8e5ad-425">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-425">Type</span></span>| <span data-ttu-id="8e5ad-426">Descripción</span><span class="sxs-lookup"><span data-stu-id="8e5ad-426">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="8e5ad-427">Objeto</span><span class="sxs-lookup"><span data-stu-id="8e5ad-427">Object</span></span> | <span data-ttu-id="8e5ad-428">Diccionario de parámetros que describen el mensaje nuevo.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-428">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="8e5ad-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="8e5ad-430">Matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios de la línea Para.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-430">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="8e5ad-431">La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-431">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="8e5ad-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="8e5ad-433">Matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios de la línea CC.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="8e5ad-434">La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="8e5ad-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="8e5ad-436">Matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios de la línea CCO.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="8e5ad-437">La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="8e5ad-438">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-438">String</span></span> | <span data-ttu-id="8e5ad-439">Cadena que contiene el asunto del mensaje.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-439">A string containing the subject of the message.</span></span> <span data-ttu-id="8e5ad-440">La cadena está limitada a un máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-440">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="8e5ad-441">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-441">String</span></span> | <span data-ttu-id="8e5ad-442">Cuerpo HTML del mensaje.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-442">The HTML body of the message.</span></span> <span data-ttu-id="8e5ad-443">El contenido del cuerpo está limitado a un tamaño máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-443">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="8e5ad-444">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-444">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="8e5ad-445">Matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-445">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="8e5ad-446">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-446">String</span></span> | <span data-ttu-id="8e5ad-p128">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="8e5ad-449">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-449">String</span></span> | <span data-ttu-id="8e5ad-450">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-450">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="8e5ad-451">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-451">String</span></span> | <span data-ttu-id="8e5ad-p129">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="8e5ad-454">Booleano</span><span class="sxs-lookup"><span data-stu-id="8e5ad-454">Boolean</span></span> | <span data-ttu-id="8e5ad-p130">Solo se usa si `type` se establece en `file`. Si `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="8e5ad-457">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-457">String</span></span> | <span data-ttu-id="8e5ad-458">Solo se usa si `type` se establece en `item`.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-458">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="8e5ad-459">El identificador de elemento EWS del correo electrónico existente que desea adjuntar al mensaje de nuevo.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-459">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="8e5ad-460">Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-460">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="8e5ad-461">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-461">Requirements</span></span>

|<span data-ttu-id="8e5ad-462">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-462">Requirement</span></span>| <span data-ttu-id="8e5ad-463">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-463">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-464">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-464">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-465">1.6</span><span class="sxs-lookup"><span data-stu-id="8e5ad-465">1.6</span></span> |
|[<span data-ttu-id="8e5ad-466">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-466">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-467">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e5ad-467">ReadItem</span></span>|
|[<span data-ttu-id="8e5ad-468">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-468">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-469">Lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-469">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e5ad-470">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-470">Example</span></span>

```
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="8e5ad-471">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="8e5ad-471">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="8e5ad-472">Obtiene una cadena que contiene un token usado para llamar a las API de REST o a los servicios Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-472">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="8e5ad-p132">El método `getCallbackTokenAsync` realiza una llamada asincrónica para obtener un token opaco desde Exchange Server que hospeda el buzón del usuario. La duración del token de devolución de llamada es de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="8e5ad-475">Se recomienda que los complementos de usar las API de REST en lugar de los servicios Web Exchange siempre que sea posible.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-475">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="8e5ad-476">**Tokens de REST**</span><span class="sxs-lookup"><span data-stu-id="8e5ad-476">**REST Tokens**</span></span>

<span data-ttu-id="8e5ad-p133">Cuando se solicita un token de REST (`options.isRest = true`), el token resultante no funcionará para autenticar las llamadas a los servicios Web Exchange. El token se limitará en su ámbito para el acceso de solo lectura al elemento actual y a sus datos adjuntos, a no ser que el complemento haya especificado el permiso [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) en su manifiesto. Si está especificado el permiso `ReadWriteMailbox`, el token resultante concederá acceso de lectura o escritura al correo, calendario y contactos, incluida la capacidad de enviar correo.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="8e5ad-480">El complemento debe usar la propiedad `restUrl` para determinar la URL correcta que se va a usar al realizar las llamadas a la API de REST.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-480">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="8e5ad-481">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="8e5ad-481">**EWS Tokens**</span></span>

<span data-ttu-id="8e5ad-p134">Cuando se solicita un token EWS (`options.isRest = false`), el token resultante no funcionará para autenticar las llamadas a las API de REST. El token estará limitado en su ámbito para el acceso al elemento actual.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="8e5ad-484">El complemento debe usar la propiedad `ewsUrl` para determinar la URL correcta que se va a usar al realizar las llamadas a EWS.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-484">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e5ad-485">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-485">Parameters:</span></span>

|<span data-ttu-id="8e5ad-486">Nombre</span><span class="sxs-lookup"><span data-stu-id="8e5ad-486">Name</span></span>| <span data-ttu-id="8e5ad-487">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-487">Type</span></span>| <span data-ttu-id="8e5ad-488">Atributos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-488">Attributes</span></span>| <span data-ttu-id="8e5ad-489">Descripción</span><span class="sxs-lookup"><span data-stu-id="8e5ad-489">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="8e5ad-490">Objeto</span><span class="sxs-lookup"><span data-stu-id="8e5ad-490">Object</span></span> | <span data-ttu-id="8e5ad-491">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-491">&lt;optional&gt;</span></span> | <span data-ttu-id="8e5ad-492">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-492">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="8e5ad-493">Booleano</span><span class="sxs-lookup"><span data-stu-id="8e5ad-493">Boolean</span></span> |  <span data-ttu-id="8e5ad-494">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-494">&lt;optional&gt;</span></span> | <span data-ttu-id="8e5ad-p135">Determina si el token proporcionado se usará para las API de REST de Outlook o para los servicios Web Exchange. El valor predeterminado es `false`.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="8e5ad-497">Objeto</span><span class="sxs-lookup"><span data-stu-id="8e5ad-497">Object</span></span> |  <span data-ttu-id="8e5ad-498">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-498">&lt;optional&gt;</span></span> | <span data-ttu-id="8e5ad-499">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-499">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="8e5ad-500">Función</span><span class="sxs-lookup"><span data-stu-id="8e5ad-500">function</span></span>||<span data-ttu-id="8e5ad-p136">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). El token se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e5ad-503">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-503">Requirements</span></span>

|<span data-ttu-id="8e5ad-504">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-504">Requirement</span></span>| <span data-ttu-id="8e5ad-505">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-506">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-507">1,5</span><span class="sxs-lookup"><span data-stu-id="8e5ad-507">1.5</span></span> |
|[<span data-ttu-id="8e5ad-508">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e5ad-509">ReadItem</span></span>|
|[<span data-ttu-id="8e5ad-510">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-511">Redacción y lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-511">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e5ad-512">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-512">Example</span></span>

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="8e5ad-513">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8e5ad-513">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="8e5ad-514">Obtiene una cadena que contiene un token usado para obtener datos adjuntos o un elemento de Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-514">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="8e5ad-p137">El método `getCallbackTokenAsync` realiza una llamada asincrónica para obtener un token opaco desde Exchange Server que hospeda el buzón del usuario. La duración del token de devolución de llamada es de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="8e5ad-p138">Puede pasar el token y un identificador de archivo adjunto o el identificador del elemento a un sistema de terceros. El sistema de terceros usa el token como token de autorización del portador para llamar a los Servicios web de Exchange (EWS) o a las operaciones [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) o [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) para devolver datos adjuntos o un elemento. Por ejemplo, puede crear un servicio remoto para [obtener datos adjuntos desde el elemento seleccionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="8e5ad-520">La aplicación debe tener el permiso **ReadItem** especificado en su manifiesto para poder llamar al método `getCallbackTokenAsync` en el modo lectura.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-520">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="8e5ad-p139">En el modo de redacción debe llamar al método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obtener un identificador de elemento para pasar al método `getCallbackTokenAsync`. Su aplicación debe tener permisos **ReadWriteItem** para llamar al método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e5ad-523">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-523">Parameters:</span></span>

|<span data-ttu-id="8e5ad-524">Nombre</span><span class="sxs-lookup"><span data-stu-id="8e5ad-524">Name</span></span>| <span data-ttu-id="8e5ad-525">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-525">Type</span></span>| <span data-ttu-id="8e5ad-526">Atributos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-526">Attributes</span></span>| <span data-ttu-id="8e5ad-527">Descripción</span><span class="sxs-lookup"><span data-stu-id="8e5ad-527">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8e5ad-528">función</span><span class="sxs-lookup"><span data-stu-id="8e5ad-528">function</span></span>||<span data-ttu-id="8e5ad-p140">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). El token se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="8e5ad-531">Objeto</span><span class="sxs-lookup"><span data-stu-id="8e5ad-531">Object</span></span>| <span data-ttu-id="8e5ad-532">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-532">&lt;optional&gt;</span></span>|<span data-ttu-id="8e5ad-533">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-533">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e5ad-534">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-534">Requirements</span></span>

|<span data-ttu-id="8e5ad-535">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-535">Requirement</span></span>| <span data-ttu-id="8e5ad-536">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-537">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-537">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-538">1.3</span><span class="sxs-lookup"><span data-stu-id="8e5ad-538">1.3</span></span>|
|[<span data-ttu-id="8e5ad-539">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-539">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e5ad-540">ReadItem</span></span>|
|[<span data-ttu-id="8e5ad-541">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-541">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-542">Redacción y lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-542">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e5ad-543">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-543">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="8e5ad-544">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8e5ad-544">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="8e5ad-545">Obtiene un token que identifica al usuario y al complemento de Office.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-545">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="8e5ad-546">El método `getUserIdentityTokenAsync` devuelve un token que puede usar para identificar y [autenticar el complemento y el usuario mediante un sistema de terceros](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="8e5ad-546">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e5ad-547">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-547">Parameters:</span></span>

|<span data-ttu-id="8e5ad-548">Nombre</span><span class="sxs-lookup"><span data-stu-id="8e5ad-548">Name</span></span>| <span data-ttu-id="8e5ad-549">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-549">Type</span></span>| <span data-ttu-id="8e5ad-550">Atributos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-550">Attributes</span></span>| <span data-ttu-id="8e5ad-551">Descripción</span><span class="sxs-lookup"><span data-stu-id="8e5ad-551">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8e5ad-552">función</span><span class="sxs-lookup"><span data-stu-id="8e5ad-552">function</span></span>||<span data-ttu-id="8e5ad-553">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e5ad-553">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8e5ad-554">El token se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-554">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="8e5ad-555">Object</span><span class="sxs-lookup"><span data-stu-id="8e5ad-555">Object</span></span>| <span data-ttu-id="8e5ad-556">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-556">&lt;optional&gt;</span></span>|<span data-ttu-id="8e5ad-557">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-557">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e5ad-558">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-558">Requirements</span></span>

|<span data-ttu-id="8e5ad-559">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-559">Requirement</span></span>| <span data-ttu-id="8e5ad-560">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-561">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-561">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-562">1.0</span><span class="sxs-lookup"><span data-stu-id="8e5ad-562">1.0</span></span>|
|[<span data-ttu-id="8e5ad-563">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-563">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e5ad-564">ReadItem</span></span>|
|[<span data-ttu-id="8e5ad-565">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-565">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-566">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-566">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e5ad-567">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-567">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="8e5ad-568">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8e5ad-568">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="8e5ad-569">Realiza una solicitud asincrónica a un servicio de Servicios Web Exchange (EWS) en el servidor Exchange que hospeda el buzón del usuario.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-569">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="8e5ad-570">Este método no se admite en los siguientes escenarios.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-570">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="8e5ad-571">En Outlook para iOS o Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="8e5ad-571">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="8e5ad-572">Cuando el complemento se carga en un buzón de Gmail</span><span class="sxs-lookup"><span data-stu-id="8e5ad-572">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="8e5ad-573">En estos casos, complementos deben [usar API de REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) para tener acceso al buzón del usuario en su lugar.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-573">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="8e5ad-574">El método `makeEwsRequestAsync` envía una solicitud de EWS en nombre del complemento a Exchange.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-574">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="8e5ad-575">Para obtener una lista de las operaciones de EWS compatibles, vea [llamar a servicios web desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="8e5ad-575">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="8e5ad-576">No puede solicitar elementos asociados de las carpetas con el método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-576">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="8e5ad-577">La solicitud XML debe especificar codificación UTF-8.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-577">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="8e5ad-p142">Su complemento debe tener el permiso **ReadWriteMailbox** para usar el método `makeEwsRequestAsync`. Para obtener información sobre cómo usar el permiso **ReadWriteMailbox** y sobre las operaciones EWS que puede llamar con el método `makeEwsRequestAsync`, consulte [Especificar permisos para el acceso del complemento de correo al buzón del usuario](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="8e5ad-580">El administrador del servidor debe establecer `OAuthAuthentication` en true en el directorio EWS del servidor de acceso de cliente para habilitar la `makeEwsRequestAsync` método para realizar EWS solicita.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-580">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="8e5ad-581">Diferencias de versión</span><span class="sxs-lookup"><span data-stu-id="8e5ad-581">Version differences</span></span>

<span data-ttu-id="8e5ad-582">Si usa el método `makeEwsRequestAsync` en aplicaciones de correo que se ejecutan en versiones de Outlook anteriores a 15.0.4535.1004, debe establecer el valor de codificación a `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-582">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="8e5ad-p143">No es necesario establecer el valor de codificación si la aplicación de correo se ejecuta en Outlook en la web. Puede averiguar si su aplicación de correo se ejecuta en Outlook o en Outlook en la web usando la propiedad mailbox.diagnostics.hostName. Para averiguar qué versión de Outlook se está ejecutando, use la propiedad mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e5ad-586">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="8e5ad-586">Parameters:</span></span>

|<span data-ttu-id="8e5ad-587">Nombre</span><span class="sxs-lookup"><span data-stu-id="8e5ad-587">Name</span></span>| <span data-ttu-id="8e5ad-588">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-588">Type</span></span>| <span data-ttu-id="8e5ad-589">Atributos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-589">Attributes</span></span>| <span data-ttu-id="8e5ad-590">Descripción</span><span class="sxs-lookup"><span data-stu-id="8e5ad-590">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="8e5ad-591">String</span><span class="sxs-lookup"><span data-stu-id="8e5ad-591">String</span></span>||<span data-ttu-id="8e5ad-592">La solicitud de EWS.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-592">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="8e5ad-593">función</span><span class="sxs-lookup"><span data-stu-id="8e5ad-593">function</span></span>||<span data-ttu-id="8e5ad-594">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e5ad-594">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8e5ad-595">El resultado XML de la llamada EWS se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-595">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="8e5ad-596">Si el resultado supera 1 MB de tamaño, se devuelve un mensaje de error en su lugar.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-596">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="8e5ad-597">Objeto</span><span class="sxs-lookup"><span data-stu-id="8e5ad-597">Object</span></span>| <span data-ttu-id="8e5ad-598">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e5ad-598">&lt;optional&gt;</span></span>|<span data-ttu-id="8e5ad-599">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-599">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e5ad-600">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e5ad-600">Requirements</span></span>

|<span data-ttu-id="8e5ad-601">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e5ad-601">Requirement</span></span>| <span data-ttu-id="8e5ad-602">Valor</span><span class="sxs-lookup"><span data-stu-id="8e5ad-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e5ad-603">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="8e5ad-603">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e5ad-604">1.0</span><span class="sxs-lookup"><span data-stu-id="8e5ad-604">1.0</span></span>|
|[<span data-ttu-id="8e5ad-605">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-605">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e5ad-606">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="8e5ad-606">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="8e5ad-607">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="8e5ad-607">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e5ad-608">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="8e5ad-608">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e5ad-609">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="8e5ad-609">Example</span></span>

<span data-ttu-id="8e5ad-610">En el siguiente ejemplo, se llama a `makeEwsRequestAsync` para usar la operación `GetItem` para obtener el asunto de un elemento.</span><span class="sxs-lookup"><span data-stu-id="8e5ad-610">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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