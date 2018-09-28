
# <a name="mailbox"></a><span data-ttu-id="1c6b7-101">buzón de correo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-101">mailbox</span></span>

### <span data-ttu-id="1c6b7-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="1c6b7-104">Proporciona acceso al modelo de objeto de complemento de Outlook para Microsoft Outlook y Microsoft Outlook en el web.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1c6b7-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-105">Requirements</span></span>

|<span data-ttu-id="1c6b7-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="1c6b7-106">Requirement</span></span>| <span data-ttu-id="1c6b7-107">Valor</span><span class="sxs-lookup"><span data-stu-id="1c6b7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1c6b7-108">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1c6b7-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1c6b7-109">1.0</span><span class="sxs-lookup"><span data-stu-id="1c6b7-109">1.0</span></span>|
|[<span data-ttu-id="1c6b7-110">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1c6b7-111">Restringido</span><span class="sxs-lookup"><span data-stu-id="1c6b7-111">Restricted</span></span>|
|[<span data-ttu-id="1c6b7-112">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1c6b7-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1c6b7-113">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1c6b7-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1c6b7-114">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-114">Members and methods</span></span>

| <span data-ttu-id="1c6b7-115">Miembro	</span><span class="sxs-lookup"><span data-stu-id="1c6b7-115">Member</span></span> | <span data-ttu-id="1c6b7-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1c6b7-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="1c6b7-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="1c6b7-118">Miembro	</span><span class="sxs-lookup"><span data-stu-id="1c6b7-118">Member</span></span> |
| [<span data-ttu-id="1c6b7-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="1c6b7-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="1c6b7-120">Miembro	</span><span class="sxs-lookup"><span data-stu-id="1c6b7-120">Member</span></span> |
| [<span data-ttu-id="1c6b7-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="1c6b7-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="1c6b7-122">Método</span><span class="sxs-lookup"><span data-stu-id="1c6b7-122">Method</span></span> |
| [<span data-ttu-id="1c6b7-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="1c6b7-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="1c6b7-124">Método</span><span class="sxs-lookup"><span data-stu-id="1c6b7-124">Method</span></span> |
| [<span data-ttu-id="1c6b7-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="1c6b7-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) | <span data-ttu-id="1c6b7-126">Método</span><span class="sxs-lookup"><span data-stu-id="1c6b7-126">Method</span></span> |
| [<span data-ttu-id="1c6b7-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="1c6b7-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="1c6b7-128">Método</span><span class="sxs-lookup"><span data-stu-id="1c6b7-128">Method</span></span> |
| [<span data-ttu-id="1c6b7-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="1c6b7-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="1c6b7-130">Método</span><span class="sxs-lookup"><span data-stu-id="1c6b7-130">Method</span></span> |
| [<span data-ttu-id="1c6b7-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="1c6b7-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="1c6b7-132">Método</span><span class="sxs-lookup"><span data-stu-id="1c6b7-132">Method</span></span> |
| [<span data-ttu-id="1c6b7-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="1c6b7-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="1c6b7-134">Método</span><span class="sxs-lookup"><span data-stu-id="1c6b7-134">Method</span></span> |
| [<span data-ttu-id="1c6b7-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="1c6b7-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="1c6b7-136">Método</span><span class="sxs-lookup"><span data-stu-id="1c6b7-136">Method</span></span> |
| [<span data-ttu-id="1c6b7-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1c6b7-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="1c6b7-138">Método</span><span class="sxs-lookup"><span data-stu-id="1c6b7-138">Method</span></span> |
| [<span data-ttu-id="1c6b7-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1c6b7-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="1c6b7-140">Método</span><span class="sxs-lookup"><span data-stu-id="1c6b7-140">Method</span></span> |
| [<span data-ttu-id="1c6b7-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1c6b7-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="1c6b7-142">Método</span><span class="sxs-lookup"><span data-stu-id="1c6b7-142">Method</span></span> |
| [<span data-ttu-id="1c6b7-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="1c6b7-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="1c6b7-144">Método</span><span class="sxs-lookup"><span data-stu-id="1c6b7-144">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="1c6b7-145">Espacios de nombres</span><span class="sxs-lookup"><span data-stu-id="1c6b7-145">Namespaces</span></span>

<span data-ttu-id="1c6b7-146">[diagnostics](Office.context.mailbox.diagnostics.md): Proporciona información de diagnóstico a un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-146">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="1c6b7-147">[item](Office.context.mailbox.item.md): Proporciona métodos y propiedades para tener acceso a un mensaje o cita en un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-147">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="1c6b7-148">[userProfile](Office.context.mailbox.userProfile.md): Proporciona información sobre el usuario en un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-148">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="1c6b7-149">Miembros</span><span class="sxs-lookup"><span data-stu-id="1c6b7-149">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="1c6b7-150">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="1c6b7-150">ewsUrl :String</span></span>

<span data-ttu-id="1c6b7-p102">Obtiene la dirección URL del punto de conexión de Servicios Web Exchange (EWS) para esta cuenta de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1c6b7-153">Este miembro no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-153">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1c6b7-p103">El valor `ewsUrl` puede usarse por un servicio remoto para realizar llamadas EWS al buzón del usuario. Por ejemplo, puede crear un servicio remoto para [obtener datos adjuntos desde el elemento seleccionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="1c6b7-156">La aplicación debe tener el permiso **ReadItem** especificado en su manifiesto para poder llamar al miembro `ewsUrl` en el modo lectura.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-156">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="1c6b7-p104">En el modo de redacción debe llamar al método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar el miembro `ewsUrl`. Su aplicación debe tener permisos **ReadWriteItem** para llamar al método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="1c6b7-159">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-159">Type:</span></span>

*   <span data-ttu-id="1c6b7-160">String</span><span class="sxs-lookup"><span data-stu-id="1c6b7-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1c6b7-161">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-161">Requirements</span></span>

|<span data-ttu-id="1c6b7-162">Requirement</span><span class="sxs-lookup"><span data-stu-id="1c6b7-162">Requirement</span></span>| <span data-ttu-id="1c6b7-163">Valor</span><span class="sxs-lookup"><span data-stu-id="1c6b7-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="1c6b7-164">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1c6b7-164">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1c6b7-165">1.0</span><span class="sxs-lookup"><span data-stu-id="1c6b7-165">1.0</span></span>|
|[<span data-ttu-id="1c6b7-166">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-166">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1c6b7-167">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1c6b7-167">ReadItem</span></span>|
|[<span data-ttu-id="1c6b7-168">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1c6b7-168">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1c6b7-169">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1c6b7-169">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="1c6b7-170">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="1c6b7-170">restUrl :String</span></span>

<span data-ttu-id="1c6b7-171">Obtiene la URL del punto de conexión REST para esta cuenta de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-171">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="1c6b7-172">El valor `restUrl` puede usarse para realizar llamadas [API de REST](https://docs.microsoft.com/outlook/rest/) al buzón del usuario.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-172">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="1c6b7-173">La aplicación debe tener el permiso **ReadItem** especificado en su manifiesto para poder llamar al miembro `restUrl` en el modo lectura.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-173">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="1c6b7-p105">En el modo de redacción debe llamar al método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar el miembro `restUrl`. Su aplicación debe tener permisos **ReadWriteItem** para llamar al método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="1c6b7-176">Los clientes de Outlook conectados a instalaciones locales de Exchange 2016 o posteriores con una dirección URL de REST personalizada configurado devolverá un valor no válido para `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-176">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="1c6b7-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-177">Type:</span></span>

*   <span data-ttu-id="1c6b7-178">String</span><span class="sxs-lookup"><span data-stu-id="1c6b7-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1c6b7-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-179">Requirements</span></span>

|<span data-ttu-id="1c6b7-180">Requirement</span><span class="sxs-lookup"><span data-stu-id="1c6b7-180">Requirement</span></span>| <span data-ttu-id="1c6b7-181">Valor</span><span class="sxs-lookup"><span data-stu-id="1c6b7-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="1c6b7-182">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1c6b7-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1c6b7-183">1,5</span><span class="sxs-lookup"><span data-stu-id="1c6b7-183">1.5</span></span> |
|[<span data-ttu-id="1c6b7-184">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1c6b7-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1c6b7-185">ReadItem</span></span>|
|[<span data-ttu-id="1c6b7-186">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1c6b7-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1c6b7-187">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1c6b7-187">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="1c6b7-188">Métodos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-188">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="1c6b7-189">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1c6b7-189">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="1c6b7-190">Agrega un controlador de eventos para un evento admitido.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-190">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="1c6b7-p106">Actualmente, el único tipo de evento admitido es `Office.EventType.ItemChanged`, que se invoca cuando el usuario selecciona un elemento nuevo. Este evento se usa por los complementos que implementan un panel de tareas anclable, y permite que el complemento actualice la interfaz de usuario del panel de tareas basándose en el elemento seleccionado actualmente.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p106">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1c6b7-193">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-193">Parameters:</span></span>

| <span data-ttu-id="1c6b7-194">Nombre</span><span class="sxs-lookup"><span data-stu-id="1c6b7-194">Name</span></span> | <span data-ttu-id="1c6b7-195">Tipo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-195">Type</span></span> | <span data-ttu-id="1c6b7-196">Atributos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-196">Attributes</span></span> | <span data-ttu-id="1c6b7-197">Descripción</span><span class="sxs-lookup"><span data-stu-id="1c6b7-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="1c6b7-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="1c6b7-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="1c6b7-199">El evento que debe invocar el controlador.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="1c6b7-200">Función</span><span class="sxs-lookup"><span data-stu-id="1c6b7-200">Function</span></span> || <span data-ttu-id="1c6b7-p107">La función que va a controlar el evento. La función debe aceptar un único parámetro, que es un literal de objeto. La propiedad `type` del parámetro coincidirá con el parámetro `eventType` que se ha pasado a `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="1c6b7-204">Objeto</span><span class="sxs-lookup"><span data-stu-id="1c6b7-204">Object</span></span> | <span data-ttu-id="1c6b7-205">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1c6b7-205">&lt;optional&gt;</span></span> | <span data-ttu-id="1c6b7-206">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1c6b7-207">Objeto</span><span class="sxs-lookup"><span data-stu-id="1c6b7-207">Object</span></span> | <span data-ttu-id="1c6b7-208">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1c6b7-208">&lt;optional&gt;</span></span> | <span data-ttu-id="1c6b7-209">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="1c6b7-210">función</span><span class="sxs-lookup"><span data-stu-id="1c6b7-210">function</span></span>| <span data-ttu-id="1c6b7-211">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1c6b7-211">&lt;optional&gt;</span></span>|<span data-ttu-id="1c6b7-212">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1c6b7-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1c6b7-213">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-213">Requirements</span></span>

|<span data-ttu-id="1c6b7-214">Requirement</span><span class="sxs-lookup"><span data-stu-id="1c6b7-214">Requirement</span></span>| <span data-ttu-id="1c6b7-215">Valor</span><span class="sxs-lookup"><span data-stu-id="1c6b7-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="1c6b7-216">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1c6b7-216">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1c6b7-217">1,5</span><span class="sxs-lookup"><span data-stu-id="1c6b7-217">1.5</span></span> |
|[<span data-ttu-id="1c6b7-218">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1c6b7-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1c6b7-219">ReadItem</span></span> |
|[<span data-ttu-id="1c6b7-220">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1c6b7-220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1c6b7-221">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1c6b7-221">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1c6b7-222">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-222">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="1c6b7-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="1c6b7-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="1c6b7-224">Convierte un identificador de elemento con formato para REST al formato EWS.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="1c6b7-225">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-225">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1c6b7-p108">Los identificadores de elemento obtenidos a través de una API de REST (como la [API de correo de Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) o [Microsoft Graph](http://graph.microsoft.io/)) usan un formato diferente al formato que usa Exchange Web Services (EWS). El método `convertToEwsId` convierte un identificador con formato REST al formato adecuado para EWS.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1c6b7-228">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-228">Parameters:</span></span>

|<span data-ttu-id="1c6b7-229">Nombre</span><span class="sxs-lookup"><span data-stu-id="1c6b7-229">Name</span></span>| <span data-ttu-id="1c6b7-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-230">Type</span></span>| <span data-ttu-id="1c6b7-231">Descripción</span><span class="sxs-lookup"><span data-stu-id="1c6b7-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1c6b7-232">String</span><span class="sxs-lookup"><span data-stu-id="1c6b7-232">String</span></span>|<span data-ttu-id="1c6b7-233">Un identificador de elemento con formato para las API de REST de Outlook</span><span class="sxs-lookup"><span data-stu-id="1c6b7-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="1c6b7-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="1c6b7-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="1c6b7-235">Un valor que indica la versión de la API de REST de Outlook que se usa para recuperar el identificador de elemento.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1c6b7-236">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-236">Requirements</span></span>

|<span data-ttu-id="1c6b7-237">Requirement</span><span class="sxs-lookup"><span data-stu-id="1c6b7-237">Requirement</span></span>| <span data-ttu-id="1c6b7-238">Valor</span><span class="sxs-lookup"><span data-stu-id="1c6b7-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="1c6b7-239">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1c6b7-239">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1c6b7-240">1.3</span><span class="sxs-lookup"><span data-stu-id="1c6b7-240">1.3</span></span>|
|[<span data-ttu-id="1c6b7-241">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-241">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1c6b7-242">Restringido</span><span class="sxs-lookup"><span data-stu-id="1c6b7-242">Restricted</span></span>|
|[<span data-ttu-id="1c6b7-243">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1c6b7-243">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1c6b7-244">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1c6b7-244">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1c6b7-245">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-245">Returns:</span></span>

<span data-ttu-id="1c6b7-246">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="1c6b7-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="1c6b7-247">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-247">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime"></a><span data-ttu-id="1c6b7-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="1c6b7-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span></span>

<span data-ttu-id="1c6b7-249">Obtiene un diccionario con información de tiempo en el tiempo del cliente local.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="1c6b7-p109">Las fechas y horas usadas por una aplicación de correo para Outlook o Outlook Web App pueden usar distintas zonas horarias. Outlook usa la zona horaria del equipo cliente; Outlook Web App usa la zona horaria definida en el Centro de administración de Exchange (EAC). Debería tratar los valores de fecha y hora de modo que los valores que aparezcan en la interfaz de usuario sean siempre coherentes con la zona horaria que el usuario espera.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="1c6b7-p110">Si se está ejecutando la aplicación de correo en Outlook, el método `convertToLocalClientTime` devolverá un objeto de diccionario con los valores establecidos para la zona horaria del equipo cliente. Si se está ejecutando la aplicación de correo en Outlook Web App, el método `convertToLocalClientTime` devolverá un objeto de diccionario con los valores establecidos para la zona horaria especificada en el CEF.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1c6b7-255">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-255">Parameters:</span></span>

|<span data-ttu-id="1c6b7-256">Nombre</span><span class="sxs-lookup"><span data-stu-id="1c6b7-256">Name</span></span>| <span data-ttu-id="1c6b7-257">Tipo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-257">Type</span></span>| <span data-ttu-id="1c6b7-258">Descripción</span><span class="sxs-lookup"><span data-stu-id="1c6b7-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="1c6b7-259">Fecha</span><span class="sxs-lookup"><span data-stu-id="1c6b7-259">Date</span></span>|<span data-ttu-id="1c6b7-260">Un objeto Date</span><span class="sxs-lookup"><span data-stu-id="1c6b7-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1c6b7-261">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-261">Requirements</span></span>

|<span data-ttu-id="1c6b7-262">Requirement</span><span class="sxs-lookup"><span data-stu-id="1c6b7-262">Requirement</span></span>| <span data-ttu-id="1c6b7-263">Valor</span><span class="sxs-lookup"><span data-stu-id="1c6b7-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="1c6b7-264">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1c6b7-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1c6b7-265">1.0</span><span class="sxs-lookup"><span data-stu-id="1c6b7-265">1.0</span></span>|
|[<span data-ttu-id="1c6b7-266">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1c6b7-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1c6b7-267">ReadItem</span></span>|
|[<span data-ttu-id="1c6b7-268">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1c6b7-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1c6b7-269">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1c6b7-269">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1c6b7-270">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-270">Returns:</span></span>

<span data-ttu-id="1c6b7-271">Tipo: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="1c6b7-271">Type: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="1c6b7-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="1c6b7-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="1c6b7-273">Convierte un identificador de elemento con formato para EWS al formato REST.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="1c6b7-274">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-274">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1c6b7-p111">Los identificadores de elemento obtenidos a través de EWS o de la propiedad `itemId` usan un formato diferente al formato que usan las API de REST (como la [API de correo de Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) o [Microsoft Graph](http://graph.microsoft.io/)). El método `convertToRestId` convierte un identificador con formato EWS al formato adecuado para REST.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1c6b7-277">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-277">Parameters:</span></span>

|<span data-ttu-id="1c6b7-278">Nombre</span><span class="sxs-lookup"><span data-stu-id="1c6b7-278">Name</span></span>| <span data-ttu-id="1c6b7-279">Tipo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-279">Type</span></span>| <span data-ttu-id="1c6b7-280">Descripción</span><span class="sxs-lookup"><span data-stu-id="1c6b7-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1c6b7-281">String</span><span class="sxs-lookup"><span data-stu-id="1c6b7-281">String</span></span>|<span data-ttu-id="1c6b7-282">Un identificador de elemento con formato para Exchange Web Services (EWS)</span><span class="sxs-lookup"><span data-stu-id="1c6b7-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="1c6b7-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="1c6b7-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="1c6b7-284">Un valor que indica la versión de la API de REST de Outlook con la que se usará el identificador convertido.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1c6b7-285">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-285">Requirements</span></span>

|<span data-ttu-id="1c6b7-286">Requirement</span><span class="sxs-lookup"><span data-stu-id="1c6b7-286">Requirement</span></span>| <span data-ttu-id="1c6b7-287">Valor</span><span class="sxs-lookup"><span data-stu-id="1c6b7-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="1c6b7-288">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1c6b7-288">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1c6b7-289">1.3</span><span class="sxs-lookup"><span data-stu-id="1c6b7-289">1.3</span></span>|
|[<span data-ttu-id="1c6b7-290">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1c6b7-291">Restringido</span><span class="sxs-lookup"><span data-stu-id="1c6b7-291">Restricted</span></span>|
|[<span data-ttu-id="1c6b7-292">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1c6b7-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1c6b7-293">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1c6b7-293">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1c6b7-294">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-294">Returns:</span></span>

<span data-ttu-id="1c6b7-295">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="1c6b7-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="1c6b7-296">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-296">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="1c6b7-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="1c6b7-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="1c6b7-298">Obtiene un objeto Date del diccionario que contiene información de tiempo.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="1c6b7-299">El método `convertToUtcClientTime` convierte un diccionario que contiene la fecha y la hora locales en un objeto Date con los valores correctos para la fecha y la hora locales.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1c6b7-300">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-300">Parameters:</span></span>

|<span data-ttu-id="1c6b7-301">Nombre</span><span class="sxs-lookup"><span data-stu-id="1c6b7-301">Name</span></span>| <span data-ttu-id="1c6b7-302">Tipo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-302">Type</span></span>| <span data-ttu-id="1c6b7-303">Descripción</span><span class="sxs-lookup"><span data-stu-id="1c6b7-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="1c6b7-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="1c6b7-304">LocalClientTime</span></span>](/javascript/api/outlook_1_5/office.LocalClientTime)|<span data-ttu-id="1c6b7-305">El valor de la hora local para convertir.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1c6b7-306">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-306">Requirements</span></span>

|<span data-ttu-id="1c6b7-307">Requirement</span><span class="sxs-lookup"><span data-stu-id="1c6b7-307">Requirement</span></span>| <span data-ttu-id="1c6b7-308">Valor</span><span class="sxs-lookup"><span data-stu-id="1c6b7-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="1c6b7-309">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1c6b7-309">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1c6b7-310">1.0</span><span class="sxs-lookup"><span data-stu-id="1c6b7-310">1.0</span></span>|
|[<span data-ttu-id="1c6b7-311">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-311">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1c6b7-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1c6b7-312">ReadItem</span></span>|
|[<span data-ttu-id="1c6b7-313">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1c6b7-313">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1c6b7-314">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1c6b7-314">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1c6b7-315">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-315">Returns:</span></span>

<span data-ttu-id="1c6b7-316">Objeto Date con el tiempo expresado en UTC.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-316">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="1c6b7-317">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="1c6b7-317">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1c6b7-318">Fecha</span><span class="sxs-lookup"><span data-stu-id="1c6b7-318">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="1c6b7-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="1c6b7-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="1c6b7-320">Muestra una cita de calendario existente.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1c6b7-321">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-321">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1c6b7-322">El método `displayAppointmentForm` abre una cita de calendario existente en una nueva ventana del escritorio o en un cuadro de diálogo en los dispositivos móviles.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="1c6b7-p112">En Outlook para Mac, puede usar este método para mostrar una cita que no forme parte de una serie periódica o la cita principal de una serie periódica, pero no puede mostrar una instancia de la serie. Esto es porque en Outlook para Mac no se puede tener acceso a las propiedades (incluido el identificador de elemento) de las instancias de una serie periódica.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="1c6b7-325">En Outlook Web App, este método abre el formato especificado solo si el cuerpo del formulario es inferior o igual a 32 KB en el número de caracteres.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-325">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="1c6b7-326">Si el identificador de elemento especificado no identifica una cita existente, se abrirá una página en blanco en el dispositivo o equipo cliente y no se generará ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1c6b7-327">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-327">Parameters:</span></span>

|<span data-ttu-id="1c6b7-328">Nombre</span><span class="sxs-lookup"><span data-stu-id="1c6b7-328">Name</span></span>| <span data-ttu-id="1c6b7-329">Tipo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-329">Type</span></span>| <span data-ttu-id="1c6b7-330">Descripción</span><span class="sxs-lookup"><span data-stu-id="1c6b7-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1c6b7-331">String</span><span class="sxs-lookup"><span data-stu-id="1c6b7-331">String</span></span>|<span data-ttu-id="1c6b7-332">Identificador de los servicios web de Exchange (EWS) para una cita de calendario existente.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1c6b7-333">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-333">Requirements</span></span>

|<span data-ttu-id="1c6b7-334">Requirement</span><span class="sxs-lookup"><span data-stu-id="1c6b7-334">Requirement</span></span>| <span data-ttu-id="1c6b7-335">Valor</span><span class="sxs-lookup"><span data-stu-id="1c6b7-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="1c6b7-336">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1c6b7-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1c6b7-337">1.0</span><span class="sxs-lookup"><span data-stu-id="1c6b7-337">1.0</span></span>|
|[<span data-ttu-id="1c6b7-338">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1c6b7-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1c6b7-339">ReadItem</span></span>|
|[<span data-ttu-id="1c6b7-340">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1c6b7-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1c6b7-341">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1c6b7-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1c6b7-342">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-342">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="1c6b7-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="1c6b7-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="1c6b7-344">Muestra un mensaje existente.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="1c6b7-345">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-345">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1c6b7-346">El método `displayMessageForm` abre un mensaje existente en una nueva ventana del escritorio o en un cuadro de diálogo en los dispositivos móviles.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="1c6b7-347">En Outlook Web App, este método abre el formato especificado solo si el cuerpo del formulario es inferior o igual a 32 KB en el número de caracteres.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-347">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="1c6b7-348">Si el identificador de elemento especificado no identifica un mensaje existente, no se mostrará ningún mensaje en el equipo cliente ni se generará ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="1c6b7-p113">No use el valor `displayMessageForm` con un `itemId` que represente una cita. Use el método `displayAppointmentForm` para mostrar una cita existente y `displayNewAppointmentForm` para mostrar un formulario para crear una cita nueva.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1c6b7-351">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-351">Parameters:</span></span>

|<span data-ttu-id="1c6b7-352">Nombre</span><span class="sxs-lookup"><span data-stu-id="1c6b7-352">Name</span></span>| <span data-ttu-id="1c6b7-353">Tipo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-353">Type</span></span>| <span data-ttu-id="1c6b7-354">Descripción</span><span class="sxs-lookup"><span data-stu-id="1c6b7-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1c6b7-355">String</span><span class="sxs-lookup"><span data-stu-id="1c6b7-355">String</span></span>|<span data-ttu-id="1c6b7-356">Identificador de los servicios web de Exchange (EWS) para un mensaje existente.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1c6b7-357">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-357">Requirements</span></span>

|<span data-ttu-id="1c6b7-358">Requirement</span><span class="sxs-lookup"><span data-stu-id="1c6b7-358">Requirement</span></span>| <span data-ttu-id="1c6b7-359">Valor</span><span class="sxs-lookup"><span data-stu-id="1c6b7-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="1c6b7-360">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1c6b7-360">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1c6b7-361">1.0</span><span class="sxs-lookup"><span data-stu-id="1c6b7-361">1.0</span></span>|
|[<span data-ttu-id="1c6b7-362">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1c6b7-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1c6b7-363">ReadItem</span></span>|
|[<span data-ttu-id="1c6b7-364">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1c6b7-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1c6b7-365">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1c6b7-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1c6b7-366">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-366">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="1c6b7-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="1c6b7-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="1c6b7-368">Muestra un formulario para crear una nueva cita de calendario.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1c6b7-369">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-369">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1c6b7-p114">El método `displayNewAppointmentForm` abre un formulario que permite al usuario crear una nueva cita o reunión. Si se especifican parámetros, los campos de formulario de cita se rellenan automáticamente con el contenido de los parámetros.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="1c6b7-p115">En Outlook Web App y OWA para dispositivos, este método muestra siempre un formulario con un campo de asistentes. Si no especifica ningún asistente como argumento de entrada, el método muestra un formulario con un botón **Guardar**. Si ha especificado asistentes, el formulario incluirá a los asistentes y un botón **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="1c6b7-p116">En el cliente enriquecido de Outlook y Outlook RT, si se especifica cualquier asistente o recurso en los parámetros `requiredAttendees`, `optionalAttendees` o `resources`, este método muestra un formulario de reunión con un botón **Enviar**. Si no se especifica ningún destinatario, este método muestra un formulario de cita con un botón **Guardar y cerrar**.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="1c6b7-377">Si cualquiera de los parámetros supera los límites de tamaño especificados o si se especifica un nombre de parámetro desconocido, se genera una excepción.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1c6b7-378">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-378">Parameters:</span></span>

|<span data-ttu-id="1c6b7-379">Nombre</span><span class="sxs-lookup"><span data-stu-id="1c6b7-379">Name</span></span>| <span data-ttu-id="1c6b7-380">Tipo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-380">Type</span></span>| <span data-ttu-id="1c6b7-381">Descripción</span><span class="sxs-lookup"><span data-stu-id="1c6b7-381">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="1c6b7-382">Object</span><span class="sxs-lookup"><span data-stu-id="1c6b7-382">Object</span></span> | <span data-ttu-id="1c6b7-383">Un diccionario de parámetros que describen la nueva cita.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-383">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="1c6b7-384">Matriz.&lt;Cadena&gt; &#124; Matriz.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1c6b7-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1c6b7-p117">Una matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los asistentes necesarios de la cita. La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="1c6b7-387">Matriz.&lt;Cadena&gt; &#124; Matriz.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1c6b7-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1c6b7-p118">Una matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los asistentes opcionales de la cita. La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="1c6b7-390">Fecha</span><span class="sxs-lookup"><span data-stu-id="1c6b7-390">Date</span></span> | <span data-ttu-id="1c6b7-391">Un objeto `Date` que especifica la fecha y hora de inicio de la cita.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-391">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="1c6b7-392">Fecha</span><span class="sxs-lookup"><span data-stu-id="1c6b7-392">Date</span></span> | <span data-ttu-id="1c6b7-393">Un objeto `Date` que especifica la fecha y hora de finalización de la cita.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-393">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="1c6b7-394">String</span><span class="sxs-lookup"><span data-stu-id="1c6b7-394">String</span></span> | <span data-ttu-id="1c6b7-p119">Una cadena que contiene la ubicación de la cita. La cadena está limitada a un máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="1c6b7-397">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="1c6b7-397">Array.&lt;String&gt;</span></span> | <span data-ttu-id="1c6b7-p120">Una matriz de cadenas que contiene los recursos necesarios para la cita. La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="1c6b7-400">String</span><span class="sxs-lookup"><span data-stu-id="1c6b7-400">String</span></span> | <span data-ttu-id="1c6b7-p121">Una cadena que contiene al asunto de la cita. La cadena está limitada a un máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="1c6b7-403">String</span><span class="sxs-lookup"><span data-stu-id="1c6b7-403">String</span></span> | <span data-ttu-id="1c6b7-p122">El cuerpo de la cita. El contenido del cuerpo está limitado a un tamaño máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1c6b7-406">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-406">Requirements</span></span>

|<span data-ttu-id="1c6b7-407">Requirement</span><span class="sxs-lookup"><span data-stu-id="1c6b7-407">Requirement</span></span>| <span data-ttu-id="1c6b7-408">Valor</span><span class="sxs-lookup"><span data-stu-id="1c6b7-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="1c6b7-409">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1c6b7-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1c6b7-410">1.0</span><span class="sxs-lookup"><span data-stu-id="1c6b7-410">1.0</span></span>|
|[<span data-ttu-id="1c6b7-411">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1c6b7-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1c6b7-412">ReadItem</span></span>|
|[<span data-ttu-id="1c6b7-413">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1c6b7-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1c6b7-414">Lectura</span><span class="sxs-lookup"><span data-stu-id="1c6b7-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1c6b7-415">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-415">Example</span></span>

```js
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="1c6b7-416">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="1c6b7-416">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="1c6b7-417">Obtiene una cadena que contiene un token usado para llamar a las API de REST o a los servicios Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-417">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="1c6b7-p123">El método `getCallbackTokenAsync` realiza una llamada asincrónica para obtener un token opaco desde Exchange Server que hospeda el buzón del usuario. La duración del token de devolución de llamada es de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p123">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="1c6b7-420">Se recomienda que los complementos de usar las API de REST en lugar de los servicios Web Exchange siempre que sea posible.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-420">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="1c6b7-421">**Tokens de REST**</span><span class="sxs-lookup"><span data-stu-id="1c6b7-421">**REST Tokens**</span></span>

<span data-ttu-id="1c6b7-p124">Cuando se solicita un token de REST (`options.isRest = true`), el token resultante no funcionará para autenticar las llamadas a los servicios Web Exchange. El token se limitará en su ámbito para el acceso de solo lectura al elemento actual y a sus datos adjuntos, a no ser que el complemento haya especificado el permiso [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) en su manifiesto. Si está especificado el permiso `ReadWriteMailbox`, el token resultante concederá acceso de lectura o escritura al correo, calendario y contactos, incluida la capacidad de enviar correo.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p124">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="1c6b7-425">El complemento debe usar la propiedad `restUrl` para determinar la URL correcta que se va a usar al realizar las llamadas a la API de REST.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-425">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="1c6b7-426">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="1c6b7-426">**EWS Tokens**</span></span>

<span data-ttu-id="1c6b7-p125">Cuando se solicita un token EWS (`options.isRest = false`), el token resultante no funcionará para autenticar las llamadas a las API de REST. El token estará limitado en su ámbito para el acceso al elemento actual.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p125">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="1c6b7-429">El complemento debe usar la propiedad `ewsUrl` para determinar la URL correcta que se va a usar al realizar las llamadas a EWS.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-429">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1c6b7-430">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-430">Parameters:</span></span>

|<span data-ttu-id="1c6b7-431">Nombre</span><span class="sxs-lookup"><span data-stu-id="1c6b7-431">Name</span></span>| <span data-ttu-id="1c6b7-432">Tipo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-432">Type</span></span>| <span data-ttu-id="1c6b7-433">Atributos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-433">Attributes</span></span>| <span data-ttu-id="1c6b7-434">Descripción</span><span class="sxs-lookup"><span data-stu-id="1c6b7-434">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="1c6b7-435">Objeto</span><span class="sxs-lookup"><span data-stu-id="1c6b7-435">Object</span></span> | <span data-ttu-id="1c6b7-436">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1c6b7-436">&lt;optional&gt;</span></span> | <span data-ttu-id="1c6b7-437">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-437">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="1c6b7-438">Booleano</span><span class="sxs-lookup"><span data-stu-id="1c6b7-438">Boolean</span></span> |  <span data-ttu-id="1c6b7-439">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1c6b7-439">&lt;optional&gt;</span></span> | <span data-ttu-id="1c6b7-p126">Determina si el token proporcionado se usará para las API de REST de Outlook o para los servicios Web Exchange. El valor predeterminado es `false`.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p126">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1c6b7-442">Objeto</span><span class="sxs-lookup"><span data-stu-id="1c6b7-442">Object</span></span> |  <span data-ttu-id="1c6b7-443">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1c6b7-443">&lt;optional&gt;</span></span> | <span data-ttu-id="1c6b7-444">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-444">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="1c6b7-445">Función</span><span class="sxs-lookup"><span data-stu-id="1c6b7-445">function</span></span>||<span data-ttu-id="1c6b7-p127">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). El token se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1c6b7-448">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-448">Requirements</span></span>

|<span data-ttu-id="1c6b7-449">Requirement</span><span class="sxs-lookup"><span data-stu-id="1c6b7-449">Requirement</span></span>| <span data-ttu-id="1c6b7-450">Valor</span><span class="sxs-lookup"><span data-stu-id="1c6b7-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="1c6b7-451">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1c6b7-451">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1c6b7-452">1,5</span><span class="sxs-lookup"><span data-stu-id="1c6b7-452">1.5</span></span> |
|[<span data-ttu-id="1c6b7-453">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-453">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1c6b7-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1c6b7-454">ReadItem</span></span>|
|[<span data-ttu-id="1c6b7-455">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1c6b7-455">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1c6b7-456">Redacción y lectura</span><span class="sxs-lookup"><span data-stu-id="1c6b7-456">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="1c6b7-457">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-457">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="1c6b7-458">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1c6b7-458">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="1c6b7-459">Obtiene una cadena que contiene un token usado para obtener datos adjuntos o un elemento de Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-459">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="1c6b7-p128">El método `getCallbackTokenAsync` realiza una llamada asincrónica para obtener un token opaco desde Exchange Server que hospeda el buzón del usuario. La duración del token de devolución de llamada es de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p128">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="1c6b7-p129">Puede pasar el token y un identificador de archivo adjunto o el identificador del elemento a un sistema de terceros. El sistema de terceros usa el token como token de autorización del portador para llamar a los Servicios web de Exchange (EWS) o a las operaciones [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) o [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) para devolver datos adjuntos o un elemento. Por ejemplo, puede crear un servicio remoto para [obtener datos adjuntos desde el elemento seleccionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p129">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="1c6b7-465">La aplicación debe tener el permiso **ReadItem** especificado en su manifiesto para poder llamar al método `getCallbackTokenAsync` en el modo lectura.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-465">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="1c6b7-p130">En el modo de redacción debe llamar al método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obtener un identificador de elemento para pasar al método `getCallbackTokenAsync`. Su aplicación debe tener permisos **ReadWriteItem** para llamar al método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p130">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1c6b7-468">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-468">Parameters:</span></span>

|<span data-ttu-id="1c6b7-469">Nombre</span><span class="sxs-lookup"><span data-stu-id="1c6b7-469">Name</span></span>| <span data-ttu-id="1c6b7-470">Tipo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-470">Type</span></span>| <span data-ttu-id="1c6b7-471">Atributos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-471">Attributes</span></span>| <span data-ttu-id="1c6b7-472">Descripción</span><span class="sxs-lookup"><span data-stu-id="1c6b7-472">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1c6b7-473">función</span><span class="sxs-lookup"><span data-stu-id="1c6b7-473">function</span></span>||<span data-ttu-id="1c6b7-p131">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). El token se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p131">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="1c6b7-476">Objeto</span><span class="sxs-lookup"><span data-stu-id="1c6b7-476">Object</span></span>| <span data-ttu-id="1c6b7-477">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1c6b7-477">&lt;optional&gt;</span></span>|<span data-ttu-id="1c6b7-478">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-478">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1c6b7-479">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-479">Requirements</span></span>

|<span data-ttu-id="1c6b7-480">Requirement</span><span class="sxs-lookup"><span data-stu-id="1c6b7-480">Requirement</span></span>| <span data-ttu-id="1c6b7-481">Valor</span><span class="sxs-lookup"><span data-stu-id="1c6b7-481">Value</span></span>|
|---|---|
|[<span data-ttu-id="1c6b7-482">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1c6b7-482">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1c6b7-483">1.3</span><span class="sxs-lookup"><span data-stu-id="1c6b7-483">1.3</span></span>|
|[<span data-ttu-id="1c6b7-484">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-484">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1c6b7-485">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1c6b7-485">ReadItem</span></span>|
|[<span data-ttu-id="1c6b7-486">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1c6b7-486">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1c6b7-487">Redacción y lectura</span><span class="sxs-lookup"><span data-stu-id="1c6b7-487">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="1c6b7-488">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-488">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="1c6b7-489">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1c6b7-489">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="1c6b7-490">Obtiene un token que identifica al usuario y al complemento de Office.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-490">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="1c6b7-491">El método `getUserIdentityTokenAsync` devuelve un token que puede usar para identificar y [autenticar el complemento y el usuario mediante un sistema de terceros](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="1c6b7-491">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="1c6b7-492">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-492">Parameters:</span></span>

|<span data-ttu-id="1c6b7-493">Nombre</span><span class="sxs-lookup"><span data-stu-id="1c6b7-493">Name</span></span>| <span data-ttu-id="1c6b7-494">Tipo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-494">Type</span></span>| <span data-ttu-id="1c6b7-495">Atributos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-495">Attributes</span></span>| <span data-ttu-id="1c6b7-496">Descripción</span><span class="sxs-lookup"><span data-stu-id="1c6b7-496">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1c6b7-497">función</span><span class="sxs-lookup"><span data-stu-id="1c6b7-497">function</span></span>||<span data-ttu-id="1c6b7-498">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1c6b7-498">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1c6b7-499">El token se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-499">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="1c6b7-500">Object</span><span class="sxs-lookup"><span data-stu-id="1c6b7-500">Object</span></span>| <span data-ttu-id="1c6b7-501">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1c6b7-501">&lt;optional&gt;</span></span>|<span data-ttu-id="1c6b7-502">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-502">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1c6b7-503">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-503">Requirements</span></span>

|<span data-ttu-id="1c6b7-504">Requirement</span><span class="sxs-lookup"><span data-stu-id="1c6b7-504">Requirement</span></span>| <span data-ttu-id="1c6b7-505">Valor</span><span class="sxs-lookup"><span data-stu-id="1c6b7-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="1c6b7-506">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1c6b7-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1c6b7-507">1.0</span><span class="sxs-lookup"><span data-stu-id="1c6b7-507">1.0</span></span>|
|[<span data-ttu-id="1c6b7-508">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1c6b7-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1c6b7-509">ReadItem</span></span>|
|[<span data-ttu-id="1c6b7-510">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1c6b7-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1c6b7-511">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1c6b7-511">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1c6b7-512">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-512">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="1c6b7-513">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1c6b7-513">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="1c6b7-514">Realiza una solicitud asincrónica a un servicio de Servicios Web Exchange (EWS) en el servidor Exchange que hospeda el buzón del usuario.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-514">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="1c6b7-515">Este método no se admite en los siguientes escenarios.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-515">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="1c6b7-516">En Outlook para iOS o Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="1c6b7-516">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="1c6b7-517">Cuando el complemento se carga en un buzón de Gmail</span><span class="sxs-lookup"><span data-stu-id="1c6b7-517">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="1c6b7-518">En estos casos, complementos deben [usar API de REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) para tener acceso al buzón del usuario en su lugar.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-518">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="1c6b7-519">El método `makeEwsRequestAsync` envía una solicitud de EWS en nombre del complemento a Exchange.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-519">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="1c6b7-520">Para obtener una lista de las operaciones de EWS compatibles, vea [llamar a servicios web desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="1c6b7-520">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="1c6b7-521">No puede solicitar elementos asociados de las carpetas con el método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-521">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="1c6b7-522">La solicitud XML debe especificar codificación UTF-8.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-522">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="1c6b7-p133">Su complemento debe tener el permiso **ReadWriteMailbox** para usar el método `makeEwsRequestAsync`. Para obtener información sobre cómo usar el permiso **ReadWriteMailbox** y sobre las operaciones EWS que puede llamar con el método `makeEwsRequestAsync`, consulte [Especificar permisos para el acceso del complemento de correo al buzón del usuario](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p133">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="1c6b7-525">El administrador del servidor debe establecer `OAuthAuthentication` en true en el directorio EWS del servidor de acceso de cliente para habilitar la `makeEwsRequestAsync` método para realizar EWS solicita.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-525">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="1c6b7-526">Diferencias de versión</span><span class="sxs-lookup"><span data-stu-id="1c6b7-526">Version differences</span></span>

<span data-ttu-id="1c6b7-527">Si usa el método `makeEwsRequestAsync` en aplicaciones de correo que se ejecutan en versiones de Outlook anteriores a 15.0.4535.1004, debe establecer el valor de codificación a `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-527">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="1c6b7-p134">No es necesario establecer el valor de codificación si la aplicación de correo se ejecuta en Outlook en la web. Puede averiguar si su aplicación de correo se ejecuta en Outlook o en Outlook en la web usando la propiedad mailbox.diagnostics.hostName. Para averiguar qué versión de Outlook se está ejecutando, use la propiedad mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-p134">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1c6b7-531">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="1c6b7-531">Parameters:</span></span>

|<span data-ttu-id="1c6b7-532">Nombre</span><span class="sxs-lookup"><span data-stu-id="1c6b7-532">Name</span></span>| <span data-ttu-id="1c6b7-533">Tipo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-533">Type</span></span>| <span data-ttu-id="1c6b7-534">Atributos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-534">Attributes</span></span>| <span data-ttu-id="1c6b7-535">Descripción</span><span class="sxs-lookup"><span data-stu-id="1c6b7-535">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="1c6b7-536">String</span><span class="sxs-lookup"><span data-stu-id="1c6b7-536">String</span></span>||<span data-ttu-id="1c6b7-537">La solicitud de EWS.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-537">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="1c6b7-538">función</span><span class="sxs-lookup"><span data-stu-id="1c6b7-538">function</span></span>||<span data-ttu-id="1c6b7-539">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1c6b7-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1c6b7-540">El resultado XML de la llamada EWS se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-540">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="1c6b7-541">Si el resultado supera 1 MB de tamaño, se devuelve un mensaje de error en su lugar.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-541">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="1c6b7-542">Objeto</span><span class="sxs-lookup"><span data-stu-id="1c6b7-542">Object</span></span>| <span data-ttu-id="1c6b7-543">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1c6b7-543">&lt;optional&gt;</span></span>|<span data-ttu-id="1c6b7-544">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-544">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1c6b7-545">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1c6b7-545">Requirements</span></span>

|<span data-ttu-id="1c6b7-546">Requirement</span><span class="sxs-lookup"><span data-stu-id="1c6b7-546">Requirement</span></span>| <span data-ttu-id="1c6b7-547">Valor</span><span class="sxs-lookup"><span data-stu-id="1c6b7-547">Value</span></span>|
|---|---|
|[<span data-ttu-id="1c6b7-548">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1c6b7-548">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1c6b7-549">1.0</span><span class="sxs-lookup"><span data-stu-id="1c6b7-549">1.0</span></span>|
|[<span data-ttu-id="1c6b7-550">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-550">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1c6b7-551">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="1c6b7-551">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="1c6b7-552">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1c6b7-552">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1c6b7-553">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1c6b7-553">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1c6b7-554">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="1c6b7-554">Example</span></span>

<span data-ttu-id="1c6b7-555">En el siguiente ejemplo, se llama a `makeEwsRequestAsync` para usar la operación `GetItem` para obtener el asunto de un elemento.</span><span class="sxs-lookup"><span data-stu-id="1c6b7-555">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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