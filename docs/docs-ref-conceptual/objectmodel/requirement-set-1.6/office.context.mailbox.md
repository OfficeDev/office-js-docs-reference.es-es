
# <a name="mailbox"></a><span data-ttu-id="96367-101">buzón de correo</span><span class="sxs-lookup"><span data-stu-id="96367-101">mailbox</span></span>

### <span data-ttu-id="96367-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="96367-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="96367-104">Proporciona acceso al modelo de objeto de complemento de Outlook para Microsoft Outlook y Microsoft Outlook en el web.</span><span class="sxs-lookup"><span data-stu-id="96367-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="96367-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-105">Requirements</span></span>

|<span data-ttu-id="96367-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-106">Requirement</span></span>| <span data-ttu-id="96367-107">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-108">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-109">1.0</span><span class="sxs-lookup"><span data-stu-id="96367-109">1.0</span></span>|
|[<span data-ttu-id="96367-110">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-111">Restringido</span><span class="sxs-lookup"><span data-stu-id="96367-111">Restricted</span></span>|
|[<span data-ttu-id="96367-112">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-113">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="96367-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="96367-114">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="96367-114">Members and methods</span></span>

| <span data-ttu-id="96367-115">Miembro	</span><span class="sxs-lookup"><span data-stu-id="96367-115">Member</span></span> | <span data-ttu-id="96367-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="96367-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="96367-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="96367-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="96367-118">Miembro	</span><span class="sxs-lookup"><span data-stu-id="96367-118">Member</span></span> |
| [<span data-ttu-id="96367-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="96367-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="96367-120">Miembro	</span><span class="sxs-lookup"><span data-stu-id="96367-120">Member</span></span> |
| [<span data-ttu-id="96367-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="96367-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="96367-122">Método</span><span class="sxs-lookup"><span data-stu-id="96367-122">Method</span></span> |
| [<span data-ttu-id="96367-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="96367-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="96367-124">Método</span><span class="sxs-lookup"><span data-stu-id="96367-124">Method</span></span> |
| [<span data-ttu-id="96367-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="96367-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) | <span data-ttu-id="96367-126">Método</span><span class="sxs-lookup"><span data-stu-id="96367-126">Method</span></span> |
| [<span data-ttu-id="96367-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="96367-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="96367-128">Método</span><span class="sxs-lookup"><span data-stu-id="96367-128">Method</span></span> |
| [<span data-ttu-id="96367-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="96367-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="96367-130">Método</span><span class="sxs-lookup"><span data-stu-id="96367-130">Method</span></span> |
| [<span data-ttu-id="96367-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="96367-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="96367-132">Método</span><span class="sxs-lookup"><span data-stu-id="96367-132">Method</span></span> |
| [<span data-ttu-id="96367-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="96367-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="96367-134">Método</span><span class="sxs-lookup"><span data-stu-id="96367-134">Method</span></span> |
| [<span data-ttu-id="96367-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="96367-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="96367-136">Método</span><span class="sxs-lookup"><span data-stu-id="96367-136">Method</span></span> |
| [<span data-ttu-id="96367-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="96367-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="96367-138">Método</span><span class="sxs-lookup"><span data-stu-id="96367-138">Method</span></span> |
| [<span data-ttu-id="96367-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="96367-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="96367-140">Método</span><span class="sxs-lookup"><span data-stu-id="96367-140">Method</span></span> |
| [<span data-ttu-id="96367-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="96367-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="96367-142">Método</span><span class="sxs-lookup"><span data-stu-id="96367-142">Method</span></span> |
| [<span data-ttu-id="96367-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="96367-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="96367-144">Método</span><span class="sxs-lookup"><span data-stu-id="96367-144">Method</span></span> |
| [<span data-ttu-id="96367-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="96367-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="96367-146">Método</span><span class="sxs-lookup"><span data-stu-id="96367-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="96367-147">Espacios de nombres</span><span class="sxs-lookup"><span data-stu-id="96367-147">Namespaces</span></span>

<span data-ttu-id="96367-148">[diagnostics](Office.context.mailbox.diagnostics.md): Proporciona información de diagnóstico a un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="96367-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="96367-149">[item](Office.context.mailbox.item.md): Proporciona métodos y propiedades para tener acceso a un mensaje o cita en un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="96367-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="96367-150">[userProfile](Office.context.mailbox.userProfile.md): Proporciona información sobre el usuario en un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="96367-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="96367-151">Miembros</span><span class="sxs-lookup"><span data-stu-id="96367-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="96367-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="96367-152">ewsUrl :String</span></span>

<span data-ttu-id="96367-p102">Obtiene la dirección URL del punto de conexión de Servicios Web Exchange (EWS) para esta cuenta de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="96367-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="96367-155">Este miembro no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="96367-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="96367-p103">El valor `ewsUrl` puede usarse por un servicio remoto para realizar llamadas EWS al buzón del usuario. Por ejemplo, puede crear un servicio remoto para [obtener datos adjuntos desde el elemento seleccionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="96367-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="96367-158">La aplicación debe tener el permiso **ReadItem** especificado en su manifiesto para poder llamar al miembro `ewsUrl` en el modo lectura.</span><span class="sxs-lookup"><span data-stu-id="96367-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="96367-p104">En el modo de redacción debe llamar al método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar el miembro `ewsUrl`. Su aplicación debe tener permisos **ReadWriteItem** para llamar al método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="96367-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="96367-161">Tipo:</span><span class="sxs-lookup"><span data-stu-id="96367-161">Type:</span></span>

*   <span data-ttu-id="96367-162">String</span><span class="sxs-lookup"><span data-stu-id="96367-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="96367-163">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-163">Requirements</span></span>

|<span data-ttu-id="96367-164">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-164">Requirement</span></span>| <span data-ttu-id="96367-165">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-166">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-166">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-167">1.0</span><span class="sxs-lookup"><span data-stu-id="96367-167">1.0</span></span>|
|[<span data-ttu-id="96367-168">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96367-169">ReadItem</span></span>|
|[<span data-ttu-id="96367-170">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-171">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="96367-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="96367-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="96367-172">restUrl :String</span></span>

<span data-ttu-id="96367-173">Obtiene la URL del punto de conexión REST para esta cuenta de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="96367-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="96367-174">El valor `restUrl` puede usarse para realizar llamadas [API de REST](https://docs.microsoft.com/outlook/rest/) al buzón del usuario.</span><span class="sxs-lookup"><span data-stu-id="96367-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="96367-175">La aplicación debe tener el permiso **ReadItem** especificado en su manifiesto para poder llamar al miembro `restUrl` en el modo lectura.</span><span class="sxs-lookup"><span data-stu-id="96367-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="96367-p105">En el modo de redacción debe llamar al método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar el miembro `restUrl`. Su aplicación debe tener permisos **ReadWriteItem** para llamar al método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="96367-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="96367-178">Tipo:</span><span class="sxs-lookup"><span data-stu-id="96367-178">Type:</span></span>

*   <span data-ttu-id="96367-179">String</span><span class="sxs-lookup"><span data-stu-id="96367-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="96367-180">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-180">Requirements</span></span>

|<span data-ttu-id="96367-181">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-181">Requirement</span></span>| <span data-ttu-id="96367-182">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-183">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-183">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-184">1,5</span><span class="sxs-lookup"><span data-stu-id="96367-184">1.5</span></span> |
|[<span data-ttu-id="96367-185">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96367-186">ReadItem</span></span>|
|[<span data-ttu-id="96367-187">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-188">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="96367-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="96367-189">Métodos</span><span class="sxs-lookup"><span data-stu-id="96367-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="96367-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="96367-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="96367-191">Agrega un controlador de eventos para un evento admitido.</span><span class="sxs-lookup"><span data-stu-id="96367-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="96367-p106">Actualmente, el único tipo de evento admitido es `Office.EventType.ItemChanged`, que se invoca cuando el usuario selecciona un elemento nuevo. Este evento se usa por los complementos que implementan un panel de tareas anclable, y permite que el complemento actualice la interfaz de usuario del panel de tareas basándose en el elemento seleccionado actualmente.</span><span class="sxs-lookup"><span data-stu-id="96367-p106">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96367-194">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="96367-194">Parameters:</span></span>

| <span data-ttu-id="96367-195">Nombre</span><span class="sxs-lookup"><span data-stu-id="96367-195">Name</span></span> | <span data-ttu-id="96367-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="96367-196">Type</span></span> | <span data-ttu-id="96367-197">Atributos</span><span class="sxs-lookup"><span data-stu-id="96367-197">Attributes</span></span> | <span data-ttu-id="96367-198">Descripción</span><span class="sxs-lookup"><span data-stu-id="96367-198">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="96367-199">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="96367-199">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="96367-200">El evento que debe invocar el controlador.</span><span class="sxs-lookup"><span data-stu-id="96367-200">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="96367-201">Función</span><span class="sxs-lookup"><span data-stu-id="96367-201">Function</span></span> || <span data-ttu-id="96367-p107">La función que va a controlar el evento. La función debe aceptar un único parámetro, que es un literal de objeto. La propiedad `type` del parámetro coincidirá con el parámetro `eventType` que se ha pasado a `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="96367-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="96367-205">Objeto</span><span class="sxs-lookup"><span data-stu-id="96367-205">Object</span></span> | <span data-ttu-id="96367-206">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-206">&lt;optional&gt;</span></span> | <span data-ttu-id="96367-207">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="96367-207">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="96367-208">Objeto</span><span class="sxs-lookup"><span data-stu-id="96367-208">Object</span></span> | <span data-ttu-id="96367-209">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-209">&lt;optional&gt;</span></span> | <span data-ttu-id="96367-210">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="96367-210">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="96367-211">función</span><span class="sxs-lookup"><span data-stu-id="96367-211">function</span></span>| <span data-ttu-id="96367-212">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-212">&lt;optional&gt;</span></span>|<span data-ttu-id="96367-213">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="96367-213">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96367-214">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-214">Requirements</span></span>

|<span data-ttu-id="96367-215">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-215">Requirement</span></span>| <span data-ttu-id="96367-216">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-217">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-217">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-218">1,5</span><span class="sxs-lookup"><span data-stu-id="96367-218">1.5</span></span> |
|[<span data-ttu-id="96367-219">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-219">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-220">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96367-220">ReadItem</span></span> |
|[<span data-ttu-id="96367-221">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-221">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-222">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="96367-222">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96367-223">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="96367-223">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="96367-224">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="96367-224">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="96367-225">Convierte un identificador de elemento con formato para REST al formato EWS.</span><span class="sxs-lookup"><span data-stu-id="96367-225">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="96367-226">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="96367-226">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="96367-p108">Los identificadores de elemento obtenidos a través de una API de REST (como la [API de correo de Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) o [Microsoft Graph](http://graph.microsoft.io/)) usan un formato diferente al formato que usa Exchange Web Services (EWS). El método `convertToEwsId` convierte un identificador con formato REST al formato adecuado para EWS.</span><span class="sxs-lookup"><span data-stu-id="96367-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96367-229">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="96367-229">Parameters:</span></span>

|<span data-ttu-id="96367-230">Nombre</span><span class="sxs-lookup"><span data-stu-id="96367-230">Name</span></span>| <span data-ttu-id="96367-231">Tipo</span><span class="sxs-lookup"><span data-stu-id="96367-231">Type</span></span>| <span data-ttu-id="96367-232">Descripción</span><span class="sxs-lookup"><span data-stu-id="96367-232">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="96367-233">String</span><span class="sxs-lookup"><span data-stu-id="96367-233">String</span></span>|<span data-ttu-id="96367-234">Un identificador de elemento con formato para las API de REST de Outlook</span><span class="sxs-lookup"><span data-stu-id="96367-234">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="96367-235">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="96367-235">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="96367-236">Un valor que indica la versión de la API de REST de Outlook que se usa para recuperar el identificador de elemento.</span><span class="sxs-lookup"><span data-stu-id="96367-236">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96367-237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-237">Requirements</span></span>

|<span data-ttu-id="96367-238">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-238">Requirement</span></span>| <span data-ttu-id="96367-239">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-240">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-241">1.3</span><span class="sxs-lookup"><span data-stu-id="96367-241">1.3</span></span>|
|[<span data-ttu-id="96367-242">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-243">Restringido</span><span class="sxs-lookup"><span data-stu-id="96367-243">Restricted</span></span>|
|[<span data-ttu-id="96367-244">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-245">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="96367-245">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="96367-246">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="96367-246">Returns:</span></span>

<span data-ttu-id="96367-247">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="96367-247">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="96367-248">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="96367-248">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="96367-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="96367-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="96367-250">Obtiene un diccionario con información de tiempo en el tiempo del cliente local.</span><span class="sxs-lookup"><span data-stu-id="96367-250">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="96367-p109">Las fechas y horas usadas por una aplicación de correo para Outlook o Outlook Web App pueden usar distintas zonas horarias. Outlook usa la zona horaria del equipo cliente; Outlook Web App usa la zona horaria definida en el Centro de administración de Exchange (EAC). Debería tratar los valores de fecha y hora de modo que los valores que aparezcan en la interfaz de usuario sean siempre coherentes con la zona horaria que el usuario espera.</span><span class="sxs-lookup"><span data-stu-id="96367-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="96367-p110">Si se está ejecutando la aplicación de correo en Outlook, el método `convertToLocalClientTime` devolverá un objeto de diccionario con los valores establecidos para la zona horaria del equipo cliente. Si se está ejecutando la aplicación de correo en Outlook Web App, el método `convertToLocalClientTime` devolverá un objeto de diccionario con los valores establecidos para la zona horaria especificada en el CEF.</span><span class="sxs-lookup"><span data-stu-id="96367-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96367-256">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="96367-256">Parameters:</span></span>

|<span data-ttu-id="96367-257">Nombre</span><span class="sxs-lookup"><span data-stu-id="96367-257">Name</span></span>| <span data-ttu-id="96367-258">Tipo</span><span class="sxs-lookup"><span data-stu-id="96367-258">Type</span></span>| <span data-ttu-id="96367-259">Descripción</span><span class="sxs-lookup"><span data-stu-id="96367-259">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="96367-260">Fecha</span><span class="sxs-lookup"><span data-stu-id="96367-260">Date</span></span>|<span data-ttu-id="96367-261">Un objeto Date</span><span class="sxs-lookup"><span data-stu-id="96367-261">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96367-262">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-262">Requirements</span></span>

|<span data-ttu-id="96367-263">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-263">Requirement</span></span>| <span data-ttu-id="96367-264">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-265">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-266">1.0</span><span class="sxs-lookup"><span data-stu-id="96367-266">1.0</span></span>|
|[<span data-ttu-id="96367-267">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96367-268">ReadItem</span></span>|
|[<span data-ttu-id="96367-269">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-270">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="96367-270">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="96367-271">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="96367-271">Returns:</span></span>

<span data-ttu-id="96367-272">Tipo: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="96367-272">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="96367-273">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="96367-273">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="96367-274">Convierte un identificador de elemento con formato para EWS al formato REST.</span><span class="sxs-lookup"><span data-stu-id="96367-274">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="96367-275">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="96367-275">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="96367-p111">Los identificadores de elemento obtenidos a través de EWS o de la propiedad `itemId` usan un formato diferente al formato que usan las API de REST (como la [API de correo de Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) o [Microsoft Graph](http://graph.microsoft.io/)). El método `convertToRestId` convierte un identificador con formato EWS al formato adecuado para REST.</span><span class="sxs-lookup"><span data-stu-id="96367-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96367-278">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="96367-278">Parameters:</span></span>

|<span data-ttu-id="96367-279">Nombre</span><span class="sxs-lookup"><span data-stu-id="96367-279">Name</span></span>| <span data-ttu-id="96367-280">Tipo</span><span class="sxs-lookup"><span data-stu-id="96367-280">Type</span></span>| <span data-ttu-id="96367-281">Descripción</span><span class="sxs-lookup"><span data-stu-id="96367-281">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="96367-282">String</span><span class="sxs-lookup"><span data-stu-id="96367-282">String</span></span>|<span data-ttu-id="96367-283">Un identificador de elemento con formato para Exchange Web Services (EWS)</span><span class="sxs-lookup"><span data-stu-id="96367-283">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="96367-284">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="96367-284">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="96367-285">Un valor que indica la versión de la API de REST de Outlook con la que se usará el identificador convertido.</span><span class="sxs-lookup"><span data-stu-id="96367-285">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96367-286">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-286">Requirements</span></span>

|<span data-ttu-id="96367-287">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-287">Requirement</span></span>| <span data-ttu-id="96367-288">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-289">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-289">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-290">1.3</span><span class="sxs-lookup"><span data-stu-id="96367-290">1.3</span></span>|
|[<span data-ttu-id="96367-291">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-291">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-292">Restringido</span><span class="sxs-lookup"><span data-stu-id="96367-292">Restricted</span></span>|
|[<span data-ttu-id="96367-293">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-293">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-294">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="96367-294">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="96367-295">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="96367-295">Returns:</span></span>

<span data-ttu-id="96367-296">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="96367-296">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="96367-297">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="96367-297">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="96367-298">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="96367-298">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="96367-299">Obtiene un objeto Date del diccionario que contiene información de tiempo.</span><span class="sxs-lookup"><span data-stu-id="96367-299">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="96367-300">El método `convertToUtcClientTime` convierte un diccionario que contiene la fecha y la hora locales en un objeto Date con los valores correctos para la fecha y la hora locales.</span><span class="sxs-lookup"><span data-stu-id="96367-300">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96367-301">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="96367-301">Parameters:</span></span>

|<span data-ttu-id="96367-302">Nombre</span><span class="sxs-lookup"><span data-stu-id="96367-302">Name</span></span>| <span data-ttu-id="96367-303">Tipo</span><span class="sxs-lookup"><span data-stu-id="96367-303">Type</span></span>| <span data-ttu-id="96367-304">Descripción</span><span class="sxs-lookup"><span data-stu-id="96367-304">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="96367-305">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="96367-305">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="96367-306">El valor de la hora local para convertir.</span><span class="sxs-lookup"><span data-stu-id="96367-306">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96367-307">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-307">Requirements</span></span>

|<span data-ttu-id="96367-308">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-308">Requirement</span></span>| <span data-ttu-id="96367-309">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-309">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-310">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-310">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-311">1.0</span><span class="sxs-lookup"><span data-stu-id="96367-311">1.0</span></span>|
|[<span data-ttu-id="96367-312">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-312">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-313">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96367-313">ReadItem</span></span>|
|[<span data-ttu-id="96367-314">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-314">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-315">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="96367-315">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="96367-316">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="96367-316">Returns:</span></span>

<span data-ttu-id="96367-317">Objeto Date con el tiempo expresado en UTC.</span><span class="sxs-lookup"><span data-stu-id="96367-317">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="96367-318">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="96367-318">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="96367-319">Fecha</span><span class="sxs-lookup"><span data-stu-id="96367-319">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="96367-320">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="96367-320">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="96367-321">Muestra una cita de calendario existente.</span><span class="sxs-lookup"><span data-stu-id="96367-321">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="96367-322">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="96367-322">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="96367-323">El método `displayAppointmentForm` abre una cita de calendario existente en una nueva ventana del escritorio o en un cuadro de diálogo en los dispositivos móviles.</span><span class="sxs-lookup"><span data-stu-id="96367-323">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="96367-p112">En Outlook para Mac, puede usar este método para mostrar una cita que no forme parte de una serie periódica o la cita principal de una serie periódica, pero no puede mostrar una instancia de la serie. Esto es porque en Outlook para Mac no se puede tener acceso a las propiedades (incluido el identificador de elemento) de las instancias de una serie periódica.</span><span class="sxs-lookup"><span data-stu-id="96367-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="96367-326">En Outlook Web App, este método abre el formato especificado solo si el cuerpo del formulario es inferior o igual a 32 KB en el número de caracteres.</span><span class="sxs-lookup"><span data-stu-id="96367-326">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="96367-327">Si el identificador de elemento especificado no identifica una cita existente, se abrirá una página en blanco en el dispositivo o equipo cliente y no se generará ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="96367-327">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96367-328">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="96367-328">Parameters:</span></span>

|<span data-ttu-id="96367-329">Nombre</span><span class="sxs-lookup"><span data-stu-id="96367-329">Name</span></span>| <span data-ttu-id="96367-330">Tipo</span><span class="sxs-lookup"><span data-stu-id="96367-330">Type</span></span>| <span data-ttu-id="96367-331">Descripción</span><span class="sxs-lookup"><span data-stu-id="96367-331">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="96367-332">String</span><span class="sxs-lookup"><span data-stu-id="96367-332">String</span></span>|<span data-ttu-id="96367-333">Identificador de los servicios web de Exchange (EWS) para una cita de calendario existente.</span><span class="sxs-lookup"><span data-stu-id="96367-333">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96367-334">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-334">Requirements</span></span>

|<span data-ttu-id="96367-335">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-335">Requirement</span></span>| <span data-ttu-id="96367-336">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-337">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-337">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-338">1.0</span><span class="sxs-lookup"><span data-stu-id="96367-338">1.0</span></span>|
|[<span data-ttu-id="96367-339">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-339">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96367-340">ReadItem</span></span>|
|[<span data-ttu-id="96367-341">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-341">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-342">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="96367-342">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96367-343">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="96367-343">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="96367-344">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="96367-344">displayMessageForm(itemId)</span></span>

<span data-ttu-id="96367-345">Muestra un mensaje existente.</span><span class="sxs-lookup"><span data-stu-id="96367-345">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="96367-346">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="96367-346">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="96367-347">El método `displayMessageForm` abre un mensaje existente en una nueva ventana del escritorio o en un cuadro de diálogo en los dispositivos móviles.</span><span class="sxs-lookup"><span data-stu-id="96367-347">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="96367-348">En Outlook Web App, este método abre el formato especificado solo si el cuerpo del formulario es inferior o igual a 32 KB en el número de caracteres.</span><span class="sxs-lookup"><span data-stu-id="96367-348">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="96367-349">Si el identificador de elemento especificado no identifica un mensaje existente, no se mostrará ningún mensaje en el equipo cliente ni se generará ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="96367-349">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="96367-p113">No use el valor `displayMessageForm` con un `itemId` que represente una cita. Use el método `displayAppointmentForm` para mostrar una cita existente y `displayNewAppointmentForm` para mostrar un formulario para crear una cita nueva.</span><span class="sxs-lookup"><span data-stu-id="96367-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96367-352">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="96367-352">Parameters:</span></span>

|<span data-ttu-id="96367-353">Nombre</span><span class="sxs-lookup"><span data-stu-id="96367-353">Name</span></span>| <span data-ttu-id="96367-354">Tipo</span><span class="sxs-lookup"><span data-stu-id="96367-354">Type</span></span>| <span data-ttu-id="96367-355">Descripción</span><span class="sxs-lookup"><span data-stu-id="96367-355">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="96367-356">String</span><span class="sxs-lookup"><span data-stu-id="96367-356">String</span></span>|<span data-ttu-id="96367-357">Identificador de los servicios web de Exchange (EWS) para un mensaje existente.</span><span class="sxs-lookup"><span data-stu-id="96367-357">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96367-358">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-358">Requirements</span></span>

|<span data-ttu-id="96367-359">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-359">Requirement</span></span>| <span data-ttu-id="96367-360">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-361">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-362">1.0</span><span class="sxs-lookup"><span data-stu-id="96367-362">1.0</span></span>|
|[<span data-ttu-id="96367-363">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96367-364">ReadItem</span></span>|
|[<span data-ttu-id="96367-365">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-366">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="96367-366">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96367-367">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="96367-367">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="96367-368">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="96367-368">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="96367-369">Muestra un formulario para crear una nueva cita de calendario.</span><span class="sxs-lookup"><span data-stu-id="96367-369">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="96367-370">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="96367-370">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="96367-p114">El método `displayNewAppointmentForm` abre un formulario que permite al usuario crear una nueva cita o reunión. Si se especifican parámetros, los campos de formulario de cita se rellenan automáticamente con el contenido de los parámetros.</span><span class="sxs-lookup"><span data-stu-id="96367-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="96367-p115">En Outlook Web App y OWA para dispositivos, este método muestra siempre un formulario con un campo de asistentes. Si no especifica ningún asistente como argumento de entrada, el método muestra un formulario con un botón **Guardar**. Si ha especificado asistentes, el formulario incluirá a los asistentes y un botón **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="96367-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="96367-p116">En el cliente enriquecido de Outlook y Outlook RT, si se especifica cualquier asistente o recurso en los parámetros `requiredAttendees`, `optionalAttendees` o `resources`, este método muestra un formulario de reunión con un botón **Enviar**. Si no se especifica ningún destinatario, este método muestra un formulario de cita con un botón **Guardar y cerrar**.</span><span class="sxs-lookup"><span data-stu-id="96367-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="96367-378">Si cualquiera de los parámetros supera los límites de tamaño especificados o si se especifica un nombre de parámetro desconocido, se genera una excepción.</span><span class="sxs-lookup"><span data-stu-id="96367-378">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96367-379">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="96367-379">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="96367-380">Todos los parámetros son opcionales.</span><span class="sxs-lookup"><span data-stu-id="96367-380">All parameters are optional.</span></span>

|<span data-ttu-id="96367-381">Nombre</span><span class="sxs-lookup"><span data-stu-id="96367-381">Name</span></span>| <span data-ttu-id="96367-382">Tipo</span><span class="sxs-lookup"><span data-stu-id="96367-382">Type</span></span>| <span data-ttu-id="96367-383">Descripción</span><span class="sxs-lookup"><span data-stu-id="96367-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="96367-384">Object</span><span class="sxs-lookup"><span data-stu-id="96367-384">Object</span></span> | <span data-ttu-id="96367-385">Un diccionario de parámetros que describen la nueva cita.</span><span class="sxs-lookup"><span data-stu-id="96367-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="96367-386">Matriz.&lt;Cadena&gt; &#124; Matriz.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="96367-p117">Una matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los asistentes necesarios de la cita. La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="96367-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="96367-389">Matriz.&lt;Cadena&gt; &#124; Matriz.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="96367-p118">Una matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los asistentes opcionales de la cita. La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="96367-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="96367-392">Fecha</span><span class="sxs-lookup"><span data-stu-id="96367-392">Date</span></span> | <span data-ttu-id="96367-393">Un objeto `Date` que especifica la fecha y hora de inicio de la cita.</span><span class="sxs-lookup"><span data-stu-id="96367-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="96367-394">Fecha</span><span class="sxs-lookup"><span data-stu-id="96367-394">Date</span></span> | <span data-ttu-id="96367-395">Un objeto `Date` que especifica la fecha y hora de finalización de la cita.</span><span class="sxs-lookup"><span data-stu-id="96367-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="96367-396">String</span><span class="sxs-lookup"><span data-stu-id="96367-396">String</span></span> | <span data-ttu-id="96367-p119">Una cadena que contiene la ubicación de la cita. La cadena está limitada a un máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="96367-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="96367-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="96367-p120">Una matriz de cadenas que contiene los recursos necesarios para la cita. La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="96367-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="96367-402">String</span><span class="sxs-lookup"><span data-stu-id="96367-402">String</span></span> | <span data-ttu-id="96367-p121">Una cadena que contiene al asunto de la cita. La cadena está limitada a un máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="96367-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="96367-405">String</span><span class="sxs-lookup"><span data-stu-id="96367-405">String</span></span> | <span data-ttu-id="96367-p122">El cuerpo de la cita. El contenido del cuerpo está limitado a un tamaño máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="96367-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="96367-408">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-408">Requirements</span></span>

|<span data-ttu-id="96367-409">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-409">Requirement</span></span>| <span data-ttu-id="96367-410">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-411">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-411">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-412">1.0</span><span class="sxs-lookup"><span data-stu-id="96367-412">1.0</span></span>|
|[<span data-ttu-id="96367-413">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96367-414">ReadItem</span></span>|
|[<span data-ttu-id="96367-415">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-416">Lectura</span><span class="sxs-lookup"><span data-stu-id="96367-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96367-417">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="96367-417">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="96367-418">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="96367-418">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="96367-419">Muestra un formulario para crear un mensaje nuevo.</span><span class="sxs-lookup"><span data-stu-id="96367-419">Displays a form for creating a new message.</span></span>

<span data-ttu-id="96367-420">El método `displayNewMessageForm` abre un formulario que permite al usuario crear un mensaje nuevo.</span><span class="sxs-lookup"><span data-stu-id="96367-420">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="96367-421">Si se especifican parámetros, los campos de formulario de mensaje se rellenan automáticamente con el contenido de los parámetros.</span><span class="sxs-lookup"><span data-stu-id="96367-421">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="96367-422">Si cualquiera de los parámetros supera los límites de tamaño especificados o si se indica un nombre de parámetro desconocido, se genera una excepción.</span><span class="sxs-lookup"><span data-stu-id="96367-422">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96367-423">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="96367-423">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="96367-424">Todos los parámetros son opcionales.</span><span class="sxs-lookup"><span data-stu-id="96367-424">All parameters are optional.</span></span>

|<span data-ttu-id="96367-425">Nombre</span><span class="sxs-lookup"><span data-stu-id="96367-425">Name</span></span>| <span data-ttu-id="96367-426">Tipo</span><span class="sxs-lookup"><span data-stu-id="96367-426">Type</span></span>| <span data-ttu-id="96367-427">Descripción</span><span class="sxs-lookup"><span data-stu-id="96367-427">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="96367-428">Objeto</span><span class="sxs-lookup"><span data-stu-id="96367-428">Object</span></span> | <span data-ttu-id="96367-429">Diccionario de parámetros que describen el mensaje nuevo.</span><span class="sxs-lookup"><span data-stu-id="96367-429">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="96367-430">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-430">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="96367-431">Matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios de la línea Para.</span><span class="sxs-lookup"><span data-stu-id="96367-431">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="96367-432">La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="96367-432">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="96367-433">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-433">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="96367-434">Matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios de la línea CC.</span><span class="sxs-lookup"><span data-stu-id="96367-434">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="96367-435">La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="96367-435">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="96367-436">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-436">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="96367-437">Matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios de la línea CCO.</span><span class="sxs-lookup"><span data-stu-id="96367-437">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="96367-438">La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="96367-438">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="96367-439">String</span><span class="sxs-lookup"><span data-stu-id="96367-439">String</span></span> | <span data-ttu-id="96367-440">Cadena que contiene el asunto del mensaje.</span><span class="sxs-lookup"><span data-stu-id="96367-440">A string containing the subject of the message.</span></span> <span data-ttu-id="96367-441">La cadena está limitada a un máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="96367-441">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="96367-442">String</span><span class="sxs-lookup"><span data-stu-id="96367-442">String</span></span> | <span data-ttu-id="96367-443">Cuerpo HTML del mensaje.</span><span class="sxs-lookup"><span data-stu-id="96367-443">The HTML body of the message.</span></span> <span data-ttu-id="96367-444">El contenido del cuerpo está limitado a un tamaño máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="96367-444">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="96367-445">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-445">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="96367-446">Matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="96367-446">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="96367-447">String</span><span class="sxs-lookup"><span data-stu-id="96367-447">String</span></span> | <span data-ttu-id="96367-p129">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="96367-p129">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="96367-450">String</span><span class="sxs-lookup"><span data-stu-id="96367-450">String</span></span> | <span data-ttu-id="96367-451">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="96367-451">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="96367-452">String</span><span class="sxs-lookup"><span data-stu-id="96367-452">String</span></span> | <span data-ttu-id="96367-p130">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="96367-p130">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="96367-455">Booleano</span><span class="sxs-lookup"><span data-stu-id="96367-455">Boolean</span></span> | <span data-ttu-id="96367-p131">Solo se usa si `type` se establece en `file`. Si `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="96367-p131">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="96367-458">String</span><span class="sxs-lookup"><span data-stu-id="96367-458">String</span></span> | <span data-ttu-id="96367-459">Solo se usa si `type` se establece en `item`.</span><span class="sxs-lookup"><span data-stu-id="96367-459">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="96367-460">El identificador de elemento EWS del correo electrónico existente que desea adjuntar al mensaje de nuevo.</span><span class="sxs-lookup"><span data-stu-id="96367-460">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="96367-461">Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="96367-461">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="96367-462">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-462">Requirements</span></span>

|<span data-ttu-id="96367-463">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-463">Requirement</span></span>| <span data-ttu-id="96367-464">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-464">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-465">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-465">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-466">1.6</span><span class="sxs-lookup"><span data-stu-id="96367-466">1.6</span></span> |
|[<span data-ttu-id="96367-467">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-467">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-468">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96367-468">ReadItem</span></span>|
|[<span data-ttu-id="96367-469">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-469">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-470">Lectura</span><span class="sxs-lookup"><span data-stu-id="96367-470">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96367-471">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="96367-471">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="96367-472">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="96367-472">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="96367-473">Obtiene una cadena que contiene un token usado para llamar a las API de REST o a los servicios Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="96367-473">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="96367-p133">El método `getCallbackTokenAsync` realiza una llamada asincrónica para obtener un token opaco desde Exchange Server que hospeda el buzón del usuario. La duración del token de devolución de llamada es de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="96367-p133">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="96367-476">Se recomienda que los complementos de usar las API de REST en lugar de los servicios Web Exchange siempre que sea posible.</span><span class="sxs-lookup"><span data-stu-id="96367-476">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="96367-477">**Tokens de REST**</span><span class="sxs-lookup"><span data-stu-id="96367-477">**REST Tokens**</span></span>

<span data-ttu-id="96367-p134">Cuando se solicita un token de REST (`options.isRest = true`), el token resultante no funcionará para autenticar las llamadas a los servicios Web Exchange. El token se limitará en su ámbito para el acceso de solo lectura al elemento actual y a sus datos adjuntos, a no ser que el complemento haya especificado el permiso [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) en su manifiesto. Si está especificado el permiso `ReadWriteMailbox`, el token resultante concederá acceso de lectura o escritura al correo, calendario y contactos, incluida la capacidad de enviar correo.</span><span class="sxs-lookup"><span data-stu-id="96367-p134">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="96367-481">El complemento debe usar la propiedad `restUrl` para determinar la URL correcta que se va a usar al realizar las llamadas a la API de REST.</span><span class="sxs-lookup"><span data-stu-id="96367-481">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="96367-482">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="96367-482">**EWS Tokens**</span></span>

<span data-ttu-id="96367-p135">Cuando se solicita un token EWS (`options.isRest = false`), el token resultante no funcionará para autenticar las llamadas a las API de REST. El token estará limitado en su ámbito para el acceso al elemento actual.</span><span class="sxs-lookup"><span data-stu-id="96367-p135">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="96367-485">El complemento debe usar la propiedad `ewsUrl` para determinar la URL correcta que se va a usar al realizar las llamadas a EWS.</span><span class="sxs-lookup"><span data-stu-id="96367-485">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96367-486">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="96367-486">Parameters:</span></span>

|<span data-ttu-id="96367-487">Nombre</span><span class="sxs-lookup"><span data-stu-id="96367-487">Name</span></span>| <span data-ttu-id="96367-488">Tipo</span><span class="sxs-lookup"><span data-stu-id="96367-488">Type</span></span>| <span data-ttu-id="96367-489">Atributos</span><span class="sxs-lookup"><span data-stu-id="96367-489">Attributes</span></span>| <span data-ttu-id="96367-490">Descripción</span><span class="sxs-lookup"><span data-stu-id="96367-490">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="96367-491">Objeto</span><span class="sxs-lookup"><span data-stu-id="96367-491">Object</span></span> | <span data-ttu-id="96367-492">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-492">&lt;optional&gt;</span></span> | <span data-ttu-id="96367-493">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="96367-493">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="96367-494">Booleano</span><span class="sxs-lookup"><span data-stu-id="96367-494">Boolean</span></span> |  <span data-ttu-id="96367-495">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-495">&lt;optional&gt;</span></span> | <span data-ttu-id="96367-p136">Determina si el token proporcionado se usará para las API de REST de Outlook o para los servicios Web Exchange. El valor predeterminado es `false`.</span><span class="sxs-lookup"><span data-stu-id="96367-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="96367-498">Objeto</span><span class="sxs-lookup"><span data-stu-id="96367-498">Object</span></span> |  <span data-ttu-id="96367-499">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-499">&lt;optional&gt;</span></span> | <span data-ttu-id="96367-500">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="96367-500">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="96367-501">Función</span><span class="sxs-lookup"><span data-stu-id="96367-501">function</span></span>||<span data-ttu-id="96367-p137">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). El token se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="96367-p137">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96367-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-504">Requirements</span></span>

|<span data-ttu-id="96367-505">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-505">Requirement</span></span>| <span data-ttu-id="96367-506">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-507">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-508">1,5</span><span class="sxs-lookup"><span data-stu-id="96367-508">1.5</span></span> |
|[<span data-ttu-id="96367-509">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96367-510">ReadItem</span></span>|
|[<span data-ttu-id="96367-511">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-512">Redacción y lectura</span><span class="sxs-lookup"><span data-stu-id="96367-512">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="96367-513">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="96367-513">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="96367-514">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="96367-514">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="96367-515">Obtiene una cadena que contiene un token usado para obtener datos adjuntos o un elemento de Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="96367-515">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="96367-p138">El método `getCallbackTokenAsync` realiza una llamada asincrónica para obtener un token opaco desde Exchange Server que hospeda el buzón del usuario. La duración del token de devolución de llamada es de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="96367-p138">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="96367-p139">Puede pasar el token y un identificador de archivo adjunto o el identificador del elemento a un sistema de terceros. El sistema de terceros usa el token como token de autorización del portador para llamar a los Servicios web de Exchange (EWS) o a las operaciones [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) o [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) para devolver datos adjuntos o un elemento. Por ejemplo, puede crear un servicio remoto para [obtener datos adjuntos desde el elemento seleccionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="96367-p139">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="96367-521">La aplicación debe tener el permiso **ReadItem** especificado en su manifiesto para poder llamar al método `getCallbackTokenAsync` en el modo lectura.</span><span class="sxs-lookup"><span data-stu-id="96367-521">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="96367-p140">En el modo de redacción debe llamar al método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obtener un identificador de elemento para pasar al método `getCallbackTokenAsync`. Su aplicación debe tener permisos **ReadWriteItem** para llamar al método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="96367-p140">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96367-524">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="96367-524">Parameters:</span></span>

|<span data-ttu-id="96367-525">Nombre</span><span class="sxs-lookup"><span data-stu-id="96367-525">Name</span></span>| <span data-ttu-id="96367-526">Tipo</span><span class="sxs-lookup"><span data-stu-id="96367-526">Type</span></span>| <span data-ttu-id="96367-527">Atributos</span><span class="sxs-lookup"><span data-stu-id="96367-527">Attributes</span></span>| <span data-ttu-id="96367-528">Descripción</span><span class="sxs-lookup"><span data-stu-id="96367-528">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="96367-529">función</span><span class="sxs-lookup"><span data-stu-id="96367-529">function</span></span>||<span data-ttu-id="96367-p141">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). El token se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="96367-p141">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="96367-532">Objeto</span><span class="sxs-lookup"><span data-stu-id="96367-532">Object</span></span>| <span data-ttu-id="96367-533">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-533">&lt;optional&gt;</span></span>|<span data-ttu-id="96367-534">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="96367-534">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96367-535">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-535">Requirements</span></span>

|<span data-ttu-id="96367-536">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-536">Requirement</span></span>| <span data-ttu-id="96367-537">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-538">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-538">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-539">1.3</span><span class="sxs-lookup"><span data-stu-id="96367-539">1.3</span></span>|
|[<span data-ttu-id="96367-540">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-540">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96367-541">ReadItem</span></span>|
|[<span data-ttu-id="96367-542">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-542">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-543">Redacción y lectura</span><span class="sxs-lookup"><span data-stu-id="96367-543">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="96367-544">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="96367-544">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="96367-545">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="96367-545">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="96367-546">Obtiene un token que identifica al usuario y al complemento de Office.</span><span class="sxs-lookup"><span data-stu-id="96367-546">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="96367-547">El método `getUserIdentityTokenAsync` devuelve un token que puede usar para identificar y [autenticar el complemento y el usuario mediante un sistema de terceros](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="96367-547">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="96367-548">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="96367-548">Parameters:</span></span>

|<span data-ttu-id="96367-549">Nombre</span><span class="sxs-lookup"><span data-stu-id="96367-549">Name</span></span>| <span data-ttu-id="96367-550">Tipo</span><span class="sxs-lookup"><span data-stu-id="96367-550">Type</span></span>| <span data-ttu-id="96367-551">Atributos</span><span class="sxs-lookup"><span data-stu-id="96367-551">Attributes</span></span>| <span data-ttu-id="96367-552">Descripción</span><span class="sxs-lookup"><span data-stu-id="96367-552">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="96367-553">función</span><span class="sxs-lookup"><span data-stu-id="96367-553">function</span></span>||<span data-ttu-id="96367-554">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="96367-554">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="96367-555">El token se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="96367-555">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="96367-556">Object</span><span class="sxs-lookup"><span data-stu-id="96367-556">Object</span></span>| <span data-ttu-id="96367-557">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-557">&lt;optional&gt;</span></span>|<span data-ttu-id="96367-558">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="96367-558">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96367-559">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-559">Requirements</span></span>

|<span data-ttu-id="96367-560">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-560">Requirement</span></span>| <span data-ttu-id="96367-561">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-562">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-562">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-563">1.0</span><span class="sxs-lookup"><span data-stu-id="96367-563">1.0</span></span>|
|[<span data-ttu-id="96367-564">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-564">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96367-565">ReadItem</span></span>|
|[<span data-ttu-id="96367-566">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-566">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-567">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="96367-567">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96367-568">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="96367-568">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="96367-569">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="96367-569">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="96367-570">Realiza una solicitud asincrónica a un servicio de Servicios Web Exchange (EWS) en el servidor Exchange que hospeda el buzón del usuario.</span><span class="sxs-lookup"><span data-stu-id="96367-570">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="96367-571">Este método no se admite en los siguientes escenarios.</span><span class="sxs-lookup"><span data-stu-id="96367-571">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="96367-572">En Outlook para iOS o Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="96367-572">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="96367-573">Cuando el complemento se carga en un buzón de Gmail</span><span class="sxs-lookup"><span data-stu-id="96367-573">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="96367-574">En estos casos, complementos deben [usar API de REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) para tener acceso al buzón del usuario en su lugar.</span><span class="sxs-lookup"><span data-stu-id="96367-574">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="96367-575">El método `makeEwsRequestAsync` envía una solicitud de EWS en nombre del complemento a Exchange.</span><span class="sxs-lookup"><span data-stu-id="96367-575">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="96367-576">Para obtener una lista de las operaciones de EWS compatibles, vea [llamar a servicios web desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="96367-576">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="96367-577">No puede solicitar elementos asociados de las carpetas con el método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="96367-577">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="96367-578">La solicitud XML debe especificar codificación UTF-8.</span><span class="sxs-lookup"><span data-stu-id="96367-578">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="96367-p143">Su complemento debe tener el permiso **ReadWriteMailbox** para usar el método `makeEwsRequestAsync`. Para obtener información sobre cómo usar el permiso **ReadWriteMailbox** y sobre las operaciones EWS que puede llamar con el método `makeEwsRequestAsync`, consulte [Especificar permisos para el acceso del complemento de correo al buzón del usuario](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="96367-p143">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="96367-581">El administrador del servidor debe establecer `OAuthAuthentication` en true en el directorio EWS del servidor de acceso de cliente para habilitar la `makeEwsRequestAsync` método para realizar EWS solicita.</span><span class="sxs-lookup"><span data-stu-id="96367-581">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="96367-582">Diferencias de versión</span><span class="sxs-lookup"><span data-stu-id="96367-582">Version differences</span></span>

<span data-ttu-id="96367-583">Si usa el método `makeEwsRequestAsync` en aplicaciones de correo que se ejecutan en versiones de Outlook anteriores a 15.0.4535.1004, debe establecer el valor de codificación a `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="96367-583">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="96367-p144">No es necesario establecer el valor de codificación si la aplicación de correo se ejecuta en Outlook en la web. Puede averiguar si su aplicación de correo se ejecuta en Outlook o en Outlook en la web usando la propiedad mailbox.diagnostics.hostName. Para averiguar qué versión de Outlook se está ejecutando, use la propiedad mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="96367-p144">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96367-587">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="96367-587">Parameters:</span></span>

|<span data-ttu-id="96367-588">Nombre</span><span class="sxs-lookup"><span data-stu-id="96367-588">Name</span></span>| <span data-ttu-id="96367-589">Tipo</span><span class="sxs-lookup"><span data-stu-id="96367-589">Type</span></span>| <span data-ttu-id="96367-590">Atributos</span><span class="sxs-lookup"><span data-stu-id="96367-590">Attributes</span></span>| <span data-ttu-id="96367-591">Descripción</span><span class="sxs-lookup"><span data-stu-id="96367-591">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="96367-592">String</span><span class="sxs-lookup"><span data-stu-id="96367-592">String</span></span>||<span data-ttu-id="96367-593">La solicitud de EWS.</span><span class="sxs-lookup"><span data-stu-id="96367-593">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="96367-594">función</span><span class="sxs-lookup"><span data-stu-id="96367-594">function</span></span>||<span data-ttu-id="96367-595">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="96367-595">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="96367-596">El resultado XML de la llamada EWS se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="96367-596">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="96367-597">Si el resultado supera 1 MB de tamaño, se devuelve un mensaje de error en su lugar.</span><span class="sxs-lookup"><span data-stu-id="96367-597">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="96367-598">Objeto</span><span class="sxs-lookup"><span data-stu-id="96367-598">Object</span></span>| <span data-ttu-id="96367-599">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="96367-599">&lt;optional&gt;</span></span>|<span data-ttu-id="96367-600">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="96367-600">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96367-601">Requisitos</span><span class="sxs-lookup"><span data-stu-id="96367-601">Requirements</span></span>

|<span data-ttu-id="96367-602">Requirement</span><span class="sxs-lookup"><span data-stu-id="96367-602">Requirement</span></span>| <span data-ttu-id="96367-603">Valor</span><span class="sxs-lookup"><span data-stu-id="96367-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="96367-604">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="96367-604">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96367-605">1.0</span><span class="sxs-lookup"><span data-stu-id="96367-605">1.0</span></span>|
|[<span data-ttu-id="96367-606">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="96367-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96367-607">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="96367-607">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="96367-608">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="96367-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96367-609">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="96367-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96367-610">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="96367-610">Example</span></span>

<span data-ttu-id="96367-611">En el siguiente ejemplo, se llama a `makeEwsRequestAsync` para usar la operación `GetItem` para obtener el asunto de un elemento.</span><span class="sxs-lookup"><span data-stu-id="96367-611">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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