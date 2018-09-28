 

# <a name="office"></a><span data-ttu-id="ba42c-101">Office</span><span class="sxs-lookup"><span data-stu-id="ba42c-101">Office</span></span>

<span data-ttu-id="ba42c-p101">El espacio de nombres de Office proporciona interfaces compartidas que los complementos usan en todas las aplicaciones de Office. Este listado documenta solo aquellas interfaces que usan los complementos de Outlook. Para obtener un listado completo del espacio de nombres de Office, vea [API compartida](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="ba42c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ba42c-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ba42c-104">Requirements</span></span>

|<span data-ttu-id="ba42c-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="ba42c-105">Requirement</span></span>| <span data-ttu-id="ba42c-106">Valor</span><span class="sxs-lookup"><span data-stu-id="ba42c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ba42c-107">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ba42c-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ba42c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ba42c-108">1.0</span></span>|
|[<span data-ttu-id="ba42c-109">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ba42c-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ba42c-110">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ba42c-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ba42c-111">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="ba42c-111">Members and methods</span></span>

| <span data-ttu-id="ba42c-112">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ba42c-112">Member</span></span> | <span data-ttu-id="ba42c-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="ba42c-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ba42c-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="ba42c-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="ba42c-115">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ba42c-115">Member</span></span> |
| [<span data-ttu-id="ba42c-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="ba42c-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="ba42c-117">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ba42c-117">Member</span></span> |
| [<span data-ttu-id="ba42c-118">EventType</span><span class="sxs-lookup"><span data-stu-id="ba42c-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="ba42c-119">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ba42c-119">Member</span></span> |
| [<span data-ttu-id="ba42c-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="ba42c-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="ba42c-121">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ba42c-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="ba42c-122">Espacios de nombres</span><span class="sxs-lookup"><span data-stu-id="ba42c-122">Namespaces</span></span>

<span data-ttu-id="ba42c-123">[contexto](office.context.md): proporciona interfaces compartidas del espacio de nombres de contexto de la API de complementos de Office para su uso en la API de complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="ba42c-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="ba42c-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Incluye las enumeraciones ItemType, EntityType, AttachmentType, RecipientType, ResponseType y ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="ba42c-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="ba42c-125">Miembros</span><span class="sxs-lookup"><span data-stu-id="ba42c-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="ba42c-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="ba42c-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="ba42c-127">Especifica el resultado de una llamada asíncrona.</span><span class="sxs-lookup"><span data-stu-id="ba42c-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ba42c-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ba42c-128">Type:</span></span>

*   <span data-ttu-id="ba42c-129">String</span><span class="sxs-lookup"><span data-stu-id="ba42c-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ba42c-130">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="ba42c-130">Properties:</span></span>

|<span data-ttu-id="ba42c-131">Nombre</span><span class="sxs-lookup"><span data-stu-id="ba42c-131">Name</span></span>| <span data-ttu-id="ba42c-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="ba42c-132">Type</span></span>| <span data-ttu-id="ba42c-133">Descripción</span><span class="sxs-lookup"><span data-stu-id="ba42c-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ba42c-134">String</span><span class="sxs-lookup"><span data-stu-id="ba42c-134">String</span></span>|<span data-ttu-id="ba42c-135">La llamada ha sido correcta.</span><span class="sxs-lookup"><span data-stu-id="ba42c-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ba42c-136">String</span><span class="sxs-lookup"><span data-stu-id="ba42c-136">String</span></span>|<span data-ttu-id="ba42c-137">La llamada ha fallado.</span><span class="sxs-lookup"><span data-stu-id="ba42c-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ba42c-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ba42c-138">Requirements</span></span>

|<span data-ttu-id="ba42c-139">Requirement</span><span class="sxs-lookup"><span data-stu-id="ba42c-139">Requirement</span></span>| <span data-ttu-id="ba42c-140">Valor</span><span class="sxs-lookup"><span data-stu-id="ba42c-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="ba42c-141">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ba42c-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ba42c-142">1.0</span><span class="sxs-lookup"><span data-stu-id="ba42c-142">1.0</span></span>|
|[<span data-ttu-id="ba42c-143">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ba42c-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ba42c-144">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ba42c-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="ba42c-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="ba42c-145">CoercionType :String</span></span>

<span data-ttu-id="ba42c-146">Especifica cómo convertir los datos que el método invocado ha devuelto o definido.</span><span class="sxs-lookup"><span data-stu-id="ba42c-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ba42c-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ba42c-147">Type:</span></span>

*   <span data-ttu-id="ba42c-148">String</span><span class="sxs-lookup"><span data-stu-id="ba42c-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ba42c-149">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="ba42c-149">Properties:</span></span>

|<span data-ttu-id="ba42c-150">Nombre</span><span class="sxs-lookup"><span data-stu-id="ba42c-150">Name</span></span>| <span data-ttu-id="ba42c-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="ba42c-151">Type</span></span>| <span data-ttu-id="ba42c-152">Descripción</span><span class="sxs-lookup"><span data-stu-id="ba42c-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ba42c-153">String</span><span class="sxs-lookup"><span data-stu-id="ba42c-153">String</span></span>|<span data-ttu-id="ba42c-154">Solicita que los datos se devuelvan en formato HTML.</span><span class="sxs-lookup"><span data-stu-id="ba42c-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ba42c-155">String</span><span class="sxs-lookup"><span data-stu-id="ba42c-155">String</span></span>|<span data-ttu-id="ba42c-156">Solicita que los datos se devuelvan en formato de texto.</span><span class="sxs-lookup"><span data-stu-id="ba42c-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ba42c-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ba42c-157">Requirements</span></span>

|<span data-ttu-id="ba42c-158">Requirement</span><span class="sxs-lookup"><span data-stu-id="ba42c-158">Requirement</span></span>| <span data-ttu-id="ba42c-159">Valor</span><span class="sxs-lookup"><span data-stu-id="ba42c-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="ba42c-160">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ba42c-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ba42c-161">1.0</span><span class="sxs-lookup"><span data-stu-id="ba42c-161">1.0</span></span>|
|[<span data-ttu-id="ba42c-162">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ba42c-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ba42c-163">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ba42c-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="ba42c-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="ba42c-164">EventType :String</span></span>

<span data-ttu-id="ba42c-165">Especifica el evento asociado con un controlador de eventos.</span><span class="sxs-lookup"><span data-stu-id="ba42c-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="ba42c-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ba42c-166">Type:</span></span>

*   <span data-ttu-id="ba42c-167">String</span><span class="sxs-lookup"><span data-stu-id="ba42c-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ba42c-168">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="ba42c-168">Properties:</span></span>

| <span data-ttu-id="ba42c-169">Nombre</span><span class="sxs-lookup"><span data-stu-id="ba42c-169">Name</span></span> | <span data-ttu-id="ba42c-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="ba42c-170">Type</span></span> | <span data-ttu-id="ba42c-171">Descripción</span><span class="sxs-lookup"><span data-stu-id="ba42c-171">Description</span></span> | <span data-ttu-id="ba42c-172">Conjunto de requisito mínimo</span><span class="sxs-lookup"><span data-stu-id="ba42c-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="ba42c-173">String</span><span class="sxs-lookup"><span data-stu-id="ba42c-173">String</span></span> | <span data-ttu-id="ba42c-174">Ha cambiado la fecha o la hora de la serie o la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="ba42c-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="ba42c-175">1.7</span><span class="sxs-lookup"><span data-stu-id="ba42c-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="ba42c-176">String</span><span class="sxs-lookup"><span data-stu-id="ba42c-176">String</span></span> | <span data-ttu-id="ba42c-177">El elemento seleccionado ha cambiado.</span><span class="sxs-lookup"><span data-stu-id="ba42c-177">The selected item has changed.</span></span> | <span data-ttu-id="ba42c-178">1,5</span><span class="sxs-lookup"><span data-stu-id="ba42c-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="ba42c-179">String</span><span class="sxs-lookup"><span data-stu-id="ba42c-179">String</span></span> | <span data-ttu-id="ba42c-180">Ha cambiado la lista de destinatarios de la ubicación del elemento o una cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="ba42c-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="ba42c-181">1.7</span><span class="sxs-lookup"><span data-stu-id="ba42c-181">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="ba42c-182">String</span><span class="sxs-lookup"><span data-stu-id="ba42c-182">String</span></span> | <span data-ttu-id="ba42c-183">Ha cambiado el patrón de periodicidad de la serie seleccionada.</span><span class="sxs-lookup"><span data-stu-id="ba42c-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="ba42c-184">1.7</span><span class="sxs-lookup"><span data-stu-id="ba42c-184">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ba42c-185">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ba42c-185">Requirements</span></span>

|<span data-ttu-id="ba42c-186">Requirement</span><span class="sxs-lookup"><span data-stu-id="ba42c-186">Requirement</span></span>| <span data-ttu-id="ba42c-187">Valor</span><span class="sxs-lookup"><span data-stu-id="ba42c-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="ba42c-188">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ba42c-188">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ba42c-189">1,5</span><span class="sxs-lookup"><span data-stu-id="ba42c-189">1.5</span></span> |
|[<span data-ttu-id="ba42c-190">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ba42c-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ba42c-191">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ba42c-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="ba42c-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="ba42c-192">SourceProperty :String</span></span>

<span data-ttu-id="ba42c-193">Especifica el origen de los datos devueltos por el método invocado.</span><span class="sxs-lookup"><span data-stu-id="ba42c-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ba42c-194">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ba42c-194">Type:</span></span>

*   <span data-ttu-id="ba42c-195">String</span><span class="sxs-lookup"><span data-stu-id="ba42c-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ba42c-196">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="ba42c-196">Properties:</span></span>

|<span data-ttu-id="ba42c-197">Nombre</span><span class="sxs-lookup"><span data-stu-id="ba42c-197">Name</span></span>| <span data-ttu-id="ba42c-198">Tipo</span><span class="sxs-lookup"><span data-stu-id="ba42c-198">Type</span></span>| <span data-ttu-id="ba42c-199">Descripción</span><span class="sxs-lookup"><span data-stu-id="ba42c-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ba42c-200">String</span><span class="sxs-lookup"><span data-stu-id="ba42c-200">String</span></span>|<span data-ttu-id="ba42c-201">El origen de los datos proviene del cuerpo de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="ba42c-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ba42c-202">String</span><span class="sxs-lookup"><span data-stu-id="ba42c-202">String</span></span>|<span data-ttu-id="ba42c-203">El origen de los datos proviene del asunto de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="ba42c-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ba42c-204">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ba42c-204">Requirements</span></span>

|<span data-ttu-id="ba42c-205">Requirement</span><span class="sxs-lookup"><span data-stu-id="ba42c-205">Requirement</span></span>| <span data-ttu-id="ba42c-206">Valor</span><span class="sxs-lookup"><span data-stu-id="ba42c-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="ba42c-207">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ba42c-207">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ba42c-208">1.0</span><span class="sxs-lookup"><span data-stu-id="ba42c-208">1.0</span></span>|
|[<span data-ttu-id="ba42c-209">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ba42c-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ba42c-210">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ba42c-210">Compose or read</span></span>|