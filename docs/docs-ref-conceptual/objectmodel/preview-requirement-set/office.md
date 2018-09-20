 

# <a name="office"></a><span data-ttu-id="f9b31-101">Office</span><span class="sxs-lookup"><span data-stu-id="f9b31-101">Office</span></span>

<span data-ttu-id="f9b31-p101">El espacio de nombres de Office proporciona interfaces compartidas que los complementos usan en todas las aplicaciones de Office. Este listado documenta solo aquellas interfaces que usan los complementos de Outlook. Para obtener un listado completo del espacio de nombres de Office, vea [API compartida](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f9b31-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9b31-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9b31-104">Requirements</span></span>

|<span data-ttu-id="f9b31-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="f9b31-105">Requirement</span></span>| <span data-ttu-id="f9b31-106">Valor</span><span class="sxs-lookup"><span data-stu-id="f9b31-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9b31-107">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="f9b31-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9b31-108">1.0</span><span class="sxs-lookup"><span data-stu-id="f9b31-108">1.0</span></span>|
|[<span data-ttu-id="f9b31-109">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="f9b31-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9b31-110">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="f9b31-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f9b31-111">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="f9b31-111">Members and methods</span></span>

| <span data-ttu-id="f9b31-112">Miembro	</span><span class="sxs-lookup"><span data-stu-id="f9b31-112">Member</span></span> | <span data-ttu-id="f9b31-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9b31-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f9b31-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f9b31-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f9b31-115">Miembro	</span><span class="sxs-lookup"><span data-stu-id="f9b31-115">Member</span></span> |
| [<span data-ttu-id="f9b31-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f9b31-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f9b31-117">Miembro	</span><span class="sxs-lookup"><span data-stu-id="f9b31-117">Member</span></span> |
| [<span data-ttu-id="f9b31-118">EventType</span><span class="sxs-lookup"><span data-stu-id="f9b31-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f9b31-119">Miembro	</span><span class="sxs-lookup"><span data-stu-id="f9b31-119">Member</span></span> |
| [<span data-ttu-id="f9b31-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f9b31-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f9b31-121">Miembro	</span><span class="sxs-lookup"><span data-stu-id="f9b31-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="f9b31-122">Espacios de nombres</span><span class="sxs-lookup"><span data-stu-id="f9b31-122">Namespaces</span></span>

<span data-ttu-id="f9b31-123">[contexto](office.context.md): proporciona interfaces compartidas del espacio de nombres de contexto de la API de complementos de Office para su uso en la API de complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="f9b31-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f9b31-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Incluye las enumeraciones ItemType, EntityType, AttachmentType, RecipientType, ResponseType y ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="f9b31-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="f9b31-125">Miembros</span><span class="sxs-lookup"><span data-stu-id="f9b31-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="f9b31-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="f9b31-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="f9b31-127">Especifica el resultado de una llamada asíncrona.</span><span class="sxs-lookup"><span data-stu-id="f9b31-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f9b31-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9b31-128">Type:</span></span>

*   <span data-ttu-id="f9b31-129">String</span><span class="sxs-lookup"><span data-stu-id="f9b31-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9b31-130">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="f9b31-130">Properties:</span></span>

|<span data-ttu-id="f9b31-131">Nombre</span><span class="sxs-lookup"><span data-stu-id="f9b31-131">Name</span></span>| <span data-ttu-id="f9b31-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9b31-132">Type</span></span>| <span data-ttu-id="f9b31-133">Descripción</span><span class="sxs-lookup"><span data-stu-id="f9b31-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f9b31-134">String</span><span class="sxs-lookup"><span data-stu-id="f9b31-134">String</span></span>|<span data-ttu-id="f9b31-135">La llamada ha sido correcta.</span><span class="sxs-lookup"><span data-stu-id="f9b31-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f9b31-136">String</span><span class="sxs-lookup"><span data-stu-id="f9b31-136">String</span></span>|<span data-ttu-id="f9b31-137">La llamada ha fallado.</span><span class="sxs-lookup"><span data-stu-id="f9b31-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9b31-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9b31-138">Requirements</span></span>

|<span data-ttu-id="f9b31-139">Requirement</span><span class="sxs-lookup"><span data-stu-id="f9b31-139">Requirement</span></span>| <span data-ttu-id="f9b31-140">Valor</span><span class="sxs-lookup"><span data-stu-id="f9b31-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9b31-141">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="f9b31-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9b31-142">1.0</span><span class="sxs-lookup"><span data-stu-id="f9b31-142">1.0</span></span>|
|[<span data-ttu-id="f9b31-143">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="f9b31-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9b31-144">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="f9b31-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="f9b31-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="f9b31-145">CoercionType :String</span></span>

<span data-ttu-id="f9b31-146">Especifica cómo convertir los datos que el método invocado ha devuelto o definido.</span><span class="sxs-lookup"><span data-stu-id="f9b31-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f9b31-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9b31-147">Type:</span></span>

*   <span data-ttu-id="f9b31-148">String</span><span class="sxs-lookup"><span data-stu-id="f9b31-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9b31-149">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="f9b31-149">Properties:</span></span>

|<span data-ttu-id="f9b31-150">Nombre</span><span class="sxs-lookup"><span data-stu-id="f9b31-150">Name</span></span>| <span data-ttu-id="f9b31-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9b31-151">Type</span></span>| <span data-ttu-id="f9b31-152">Descripción</span><span class="sxs-lookup"><span data-stu-id="f9b31-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f9b31-153">String</span><span class="sxs-lookup"><span data-stu-id="f9b31-153">String</span></span>|<span data-ttu-id="f9b31-154">Solicita que los datos se devuelvan en formato HTML.</span><span class="sxs-lookup"><span data-stu-id="f9b31-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f9b31-155">String</span><span class="sxs-lookup"><span data-stu-id="f9b31-155">String</span></span>|<span data-ttu-id="f9b31-156">Solicita que los datos se devuelvan en formato de texto.</span><span class="sxs-lookup"><span data-stu-id="f9b31-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9b31-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9b31-157">Requirements</span></span>

|<span data-ttu-id="f9b31-158">Requirement</span><span class="sxs-lookup"><span data-stu-id="f9b31-158">Requirement</span></span>| <span data-ttu-id="f9b31-159">Valor</span><span class="sxs-lookup"><span data-stu-id="f9b31-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9b31-160">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="f9b31-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9b31-161">1.0</span><span class="sxs-lookup"><span data-stu-id="f9b31-161">1.0</span></span>|
|[<span data-ttu-id="f9b31-162">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="f9b31-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9b31-163">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="f9b31-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="f9b31-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="f9b31-164">EventType :String</span></span>

<span data-ttu-id="f9b31-165">Especifica el evento asociado con un controlador de eventos.</span><span class="sxs-lookup"><span data-stu-id="f9b31-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f9b31-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9b31-166">Type:</span></span>

*   <span data-ttu-id="f9b31-167">String</span><span class="sxs-lookup"><span data-stu-id="f9b31-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9b31-168">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="f9b31-168">Properties:</span></span>

| <span data-ttu-id="f9b31-169">Nombre</span><span class="sxs-lookup"><span data-stu-id="f9b31-169">Name</span></span> | <span data-ttu-id="f9b31-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9b31-170">Type</span></span> | <span data-ttu-id="f9b31-171">Descripción</span><span class="sxs-lookup"><span data-stu-id="f9b31-171">Description</span></span> | <span data-ttu-id="f9b31-172">Conjunto de requisito mínimo</span><span class="sxs-lookup"><span data-stu-id="f9b31-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="f9b31-173">String</span><span class="sxs-lookup"><span data-stu-id="f9b31-173">String</span></span> | <span data-ttu-id="f9b31-174">Ha cambiado la cita de fecha u hora de la serie seleccionada.</span><span class="sxs-lookup"><span data-stu-id="f9b31-174">The appointment date or time of the selected series has changed.</span></span> | <span data-ttu-id="f9b31-175">Preview</span><span class="sxs-lookup"><span data-stu-id="f9b31-175">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="f9b31-176">String</span><span class="sxs-lookup"><span data-stu-id="f9b31-176">String</span></span> | <span data-ttu-id="f9b31-177">El elemento seleccionado ha cambiado.</span><span class="sxs-lookup"><span data-stu-id="f9b31-177">The selected item has changed.</span></span> | <span data-ttu-id="f9b31-178">1,5</span><span class="sxs-lookup"><span data-stu-id="f9b31-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="f9b31-179">String</span><span class="sxs-lookup"><span data-stu-id="f9b31-179">String</span></span> | <span data-ttu-id="f9b31-180">El elemento seleccionado ha cambiado.</span><span class="sxs-lookup"><span data-stu-id="f9b31-180">The selected item has changed.</span></span> | <span data-ttu-id="f9b31-181">Preview</span><span class="sxs-lookup"><span data-stu-id="f9b31-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="f9b31-182">String</span><span class="sxs-lookup"><span data-stu-id="f9b31-182">String</span></span> | <span data-ttu-id="f9b31-183">Ha cambiado la lista de destinatarios del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="f9b31-183">The recipient list of the selected item has changed.</span></span> | <span data-ttu-id="f9b31-184">Preview</span><span class="sxs-lookup"><span data-stu-id="f9b31-184">Preview</span></span> |
|`RecurrencePatternChanged`| <span data-ttu-id="f9b31-185">String</span><span class="sxs-lookup"><span data-stu-id="f9b31-185">String</span></span> | <span data-ttu-id="f9b31-186">Ha cambiado el patrón de periodicidad de la serie seleccionada.</span><span class="sxs-lookup"><span data-stu-id="f9b31-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="f9b31-187">Preview</span><span class="sxs-lookup"><span data-stu-id="f9b31-187">Preview</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f9b31-188">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9b31-188">Requirements</span></span>

|<span data-ttu-id="f9b31-189">Requirement</span><span class="sxs-lookup"><span data-stu-id="f9b31-189">Requirement</span></span>| <span data-ttu-id="f9b31-190">Valor</span><span class="sxs-lookup"><span data-stu-id="f9b31-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9b31-191">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="f9b31-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9b31-192">1,5</span><span class="sxs-lookup"><span data-stu-id="f9b31-192">1.5</span></span> |
|[<span data-ttu-id="f9b31-193">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="f9b31-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9b31-194">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="f9b31-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="f9b31-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="f9b31-195">SourceProperty :String</span></span>

<span data-ttu-id="f9b31-196">Especifica el origen de los datos devueltos por el método invocado.</span><span class="sxs-lookup"><span data-stu-id="f9b31-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f9b31-197">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9b31-197">Type:</span></span>

*   <span data-ttu-id="f9b31-198">String</span><span class="sxs-lookup"><span data-stu-id="f9b31-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9b31-199">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="f9b31-199">Properties:</span></span>

|<span data-ttu-id="f9b31-200">Nombre</span><span class="sxs-lookup"><span data-stu-id="f9b31-200">Name</span></span>| <span data-ttu-id="f9b31-201">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9b31-201">Type</span></span>| <span data-ttu-id="f9b31-202">Descripción</span><span class="sxs-lookup"><span data-stu-id="f9b31-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f9b31-203">String</span><span class="sxs-lookup"><span data-stu-id="f9b31-203">String</span></span>|<span data-ttu-id="f9b31-204">El origen de los datos proviene del cuerpo de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="f9b31-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f9b31-205">String</span><span class="sxs-lookup"><span data-stu-id="f9b31-205">String</span></span>|<span data-ttu-id="f9b31-206">El origen de los datos proviene del asunto de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="f9b31-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9b31-207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9b31-207">Requirements</span></span>

|<span data-ttu-id="f9b31-208">Requirement</span><span class="sxs-lookup"><span data-stu-id="f9b31-208">Requirement</span></span>| <span data-ttu-id="f9b31-209">Valor</span><span class="sxs-lookup"><span data-stu-id="f9b31-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9b31-210">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="f9b31-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9b31-211">1.0</span><span class="sxs-lookup"><span data-stu-id="f9b31-211">1.0</span></span>|
|[<span data-ttu-id="f9b31-212">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="f9b31-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9b31-213">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="f9b31-213">Compose or read</span></span>|