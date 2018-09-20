 

# <a name="office"></a><span data-ttu-id="143f4-101">Office</span><span class="sxs-lookup"><span data-stu-id="143f4-101">Office</span></span>

<span data-ttu-id="143f4-p101">El espacio de nombres de Office proporciona interfaces compartidas que los complementos usan en todas las aplicaciones de Office. Este listado documenta solo aquellas interfaces que usan los complementos de Outlook. Para obtener un listado completo del espacio de nombres de Office, vea [API compartida](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="143f4-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="143f4-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="143f4-104">Requirements</span></span>

|<span data-ttu-id="143f4-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="143f4-105">Requirement</span></span>| <span data-ttu-id="143f4-106">Valor</span><span class="sxs-lookup"><span data-stu-id="143f4-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="143f4-107">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="143f4-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="143f4-108">1.0</span><span class="sxs-lookup"><span data-stu-id="143f4-108">1.0</span></span>|
|[<span data-ttu-id="143f4-109">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="143f4-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="143f4-110">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="143f4-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="143f4-111">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="143f4-111">Members and methods</span></span>

| <span data-ttu-id="143f4-112">Miembro	</span><span class="sxs-lookup"><span data-stu-id="143f4-112">Member</span></span> | <span data-ttu-id="143f4-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="143f4-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="143f4-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="143f4-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="143f4-115">Miembro	</span><span class="sxs-lookup"><span data-stu-id="143f4-115">Member</span></span> |
| [<span data-ttu-id="143f4-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="143f4-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="143f4-117">Miembro	</span><span class="sxs-lookup"><span data-stu-id="143f4-117">Member</span></span> |
| [<span data-ttu-id="143f4-118">EventType</span><span class="sxs-lookup"><span data-stu-id="143f4-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="143f4-119">Miembro	</span><span class="sxs-lookup"><span data-stu-id="143f4-119">Member</span></span> |
| [<span data-ttu-id="143f4-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="143f4-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="143f4-121">Miembro	</span><span class="sxs-lookup"><span data-stu-id="143f4-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="143f4-122">Espacios de nombres</span><span class="sxs-lookup"><span data-stu-id="143f4-122">Namespaces</span></span>

<span data-ttu-id="143f4-123">[contexto](office.context.md): proporciona interfaces compartidas del espacio de nombres de contexto de la API de complementos de Office para su uso en la API de complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="143f4-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="143f4-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Incluye las enumeraciones ItemType, EntityType, AttachmentType, RecipientType, ResponseType y ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="143f4-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="143f4-125">Miembros</span><span class="sxs-lookup"><span data-stu-id="143f4-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="143f4-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="143f4-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="143f4-127">Especifica el resultado de una llamada asíncrona.</span><span class="sxs-lookup"><span data-stu-id="143f4-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="143f4-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="143f4-128">Type:</span></span>

*   <span data-ttu-id="143f4-129">String</span><span class="sxs-lookup"><span data-stu-id="143f4-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="143f4-130">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="143f4-130">Properties:</span></span>

|<span data-ttu-id="143f4-131">Nombre</span><span class="sxs-lookup"><span data-stu-id="143f4-131">Name</span></span>| <span data-ttu-id="143f4-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="143f4-132">Type</span></span>| <span data-ttu-id="143f4-133">Descripción</span><span class="sxs-lookup"><span data-stu-id="143f4-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="143f4-134">String</span><span class="sxs-lookup"><span data-stu-id="143f4-134">String</span></span>|<span data-ttu-id="143f4-135">La llamada ha sido correcta.</span><span class="sxs-lookup"><span data-stu-id="143f4-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="143f4-136">String</span><span class="sxs-lookup"><span data-stu-id="143f4-136">String</span></span>|<span data-ttu-id="143f4-137">La llamada ha fallado.</span><span class="sxs-lookup"><span data-stu-id="143f4-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="143f4-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="143f4-138">Requirements</span></span>

|<span data-ttu-id="143f4-139">Requirement</span><span class="sxs-lookup"><span data-stu-id="143f4-139">Requirement</span></span>| <span data-ttu-id="143f4-140">Valor</span><span class="sxs-lookup"><span data-stu-id="143f4-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="143f4-141">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="143f4-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="143f4-142">1.0</span><span class="sxs-lookup"><span data-stu-id="143f4-142">1.0</span></span>|
|[<span data-ttu-id="143f4-143">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="143f4-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="143f4-144">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="143f4-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="143f4-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="143f4-145">CoercionType :String</span></span>

<span data-ttu-id="143f4-146">Especifica cómo convertir los datos que el método invocado ha devuelto o definido.</span><span class="sxs-lookup"><span data-stu-id="143f4-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="143f4-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="143f4-147">Type:</span></span>

*   <span data-ttu-id="143f4-148">String</span><span class="sxs-lookup"><span data-stu-id="143f4-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="143f4-149">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="143f4-149">Properties:</span></span>

|<span data-ttu-id="143f4-150">Nombre</span><span class="sxs-lookup"><span data-stu-id="143f4-150">Name</span></span>| <span data-ttu-id="143f4-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="143f4-151">Type</span></span>| <span data-ttu-id="143f4-152">Descripción</span><span class="sxs-lookup"><span data-stu-id="143f4-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="143f4-153">String</span><span class="sxs-lookup"><span data-stu-id="143f4-153">String</span></span>|<span data-ttu-id="143f4-154">Solicita que los datos se devuelvan en formato HTML.</span><span class="sxs-lookup"><span data-stu-id="143f4-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="143f4-155">String</span><span class="sxs-lookup"><span data-stu-id="143f4-155">String</span></span>|<span data-ttu-id="143f4-156">Solicita que los datos se devuelvan en formato de texto.</span><span class="sxs-lookup"><span data-stu-id="143f4-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="143f4-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="143f4-157">Requirements</span></span>

|<span data-ttu-id="143f4-158">Requirement</span><span class="sxs-lookup"><span data-stu-id="143f4-158">Requirement</span></span>| <span data-ttu-id="143f4-159">Valor</span><span class="sxs-lookup"><span data-stu-id="143f4-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="143f4-160">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="143f4-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="143f4-161">1.0</span><span class="sxs-lookup"><span data-stu-id="143f4-161">1.0</span></span>|
|[<span data-ttu-id="143f4-162">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="143f4-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="143f4-163">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="143f4-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="143f4-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="143f4-164">EventType :String</span></span>

<span data-ttu-id="143f4-165">Especifica el evento asociado con un controlador de eventos.</span><span class="sxs-lookup"><span data-stu-id="143f4-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="143f4-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="143f4-166">Type:</span></span>

*   <span data-ttu-id="143f4-167">String</span><span class="sxs-lookup"><span data-stu-id="143f4-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="143f4-168">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="143f4-168">Properties:</span></span>

| <span data-ttu-id="143f4-169">Nombre</span><span class="sxs-lookup"><span data-stu-id="143f4-169">Name</span></span> | <span data-ttu-id="143f4-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="143f4-170">Type</span></span> | <span data-ttu-id="143f4-171">Descripción</span><span class="sxs-lookup"><span data-stu-id="143f4-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="143f4-172">String</span><span class="sxs-lookup"><span data-stu-id="143f4-172">String</span></span> | <span data-ttu-id="143f4-173">El elemento seleccionado ha cambiado.</span><span class="sxs-lookup"><span data-stu-id="143f4-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="143f4-174">Requisitos</span><span class="sxs-lookup"><span data-stu-id="143f4-174">Requirements</span></span>

|<span data-ttu-id="143f4-175">Requirement</span><span class="sxs-lookup"><span data-stu-id="143f4-175">Requirement</span></span>| <span data-ttu-id="143f4-176">Valor</span><span class="sxs-lookup"><span data-stu-id="143f4-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="143f4-177">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="143f4-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="143f4-178">1,5</span><span class="sxs-lookup"><span data-stu-id="143f4-178">1.5</span></span> |
|[<span data-ttu-id="143f4-179">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="143f4-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="143f4-180">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="143f4-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="143f4-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="143f4-181">SourceProperty :String</span></span>

<span data-ttu-id="143f4-182">Especifica el origen de los datos devueltos por el método invocado.</span><span class="sxs-lookup"><span data-stu-id="143f4-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="143f4-183">Tipo:</span><span class="sxs-lookup"><span data-stu-id="143f4-183">Type:</span></span>

*   <span data-ttu-id="143f4-184">String</span><span class="sxs-lookup"><span data-stu-id="143f4-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="143f4-185">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="143f4-185">Properties:</span></span>

|<span data-ttu-id="143f4-186">Nombre</span><span class="sxs-lookup"><span data-stu-id="143f4-186">Name</span></span>| <span data-ttu-id="143f4-187">Tipo</span><span class="sxs-lookup"><span data-stu-id="143f4-187">Type</span></span>| <span data-ttu-id="143f4-188">Descripción</span><span class="sxs-lookup"><span data-stu-id="143f4-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="143f4-189">String</span><span class="sxs-lookup"><span data-stu-id="143f4-189">String</span></span>|<span data-ttu-id="143f4-190">El origen de los datos proviene del cuerpo de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="143f4-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="143f4-191">String</span><span class="sxs-lookup"><span data-stu-id="143f4-191">String</span></span>|<span data-ttu-id="143f4-192">El origen de los datos proviene del asunto de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="143f4-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="143f4-193">Requisitos</span><span class="sxs-lookup"><span data-stu-id="143f4-193">Requirements</span></span>

|<span data-ttu-id="143f4-194">Requirement</span><span class="sxs-lookup"><span data-stu-id="143f4-194">Requirement</span></span>| <span data-ttu-id="143f4-195">Valor</span><span class="sxs-lookup"><span data-stu-id="143f4-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="143f4-196">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="143f4-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="143f4-197">1.0</span><span class="sxs-lookup"><span data-stu-id="143f4-197">1.0</span></span>|
|[<span data-ttu-id="143f4-198">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="143f4-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="143f4-199">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="143f4-199">Compose or read</span></span>|