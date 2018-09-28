 

# <a name="office"></a><span data-ttu-id="21827-101">Office</span><span class="sxs-lookup"><span data-stu-id="21827-101">Office</span></span>

<span data-ttu-id="21827-p101">El espacio de nombres de Office proporciona interfaces compartidas que los complementos usan en todas las aplicaciones de Office. Este listado documenta solo aquellas interfaces que usan los complementos de Outlook. Para obtener un listado completo del espacio de nombres de Office, vea [API compartida](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="21827-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="21827-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="21827-104">Requirements</span></span>

|<span data-ttu-id="21827-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="21827-105">Requirement</span></span>| <span data-ttu-id="21827-106">Valor</span><span class="sxs-lookup"><span data-stu-id="21827-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="21827-107">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="21827-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="21827-108">1.0</span><span class="sxs-lookup"><span data-stu-id="21827-108">1.0</span></span>|
|[<span data-ttu-id="21827-109">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="21827-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="21827-110">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="21827-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="21827-111">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="21827-111">Members and methods</span></span>

| <span data-ttu-id="21827-112">Miembro	</span><span class="sxs-lookup"><span data-stu-id="21827-112">Member</span></span> | <span data-ttu-id="21827-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="21827-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="21827-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="21827-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="21827-115">Miembro	</span><span class="sxs-lookup"><span data-stu-id="21827-115">Member</span></span> |
| [<span data-ttu-id="21827-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="21827-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="21827-117">Miembro	</span><span class="sxs-lookup"><span data-stu-id="21827-117">Member</span></span> |
| [<span data-ttu-id="21827-118">EventType</span><span class="sxs-lookup"><span data-stu-id="21827-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="21827-119">Miembro	</span><span class="sxs-lookup"><span data-stu-id="21827-119">Member</span></span> |
| [<span data-ttu-id="21827-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="21827-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="21827-121">Miembro	</span><span class="sxs-lookup"><span data-stu-id="21827-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="21827-122">Espacios de nombres</span><span class="sxs-lookup"><span data-stu-id="21827-122">Namespaces</span></span>

<span data-ttu-id="21827-123">[contexto](office.context.md): proporciona interfaces compartidas del espacio de nombres de contexto de la API de complementos de Office para su uso en la API de complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="21827-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="21827-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Incluye las enumeraciones ItemType, EntityType, AttachmentType, RecipientType, ResponseType y ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="21827-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="21827-125">Miembros</span><span class="sxs-lookup"><span data-stu-id="21827-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="21827-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="21827-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="21827-127">Especifica el resultado de una llamada asíncrona.</span><span class="sxs-lookup"><span data-stu-id="21827-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="21827-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="21827-128">Type:</span></span>

*   <span data-ttu-id="21827-129">String</span><span class="sxs-lookup"><span data-stu-id="21827-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="21827-130">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="21827-130">Properties:</span></span>

|<span data-ttu-id="21827-131">Nombre</span><span class="sxs-lookup"><span data-stu-id="21827-131">Name</span></span>| <span data-ttu-id="21827-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="21827-132">Type</span></span>| <span data-ttu-id="21827-133">Descripción</span><span class="sxs-lookup"><span data-stu-id="21827-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="21827-134">String</span><span class="sxs-lookup"><span data-stu-id="21827-134">String</span></span>|<span data-ttu-id="21827-135">La llamada ha sido correcta.</span><span class="sxs-lookup"><span data-stu-id="21827-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="21827-136">String</span><span class="sxs-lookup"><span data-stu-id="21827-136">String</span></span>|<span data-ttu-id="21827-137">La llamada ha fallado.</span><span class="sxs-lookup"><span data-stu-id="21827-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21827-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="21827-138">Requirements</span></span>

|<span data-ttu-id="21827-139">Requirement</span><span class="sxs-lookup"><span data-stu-id="21827-139">Requirement</span></span>| <span data-ttu-id="21827-140">Valor</span><span class="sxs-lookup"><span data-stu-id="21827-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="21827-141">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="21827-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="21827-142">1.0</span><span class="sxs-lookup"><span data-stu-id="21827-142">1.0</span></span>|
|[<span data-ttu-id="21827-143">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="21827-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="21827-144">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="21827-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="21827-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="21827-145">CoercionType :String</span></span>

<span data-ttu-id="21827-146">Especifica cómo convertir los datos que el método invocado ha devuelto o definido.</span><span class="sxs-lookup"><span data-stu-id="21827-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="21827-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="21827-147">Type:</span></span>

*   <span data-ttu-id="21827-148">String</span><span class="sxs-lookup"><span data-stu-id="21827-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="21827-149">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="21827-149">Properties:</span></span>

|<span data-ttu-id="21827-150">Nombre</span><span class="sxs-lookup"><span data-stu-id="21827-150">Name</span></span>| <span data-ttu-id="21827-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="21827-151">Type</span></span>| <span data-ttu-id="21827-152">Descripción</span><span class="sxs-lookup"><span data-stu-id="21827-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="21827-153">String</span><span class="sxs-lookup"><span data-stu-id="21827-153">String</span></span>|<span data-ttu-id="21827-154">Solicita que los datos se devuelvan en formato HTML.</span><span class="sxs-lookup"><span data-stu-id="21827-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="21827-155">String</span><span class="sxs-lookup"><span data-stu-id="21827-155">String</span></span>|<span data-ttu-id="21827-156">Solicita que los datos se devuelvan en formato de texto.</span><span class="sxs-lookup"><span data-stu-id="21827-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21827-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="21827-157">Requirements</span></span>

|<span data-ttu-id="21827-158">Requirement</span><span class="sxs-lookup"><span data-stu-id="21827-158">Requirement</span></span>| <span data-ttu-id="21827-159">Valor</span><span class="sxs-lookup"><span data-stu-id="21827-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="21827-160">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="21827-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="21827-161">1.0</span><span class="sxs-lookup"><span data-stu-id="21827-161">1.0</span></span>|
|[<span data-ttu-id="21827-162">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="21827-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="21827-163">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="21827-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="21827-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="21827-164">EventType :String</span></span>

<span data-ttu-id="21827-165">Especifica el evento asociado con un controlador de eventos.</span><span class="sxs-lookup"><span data-stu-id="21827-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="21827-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="21827-166">Type:</span></span>

*   <span data-ttu-id="21827-167">String</span><span class="sxs-lookup"><span data-stu-id="21827-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="21827-168">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="21827-168">Properties:</span></span>

| <span data-ttu-id="21827-169">Nombre</span><span class="sxs-lookup"><span data-stu-id="21827-169">Name</span></span> | <span data-ttu-id="21827-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="21827-170">Type</span></span> | <span data-ttu-id="21827-171">Descripción</span><span class="sxs-lookup"><span data-stu-id="21827-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="21827-172">String</span><span class="sxs-lookup"><span data-stu-id="21827-172">String</span></span> | <span data-ttu-id="21827-173">El elemento seleccionado ha cambiado.</span><span class="sxs-lookup"><span data-stu-id="21827-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="21827-174">Requisitos</span><span class="sxs-lookup"><span data-stu-id="21827-174">Requirements</span></span>

|<span data-ttu-id="21827-175">Requirement</span><span class="sxs-lookup"><span data-stu-id="21827-175">Requirement</span></span>| <span data-ttu-id="21827-176">Valor</span><span class="sxs-lookup"><span data-stu-id="21827-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="21827-177">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="21827-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="21827-178">1,5</span><span class="sxs-lookup"><span data-stu-id="21827-178">1.5</span></span> |
|[<span data-ttu-id="21827-179">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="21827-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="21827-180">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="21827-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="21827-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="21827-181">SourceProperty :String</span></span>

<span data-ttu-id="21827-182">Especifica el origen de los datos devueltos por el método invocado.</span><span class="sxs-lookup"><span data-stu-id="21827-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="21827-183">Tipo:</span><span class="sxs-lookup"><span data-stu-id="21827-183">Type:</span></span>

*   <span data-ttu-id="21827-184">String</span><span class="sxs-lookup"><span data-stu-id="21827-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="21827-185">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="21827-185">Properties:</span></span>

|<span data-ttu-id="21827-186">Nombre</span><span class="sxs-lookup"><span data-stu-id="21827-186">Name</span></span>| <span data-ttu-id="21827-187">Tipo</span><span class="sxs-lookup"><span data-stu-id="21827-187">Type</span></span>| <span data-ttu-id="21827-188">Descripción</span><span class="sxs-lookup"><span data-stu-id="21827-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="21827-189">String</span><span class="sxs-lookup"><span data-stu-id="21827-189">String</span></span>|<span data-ttu-id="21827-190">El origen de los datos proviene del cuerpo de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="21827-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="21827-191">String</span><span class="sxs-lookup"><span data-stu-id="21827-191">String</span></span>|<span data-ttu-id="21827-192">El origen de los datos proviene del asunto de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="21827-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="21827-193">Requisitos</span><span class="sxs-lookup"><span data-stu-id="21827-193">Requirements</span></span>

|<span data-ttu-id="21827-194">Requirement</span><span class="sxs-lookup"><span data-stu-id="21827-194">Requirement</span></span>| <span data-ttu-id="21827-195">Valor</span><span class="sxs-lookup"><span data-stu-id="21827-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="21827-196">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="21827-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="21827-197">1.0</span><span class="sxs-lookup"><span data-stu-id="21827-197">1.0</span></span>|
|[<span data-ttu-id="21827-198">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="21827-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="21827-199">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="21827-199">Compose or read</span></span>|