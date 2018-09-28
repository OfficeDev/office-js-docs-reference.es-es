 

# <a name="office"></a><span data-ttu-id="66e51-101">Office</span><span class="sxs-lookup"><span data-stu-id="66e51-101">Office</span></span>

<span data-ttu-id="66e51-p101">El espacio de nombres de Office proporciona interfaces compartidas que los complementos usan en todas las aplicaciones de Office. Este listado documenta solo aquellas interfaces que usan los complementos de Outlook. Para obtener un listado completo del espacio de nombres de Office, vea [API compartida](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="66e51-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="66e51-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="66e51-104">Requirements</span></span>

|<span data-ttu-id="66e51-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="66e51-105">Requirement</span></span>| <span data-ttu-id="66e51-106">Valor</span><span class="sxs-lookup"><span data-stu-id="66e51-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="66e51-107">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="66e51-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66e51-108">1.0</span><span class="sxs-lookup"><span data-stu-id="66e51-108">1.0</span></span>|
|[<span data-ttu-id="66e51-109">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="66e51-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="66e51-110">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="66e51-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="66e51-111">Espacios de nombres</span><span class="sxs-lookup"><span data-stu-id="66e51-111">Namespaces</span></span>

<span data-ttu-id="66e51-112">[contexto](office.context.md): proporciona interfaces compartidas del espacio de nombres de contexto de la API de complementos de Office para su uso en la API de complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="66e51-112">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="66e51-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Incluye las enumeraciones ItemType, EntityType, AttachmentType, RecipientType, ResponseType y ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="66e51-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="66e51-114">Miembros</span><span class="sxs-lookup"><span data-stu-id="66e51-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="66e51-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="66e51-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="66e51-116">Especifica el resultado de una llamada asíncrona.</span><span class="sxs-lookup"><span data-stu-id="66e51-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="66e51-117">Tipo:</span><span class="sxs-lookup"><span data-stu-id="66e51-117">Type:</span></span>

*   <span data-ttu-id="66e51-118">String</span><span class="sxs-lookup"><span data-stu-id="66e51-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="66e51-119">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="66e51-119">Properties:</span></span>

|<span data-ttu-id="66e51-120">Nombre</span><span class="sxs-lookup"><span data-stu-id="66e51-120">Name</span></span>| <span data-ttu-id="66e51-121">Tipo</span><span class="sxs-lookup"><span data-stu-id="66e51-121">Type</span></span>| <span data-ttu-id="66e51-122">Descripción</span><span class="sxs-lookup"><span data-stu-id="66e51-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="66e51-123">String</span><span class="sxs-lookup"><span data-stu-id="66e51-123">String</span></span>|<span data-ttu-id="66e51-124">La llamada ha sido correcta.</span><span class="sxs-lookup"><span data-stu-id="66e51-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="66e51-125">String</span><span class="sxs-lookup"><span data-stu-id="66e51-125">String</span></span>|<span data-ttu-id="66e51-126">La llamada ha fallado.</span><span class="sxs-lookup"><span data-stu-id="66e51-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="66e51-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="66e51-127">Requirements</span></span>

|<span data-ttu-id="66e51-128">Requirement</span><span class="sxs-lookup"><span data-stu-id="66e51-128">Requirement</span></span>| <span data-ttu-id="66e51-129">Valor</span><span class="sxs-lookup"><span data-stu-id="66e51-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="66e51-130">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="66e51-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66e51-131">1.0</span><span class="sxs-lookup"><span data-stu-id="66e51-131">1.0</span></span>|
|[<span data-ttu-id="66e51-132">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="66e51-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="66e51-133">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="66e51-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="66e51-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="66e51-134">CoercionType :String</span></span>

<span data-ttu-id="66e51-135">Especifica cómo convertir los datos que el método invocado ha devuelto o definido.</span><span class="sxs-lookup"><span data-stu-id="66e51-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="66e51-136">Tipo:</span><span class="sxs-lookup"><span data-stu-id="66e51-136">Type:</span></span>

*   <span data-ttu-id="66e51-137">String</span><span class="sxs-lookup"><span data-stu-id="66e51-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="66e51-138">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="66e51-138">Properties:</span></span>

|<span data-ttu-id="66e51-139">Nombre</span><span class="sxs-lookup"><span data-stu-id="66e51-139">Name</span></span>| <span data-ttu-id="66e51-140">Tipo</span><span class="sxs-lookup"><span data-stu-id="66e51-140">Type</span></span>| <span data-ttu-id="66e51-141">Descripción</span><span class="sxs-lookup"><span data-stu-id="66e51-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="66e51-142">String</span><span class="sxs-lookup"><span data-stu-id="66e51-142">String</span></span>|<span data-ttu-id="66e51-143">Solicita que los datos se devuelvan en formato HTML.</span><span class="sxs-lookup"><span data-stu-id="66e51-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="66e51-144">String</span><span class="sxs-lookup"><span data-stu-id="66e51-144">String</span></span>|<span data-ttu-id="66e51-145">Solicita que los datos se devuelvan en formato de texto.</span><span class="sxs-lookup"><span data-stu-id="66e51-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="66e51-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="66e51-146">Requirements</span></span>

|<span data-ttu-id="66e51-147">Requirement</span><span class="sxs-lookup"><span data-stu-id="66e51-147">Requirement</span></span>| <span data-ttu-id="66e51-148">Valor</span><span class="sxs-lookup"><span data-stu-id="66e51-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="66e51-149">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="66e51-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66e51-150">1.0</span><span class="sxs-lookup"><span data-stu-id="66e51-150">1.0</span></span>|
|[<span data-ttu-id="66e51-151">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="66e51-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="66e51-152">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="66e51-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="66e51-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="66e51-153">SourceProperty :String</span></span>

<span data-ttu-id="66e51-154">Especifica el origen de los datos devueltos por el método invocado.</span><span class="sxs-lookup"><span data-stu-id="66e51-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="66e51-155">Tipo:</span><span class="sxs-lookup"><span data-stu-id="66e51-155">Type:</span></span>

*   <span data-ttu-id="66e51-156">String</span><span class="sxs-lookup"><span data-stu-id="66e51-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="66e51-157">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="66e51-157">Properties:</span></span>

|<span data-ttu-id="66e51-158">Nombre</span><span class="sxs-lookup"><span data-stu-id="66e51-158">Name</span></span>| <span data-ttu-id="66e51-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="66e51-159">Type</span></span>| <span data-ttu-id="66e51-160">Descripción</span><span class="sxs-lookup"><span data-stu-id="66e51-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="66e51-161">String</span><span class="sxs-lookup"><span data-stu-id="66e51-161">String</span></span>|<span data-ttu-id="66e51-162">El origen de los datos proviene del cuerpo de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="66e51-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="66e51-163">String</span><span class="sxs-lookup"><span data-stu-id="66e51-163">String</span></span>|<span data-ttu-id="66e51-164">El origen de los datos proviene del asunto de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="66e51-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="66e51-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="66e51-165">Requirements</span></span>

|<span data-ttu-id="66e51-166">Requirement</span><span class="sxs-lookup"><span data-stu-id="66e51-166">Requirement</span></span>| <span data-ttu-id="66e51-167">Valor</span><span class="sxs-lookup"><span data-stu-id="66e51-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="66e51-168">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="66e51-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66e51-169">1.0</span><span class="sxs-lookup"><span data-stu-id="66e51-169">1.0</span></span>|
|[<span data-ttu-id="66e51-170">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="66e51-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="66e51-171">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="66e51-171">Compose or read</span></span>|