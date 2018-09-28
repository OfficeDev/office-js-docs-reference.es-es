 

# <a name="office"></a><span data-ttu-id="dfd91-101">Office</span><span class="sxs-lookup"><span data-stu-id="dfd91-101">Office</span></span>

<span data-ttu-id="dfd91-p101">El espacio de nombres de Office proporciona interfaces compartidas que los complementos usan en todas las aplicaciones de Office. Este listado documenta solo aquellas interfaces que usan los complementos de Outlook. Para obtener un listado completo del espacio de nombres de Office, vea [API compartida](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="dfd91-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="dfd91-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfd91-104">Requirements</span></span>

|<span data-ttu-id="dfd91-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="dfd91-105">Requirement</span></span>| <span data-ttu-id="dfd91-106">Valor</span><span class="sxs-lookup"><span data-stu-id="dfd91-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfd91-107">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="dfd91-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dfd91-108">1.0</span><span class="sxs-lookup"><span data-stu-id="dfd91-108">1.0</span></span>|
|[<span data-ttu-id="dfd91-109">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="dfd91-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dfd91-110">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="dfd91-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="dfd91-111">Espacios de nombres</span><span class="sxs-lookup"><span data-stu-id="dfd91-111">Namespaces</span></span>

<span data-ttu-id="dfd91-112">[contexto](Office.context.md): proporciona interfaces compartidas del espacio de nombres de contexto de la API de complementos de Office para su uso en la API de complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="dfd91-112">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="dfd91-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Incluye las enumeraciones ItemType, EntityType, AttachmentType, RecipientType, ResponseType y ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="dfd91-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="dfd91-114">Miembros</span><span class="sxs-lookup"><span data-stu-id="dfd91-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="dfd91-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="dfd91-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="dfd91-116">Especifica el resultado de una llamada asíncrona.</span><span class="sxs-lookup"><span data-stu-id="dfd91-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="dfd91-117">Tipo:</span><span class="sxs-lookup"><span data-stu-id="dfd91-117">Type:</span></span>

*   <span data-ttu-id="dfd91-118">String</span><span class="sxs-lookup"><span data-stu-id="dfd91-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dfd91-119">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="dfd91-119">Properties:</span></span>

|<span data-ttu-id="dfd91-120">Nombre</span><span class="sxs-lookup"><span data-stu-id="dfd91-120">Name</span></span>| <span data-ttu-id="dfd91-121">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfd91-121">Type</span></span>| <span data-ttu-id="dfd91-122">Descripción</span><span class="sxs-lookup"><span data-stu-id="dfd91-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="dfd91-123">String</span><span class="sxs-lookup"><span data-stu-id="dfd91-123">String</span></span>|<span data-ttu-id="dfd91-124">La llamada ha sido correcta.</span><span class="sxs-lookup"><span data-stu-id="dfd91-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="dfd91-125">String</span><span class="sxs-lookup"><span data-stu-id="dfd91-125">String</span></span>|<span data-ttu-id="dfd91-126">La llamada ha fallado.</span><span class="sxs-lookup"><span data-stu-id="dfd91-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dfd91-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfd91-127">Requirements</span></span>

|<span data-ttu-id="dfd91-128">Requirement</span><span class="sxs-lookup"><span data-stu-id="dfd91-128">Requirement</span></span>| <span data-ttu-id="dfd91-129">Valor</span><span class="sxs-lookup"><span data-stu-id="dfd91-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfd91-130">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="dfd91-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dfd91-131">1.0</span><span class="sxs-lookup"><span data-stu-id="dfd91-131">1.0</span></span>|
|[<span data-ttu-id="dfd91-132">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="dfd91-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dfd91-133">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="dfd91-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="dfd91-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="dfd91-134">CoercionType :String</span></span>

<span data-ttu-id="dfd91-135">Especifica cómo convertir los datos que el método invocado ha devuelto o definido.</span><span class="sxs-lookup"><span data-stu-id="dfd91-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dfd91-136">Tipo:</span><span class="sxs-lookup"><span data-stu-id="dfd91-136">Type:</span></span>

*   <span data-ttu-id="dfd91-137">String</span><span class="sxs-lookup"><span data-stu-id="dfd91-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dfd91-138">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="dfd91-138">Properties:</span></span>

|<span data-ttu-id="dfd91-139">Nombre</span><span class="sxs-lookup"><span data-stu-id="dfd91-139">Name</span></span>| <span data-ttu-id="dfd91-140">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfd91-140">Type</span></span>| <span data-ttu-id="dfd91-141">Descripción</span><span class="sxs-lookup"><span data-stu-id="dfd91-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="dfd91-142">String</span><span class="sxs-lookup"><span data-stu-id="dfd91-142">String</span></span>|<span data-ttu-id="dfd91-143">Solicita que los datos se devuelvan en formato HTML.</span><span class="sxs-lookup"><span data-stu-id="dfd91-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="dfd91-144">String</span><span class="sxs-lookup"><span data-stu-id="dfd91-144">String</span></span>|<span data-ttu-id="dfd91-145">Solicita que los datos se devuelvan en formato de texto.</span><span class="sxs-lookup"><span data-stu-id="dfd91-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dfd91-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfd91-146">Requirements</span></span>

|<span data-ttu-id="dfd91-147">Requirement</span><span class="sxs-lookup"><span data-stu-id="dfd91-147">Requirement</span></span>| <span data-ttu-id="dfd91-148">Valor</span><span class="sxs-lookup"><span data-stu-id="dfd91-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfd91-149">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="dfd91-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dfd91-150">1.0</span><span class="sxs-lookup"><span data-stu-id="dfd91-150">1.0</span></span>|
|[<span data-ttu-id="dfd91-151">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="dfd91-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dfd91-152">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="dfd91-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="dfd91-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="dfd91-153">SourceProperty :String</span></span>

<span data-ttu-id="dfd91-154">Especifica el origen de los datos devueltos por el método invocado.</span><span class="sxs-lookup"><span data-stu-id="dfd91-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dfd91-155">Tipo:</span><span class="sxs-lookup"><span data-stu-id="dfd91-155">Type:</span></span>

*   <span data-ttu-id="dfd91-156">String</span><span class="sxs-lookup"><span data-stu-id="dfd91-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dfd91-157">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="dfd91-157">Properties:</span></span>

|<span data-ttu-id="dfd91-158">Nombre</span><span class="sxs-lookup"><span data-stu-id="dfd91-158">Name</span></span>| <span data-ttu-id="dfd91-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfd91-159">Type</span></span>| <span data-ttu-id="dfd91-160">Descripción</span><span class="sxs-lookup"><span data-stu-id="dfd91-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="dfd91-161">String</span><span class="sxs-lookup"><span data-stu-id="dfd91-161">String</span></span>|<span data-ttu-id="dfd91-162">El origen de los datos proviene del cuerpo de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="dfd91-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="dfd91-163">String</span><span class="sxs-lookup"><span data-stu-id="dfd91-163">String</span></span>|<span data-ttu-id="dfd91-164">El origen de los datos proviene del asunto de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="dfd91-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dfd91-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfd91-165">Requirements</span></span>

|<span data-ttu-id="dfd91-166">Requirement</span><span class="sxs-lookup"><span data-stu-id="dfd91-166">Requirement</span></span>| <span data-ttu-id="dfd91-167">Valor</span><span class="sxs-lookup"><span data-stu-id="dfd91-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfd91-168">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="dfd91-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dfd91-169">1.0</span><span class="sxs-lookup"><span data-stu-id="dfd91-169">1.0</span></span>|
|[<span data-ttu-id="dfd91-170">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="dfd91-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dfd91-171">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="dfd91-171">Compose or read</span></span>|