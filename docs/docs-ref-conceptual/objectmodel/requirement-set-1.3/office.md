 

# <a name="office"></a><span data-ttu-id="674cd-101">Office</span><span class="sxs-lookup"><span data-stu-id="674cd-101">Office</span></span>

<span data-ttu-id="674cd-p101">El espacio de nombres de Office proporciona interfaces compartidas que los complementos usan en todas las aplicaciones de Office. Este listado documenta solo aquellas interfaces que usan los complementos de Outlook. Para obtener un listado completo del espacio de nombres de Office, vea [API compartida](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="674cd-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="674cd-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="674cd-104">Requirements</span></span>

|<span data-ttu-id="674cd-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="674cd-105">Requirement</span></span>| <span data-ttu-id="674cd-106">Valor</span><span class="sxs-lookup"><span data-stu-id="674cd-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="674cd-107">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="674cd-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="674cd-108">1.0</span><span class="sxs-lookup"><span data-stu-id="674cd-108">1.0</span></span>|
|[<span data-ttu-id="674cd-109">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="674cd-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="674cd-110">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="674cd-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="674cd-111">Espacios de nombres</span><span class="sxs-lookup"><span data-stu-id="674cd-111">Namespaces</span></span>

<span data-ttu-id="674cd-112">[contexto](office.context.md): proporciona interfaces compartidas del espacio de nombres de contexto de la API de complementos de Office para su uso en la API de complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="674cd-112">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="674cd-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Incluye las enumeraciones ItemType, EntityType, AttachmentType, RecipientType, ResponseType y ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="674cd-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="674cd-114">Miembros</span><span class="sxs-lookup"><span data-stu-id="674cd-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="674cd-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="674cd-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="674cd-116">Especifica el resultado de una llamada asíncrona.</span><span class="sxs-lookup"><span data-stu-id="674cd-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="674cd-117">Tipo:</span><span class="sxs-lookup"><span data-stu-id="674cd-117">Type:</span></span>

*   <span data-ttu-id="674cd-118">String</span><span class="sxs-lookup"><span data-stu-id="674cd-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="674cd-119">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="674cd-119">Properties:</span></span>

|<span data-ttu-id="674cd-120">Nombre</span><span class="sxs-lookup"><span data-stu-id="674cd-120">Name</span></span>| <span data-ttu-id="674cd-121">Tipo</span><span class="sxs-lookup"><span data-stu-id="674cd-121">Type</span></span>| <span data-ttu-id="674cd-122">Descripción</span><span class="sxs-lookup"><span data-stu-id="674cd-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="674cd-123">String</span><span class="sxs-lookup"><span data-stu-id="674cd-123">String</span></span>|<span data-ttu-id="674cd-124">La llamada ha sido correcta.</span><span class="sxs-lookup"><span data-stu-id="674cd-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="674cd-125">String</span><span class="sxs-lookup"><span data-stu-id="674cd-125">String</span></span>|<span data-ttu-id="674cd-126">La llamada ha fallado.</span><span class="sxs-lookup"><span data-stu-id="674cd-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="674cd-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="674cd-127">Requirements</span></span>

|<span data-ttu-id="674cd-128">Requirement</span><span class="sxs-lookup"><span data-stu-id="674cd-128">Requirement</span></span>| <span data-ttu-id="674cd-129">Valor</span><span class="sxs-lookup"><span data-stu-id="674cd-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="674cd-130">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="674cd-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="674cd-131">1.0</span><span class="sxs-lookup"><span data-stu-id="674cd-131">1.0</span></span>|
|[<span data-ttu-id="674cd-132">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="674cd-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="674cd-133">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="674cd-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="674cd-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="674cd-134">CoercionType :String</span></span>

<span data-ttu-id="674cd-135">Especifica cómo convertir los datos que el método invocado ha devuelto o definido.</span><span class="sxs-lookup"><span data-stu-id="674cd-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="674cd-136">Tipo:</span><span class="sxs-lookup"><span data-stu-id="674cd-136">Type:</span></span>

*   <span data-ttu-id="674cd-137">String</span><span class="sxs-lookup"><span data-stu-id="674cd-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="674cd-138">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="674cd-138">Properties:</span></span>

|<span data-ttu-id="674cd-139">Nombre</span><span class="sxs-lookup"><span data-stu-id="674cd-139">Name</span></span>| <span data-ttu-id="674cd-140">Tipo</span><span class="sxs-lookup"><span data-stu-id="674cd-140">Type</span></span>| <span data-ttu-id="674cd-141">Descripción</span><span class="sxs-lookup"><span data-stu-id="674cd-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="674cd-142">String</span><span class="sxs-lookup"><span data-stu-id="674cd-142">String</span></span>|<span data-ttu-id="674cd-143">Solicita que los datos se devuelvan en formato HTML.</span><span class="sxs-lookup"><span data-stu-id="674cd-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="674cd-144">String</span><span class="sxs-lookup"><span data-stu-id="674cd-144">String</span></span>|<span data-ttu-id="674cd-145">Solicita que los datos se devuelvan en formato de texto.</span><span class="sxs-lookup"><span data-stu-id="674cd-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="674cd-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="674cd-146">Requirements</span></span>

|<span data-ttu-id="674cd-147">Requirement</span><span class="sxs-lookup"><span data-stu-id="674cd-147">Requirement</span></span>| <span data-ttu-id="674cd-148">Valor</span><span class="sxs-lookup"><span data-stu-id="674cd-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="674cd-149">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="674cd-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="674cd-150">1.0</span><span class="sxs-lookup"><span data-stu-id="674cd-150">1.0</span></span>|
|[<span data-ttu-id="674cd-151">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="674cd-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="674cd-152">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="674cd-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="674cd-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="674cd-153">SourceProperty :String</span></span>

<span data-ttu-id="674cd-154">Especifica el origen de los datos devueltos por el método invocado.</span><span class="sxs-lookup"><span data-stu-id="674cd-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="674cd-155">Tipo:</span><span class="sxs-lookup"><span data-stu-id="674cd-155">Type:</span></span>

*   <span data-ttu-id="674cd-156">String</span><span class="sxs-lookup"><span data-stu-id="674cd-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="674cd-157">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="674cd-157">Properties:</span></span>

|<span data-ttu-id="674cd-158">Nombre</span><span class="sxs-lookup"><span data-stu-id="674cd-158">Name</span></span>| <span data-ttu-id="674cd-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="674cd-159">Type</span></span>| <span data-ttu-id="674cd-160">Descripción</span><span class="sxs-lookup"><span data-stu-id="674cd-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="674cd-161">String</span><span class="sxs-lookup"><span data-stu-id="674cd-161">String</span></span>|<span data-ttu-id="674cd-162">El origen de los datos proviene del cuerpo de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="674cd-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="674cd-163">String</span><span class="sxs-lookup"><span data-stu-id="674cd-163">String</span></span>|<span data-ttu-id="674cd-164">El origen de los datos proviene del asunto de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="674cd-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="674cd-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="674cd-165">Requirements</span></span>

|<span data-ttu-id="674cd-166">Requirement</span><span class="sxs-lookup"><span data-stu-id="674cd-166">Requirement</span></span>| <span data-ttu-id="674cd-167">Valor</span><span class="sxs-lookup"><span data-stu-id="674cd-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="674cd-168">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="674cd-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="674cd-169">1.0</span><span class="sxs-lookup"><span data-stu-id="674cd-169">1.0</span></span>|
|[<span data-ttu-id="674cd-170">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="674cd-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="674cd-171">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="674cd-171">Compose or read</span></span>|