
# <a name="userprofile"></a><span data-ttu-id="4efe5-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="4efe5-101">userProfile</span></span>

### <span data-ttu-id="4efe5-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="4efe5-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="4efe5-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4efe5-104">Requirements</span></span>

|<span data-ttu-id="4efe5-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="4efe5-105">Requirement</span></span>| <span data-ttu-id="4efe5-106">Valor</span><span class="sxs-lookup"><span data-stu-id="4efe5-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="4efe5-107">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="4efe5-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4efe5-108">1.0</span><span class="sxs-lookup"><span data-stu-id="4efe5-108">1.0</span></span>|
|[<span data-ttu-id="4efe5-109">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="4efe5-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4efe5-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4efe5-110">ReadItem</span></span>|
|[<span data-ttu-id="4efe5-111">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="4efe5-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4efe5-112">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="4efe5-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="4efe5-113">Miembros</span><span class="sxs-lookup"><span data-stu-id="4efe5-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="4efe5-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="4efe5-114">displayName :String</span></span>

<span data-ttu-id="4efe5-115">Obtiene el nombre para mostrar del usuario.</span><span class="sxs-lookup"><span data-stu-id="4efe5-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="4efe5-116">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4efe5-116">Type:</span></span>

*   <span data-ttu-id="4efe5-117">String</span><span class="sxs-lookup"><span data-stu-id="4efe5-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4efe5-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4efe5-118">Requirements</span></span>

|<span data-ttu-id="4efe5-119">Requirement</span><span class="sxs-lookup"><span data-stu-id="4efe5-119">Requirement</span></span>| <span data-ttu-id="4efe5-120">Valor</span><span class="sxs-lookup"><span data-stu-id="4efe5-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="4efe5-121">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="4efe5-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4efe5-122">1.0</span><span class="sxs-lookup"><span data-stu-id="4efe5-122">1.0</span></span>|
|[<span data-ttu-id="4efe5-123">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="4efe5-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4efe5-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4efe5-124">ReadItem</span></span>|
|[<span data-ttu-id="4efe5-125">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="4efe5-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4efe5-126">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="4efe5-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4efe5-127">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="4efe5-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="4efe5-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="4efe5-128">emailAddress :String</span></span>

<span data-ttu-id="4efe5-129">Obtiene la dirección de correo electrónico SMTP del usuario.</span><span class="sxs-lookup"><span data-stu-id="4efe5-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="4efe5-130">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4efe5-130">Type:</span></span>

*   <span data-ttu-id="4efe5-131">String</span><span class="sxs-lookup"><span data-stu-id="4efe5-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4efe5-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4efe5-132">Requirements</span></span>

|<span data-ttu-id="4efe5-133">Requirement</span><span class="sxs-lookup"><span data-stu-id="4efe5-133">Requirement</span></span>| <span data-ttu-id="4efe5-134">Valor</span><span class="sxs-lookup"><span data-stu-id="4efe5-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="4efe5-135">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="4efe5-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4efe5-136">1.0</span><span class="sxs-lookup"><span data-stu-id="4efe5-136">1.0</span></span>|
|[<span data-ttu-id="4efe5-137">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="4efe5-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4efe5-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4efe5-138">ReadItem</span></span>|
|[<span data-ttu-id="4efe5-139">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="4efe5-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4efe5-140">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="4efe5-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4efe5-141">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="4efe5-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="4efe5-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="4efe5-142">timeZone :String</span></span>

<span data-ttu-id="4efe5-143">Obtiene la zona horaria predeterminada del usuario.</span><span class="sxs-lookup"><span data-stu-id="4efe5-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="4efe5-144">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4efe5-144">Type:</span></span>

*   <span data-ttu-id="4efe5-145">String</span><span class="sxs-lookup"><span data-stu-id="4efe5-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4efe5-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4efe5-146">Requirements</span></span>

|<span data-ttu-id="4efe5-147">Requirement</span><span class="sxs-lookup"><span data-stu-id="4efe5-147">Requirement</span></span>| <span data-ttu-id="4efe5-148">Valor</span><span class="sxs-lookup"><span data-stu-id="4efe5-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="4efe5-149">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="4efe5-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4efe5-150">1.0</span><span class="sxs-lookup"><span data-stu-id="4efe5-150">1.0</span></span>|
|[<span data-ttu-id="4efe5-151">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="4efe5-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4efe5-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4efe5-152">ReadItem</span></span>|
|[<span data-ttu-id="4efe5-153">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="4efe5-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4efe5-154">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="4efe5-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4efe5-155">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="4efe5-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```