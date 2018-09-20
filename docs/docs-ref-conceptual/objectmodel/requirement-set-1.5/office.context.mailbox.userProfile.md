# <a name="userprofile"></a><span data-ttu-id="1dd9a-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="1dd9a-101">userProfile</span></span>

### <span data-ttu-id="1dd9a-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="1dd9a-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dd9a-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1dd9a-104">Requirements</span></span>

|<span data-ttu-id="1dd9a-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="1dd9a-105">Requirement</span></span>| <span data-ttu-id="1dd9a-106">Valor</span><span class="sxs-lookup"><span data-stu-id="1dd9a-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dd9a-107">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1dd9a-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dd9a-108">1.0</span><span class="sxs-lookup"><span data-stu-id="1dd9a-108">1.0</span></span>|
|[<span data-ttu-id="1dd9a-109">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1dd9a-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dd9a-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dd9a-110">ReadItem</span></span>|
|[<span data-ttu-id="1dd9a-111">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1dd9a-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dd9a-112">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1dd9a-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1dd9a-113">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="1dd9a-113">Members and methods</span></span>

| <span data-ttu-id="1dd9a-114">Miembro	</span><span class="sxs-lookup"><span data-stu-id="1dd9a-114">Member</span></span> | <span data-ttu-id="1dd9a-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="1dd9a-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1dd9a-116">displayName</span><span class="sxs-lookup"><span data-stu-id="1dd9a-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="1dd9a-117">Miembro	</span><span class="sxs-lookup"><span data-stu-id="1dd9a-117">Member</span></span> |
| [<span data-ttu-id="1dd9a-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="1dd9a-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="1dd9a-119">Miembro	</span><span class="sxs-lookup"><span data-stu-id="1dd9a-119">Member</span></span> |
| [<span data-ttu-id="1dd9a-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="1dd9a-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="1dd9a-121">Miembro	</span><span class="sxs-lookup"><span data-stu-id="1dd9a-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="1dd9a-122">Miembros</span><span class="sxs-lookup"><span data-stu-id="1dd9a-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="1dd9a-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="1dd9a-123">displayName :String</span></span>

<span data-ttu-id="1dd9a-124">Obtiene el nombre para mostrar del usuario.</span><span class="sxs-lookup"><span data-stu-id="1dd9a-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="1dd9a-125">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1dd9a-125">Type:</span></span>

*   <span data-ttu-id="1dd9a-126">String</span><span class="sxs-lookup"><span data-stu-id="1dd9a-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dd9a-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1dd9a-127">Requirements</span></span>

|<span data-ttu-id="1dd9a-128">Requirement</span><span class="sxs-lookup"><span data-stu-id="1dd9a-128">Requirement</span></span>| <span data-ttu-id="1dd9a-129">Valor</span><span class="sxs-lookup"><span data-stu-id="1dd9a-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dd9a-130">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1dd9a-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dd9a-131">1.0</span><span class="sxs-lookup"><span data-stu-id="1dd9a-131">1.0</span></span>|
|[<span data-ttu-id="1dd9a-132">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1dd9a-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dd9a-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dd9a-133">ReadItem</span></span>|
|[<span data-ttu-id="1dd9a-134">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1dd9a-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dd9a-135">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1dd9a-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dd9a-136">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="1dd9a-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="1dd9a-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="1dd9a-137">emailAddress :String</span></span>

<span data-ttu-id="1dd9a-138">Obtiene la dirección de correo electrónico SMTP del usuario.</span><span class="sxs-lookup"><span data-stu-id="1dd9a-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="1dd9a-139">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1dd9a-139">Type:</span></span>

*   <span data-ttu-id="1dd9a-140">String</span><span class="sxs-lookup"><span data-stu-id="1dd9a-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dd9a-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1dd9a-141">Requirements</span></span>

|<span data-ttu-id="1dd9a-142">Requirement</span><span class="sxs-lookup"><span data-stu-id="1dd9a-142">Requirement</span></span>| <span data-ttu-id="1dd9a-143">Valor</span><span class="sxs-lookup"><span data-stu-id="1dd9a-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dd9a-144">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1dd9a-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dd9a-145">1.0</span><span class="sxs-lookup"><span data-stu-id="1dd9a-145">1.0</span></span>|
|[<span data-ttu-id="1dd9a-146">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1dd9a-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dd9a-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dd9a-147">ReadItem</span></span>|
|[<span data-ttu-id="1dd9a-148">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1dd9a-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dd9a-149">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1dd9a-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dd9a-150">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="1dd9a-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="1dd9a-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="1dd9a-151">timeZone :String</span></span>

<span data-ttu-id="1dd9a-152">Obtiene la zona horaria predeterminada del usuario.</span><span class="sxs-lookup"><span data-stu-id="1dd9a-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="1dd9a-153">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1dd9a-153">Type:</span></span>

*   <span data-ttu-id="1dd9a-154">String</span><span class="sxs-lookup"><span data-stu-id="1dd9a-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dd9a-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1dd9a-155">Requirements</span></span>

|<span data-ttu-id="1dd9a-156">Requirement</span><span class="sxs-lookup"><span data-stu-id="1dd9a-156">Requirement</span></span>| <span data-ttu-id="1dd9a-157">Valor</span><span class="sxs-lookup"><span data-stu-id="1dd9a-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dd9a-158">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1dd9a-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dd9a-159">1.0</span><span class="sxs-lookup"><span data-stu-id="1dd9a-159">1.0</span></span>|
|[<span data-ttu-id="1dd9a-160">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1dd9a-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dd9a-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dd9a-161">ReadItem</span></span>|
|[<span data-ttu-id="1dd9a-162">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1dd9a-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dd9a-163">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1dd9a-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dd9a-164">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="1dd9a-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```