# <a name="userprofile"></a><span data-ttu-id="04c62-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="04c62-101">userProfile</span></span>

### <span data-ttu-id="04c62-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="04c62-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="04c62-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="04c62-104">Requirements</span></span>

|<span data-ttu-id="04c62-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="04c62-105">Requirement</span></span>| <span data-ttu-id="04c62-106">Valor</span><span class="sxs-lookup"><span data-stu-id="04c62-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="04c62-107">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="04c62-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04c62-108">1.0</span><span class="sxs-lookup"><span data-stu-id="04c62-108">1.0</span></span>|
|[<span data-ttu-id="04c62-109">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="04c62-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04c62-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04c62-110">ReadItem</span></span>|
|[<span data-ttu-id="04c62-111">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="04c62-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="04c62-112">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="04c62-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="04c62-113">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="04c62-113">Members and methods</span></span>

| <span data-ttu-id="04c62-114">Miembro	</span><span class="sxs-lookup"><span data-stu-id="04c62-114">Member</span></span> | <span data-ttu-id="04c62-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="04c62-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="04c62-116">displayName</span><span class="sxs-lookup"><span data-stu-id="04c62-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="04c62-117">Miembro	</span><span class="sxs-lookup"><span data-stu-id="04c62-117">Member</span></span> |
| [<span data-ttu-id="04c62-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="04c62-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="04c62-119">Miembro	</span><span class="sxs-lookup"><span data-stu-id="04c62-119">Member</span></span> |
| [<span data-ttu-id="04c62-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="04c62-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="04c62-121">Miembro	</span><span class="sxs-lookup"><span data-stu-id="04c62-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="04c62-122">Miembros</span><span class="sxs-lookup"><span data-stu-id="04c62-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="04c62-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="04c62-123">displayName :String</span></span>

<span data-ttu-id="04c62-124">Obtiene el nombre para mostrar del usuario.</span><span class="sxs-lookup"><span data-stu-id="04c62-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="04c62-125">Tipo:</span><span class="sxs-lookup"><span data-stu-id="04c62-125">Type:</span></span>

*   <span data-ttu-id="04c62-126">String</span><span class="sxs-lookup"><span data-stu-id="04c62-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04c62-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="04c62-127">Requirements</span></span>

|<span data-ttu-id="04c62-128">Requirement</span><span class="sxs-lookup"><span data-stu-id="04c62-128">Requirement</span></span>| <span data-ttu-id="04c62-129">Valor</span><span class="sxs-lookup"><span data-stu-id="04c62-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="04c62-130">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="04c62-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04c62-131">1.0</span><span class="sxs-lookup"><span data-stu-id="04c62-131">1.0</span></span>|
|[<span data-ttu-id="04c62-132">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="04c62-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04c62-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04c62-133">ReadItem</span></span>|
|[<span data-ttu-id="04c62-134">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="04c62-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="04c62-135">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="04c62-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="04c62-136">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="04c62-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="04c62-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="04c62-137">emailAddress :String</span></span>

<span data-ttu-id="04c62-138">Obtiene la dirección de correo electrónico SMTP del usuario.</span><span class="sxs-lookup"><span data-stu-id="04c62-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="04c62-139">Tipo:</span><span class="sxs-lookup"><span data-stu-id="04c62-139">Type:</span></span>

*   <span data-ttu-id="04c62-140">String</span><span class="sxs-lookup"><span data-stu-id="04c62-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04c62-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="04c62-141">Requirements</span></span>

|<span data-ttu-id="04c62-142">Requirement</span><span class="sxs-lookup"><span data-stu-id="04c62-142">Requirement</span></span>| <span data-ttu-id="04c62-143">Valor</span><span class="sxs-lookup"><span data-stu-id="04c62-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="04c62-144">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="04c62-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04c62-145">1.0</span><span class="sxs-lookup"><span data-stu-id="04c62-145">1.0</span></span>|
|[<span data-ttu-id="04c62-146">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="04c62-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04c62-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04c62-147">ReadItem</span></span>|
|[<span data-ttu-id="04c62-148">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="04c62-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="04c62-149">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="04c62-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="04c62-150">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="04c62-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="04c62-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="04c62-151">timeZone :String</span></span>

<span data-ttu-id="04c62-152">Obtiene la zona horaria predeterminada del usuario.</span><span class="sxs-lookup"><span data-stu-id="04c62-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="04c62-153">Tipo:</span><span class="sxs-lookup"><span data-stu-id="04c62-153">Type:</span></span>

*   <span data-ttu-id="04c62-154">String</span><span class="sxs-lookup"><span data-stu-id="04c62-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04c62-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="04c62-155">Requirements</span></span>

|<span data-ttu-id="04c62-156">Requirement</span><span class="sxs-lookup"><span data-stu-id="04c62-156">Requirement</span></span>| <span data-ttu-id="04c62-157">Valor</span><span class="sxs-lookup"><span data-stu-id="04c62-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="04c62-158">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="04c62-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04c62-159">1.0</span><span class="sxs-lookup"><span data-stu-id="04c62-159">1.0</span></span>|
|[<span data-ttu-id="04c62-160">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="04c62-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04c62-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04c62-161">ReadItem</span></span>|
|[<span data-ttu-id="04c62-162">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="04c62-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="04c62-163">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="04c62-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="04c62-164">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="04c62-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```