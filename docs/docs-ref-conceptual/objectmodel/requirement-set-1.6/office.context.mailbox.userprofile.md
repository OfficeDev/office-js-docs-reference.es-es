
# <a name="userprofile"></a><span data-ttu-id="e2498-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="e2498-101">userProfile</span></span>

### <span data-ttu-id="e2498-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="e2498-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2498-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e2498-104">Requirements</span></span>

|<span data-ttu-id="e2498-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="e2498-105">Requirement</span></span>| <span data-ttu-id="e2498-106">Valor</span><span class="sxs-lookup"><span data-stu-id="e2498-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2498-107">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="e2498-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2498-108">1.0</span><span class="sxs-lookup"><span data-stu-id="e2498-108">1.0</span></span>|
|[<span data-ttu-id="e2498-109">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="e2498-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e2498-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2498-110">ReadItem</span></span>|
|[<span data-ttu-id="e2498-111">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="e2498-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e2498-112">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="e2498-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e2498-113">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="e2498-113">Members and methods</span></span>

| <span data-ttu-id="e2498-114">Miembro	</span><span class="sxs-lookup"><span data-stu-id="e2498-114">Member</span></span> | <span data-ttu-id="e2498-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="e2498-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e2498-116">accountType</span><span class="sxs-lookup"><span data-stu-id="e2498-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="e2498-117">Member</span><span class="sxs-lookup"><span data-stu-id="e2498-117">Member</span></span> |
| [<span data-ttu-id="e2498-118">displayName</span><span class="sxs-lookup"><span data-stu-id="e2498-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="e2498-119">Miembro	</span><span class="sxs-lookup"><span data-stu-id="e2498-119">Member</span></span> |
| [<span data-ttu-id="e2498-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="e2498-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="e2498-121">Miembro	</span><span class="sxs-lookup"><span data-stu-id="e2498-121">Member</span></span> |
| [<span data-ttu-id="e2498-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="e2498-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="e2498-123">Miembro	</span><span class="sxs-lookup"><span data-stu-id="e2498-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="e2498-124">Members</span><span class="sxs-lookup"><span data-stu-id="e2498-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="e2498-125">accountType: cadena</span><span class="sxs-lookup"><span data-stu-id="e2498-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="e2498-126">Este miembro actualmente sólo es compatible en Outlook 2016 o posterior para Mac (compilación 16.9.1212 o posterior).</span><span class="sxs-lookup"><span data-stu-id="e2498-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="e2498-127">Obtiene el tipo de cuenta del usuario asociado con el buzón de correo.</span><span class="sxs-lookup"><span data-stu-id="e2498-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="e2498-128">Los valores posibles se enumeran en la siguiente tabla.</span><span class="sxs-lookup"><span data-stu-id="e2498-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="e2498-129">Valor</span><span class="sxs-lookup"><span data-stu-id="e2498-129">Value</span></span> | <span data-ttu-id="e2498-130">Descripción</span><span class="sxs-lookup"><span data-stu-id="e2498-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="e2498-131">El buzón se encuentra en un servidor de Exchange local.</span><span class="sxs-lookup"><span data-stu-id="e2498-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="e2498-132">El buzón de correo está asociado con una cuenta de Gmail.</span><span class="sxs-lookup"><span data-stu-id="e2498-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="e2498-133">El buzón de correo está asociada con un Office 365 funciona o escuela cuenta.</span><span class="sxs-lookup"><span data-stu-id="e2498-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="e2498-134">El buzón de correo está asociado con una cuenta de Outlook.com personal.</span><span class="sxs-lookup"><span data-stu-id="e2498-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="e2498-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e2498-135">Type:</span></span>

*   <span data-ttu-id="e2498-136">String</span><span class="sxs-lookup"><span data-stu-id="e2498-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2498-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e2498-137">Requirements</span></span>

|<span data-ttu-id="e2498-138">Requirement</span><span class="sxs-lookup"><span data-stu-id="e2498-138">Requirement</span></span>| <span data-ttu-id="e2498-139">Valor</span><span class="sxs-lookup"><span data-stu-id="e2498-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2498-140">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="e2498-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2498-141">1.6</span><span class="sxs-lookup"><span data-stu-id="e2498-141">1.6</span></span> |
|[<span data-ttu-id="e2498-142">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="e2498-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e2498-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2498-143">ReadItem</span></span>|
|[<span data-ttu-id="e2498-144">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="e2498-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e2498-145">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="e2498-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e2498-146">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="e2498-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="e2498-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="e2498-147">displayName :String</span></span>

<span data-ttu-id="e2498-148">Obtiene el nombre para mostrar del usuario.</span><span class="sxs-lookup"><span data-stu-id="e2498-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="e2498-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e2498-149">Type:</span></span>

*   <span data-ttu-id="e2498-150">String</span><span class="sxs-lookup"><span data-stu-id="e2498-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2498-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e2498-151">Requirements</span></span>

|<span data-ttu-id="e2498-152">Requirement</span><span class="sxs-lookup"><span data-stu-id="e2498-152">Requirement</span></span>| <span data-ttu-id="e2498-153">Valor</span><span class="sxs-lookup"><span data-stu-id="e2498-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2498-154">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="e2498-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2498-155">1.0</span><span class="sxs-lookup"><span data-stu-id="e2498-155">1.0</span></span>|
|[<span data-ttu-id="e2498-156">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="e2498-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e2498-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2498-157">ReadItem</span></span>|
|[<span data-ttu-id="e2498-158">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="e2498-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e2498-159">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="e2498-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e2498-160">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="e2498-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="e2498-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="e2498-161">emailAddress :String</span></span>

<span data-ttu-id="e2498-162">Obtiene la dirección de correo electrónico SMTP del usuario.</span><span class="sxs-lookup"><span data-stu-id="e2498-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="e2498-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e2498-163">Type:</span></span>

*   <span data-ttu-id="e2498-164">String</span><span class="sxs-lookup"><span data-stu-id="e2498-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2498-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e2498-165">Requirements</span></span>

|<span data-ttu-id="e2498-166">Requirement</span><span class="sxs-lookup"><span data-stu-id="e2498-166">Requirement</span></span>| <span data-ttu-id="e2498-167">Valor</span><span class="sxs-lookup"><span data-stu-id="e2498-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2498-168">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="e2498-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2498-169">1.0</span><span class="sxs-lookup"><span data-stu-id="e2498-169">1.0</span></span>|
|[<span data-ttu-id="e2498-170">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="e2498-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e2498-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2498-171">ReadItem</span></span>|
|[<span data-ttu-id="e2498-172">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="e2498-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e2498-173">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="e2498-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e2498-174">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="e2498-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="e2498-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="e2498-175">timeZone :String</span></span>

<span data-ttu-id="e2498-176">Obtiene la zona horaria predeterminada del usuario.</span><span class="sxs-lookup"><span data-stu-id="e2498-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="e2498-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e2498-177">Type:</span></span>

*   <span data-ttu-id="e2498-178">String</span><span class="sxs-lookup"><span data-stu-id="e2498-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2498-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e2498-179">Requirements</span></span>

|<span data-ttu-id="e2498-180">Requirement</span><span class="sxs-lookup"><span data-stu-id="e2498-180">Requirement</span></span>| <span data-ttu-id="e2498-181">Valor</span><span class="sxs-lookup"><span data-stu-id="e2498-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2498-182">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="e2498-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2498-183">1.0</span><span class="sxs-lookup"><span data-stu-id="e2498-183">1.0</span></span>|
|[<span data-ttu-id="e2498-184">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="e2498-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e2498-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e2498-185">ReadItem</span></span>|
|[<span data-ttu-id="e2498-186">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="e2498-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e2498-187">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="e2498-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e2498-188">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="e2498-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```