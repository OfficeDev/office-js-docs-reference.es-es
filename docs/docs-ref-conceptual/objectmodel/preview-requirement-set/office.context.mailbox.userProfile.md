
# <a name="userprofile"></a><span data-ttu-id="02a13-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="02a13-101">userProfile</span></span>

### <span data-ttu-id="02a13-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="02a13-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="02a13-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="02a13-104">Requirements</span></span>

|<span data-ttu-id="02a13-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="02a13-105">Requirement</span></span>| <span data-ttu-id="02a13-106">Valor</span><span class="sxs-lookup"><span data-stu-id="02a13-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="02a13-107">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="02a13-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02a13-108">1.0</span><span class="sxs-lookup"><span data-stu-id="02a13-108">1.0</span></span>|
|[<span data-ttu-id="02a13-109">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="02a13-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02a13-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02a13-110">ReadItem</span></span>|
|[<span data-ttu-id="02a13-111">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="02a13-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="02a13-112">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="02a13-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="02a13-113">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="02a13-113">Members and methods</span></span>

| <span data-ttu-id="02a13-114">Miembro	</span><span class="sxs-lookup"><span data-stu-id="02a13-114">Member</span></span> | <span data-ttu-id="02a13-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="02a13-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="02a13-116">accountType</span><span class="sxs-lookup"><span data-stu-id="02a13-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="02a13-117">Member</span><span class="sxs-lookup"><span data-stu-id="02a13-117">Member</span></span> |
| [<span data-ttu-id="02a13-118">displayName</span><span class="sxs-lookup"><span data-stu-id="02a13-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="02a13-119">Miembro	</span><span class="sxs-lookup"><span data-stu-id="02a13-119">Member</span></span> |
| [<span data-ttu-id="02a13-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="02a13-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="02a13-121">Miembro	</span><span class="sxs-lookup"><span data-stu-id="02a13-121">Member</span></span> |
| [<span data-ttu-id="02a13-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="02a13-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="02a13-123">Miembro	</span><span class="sxs-lookup"><span data-stu-id="02a13-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="02a13-124">Members</span><span class="sxs-lookup"><span data-stu-id="02a13-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="02a13-125">accountType: cadena</span><span class="sxs-lookup"><span data-stu-id="02a13-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="02a13-126">Este miembro actualmente sólo es compatible en Outlook 2016 o posterior para Mac (compilación 16.9.1212 o posterior).</span><span class="sxs-lookup"><span data-stu-id="02a13-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="02a13-127">Obtiene el tipo de cuenta del usuario asociado con el buzón de correo.</span><span class="sxs-lookup"><span data-stu-id="02a13-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="02a13-128">Los valores posibles se enumeran en la siguiente tabla.</span><span class="sxs-lookup"><span data-stu-id="02a13-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="02a13-129">Valor</span><span class="sxs-lookup"><span data-stu-id="02a13-129">Value</span></span> | <span data-ttu-id="02a13-130">Descripción</span><span class="sxs-lookup"><span data-stu-id="02a13-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="02a13-131">El buzón se encuentra en un servidor de Exchange local.</span><span class="sxs-lookup"><span data-stu-id="02a13-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="02a13-132">El buzón de correo está asociado con una cuenta de Gmail.</span><span class="sxs-lookup"><span data-stu-id="02a13-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="02a13-133">El buzón de correo está asociada con un Office 365 funciona o escuela cuenta.</span><span class="sxs-lookup"><span data-stu-id="02a13-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="02a13-134">El buzón de correo está asociado con una cuenta de Outlook.com personal.</span><span class="sxs-lookup"><span data-stu-id="02a13-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="02a13-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="02a13-135">Type:</span></span>

*   <span data-ttu-id="02a13-136">String</span><span class="sxs-lookup"><span data-stu-id="02a13-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02a13-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="02a13-137">Requirements</span></span>

|<span data-ttu-id="02a13-138">Requirement</span><span class="sxs-lookup"><span data-stu-id="02a13-138">Requirement</span></span>| <span data-ttu-id="02a13-139">Valor</span><span class="sxs-lookup"><span data-stu-id="02a13-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="02a13-140">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="02a13-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02a13-141">1.6</span><span class="sxs-lookup"><span data-stu-id="02a13-141">1.6</span></span> |
|[<span data-ttu-id="02a13-142">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="02a13-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02a13-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02a13-143">ReadItem</span></span>|
|[<span data-ttu-id="02a13-144">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="02a13-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="02a13-145">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="02a13-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="02a13-146">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="02a13-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="02a13-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="02a13-147">displayName :String</span></span>

<span data-ttu-id="02a13-148">Obtiene el nombre para mostrar del usuario.</span><span class="sxs-lookup"><span data-stu-id="02a13-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="02a13-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="02a13-149">Type:</span></span>

*   <span data-ttu-id="02a13-150">String</span><span class="sxs-lookup"><span data-stu-id="02a13-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02a13-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="02a13-151">Requirements</span></span>

|<span data-ttu-id="02a13-152">Requirement</span><span class="sxs-lookup"><span data-stu-id="02a13-152">Requirement</span></span>| <span data-ttu-id="02a13-153">Valor</span><span class="sxs-lookup"><span data-stu-id="02a13-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="02a13-154">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="02a13-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02a13-155">1.0</span><span class="sxs-lookup"><span data-stu-id="02a13-155">1.0</span></span>|
|[<span data-ttu-id="02a13-156">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="02a13-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02a13-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02a13-157">ReadItem</span></span>|
|[<span data-ttu-id="02a13-158">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="02a13-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="02a13-159">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="02a13-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="02a13-160">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="02a13-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="02a13-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="02a13-161">emailAddress :String</span></span>

<span data-ttu-id="02a13-162">Obtiene la dirección de correo electrónico SMTP del usuario.</span><span class="sxs-lookup"><span data-stu-id="02a13-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="02a13-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="02a13-163">Type:</span></span>

*   <span data-ttu-id="02a13-164">String</span><span class="sxs-lookup"><span data-stu-id="02a13-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02a13-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="02a13-165">Requirements</span></span>

|<span data-ttu-id="02a13-166">Requirement</span><span class="sxs-lookup"><span data-stu-id="02a13-166">Requirement</span></span>| <span data-ttu-id="02a13-167">Valor</span><span class="sxs-lookup"><span data-stu-id="02a13-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="02a13-168">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="02a13-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02a13-169">1.0</span><span class="sxs-lookup"><span data-stu-id="02a13-169">1.0</span></span>|
|[<span data-ttu-id="02a13-170">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="02a13-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02a13-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02a13-171">ReadItem</span></span>|
|[<span data-ttu-id="02a13-172">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="02a13-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="02a13-173">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="02a13-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="02a13-174">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="02a13-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="02a13-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="02a13-175">timeZone :String</span></span>

<span data-ttu-id="02a13-176">Obtiene la zona horaria predeterminada del usuario.</span><span class="sxs-lookup"><span data-stu-id="02a13-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="02a13-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="02a13-177">Type:</span></span>

*   <span data-ttu-id="02a13-178">String</span><span class="sxs-lookup"><span data-stu-id="02a13-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02a13-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="02a13-179">Requirements</span></span>

|<span data-ttu-id="02a13-180">Requirement</span><span class="sxs-lookup"><span data-stu-id="02a13-180">Requirement</span></span>| <span data-ttu-id="02a13-181">Valor</span><span class="sxs-lookup"><span data-stu-id="02a13-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="02a13-182">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="02a13-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02a13-183">1.0</span><span class="sxs-lookup"><span data-stu-id="02a13-183">1.0</span></span>|
|[<span data-ttu-id="02a13-184">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="02a13-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02a13-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02a13-185">ReadItem</span></span>|
|[<span data-ttu-id="02a13-186">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="02a13-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="02a13-187">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="02a13-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="02a13-188">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="02a13-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```