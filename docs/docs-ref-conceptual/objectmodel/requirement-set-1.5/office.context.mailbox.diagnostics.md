# <a name="diagnostics"></a><span data-ttu-id="165b2-101">diagnostics</span><span class="sxs-lookup"><span data-stu-id="165b2-101">diagnostics</span></span>

### <span data-ttu-id="165b2-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="165b2-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="165b2-104">Proporciona información de diagnóstico a un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="165b2-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="165b2-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="165b2-105">Requirements</span></span>

|<span data-ttu-id="165b2-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="165b2-106">Requirement</span></span>| <span data-ttu-id="165b2-107">Valor</span><span class="sxs-lookup"><span data-stu-id="165b2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="165b2-108">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="165b2-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="165b2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="165b2-109">1.0</span></span>|
|[<span data-ttu-id="165b2-110">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="165b2-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="165b2-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="165b2-111">ReadItem</span></span>|
|[<span data-ttu-id="165b2-112">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="165b2-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="165b2-113">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="165b2-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="165b2-114">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="165b2-114">Members and methods</span></span>

| <span data-ttu-id="165b2-115">Miembro	</span><span class="sxs-lookup"><span data-stu-id="165b2-115">Member</span></span> | <span data-ttu-id="165b2-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="165b2-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="165b2-117">hostName</span><span class="sxs-lookup"><span data-stu-id="165b2-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="165b2-118">Miembro	</span><span class="sxs-lookup"><span data-stu-id="165b2-118">Member</span></span> |
| [<span data-ttu-id="165b2-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="165b2-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="165b2-120">Miembro	</span><span class="sxs-lookup"><span data-stu-id="165b2-120">Member</span></span> |
| [<span data-ttu-id="165b2-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="165b2-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="165b2-122">Miembro	</span><span class="sxs-lookup"><span data-stu-id="165b2-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="165b2-123">Miembros</span><span class="sxs-lookup"><span data-stu-id="165b2-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="165b2-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="165b2-124">hostName :String</span></span>

<span data-ttu-id="165b2-125">Obtiene una cadena que representa el nombre de la aplicación host.</span><span class="sxs-lookup"><span data-stu-id="165b2-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="165b2-126">Una cadena que puede ser uno de los siguientes valores: `Outlook`, `OutlookIOS`, o `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="165b2-126">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="165b2-127">Tipo:</span><span class="sxs-lookup"><span data-stu-id="165b2-127">Type:</span></span>

*   <span data-ttu-id="165b2-128">String</span><span class="sxs-lookup"><span data-stu-id="165b2-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="165b2-129">Requisitos</span><span class="sxs-lookup"><span data-stu-id="165b2-129">Requirements</span></span>

|<span data-ttu-id="165b2-130">Requirement</span><span class="sxs-lookup"><span data-stu-id="165b2-130">Requirement</span></span>| <span data-ttu-id="165b2-131">Valor</span><span class="sxs-lookup"><span data-stu-id="165b2-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="165b2-132">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="165b2-132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="165b2-133">1.0</span><span class="sxs-lookup"><span data-stu-id="165b2-133">1.0</span></span>|
|[<span data-ttu-id="165b2-134">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="165b2-134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="165b2-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="165b2-135">ReadItem</span></span>|
|[<span data-ttu-id="165b2-136">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="165b2-136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="165b2-137">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="165b2-137">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="165b2-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="165b2-138">hostVersion :String</span></span>

<span data-ttu-id="165b2-139">Obtiene una cadena que representa la versión de la aplicación host o de Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="165b2-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="165b2-p102">Si el complemento de correo se ejecuta en el cliente de escritorio de Outlook o en Outlook para iOS, la propiedad `hostVersion` devuelve la versión de la aplicación host, Outlook. En Outlook Web App, la propiedad devuelve la versión de Exchange Server. Un ejemplo es la cadena `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="165b2-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="165b2-143">Tipo:</span><span class="sxs-lookup"><span data-stu-id="165b2-143">Type:</span></span>

*   <span data-ttu-id="165b2-144">String</span><span class="sxs-lookup"><span data-stu-id="165b2-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="165b2-145">Requisitos</span><span class="sxs-lookup"><span data-stu-id="165b2-145">Requirements</span></span>

|<span data-ttu-id="165b2-146">Requirement</span><span class="sxs-lookup"><span data-stu-id="165b2-146">Requirement</span></span>| <span data-ttu-id="165b2-147">Valor</span><span class="sxs-lookup"><span data-stu-id="165b2-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="165b2-148">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="165b2-148">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="165b2-149">1.0</span><span class="sxs-lookup"><span data-stu-id="165b2-149">1.0</span></span>|
|[<span data-ttu-id="165b2-150">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="165b2-150">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="165b2-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="165b2-151">ReadItem</span></span>|
|[<span data-ttu-id="165b2-152">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="165b2-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="165b2-153">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="165b2-153">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="165b2-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="165b2-154">OWAView :String</span></span>

<span data-ttu-id="165b2-155">Obtiene una cadena que representa la vista actual de Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="165b2-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="165b2-156">La cadena devuelta puede ser uno de los valores siguientes: `OneColumn`, `TwoColumns` o `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="165b2-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="165b2-157">Si la aplicación host no es Outlook Web App, obtener acceso a esta propiedad será `undefined`.</span><span class="sxs-lookup"><span data-stu-id="165b2-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="165b2-158">Outlook Web App tiene tres vistas que se corresponden con el ancho de la pantalla, con la ventana y con el número de columnas que pueden mostrarse:</span><span class="sxs-lookup"><span data-stu-id="165b2-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="165b2-p103">`OneColumn` se muestra cuando la pantalla es estrecha. Outlook Web App usa este diseño de una sola columna en la pantalla completa de un smartphone.</span><span class="sxs-lookup"><span data-stu-id="165b2-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="165b2-p104">`TwoColumns` se muestra cuando la pantalla es más ancha. Outlook Web App usa esta vista en la mayor parte de las tabletas.</span><span class="sxs-lookup"><span data-stu-id="165b2-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="165b2-p105">`ThreeColumns` se muestra cuando la pantalla es ancha. Por ejemplo, Outlook Web App usa esta vista en una ventana de pantalla completa en un equipo de escritorio.</span><span class="sxs-lookup"><span data-stu-id="165b2-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="165b2-165">Tipo:</span><span class="sxs-lookup"><span data-stu-id="165b2-165">Type:</span></span>

*   <span data-ttu-id="165b2-166">String</span><span class="sxs-lookup"><span data-stu-id="165b2-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="165b2-167">Requisitos</span><span class="sxs-lookup"><span data-stu-id="165b2-167">Requirements</span></span>

|<span data-ttu-id="165b2-168">Requirement</span><span class="sxs-lookup"><span data-stu-id="165b2-168">Requirement</span></span>| <span data-ttu-id="165b2-169">Valor</span><span class="sxs-lookup"><span data-stu-id="165b2-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="165b2-170">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="165b2-170">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="165b2-171">1.0</span><span class="sxs-lookup"><span data-stu-id="165b2-171">1.0</span></span>|
|[<span data-ttu-id="165b2-172">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="165b2-172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="165b2-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="165b2-173">ReadItem</span></span>|
|[<span data-ttu-id="165b2-174">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="165b2-174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="165b2-175">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="165b2-175">Compose or read</span></span>|