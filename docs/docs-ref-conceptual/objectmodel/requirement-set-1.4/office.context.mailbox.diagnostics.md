
# <a name="diagnostics"></a><span data-ttu-id="2c22c-101">diagnostics</span><span class="sxs-lookup"><span data-stu-id="2c22c-101">diagnostics</span></span>

### <span data-ttu-id="2c22c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="2c22c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="2c22c-104">Proporciona información de diagnóstico a un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="2c22c-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2c22c-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2c22c-105">Requirements</span></span>

|<span data-ttu-id="2c22c-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="2c22c-106">Requirement</span></span>| <span data-ttu-id="2c22c-107">Valor</span><span class="sxs-lookup"><span data-stu-id="2c22c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c22c-108">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2c22c-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c22c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="2c22c-109">1.0</span></span>|
|[<span data-ttu-id="2c22c-110">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2c22c-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2c22c-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2c22c-111">ReadItem</span></span>|
|[<span data-ttu-id="2c22c-112">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2c22c-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2c22c-113">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2c22c-113">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="2c22c-114">Miembros</span><span class="sxs-lookup"><span data-stu-id="2c22c-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="2c22c-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="2c22c-115">hostName :String</span></span>

<span data-ttu-id="2c22c-116">Obtiene una cadena que representa el nombre de la aplicación host.</span><span class="sxs-lookup"><span data-stu-id="2c22c-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="2c22c-117">Una cadena que puede ser uno de los siguientes valores: `Outlook`, `OutlookIOS`, o `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="2c22c-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="2c22c-118">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2c22c-118">Type:</span></span>

*   <span data-ttu-id="2c22c-119">String</span><span class="sxs-lookup"><span data-stu-id="2c22c-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2c22c-120">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2c22c-120">Requirements</span></span>

|<span data-ttu-id="2c22c-121">Requirement</span><span class="sxs-lookup"><span data-stu-id="2c22c-121">Requirement</span></span>| <span data-ttu-id="2c22c-122">Valor</span><span class="sxs-lookup"><span data-stu-id="2c22c-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c22c-123">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2c22c-123">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c22c-124">1.0</span><span class="sxs-lookup"><span data-stu-id="2c22c-124">1.0</span></span>|
|[<span data-ttu-id="2c22c-125">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2c22c-125">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2c22c-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2c22c-126">ReadItem</span></span>|
|[<span data-ttu-id="2c22c-127">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2c22c-127">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2c22c-128">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2c22c-128">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="2c22c-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="2c22c-129">hostVersion :String</span></span>

<span data-ttu-id="2c22c-130">Obtiene una cadena que representa la versión de la aplicación host o de Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="2c22c-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="2c22c-p102">Si el complemento de correo se ejecuta en el cliente de escritorio de Outlook o en Outlook para iOS, la propiedad `hostVersion` devuelve la versión de la aplicación host, Outlook. En Outlook Web App, la propiedad devuelve la versión de Exchange Server. Un ejemplo es la cadena `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="2c22c-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="2c22c-134">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2c22c-134">Type:</span></span>

*   <span data-ttu-id="2c22c-135">String</span><span class="sxs-lookup"><span data-stu-id="2c22c-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2c22c-136">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2c22c-136">Requirements</span></span>

|<span data-ttu-id="2c22c-137">Requirement</span><span class="sxs-lookup"><span data-stu-id="2c22c-137">Requirement</span></span>| <span data-ttu-id="2c22c-138">Valor</span><span class="sxs-lookup"><span data-stu-id="2c22c-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c22c-139">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2c22c-139">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c22c-140">1.0</span><span class="sxs-lookup"><span data-stu-id="2c22c-140">1.0</span></span>|
|[<span data-ttu-id="2c22c-141">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2c22c-141">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2c22c-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2c22c-142">ReadItem</span></span>|
|[<span data-ttu-id="2c22c-143">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2c22c-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2c22c-144">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2c22c-144">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="2c22c-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="2c22c-145">OWAView :String</span></span>

<span data-ttu-id="2c22c-146">Obtiene una cadena que representa la vista actual de Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="2c22c-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="2c22c-147">La cadena devuelta puede ser uno de los valores siguientes: `OneColumn`, `TwoColumns` o `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="2c22c-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="2c22c-148">Si la aplicación host no es Outlook Web App, obtener acceso a esta propiedad será `undefined`.</span><span class="sxs-lookup"><span data-stu-id="2c22c-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="2c22c-149">Outlook Web App tiene tres vistas que se corresponden con el ancho de la pantalla, con la ventana y con el número de columnas que pueden mostrarse:</span><span class="sxs-lookup"><span data-stu-id="2c22c-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="2c22c-p103">`OneColumn` se muestra cuando la pantalla es estrecha. Outlook Web App usa este diseño de una sola columna en la pantalla completa de un smartphone.</span><span class="sxs-lookup"><span data-stu-id="2c22c-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="2c22c-p104">`TwoColumns` se muestra cuando la pantalla es más ancha. Outlook Web App usa esta vista en la mayor parte de las tabletas.</span><span class="sxs-lookup"><span data-stu-id="2c22c-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="2c22c-p105">`ThreeColumns` se muestra cuando la pantalla es ancha. Por ejemplo, Outlook Web App usa esta vista en una ventana de pantalla completa en un equipo de escritorio.</span><span class="sxs-lookup"><span data-stu-id="2c22c-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="2c22c-156">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2c22c-156">Type:</span></span>

*   <span data-ttu-id="2c22c-157">String</span><span class="sxs-lookup"><span data-stu-id="2c22c-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2c22c-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2c22c-158">Requirements</span></span>

|<span data-ttu-id="2c22c-159">Requirement</span><span class="sxs-lookup"><span data-stu-id="2c22c-159">Requirement</span></span>| <span data-ttu-id="2c22c-160">Valor</span><span class="sxs-lookup"><span data-stu-id="2c22c-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c22c-161">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2c22c-161">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c22c-162">1.0</span><span class="sxs-lookup"><span data-stu-id="2c22c-162">1.0</span></span>|
|[<span data-ttu-id="2c22c-163">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2c22c-163">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2c22c-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2c22c-164">ReadItem</span></span>|
|[<span data-ttu-id="2c22c-165">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2c22c-165">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2c22c-166">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2c22c-166">Compose or read</span></span>|