
# <a name="context"></a><span data-ttu-id="1aae5-101">contexto</span><span class="sxs-lookup"><span data-stu-id="1aae5-101">context</span></span>

### <span data-ttu-id="1aae5-p101">[Office](Office.md). context</span><span class="sxs-lookup"><span data-stu-id="1aae5-p101">[Office](Office.md). context</span></span>

<span data-ttu-id="1aae5-p102">El espacio de nombres de Office .context proporciona interfaces compartidas que los complementos usan en todas las aplicaciones de Office. Este listado documenta solo aquellas interfaces que usan los complementos de Outlook. Para obtener un listado completo del espacio de nombres Office.context, vea [Referencia de Office.context de referencia de la API compartida](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="1aae5-p102">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>


##### <a name="requirements"></a><span data-ttu-id="1aae5-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1aae5-106">Requirements</span></span>

|<span data-ttu-id="1aae5-107">Requirement</span><span class="sxs-lookup"><span data-stu-id="1aae5-107">Requirement</span></span>| <span data-ttu-id="1aae5-108">Valor</span><span class="sxs-lookup"><span data-stu-id="1aae5-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="1aae5-109">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1aae5-109">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1aae5-110">1.0</span><span class="sxs-lookup"><span data-stu-id="1aae5-110">1.0</span></span>|
|[<span data-ttu-id="1aae5-111">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1aae5-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1aae5-112">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1aae5-112">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="1aae5-113">Espacios de nombres</span><span class="sxs-lookup"><span data-stu-id="1aae5-113">Namespaces</span></span>

<span data-ttu-id="1aae5-114">[buzón de correo](office.context.mailbox.md) - proporciona acceso al modelo de objeto de complemento de Outlook para Microsoft Outlook y Microsoft Outlook en el web.</span><span class="sxs-lookup"><span data-stu-id="1aae5-114">[mailbox](office.context.mailbox.md) - Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="1aae5-115">Miembros</span><span class="sxs-lookup"><span data-stu-id="1aae5-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="1aae5-116">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="1aae5-116">displayLanguage :String</span></span>

<span data-ttu-id="1aae5-117">Obtiene la configuración local (de idioma) en un formato de etiqueta de idioma RFC 1766 especificado por el usuario para la interfaz de usuario de la aplicación host de Office.</span><span class="sxs-lookup"><span data-stu-id="1aae5-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="1aae5-118">El valor `displayLanguage` refleja la configuración de **Mostrar idioma** actual que se ha especificado desde **Archivo > Opciones > Idioma**, en la aplicación host de Office.</span><span class="sxs-lookup"><span data-stu-id="1aae5-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="1aae5-119">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1aae5-119">Type:</span></span>

*   <span data-ttu-id="1aae5-120">String</span><span class="sxs-lookup"><span data-stu-id="1aae5-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1aae5-121">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1aae5-121">Requirements</span></span>

|<span data-ttu-id="1aae5-122">Requirement</span><span class="sxs-lookup"><span data-stu-id="1aae5-122">Requirement</span></span>| <span data-ttu-id="1aae5-123">Valor</span><span class="sxs-lookup"><span data-stu-id="1aae5-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="1aae5-124">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1aae5-124">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1aae5-125">1.0</span><span class="sxs-lookup"><span data-stu-id="1aae5-125">1.0</span></span>|
|[<span data-ttu-id="1aae5-126">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1aae5-126">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1aae5-127">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1aae5-127">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1aae5-128">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="1aae5-128">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook12officeroamingsettings"></a><span data-ttu-id="1aae5-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="1aae5-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)</span></span>

<span data-ttu-id="1aae5-130">Obtiene un objeto que representa la configuración o el estado personalizado de un complemento de correo que se guardó en el buzón de un usuario.</span><span class="sxs-lookup"><span data-stu-id="1aae5-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="1aae5-131">El objeto `RoamingSettings` le permite almacenar y tener acceso a datos para un complemento de correo almacenado en el buzón de un usuario, de forma que esté disponible para ese complemento cuando se ejecute desde cualquier aplicación de cliente host usada para tener acceso a ese buzón.</span><span class="sxs-lookup"><span data-stu-id="1aae5-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="1aae5-132">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1aae5-132">Type:</span></span>

*   [<span data-ttu-id="1aae5-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1aae5-133">RoamingSettings</span></span>](/javascript/api/outlook_1_2/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="1aae5-134">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1aae5-134">Requirements</span></span>

|<span data-ttu-id="1aae5-135">Requirement</span><span class="sxs-lookup"><span data-stu-id="1aae5-135">Requirement</span></span>| <span data-ttu-id="1aae5-136">Valor</span><span class="sxs-lookup"><span data-stu-id="1aae5-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="1aae5-137">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="1aae5-137">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1aae5-138">1.0</span><span class="sxs-lookup"><span data-stu-id="1aae5-138">1.0</span></span>|
|[<span data-ttu-id="1aae5-139">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="1aae5-139">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1aae5-140">Restringido</span><span class="sxs-lookup"><span data-stu-id="1aae5-140">Restricted</span></span>|
|[<span data-ttu-id="1aae5-141">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="1aae5-141">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1aae5-142">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="1aae5-142">Compose or read</span></span>|