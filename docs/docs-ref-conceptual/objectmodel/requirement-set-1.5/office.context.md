# <a name="context"></a><span data-ttu-id="79dd2-101">contexto</span><span class="sxs-lookup"><span data-stu-id="79dd2-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="79dd2-102">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="79dd2-102">[Office](Office.md).context</span></span>

<span data-ttu-id="79dd2-p101">El espacio de nombres de Office .context proporciona interfaces compartidas que los complementos usan en todas las aplicaciones de Office. Este listado documenta solo aquellas interfaces que usan los complementos de Outlook. Para obtener un listado completo del espacio de nombres Office.context, vea [Referencia de Office.context de referencia de la API compartida](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="79dd2-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="79dd2-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="79dd2-105">Requirements</span></span>

|<span data-ttu-id="79dd2-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="79dd2-106">Requirement</span></span>| <span data-ttu-id="79dd2-107">Valor</span><span class="sxs-lookup"><span data-stu-id="79dd2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="79dd2-108">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="79dd2-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79dd2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="79dd2-109">1.0</span></span>|
|[<span data-ttu-id="79dd2-110">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="79dd2-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79dd2-111">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="79dd2-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="79dd2-112">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="79dd2-112">Members and methods</span></span>

| <span data-ttu-id="79dd2-113">Miembro	</span><span class="sxs-lookup"><span data-stu-id="79dd2-113">Member</span></span> | <span data-ttu-id="79dd2-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="79dd2-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="79dd2-115">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="79dd2-115">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="79dd2-116">Miembro	</span><span class="sxs-lookup"><span data-stu-id="79dd2-116">Member</span></span> |
| [<span data-ttu-id="79dd2-117">officeTheme</span><span class="sxs-lookup"><span data-stu-id="79dd2-117">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="79dd2-118">Miembro	</span><span class="sxs-lookup"><span data-stu-id="79dd2-118">Member</span></span> |
| [<span data-ttu-id="79dd2-119">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="79dd2-119">roamingSettings</span></span>](#roamingsettings-roamingsettingsjavascriptapioutlook15officeroamingsettings) | <span data-ttu-id="79dd2-120">Miembro	</span><span class="sxs-lookup"><span data-stu-id="79dd2-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="79dd2-121">Espacios de nombres</span><span class="sxs-lookup"><span data-stu-id="79dd2-121">Namespaces</span></span>

<span data-ttu-id="79dd2-122">[buzón de correo](office.context.mailbox.md): proporciona acceso al modelo de objeto de complemento de Outlook para Microsoft Outlook y Microsoft Outlook en el web.</span><span class="sxs-lookup"><span data-stu-id="79dd2-122">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="79dd2-123">Miembros</span><span class="sxs-lookup"><span data-stu-id="79dd2-123">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="79dd2-124">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="79dd2-124">displayLanguage :String</span></span>

<span data-ttu-id="79dd2-125">Obtiene la configuración local (de idioma) en un formato de etiqueta de idioma RFC 1766 especificado por el usuario para la interfaz de usuario de la aplicación host de Office.</span><span class="sxs-lookup"><span data-stu-id="79dd2-125">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="79dd2-126">El valor `displayLanguage` refleja la configuración de **Mostrar idioma** actual que se ha especificado desde **Archivo > Opciones > Idioma**, en la aplicación host de Office.</span><span class="sxs-lookup"><span data-stu-id="79dd2-126">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="79dd2-127">Tipo:</span><span class="sxs-lookup"><span data-stu-id="79dd2-127">Type:</span></span>

*   <span data-ttu-id="79dd2-128">String</span><span class="sxs-lookup"><span data-stu-id="79dd2-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="79dd2-129">Requisitos</span><span class="sxs-lookup"><span data-stu-id="79dd2-129">Requirements</span></span>

|<span data-ttu-id="79dd2-130">Requirement</span><span class="sxs-lookup"><span data-stu-id="79dd2-130">Requirement</span></span>| <span data-ttu-id="79dd2-131">Valor</span><span class="sxs-lookup"><span data-stu-id="79dd2-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="79dd2-132">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="79dd2-132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79dd2-133">1.0</span><span class="sxs-lookup"><span data-stu-id="79dd2-133">1.0</span></span>|
|[<span data-ttu-id="79dd2-134">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="79dd2-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79dd2-135">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="79dd2-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="79dd2-136">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="79dd2-136">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="79dd2-137">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="79dd2-137">officeTheme :Object</span></span>

<span data-ttu-id="79dd2-138">Proporciona acceso a las propiedades de los colores del tema de Office.</span><span class="sxs-lookup"><span data-stu-id="79dd2-138">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="79dd2-139">Este miembro no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="79dd2-139">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="79dd2-p102">El uso de los colores del tema de Office le permite coordinar la combinación de colores del complemento con el tema actual de Office seleccionado por el usuario mediante **Archivo > Cuenta de Office > Interfaz de usuario Tema de Office**, que se aplica a todas las aplicaciones host de Office. El uso de colores del tema de Office es idóneo para los complementos de correo y panel de tareas.</span><span class="sxs-lookup"><span data-stu-id="79dd2-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="79dd2-142">Tipo:</span><span class="sxs-lookup"><span data-stu-id="79dd2-142">Type:</span></span>

*   <span data-ttu-id="79dd2-143">Objeto</span><span class="sxs-lookup"><span data-stu-id="79dd2-143">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="79dd2-144">Propiedades:</span><span class="sxs-lookup"><span data-stu-id="79dd2-144">Properties:</span></span>

|<span data-ttu-id="79dd2-145">Nombre</span><span class="sxs-lookup"><span data-stu-id="79dd2-145">Name</span></span>| <span data-ttu-id="79dd2-146">Tipo</span><span class="sxs-lookup"><span data-stu-id="79dd2-146">Type</span></span>| <span data-ttu-id="79dd2-147">Descripción</span><span class="sxs-lookup"><span data-stu-id="79dd2-147">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="79dd2-148">String</span><span class="sxs-lookup"><span data-stu-id="79dd2-148">String</span></span>|<span data-ttu-id="79dd2-149">Obtiene el color de fondo del cuerpo del tema de Office como un tríptico de color hexadecimal.</span><span class="sxs-lookup"><span data-stu-id="79dd2-149">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="79dd2-150">String</span><span class="sxs-lookup"><span data-stu-id="79dd2-150">String</span></span>|<span data-ttu-id="79dd2-151">Obtiene el color de primer plano del cuerpo del tema de Office como un tríptico de color hexadecimal.</span><span class="sxs-lookup"><span data-stu-id="79dd2-151">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="79dd2-152">String</span><span class="sxs-lookup"><span data-stu-id="79dd2-152">String</span></span>|<span data-ttu-id="79dd2-153">Obtiene el color de fondo del control del tema de Office como un tríptico de color hexadecimal.</span><span class="sxs-lookup"><span data-stu-id="79dd2-153">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="79dd2-154">String</span><span class="sxs-lookup"><span data-stu-id="79dd2-154">String</span></span>|<span data-ttu-id="79dd2-155">Obtiene el color del control del cuerpo del tema de Office como un tríptico de color hexadecimal.</span><span class="sxs-lookup"><span data-stu-id="79dd2-155">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="79dd2-156">Requisitos</span><span class="sxs-lookup"><span data-stu-id="79dd2-156">Requirements</span></span>

|<span data-ttu-id="79dd2-157">Requirement</span><span class="sxs-lookup"><span data-stu-id="79dd2-157">Requirement</span></span>| <span data-ttu-id="79dd2-158">Valor</span><span class="sxs-lookup"><span data-stu-id="79dd2-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="79dd2-159">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="79dd2-159">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79dd2-160">1.3</span><span class="sxs-lookup"><span data-stu-id="79dd2-160">1.3</span></span>|
|[<span data-ttu-id="79dd2-161">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="79dd2-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79dd2-162">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="79dd2-162">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="79dd2-163">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="79dd2-163">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook15officeroamingsettings"></a><span data-ttu-id="79dd2-164">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_5/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="79dd2-164">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_5/office.RoamingSettings)</span></span>

<span data-ttu-id="79dd2-165">Obtiene un objeto que representa la configuración o el estado personalizado de un complemento de correo que se guardó en el buzón de un usuario.</span><span class="sxs-lookup"><span data-stu-id="79dd2-165">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="79dd2-166">El objeto `RoamingSettings` le permite almacenar y tener acceso a datos para un complemento de correo almacenado en el buzón de un usuario, de forma que esté disponible para ese complemento cuando se ejecute desde cualquier aplicación de cliente host usada para tener acceso a ese buzón.</span><span class="sxs-lookup"><span data-stu-id="79dd2-166">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="79dd2-167">Tipo:</span><span class="sxs-lookup"><span data-stu-id="79dd2-167">Type:</span></span>

*   [<span data-ttu-id="79dd2-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="79dd2-168">RoamingSettings</span></span>](/javascript/api/outlook_1_5/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="79dd2-169">Requisitos</span><span class="sxs-lookup"><span data-stu-id="79dd2-169">Requirements</span></span>

|<span data-ttu-id="79dd2-170">Requirement</span><span class="sxs-lookup"><span data-stu-id="79dd2-170">Requirement</span></span>| <span data-ttu-id="79dd2-171">Valor</span><span class="sxs-lookup"><span data-stu-id="79dd2-171">Value</span></span>|
|---|---|
|[<span data-ttu-id="79dd2-172">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="79dd2-172">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79dd2-173">1.0</span><span class="sxs-lookup"><span data-stu-id="79dd2-173">1.0</span></span>|
|[<span data-ttu-id="79dd2-174">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="79dd2-174">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79dd2-175">Restringido</span><span class="sxs-lookup"><span data-stu-id="79dd2-175">Restricted</span></span>|
|[<span data-ttu-id="79dd2-176">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="79dd2-176">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79dd2-177">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="79dd2-177">Compose or read</span></span>|