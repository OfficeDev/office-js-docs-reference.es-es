
# <a name="item"></a><span data-ttu-id="bf764-101">elemento</span><span class="sxs-lookup"><span data-stu-id="bf764-101">item</span></span>

### <span data-ttu-id="bf764-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="bf764-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="bf764-p102">El espacio de nombres `item` se usa para acceder al mensaje seleccionado actualmente, a una convocatoria de reunión o a una cita. Puede determinar el tipo del `item` mediante la propiedad [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="bf764-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-106">Requirements</span></span>

|<span data-ttu-id="bf764-107">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-107">Requirement</span></span>| <span data-ttu-id="bf764-108">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-109">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-109">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-110">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-110">1.0</span></span>|
|[<span data-ttu-id="bf764-111">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-112">Restringido</span><span class="sxs-lookup"><span data-stu-id="bf764-112">Restricted</span></span>|
|[<span data-ttu-id="bf764-113">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-114">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="bf764-115">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-115">Example</span></span>

<span data-ttu-id="bf764-116">En el siguiente ejemplo de código de JavaScript, se muestra cómo tener acceso a la propiedad `subject` del elemento actual en Outlook.</span><span class="sxs-lookup"><span data-stu-id="bf764-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a><span data-ttu-id="bf764-117">Miembros</span><span class="sxs-lookup"><span data-stu-id="bf764-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="bf764-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="bf764-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="bf764-p103">Obtiene una matriz de datos adjuntos para el elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="bf764-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bf764-121">Ciertos tipos de archivos están bloqueados por Outlook debido a problemas de seguridad y, por tanto, no se devuelven.</span><span class="sxs-lookup"><span data-stu-id="bf764-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="bf764-122">Para obtener más información, vea [datos adjuntos bloqueados en Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="bf764-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-123">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-123">Type:</span></span>

*   <span data-ttu-id="bf764-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="bf764-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-125">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-125">Requirements</span></span>

|<span data-ttu-id="bf764-126">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-126">Requirement</span></span>| <span data-ttu-id="bf764-127">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-128">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-128">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-129">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-129">1.0</span></span>|
|[<span data-ttu-id="bf764-130">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-131">ReadItem</span></span>|
|[<span data-ttu-id="bf764-132">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-133">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-134">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-134">Example</span></span>

<span data-ttu-id="bf764-135">El siguiente código crea una cadena HTML con los detalles de todos los datos adjuntos en el elemento actual.</span><span class="sxs-lookup"><span data-stu-id="bf764-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```JavaScript
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="bf764-136">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bf764-136">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="bf764-137">Obtiene un objeto que proporciona métodos para obtener o actualizar a los destinatarios de la línea de CCO (con copia oculta) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="bf764-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="bf764-138">Solo modo Redacción.</span><span class="sxs-lookup"><span data-stu-id="bf764-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-139">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-139">Type:</span></span>

*   [<span data-ttu-id="bf764-140">Destinatarios</span><span class="sxs-lookup"><span data-stu-id="bf764-140">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="bf764-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-141">Requirements</span></span>

|<span data-ttu-id="bf764-142">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-142">Requirement</span></span>| <span data-ttu-id="bf764-143">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-144">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-145">1.1</span><span class="sxs-lookup"><span data-stu-id="bf764-145">1.1</span></span>|
|[<span data-ttu-id="bf764-146">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-147">ReadItem</span></span>|
|[<span data-ttu-id="bf764-148">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-149">Redacción</span><span class="sxs-lookup"><span data-stu-id="bf764-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-150">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="bf764-151">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="bf764-151">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="bf764-152">Obtiene un objeto que proporciona métodos para manipular el cuerpo de un elemento.</span><span class="sxs-lookup"><span data-stu-id="bf764-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-153">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-153">Type:</span></span>

*   [<span data-ttu-id="bf764-154">Cuerpo</span><span class="sxs-lookup"><span data-stu-id="bf764-154">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="bf764-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-155">Requirements</span></span>

|<span data-ttu-id="bf764-156">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-156">Requirement</span></span>| <span data-ttu-id="bf764-157">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-158">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-159">1.1</span><span class="sxs-lookup"><span data-stu-id="bf764-159">1.1</span></span>|
|[<span data-ttu-id="bf764-160">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-161">ReadItem</span></span>|
|[<span data-ttu-id="bf764-162">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-163">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="bf764-164">cc: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bf764-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="bf764-165">Proporciona acceso a los destinatarios de Cc (con copia) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="bf764-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="bf764-166">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="bf764-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bf764-167">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-167">Read mode</span></span>

<span data-ttu-id="bf764-p107">La propiedad `cc` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **CC** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="bf764-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="bf764-170">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="bf764-170">Compose mode</span></span>

<span data-ttu-id="bf764-171">El `cc` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **Cc** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="bf764-171">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-172">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-172">Type:</span></span>

*   <span data-ttu-id="bf764-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bf764-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-174">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-174">Requirements</span></span>

|<span data-ttu-id="bf764-175">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-175">Requirement</span></span>| <span data-ttu-id="bf764-176">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-177">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-178">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-178">1.0</span></span>|
|[<span data-ttu-id="bf764-179">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-180">ReadItem</span></span>|
|[<span data-ttu-id="bf764-181">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-182">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-183">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="bf764-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="bf764-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="bf764-185">Obtiene un identificador de la conversación de correo electrónico que contiene un mensaje determinado.</span><span class="sxs-lookup"><span data-stu-id="bf764-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="bf764-p108">Puede obtener un entero para esta propiedad si la aplicación de correo se activa en respuestas o formularios de lectura o en formularios de redacción. Si, después, el usuario cambia el asunto del mensaje de respuesta, al enviar dicha respuesta, el Id. de conversación de ese mensaje cambiará y el valor que se obtuvo anteriormente ya no será de aplicación.</span><span class="sxs-lookup"><span data-stu-id="bf764-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="bf764-p109">Obtiene un valor NULL para esta propiedad para un nuevo elemento de un formulario de redacción. Si el usuario establece un asunto y guarda el elemento, la propiedad `conversationId` devolverá un valor.</span><span class="sxs-lookup"><span data-stu-id="bf764-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-190">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-190">Type:</span></span>

*   <span data-ttu-id="bf764-191">String</span><span class="sxs-lookup"><span data-stu-id="bf764-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-192">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-192">Requirements</span></span>

|<span data-ttu-id="bf764-193">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-193">Requirement</span></span>| <span data-ttu-id="bf764-194">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-195">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-195">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-196">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-196">1.0</span></span>|
|[<span data-ttu-id="bf764-197">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-198">ReadItem</span></span>|
|[<span data-ttu-id="bf764-199">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-200">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="bf764-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="bf764-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="bf764-p110">Obtiene la fecha y hora de creación de un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="bf764-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-204">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-204">Type:</span></span>

*   <span data-ttu-id="bf764-205">Fecha</span><span class="sxs-lookup"><span data-stu-id="bf764-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-206">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-206">Requirements</span></span>

|<span data-ttu-id="bf764-207">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-207">Requirement</span></span>| <span data-ttu-id="bf764-208">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-209">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-209">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-210">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-210">1.0</span></span>|
|[<span data-ttu-id="bf764-211">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-212">ReadItem</span></span>|
|[<span data-ttu-id="bf764-213">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-214">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-215">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="bf764-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="bf764-216">dateTimeModified :Date</span></span>

<span data-ttu-id="bf764-p111">Obtiene la fecha y hora en que se modificó por última vez un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="bf764-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bf764-219">Este miembro no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="bf764-219">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-220">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-220">Type:</span></span>

*   <span data-ttu-id="bf764-221">Fecha</span><span class="sxs-lookup"><span data-stu-id="bf764-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-222">Requirements</span></span>

|<span data-ttu-id="bf764-223">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-223">Requirement</span></span>| <span data-ttu-id="bf764-224">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-225">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-225">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-226">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-226">1.0</span></span>|
|[<span data-ttu-id="bf764-227">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-228">ReadItem</span></span>|
|[<span data-ttu-id="bf764-229">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-230">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-231">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="bf764-232">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="bf764-232">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="bf764-233">Obtiene o establece la fecha y la hora de finalización de la cita.</span><span class="sxs-lookup"><span data-stu-id="bf764-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="bf764-p112">La propiedad `end` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) para convertir el valor de la propiedad final a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="bf764-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bf764-236">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-236">Read mode</span></span>

<span data-ttu-id="bf764-237">La propiedad `end` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="bf764-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="bf764-238">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="bf764-238">Compose mode</span></span>

<span data-ttu-id="bf764-239">La propiedad `end` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="bf764-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="bf764-240">Si usa el método [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) para establecer la hora de finalización, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="bf764-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-241">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-241">Type:</span></span>

*   <span data-ttu-id="bf764-242">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="bf764-242">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-243">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-243">Requirements</span></span>

|<span data-ttu-id="bf764-244">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-244">Requirement</span></span>| <span data-ttu-id="bf764-245">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-246">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-246">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-247">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-247">1.0</span></span>|
|[<span data-ttu-id="bf764-248">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-249">ReadItem</span></span>|
|[<span data-ttu-id="bf764-250">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-251">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-252">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-252">Example</span></span>

<span data-ttu-id="bf764-253">En el ejemplo siguiente, se establece la hora de finalización de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="bf764-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="bf764-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="bf764-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="bf764-p113">Obtiene la dirección de correo electrónico del remitente de un mensaje. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="bf764-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="bf764-p114">Las propiedades `from` y [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="bf764-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="bf764-259">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `from` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="bf764-259">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-260">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-260">Type:</span></span>

*   [<span data-ttu-id="bf764-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bf764-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="bf764-262">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-262">Requirements</span></span>

|<span data-ttu-id="bf764-263">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-263">Requirement</span></span>| <span data-ttu-id="bf764-264">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-265">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-266">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-266">1.0</span></span>|
|[<span data-ttu-id="bf764-267">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-268">ReadItem</span></span>|
|[<span data-ttu-id="bf764-269">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-270">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="bf764-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="bf764-271">internetMessageId :String</span></span>

<span data-ttu-id="bf764-p115">Obtiene el identificador de mensaje de Internet de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="bf764-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-274">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-274">Type:</span></span>

*   <span data-ttu-id="bf764-275">String</span><span class="sxs-lookup"><span data-stu-id="bf764-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-276">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-276">Requirements</span></span>

|<span data-ttu-id="bf764-277">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-277">Requirement</span></span>| <span data-ttu-id="bf764-278">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-279">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-279">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-280">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-280">1.0</span></span>|
|[<span data-ttu-id="bf764-281">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-282">ReadItem</span></span>|
|[<span data-ttu-id="bf764-283">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-284">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-285">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="bf764-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="bf764-286">itemClass :String</span></span>

<span data-ttu-id="bf764-p116">Obtiene la clase de elemento Servicios Web Exchange del elemento seleccionado. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="bf764-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="bf764-p117">La propiedad `itemClass` especifica la clase de mensaje del elemento seleccionado. Las siguientes son las clases de mensaje predeterminadas para el elemento de mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="bf764-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="bf764-291">Tipo</span><span class="sxs-lookup"><span data-stu-id="bf764-291">Type</span></span> | <span data-ttu-id="bf764-292">Descripción</span><span class="sxs-lookup"><span data-stu-id="bf764-292">Description</span></span> | <span data-ttu-id="bf764-293">Clase de elemento</span><span class="sxs-lookup"><span data-stu-id="bf764-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="bf764-294">Elementos de cita</span><span class="sxs-lookup"><span data-stu-id="bf764-294">Appointment items</span></span> | <span data-ttu-id="bf764-295">Estos son los elementos de calendario de la clase de elemento `IPM.Appointment` o `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="bf764-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="bf764-296">Elementos de mensaje</span><span class="sxs-lookup"><span data-stu-id="bf764-296">Message items</span></span> | <span data-ttu-id="bf764-297">Entre estos se encuentran los mensajes de correo electrónico que tienen la clase de mensaje predeterminada `IPM.Note` y las convocatorias de reunión, respuestas y cancelaciones que usan `IPM.Schedule.Meeting` como clase de mensaje base.</span><span class="sxs-lookup"><span data-stu-id="bf764-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="bf764-298">Puede crear clases de mensaje personalizadas para ampliar una clase de mensaje predeterminada. Por ejemplo, puede crear una clase de mensaje de cita personalizada `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="bf764-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-299">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-299">Type:</span></span>

*   <span data-ttu-id="bf764-300">String</span><span class="sxs-lookup"><span data-stu-id="bf764-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-301">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-301">Requirements</span></span>

|<span data-ttu-id="bf764-302">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-302">Requirement</span></span>| <span data-ttu-id="bf764-303">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-304">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-304">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-305">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-305">1.0</span></span>|
|[<span data-ttu-id="bf764-306">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-307">ReadItem</span></span>|
|[<span data-ttu-id="bf764-308">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-309">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-310">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="bf764-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="bf764-311">(nullable) itemId :String</span></span>

<span data-ttu-id="bf764-p118">Obtiene el identificador de elemento de Servicios Web Exchange para el elemento actual. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="bf764-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bf764-314">El identificador que devuelve la propiedad `itemId` es el mismo que el identificador de elemento de Servicios Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="bf764-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="bf764-315">El `itemId` (propiedad) no es idéntico al identificador de entrada de Outlook o el identificador de la API de REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="bf764-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="bf764-316">Antes de realizar llamadas a API de REST utilizando este valor, se debe convertir mediante `Office.context.mailbox.convertToRestId`, que está disponible a partir de requisito establece 1.3.</span><span class="sxs-lookup"><span data-stu-id="bf764-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="bf764-317">Para obtener más información, vea [usar las API de REST de Outlook desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="bf764-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-318">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-318">Type:</span></span>

*   <span data-ttu-id="bf764-319">String</span><span class="sxs-lookup"><span data-stu-id="bf764-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-320">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-320">Requirements</span></span>

|<span data-ttu-id="bf764-321">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-321">Requirement</span></span>| <span data-ttu-id="bf764-322">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-323">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-323">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-324">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-324">1.0</span></span>|
|[<span data-ttu-id="bf764-325">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-326">ReadItem</span></span>|
|[<span data-ttu-id="bf764-327">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-328">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-329">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-329">Example</span></span>

<span data-ttu-id="bf764-p120">El siguiente código comprueba la presencia de un identificador de elemento. Si la propiedad `itemId` devuelve `null` o `undefined`, guarda el elemento en el servidor y obtiene el identificador de elemento desde el resultado asincrónico.</span><span class="sxs-lookup"><span data-stu-id="bf764-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="bf764-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="bf764-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="bf764-333">Obtiene el tipo de elemento que representa una instancia.</span><span class="sxs-lookup"><span data-stu-id="bf764-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="bf764-334">La propiedad `itemType` devuelve uno de los valores de enumeración de `ItemType`, lo que indica si la instancia del objeto `item` es un mensaje o una cita.</span><span class="sxs-lookup"><span data-stu-id="bf764-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-335">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-335">Type:</span></span>

*   [<span data-ttu-id="bf764-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="bf764-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="bf764-337">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-337">Requirements</span></span>

|<span data-ttu-id="bf764-338">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-338">Requirement</span></span>| <span data-ttu-id="bf764-339">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-340">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-341">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-341">1.0</span></span>|
|[<span data-ttu-id="bf764-342">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-343">ReadItem</span></span>|
|[<span data-ttu-id="bf764-344">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-345">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-346">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="bf764-347">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="bf764-347">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="bf764-348">Obtiene o establece la ubicación de una cita.</span><span class="sxs-lookup"><span data-stu-id="bf764-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bf764-349">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-349">Read mode</span></span>

<span data-ttu-id="bf764-350">La propiedad `location` devuelve una cadena que contiene la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="bf764-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="bf764-351">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="bf764-351">Compose mode</span></span>

<span data-ttu-id="bf764-352">La propiedad `location` devuelve un objeto `Location` que proporciona métodos para obtener y establecer la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="bf764-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-353">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-353">Type:</span></span>

*   <span data-ttu-id="bf764-354">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="bf764-354">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-355">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-355">Requirements</span></span>

|<span data-ttu-id="bf764-356">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-356">Requirement</span></span>| <span data-ttu-id="bf764-357">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-358">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-358">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-359">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-359">1.0</span></span>|
|[<span data-ttu-id="bf764-360">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-361">ReadItem</span></span>|
|[<span data-ttu-id="bf764-362">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-363">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-364">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="bf764-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="bf764-365">normalizedSubject :String</span></span>

<span data-ttu-id="bf764-p121">Obtiene el asunto de un elemento, con todos los prefijos quitados (incluidos `RE:` y `FWD:`). Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="bf764-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="bf764-p122">La propiedad normalizedSubject obtiene el asunto del elemento, con los prefijos estándar (como `RE:` y `FW:`) agregados por los programas de correo electrónico. Para obtener el asunto del elemento con los prefijos intactos, use la propiedad [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject).</span><span class="sxs-lookup"><span data-stu-id="bf764-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-370">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-370">Type:</span></span>

*   <span data-ttu-id="bf764-371">String</span><span class="sxs-lookup"><span data-stu-id="bf764-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-372">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-372">Requirements</span></span>

|<span data-ttu-id="bf764-373">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-373">Requirement</span></span>| <span data-ttu-id="bf764-374">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-375">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-375">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-376">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-376">1.0</span></span>|
|[<span data-ttu-id="bf764-377">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-378">ReadItem</span></span>|
|[<span data-ttu-id="bf764-379">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-380">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-381">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="bf764-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bf764-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="bf764-383">Proporciona acceso a los asistentes opcionales de un evento.</span><span class="sxs-lookup"><span data-stu-id="bf764-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="bf764-384">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="bf764-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bf764-385">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-385">Read mode</span></span>

<span data-ttu-id="bf764-386">La propiedad `optionalAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente opcional a la reunión.</span><span class="sxs-lookup"><span data-stu-id="bf764-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="bf764-387">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="bf764-387">Compose mode</span></span>

<span data-ttu-id="bf764-388">El `optionalAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes opcionales de una reunión.</span><span class="sxs-lookup"><span data-stu-id="bf764-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-389">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-389">Type:</span></span>

*   <span data-ttu-id="bf764-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bf764-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-391">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-391">Requirements</span></span>

|<span data-ttu-id="bf764-392">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-392">Requirement</span></span>| <span data-ttu-id="bf764-393">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-394">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-394">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-395">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-395">1.0</span></span>|
|[<span data-ttu-id="bf764-396">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-397">ReadItem</span></span>|
|[<span data-ttu-id="bf764-398">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-399">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-400">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="bf764-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="bf764-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="bf764-p124">Obtiene la dirección de correo electrónico del organizador de una reunión especificada. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="bf764-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-404">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-404">Type:</span></span>

*   [<span data-ttu-id="bf764-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bf764-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="bf764-406">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-406">Requirements</span></span>

|<span data-ttu-id="bf764-407">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-407">Requirement</span></span>| <span data-ttu-id="bf764-408">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-409">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-410">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-410">1.0</span></span>|
|[<span data-ttu-id="bf764-411">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-412">ReadItem</span></span>|
|[<span data-ttu-id="bf764-413">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-414">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-415">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="bf764-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bf764-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="bf764-417">Proporciona acceso a los asistentes necesarios de un evento.</span><span class="sxs-lookup"><span data-stu-id="bf764-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="bf764-418">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="bf764-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bf764-419">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-419">Read mode</span></span>

<span data-ttu-id="bf764-420">La propiedad `requiredAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente requerido a la reunión.</span><span class="sxs-lookup"><span data-stu-id="bf764-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="bf764-421">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="bf764-421">Compose mode</span></span>

<span data-ttu-id="bf764-422">El `requiredAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes necesarios de una reunión.</span><span class="sxs-lookup"><span data-stu-id="bf764-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-423">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-423">Type:</span></span>

*   <span data-ttu-id="bf764-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bf764-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-425">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-425">Requirements</span></span>

|<span data-ttu-id="bf764-426">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-426">Requirement</span></span>| <span data-ttu-id="bf764-427">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-428">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-428">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-429">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-429">1.0</span></span>|
|[<span data-ttu-id="bf764-430">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-431">ReadItem</span></span>|
|[<span data-ttu-id="bf764-432">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-433">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-434">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="bf764-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="bf764-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="bf764-p126">Obtiene la dirección de correo electrónico del remitente de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="bf764-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="bf764-p127">Las propiedades [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) y `sender` representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="bf764-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="bf764-440">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `from` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="bf764-440">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-441">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-441">Type:</span></span>

*   [<span data-ttu-id="bf764-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bf764-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="bf764-443">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-443">Requirements</span></span>

|<span data-ttu-id="bf764-444">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-444">Requirement</span></span>| <span data-ttu-id="bf764-445">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-446">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-446">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-447">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-447">1.0</span></span>|
|[<span data-ttu-id="bf764-448">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-449">ReadItem</span></span>|
|[<span data-ttu-id="bf764-450">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-451">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-452">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="bf764-453">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="bf764-453">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="bf764-454">Obtiene o establece la fecha y la hora de inicio de la cita.</span><span class="sxs-lookup"><span data-stu-id="bf764-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="bf764-p128">La propiedad `start` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) para convertir el valor a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="bf764-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bf764-457">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-457">Read mode</span></span>

<span data-ttu-id="bf764-458">La propiedad `start` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="bf764-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="bf764-459">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="bf764-459">Compose mode</span></span>

<span data-ttu-id="bf764-460">La propiedad `start` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="bf764-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="bf764-461">Si usa el método [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) para establecer la hora de inicio, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="bf764-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-462">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-462">Type:</span></span>

*   <span data-ttu-id="bf764-463">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="bf764-463">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-464">Requirements</span></span>

|<span data-ttu-id="bf764-465">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-465">Requirement</span></span>| <span data-ttu-id="bf764-466">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-467">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-467">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-468">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-468">1.0</span></span>|
|[<span data-ttu-id="bf764-469">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-470">ReadItem</span></span>|
|[<span data-ttu-id="bf764-471">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-472">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-473">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-473">Example</span></span>

<span data-ttu-id="bf764-474">En el ejemplo siguiente, se establece la hora de inicio de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="bf764-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="bf764-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="bf764-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="bf764-476">Obtiene o establece la descripción que se muestra en el campo de asunto de un elemento.</span><span class="sxs-lookup"><span data-stu-id="bf764-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="bf764-477">La propiedad `subject` obtiene o establece el asunto completo del elemento, como lo ha enviado el servidor de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="bf764-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bf764-478">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-478">Read mode</span></span>

<span data-ttu-id="bf764-p129">La propiedad `subject` devuelve una cadena. Use la propiedad [`normalizedSubject`](#normalizedsubject-string) para obtener el asunto menos cualquier prefijo inicial como `RE:` y `FW:`.</span><span class="sxs-lookup"><span data-stu-id="bf764-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="bf764-481">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="bf764-481">Compose mode</span></span>

<span data-ttu-id="bf764-482">La propiedad `subject` devuelve un objeto `Subject` que proporciona métodos para obtener y establecer el asunto.</span><span class="sxs-lookup"><span data-stu-id="bf764-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="bf764-483">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-483">Type:</span></span>

*   <span data-ttu-id="bf764-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="bf764-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-485">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-485">Requirements</span></span>

|<span data-ttu-id="bf764-486">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-486">Requirement</span></span>| <span data-ttu-id="bf764-487">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-488">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-488">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-489">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-489">1.0</span></span>|
|[<span data-ttu-id="bf764-490">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-491">ReadItem</span></span>|
|[<span data-ttu-id="bf764-492">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-493">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="bf764-494">para: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bf764-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="bf764-495">Proporciona acceso a los destinatarios en la línea **para** de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="bf764-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="bf764-496">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="bf764-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bf764-497">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-497">Read mode</span></span>

<span data-ttu-id="bf764-p131">La propiedad `to` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **Para** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="bf764-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="bf764-500">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="bf764-500">Compose mode</span></span>

<span data-ttu-id="bf764-501">El `to` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **para** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="bf764-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="bf764-502">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bf764-502">Type:</span></span>

*   <span data-ttu-id="bf764-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bf764-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-504">Requirements</span></span>

|<span data-ttu-id="bf764-505">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-505">Requirement</span></span>| <span data-ttu-id="bf764-506">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-507">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-508">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-508">1.0</span></span>|
|[<span data-ttu-id="bf764-509">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-510">ReadItem</span></span>|
|[<span data-ttu-id="bf764-511">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-512">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-513">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="bf764-514">Métodos</span><span class="sxs-lookup"><span data-stu-id="bf764-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="bf764-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bf764-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="bf764-516">Agrega un archivo a un mensaje o cita como datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="bf764-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="bf764-517">El método `addFileAttachmentAsync` carga el archivo en el URI especificado y lo asocia al elemento en el formulario de redacción.</span><span class="sxs-lookup"><span data-stu-id="bf764-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="bf764-518">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="bf764-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bf764-519">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="bf764-519">Parameters:</span></span>

|<span data-ttu-id="bf764-520">Nombre</span><span class="sxs-lookup"><span data-stu-id="bf764-520">Name</span></span>| <span data-ttu-id="bf764-521">Tipo</span><span class="sxs-lookup"><span data-stu-id="bf764-521">Type</span></span>| <span data-ttu-id="bf764-522">Atributos</span><span class="sxs-lookup"><span data-stu-id="bf764-522">Attributes</span></span>| <span data-ttu-id="bf764-523">Descripción</span><span class="sxs-lookup"><span data-stu-id="bf764-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="bf764-524">String</span><span class="sxs-lookup"><span data-stu-id="bf764-524">String</span></span>||<span data-ttu-id="bf764-p132">El URI que proporciona la ubicación del archivo que se va a adjuntar al mensaje o a la cita. La longitud máxima es de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="bf764-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="bf764-527">String</span><span class="sxs-lookup"><span data-stu-id="bf764-527">String</span></span>||<span data-ttu-id="bf764-p133">El nombre de los datos adjuntos que se muestra mientras estos se cargan. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="bf764-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="bf764-530">Object</span><span class="sxs-lookup"><span data-stu-id="bf764-530">Object</span></span>| <span data-ttu-id="bf764-531">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bf764-531">&lt;optional&gt;</span></span>|<span data-ttu-id="bf764-532">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="bf764-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bf764-533">Objeto</span><span class="sxs-lookup"><span data-stu-id="bf764-533">Object</span></span>| <span data-ttu-id="bf764-534">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bf764-534">&lt;optional&gt;</span></span>|<span data-ttu-id="bf764-535">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="bf764-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="bf764-536">función</span><span class="sxs-lookup"><span data-stu-id="bf764-536">function</span></span>| <span data-ttu-id="bf764-537">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bf764-537">&lt;optional&gt;</span></span>|<span data-ttu-id="bf764-538">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bf764-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="bf764-539">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="bf764-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="bf764-540">Si se produce un error en la carga de los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="bf764-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bf764-541">Errores</span><span class="sxs-lookup"><span data-stu-id="bf764-541">Errors</span></span>

| <span data-ttu-id="bf764-542">Código de error</span><span class="sxs-lookup"><span data-stu-id="bf764-542">Error code</span></span> | <span data-ttu-id="bf764-543">Descripción</span><span class="sxs-lookup"><span data-stu-id="bf764-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="bf764-544">Los datos adjuntos son más grandes de lo permitido.</span><span class="sxs-lookup"><span data-stu-id="bf764-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="bf764-545">Los datos adjuntos tienen una extensión que no está permitida.</span><span class="sxs-lookup"><span data-stu-id="bf764-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="bf764-546">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="bf764-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bf764-547">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-547">Requirements</span></span>

|<span data-ttu-id="bf764-548">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-548">Requirement</span></span>| <span data-ttu-id="bf764-549">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-550">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-550">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-551">1.1</span><span class="sxs-lookup"><span data-stu-id="bf764-551">1.1</span></span>|
|[<span data-ttu-id="bf764-552">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bf764-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="bf764-554">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-555">Redacción</span><span class="sxs-lookup"><span data-stu-id="bf764-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-556">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-556">Example</span></span>

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="bf764-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bf764-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="bf764-558">Agrega un elemento de Exchange (por ejemplo, un mensaje) como datos adjuntos al mensaje o a la cita.</span><span class="sxs-lookup"><span data-stu-id="bf764-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="bf764-p134">El método `addItemAttachmentAsync` asocia el elemento con el identificador especificado de Exchange al elemento en el formulario de redacción. Si se especifica un método de devolución de llamada, se llama al método con un parámetro, `asyncResult`, que contiene el identificador de datos adjuntos o un código que indique cualquier error que se ha producido al adjuntar el elemento. Si es necesario, puede usar el parámetro `options` para pasar información de estado al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="bf764-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="bf764-562">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="bf764-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="bf764-563">Si el complemento de Office se está ejecutando en Outlook Web App, el `addItemAttachmentAsync` método puede adjuntar elementos a los elementos que no sea el elemento que se está editando; Sin embargo, esto no se admite y no se recomienda.</span><span class="sxs-lookup"><span data-stu-id="bf764-563">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bf764-564">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="bf764-564">Parameters:</span></span>

|<span data-ttu-id="bf764-565">Nombre</span><span class="sxs-lookup"><span data-stu-id="bf764-565">Name</span></span>| <span data-ttu-id="bf764-566">Tipo</span><span class="sxs-lookup"><span data-stu-id="bf764-566">Type</span></span>| <span data-ttu-id="bf764-567">Atributos</span><span class="sxs-lookup"><span data-stu-id="bf764-567">Attributes</span></span>| <span data-ttu-id="bf764-568">Descripción</span><span class="sxs-lookup"><span data-stu-id="bf764-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="bf764-569">String</span><span class="sxs-lookup"><span data-stu-id="bf764-569">String</span></span>||<span data-ttu-id="bf764-p135">El identificador de Exchange del elemento que debe adjuntarse. La longitud máxima es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="bf764-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="bf764-572">String</span><span class="sxs-lookup"><span data-stu-id="bf764-572">String</span></span>||<span data-ttu-id="bf764-p136">El asunto del elemento que debe adjuntarse. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="bf764-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="bf764-575">Object</span><span class="sxs-lookup"><span data-stu-id="bf764-575">Object</span></span>| <span data-ttu-id="bf764-576">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bf764-576">&lt;optional&gt;</span></span>|<span data-ttu-id="bf764-577">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="bf764-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bf764-578">Objeto</span><span class="sxs-lookup"><span data-stu-id="bf764-578">Object</span></span>| <span data-ttu-id="bf764-579">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bf764-579">&lt;optional&gt;</span></span>|<span data-ttu-id="bf764-580">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="bf764-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="bf764-581">función</span><span class="sxs-lookup"><span data-stu-id="bf764-581">function</span></span>| <span data-ttu-id="bf764-582">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bf764-582">&lt;optional&gt;</span></span>|<span data-ttu-id="bf764-583">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bf764-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="bf764-584">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="bf764-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="bf764-585">Si se produce un error al agregar los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="bf764-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bf764-586">Errores</span><span class="sxs-lookup"><span data-stu-id="bf764-586">Errors</span></span>

| <span data-ttu-id="bf764-587">Código de error</span><span class="sxs-lookup"><span data-stu-id="bf764-587">Error code</span></span> | <span data-ttu-id="bf764-588">Descripción</span><span class="sxs-lookup"><span data-stu-id="bf764-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="bf764-589">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="bf764-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bf764-590">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-590">Requirements</span></span>

|<span data-ttu-id="bf764-591">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-591">Requirement</span></span>| <span data-ttu-id="bf764-592">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-593">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-593">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-594">1.1</span><span class="sxs-lookup"><span data-stu-id="bf764-594">1.1</span></span>|
|[<span data-ttu-id="bf764-595">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bf764-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="bf764-597">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-598">Redacción</span><span class="sxs-lookup"><span data-stu-id="bf764-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-599">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-599">Example</span></span>

<span data-ttu-id="bf764-600">En el ejemplo siguiente, se agrega un elemento existente de Outlook como un archivo adjunto con el nombre `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="bf764-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="bf764-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="bf764-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="bf764-602">Muestra un formulario de respuesta que incluye el remitente y todos los destinatarios del mensaje seleccionado o el organizador y todos los asistentes de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="bf764-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="bf764-603">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="bf764-603">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bf764-604">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="bf764-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="bf764-605">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyAllForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="bf764-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="bf764-606">La capacidad de incluir datos adjuntos en la llamada a `displayReplyAllForm` no se admite en el conjunto de requisitos 1.1.</span><span class="sxs-lookup"><span data-stu-id="bf764-606">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="bf764-607">Se ha agregado la compatibilidad con datos adjuntos a `displayReplyAllForm` en el conjunto de requisitos 1.2 y versiones superiores.</span><span class="sxs-lookup"><span data-stu-id="bf764-607">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bf764-608">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="bf764-608">Parameters:</span></span>

|<span data-ttu-id="bf764-609">Nombre</span><span class="sxs-lookup"><span data-stu-id="bf764-609">Name</span></span>| <span data-ttu-id="bf764-610">Tipo</span><span class="sxs-lookup"><span data-stu-id="bf764-610">Type</span></span>| <span data-ttu-id="bf764-611">Descripción</span><span class="sxs-lookup"><span data-stu-id="bf764-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="bf764-612">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="bf764-612">String &#124; Object</span></span>| |<span data-ttu-id="bf764-p138">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="bf764-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="bf764-615">**O**</span><span class="sxs-lookup"><span data-stu-id="bf764-615">**OR**</span></span><br/><span data-ttu-id="bf764-p139">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="bf764-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="bf764-618">String</span><span class="sxs-lookup"><span data-stu-id="bf764-618">String</span></span> | <span data-ttu-id="bf764-619">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bf764-619">&lt;optional&gt;</span></span> | <span data-ttu-id="bf764-p140">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="bf764-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="bf764-622">función</span><span class="sxs-lookup"><span data-stu-id="bf764-622">function</span></span> | <span data-ttu-id="bf764-623">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bf764-623">&lt;optional&gt;</span></span> | <span data-ttu-id="bf764-624">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bf764-624">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bf764-625">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-625">Requirements</span></span>

|<span data-ttu-id="bf764-626">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-626">Requirement</span></span>| <span data-ttu-id="bf764-627">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-628">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-628">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-629">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-629">1.0</span></span>|
|[<span data-ttu-id="bf764-630">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-631">ReadItem</span></span>|
|[<span data-ttu-id="bf764-632">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-633">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-633">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="bf764-634">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="bf764-634">Examples</span></span>

<span data-ttu-id="bf764-635">El código siguiente pasa una cadena a la función `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="bf764-635">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="bf764-636">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="bf764-636">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="bf764-637">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="bf764-637">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="bf764-638">Responder con un cuerpo y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="bf764-638">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="bf764-639">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="bf764-639">displayReplyForm(formData)</span></span>

<span data-ttu-id="bf764-640">Muestra un formulario de respuesta que incluye solo el remitente del mensaje seleccionado o el organizador de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="bf764-640">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="bf764-641">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="bf764-641">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bf764-642">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="bf764-642">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="bf764-643">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="bf764-643">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="bf764-644">La capacidad de incluir datos adjuntos en la llamada a `displayReplyForm` no se admite en el conjunto de requisitos 1.1.</span><span class="sxs-lookup"><span data-stu-id="bf764-644">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="bf764-645">Se ha agregado la compatibilidad con datos adjuntos a `displayReplyForm` en el conjunto de requisitos 1.2 y versiones superiores.</span><span class="sxs-lookup"><span data-stu-id="bf764-645">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bf764-646">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="bf764-646">Parameters:</span></span>

|<span data-ttu-id="bf764-647">Nombre</span><span class="sxs-lookup"><span data-stu-id="bf764-647">Name</span></span>| <span data-ttu-id="bf764-648">Tipo</span><span class="sxs-lookup"><span data-stu-id="bf764-648">Type</span></span>| <span data-ttu-id="bf764-649">Descripción</span><span class="sxs-lookup"><span data-stu-id="bf764-649">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="bf764-650">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="bf764-650">String &#124; Object</span></span>| | <span data-ttu-id="bf764-p142">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="bf764-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="bf764-653">**O**</span><span class="sxs-lookup"><span data-stu-id="bf764-653">**OR**</span></span><br/><span data-ttu-id="bf764-p143">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="bf764-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="bf764-656">String</span><span class="sxs-lookup"><span data-stu-id="bf764-656">String</span></span> | <span data-ttu-id="bf764-657">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bf764-657">&lt;optional&gt;</span></span> | <span data-ttu-id="bf764-p144">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="bf764-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="bf764-660">función</span><span class="sxs-lookup"><span data-stu-id="bf764-660">function</span></span> | <span data-ttu-id="bf764-661">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bf764-661">&lt;optional&gt;</span></span> | <span data-ttu-id="bf764-662">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bf764-662">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bf764-663">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-663">Requirements</span></span>

|<span data-ttu-id="bf764-664">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-664">Requirement</span></span>| <span data-ttu-id="bf764-665">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-665">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-666">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-666">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-667">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-667">1.0</span></span>|
|[<span data-ttu-id="bf764-668">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-668">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-669">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-669">ReadItem</span></span>|
|[<span data-ttu-id="bf764-670">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-670">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-671">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-671">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="bf764-672">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="bf764-672">Examples</span></span>

<span data-ttu-id="bf764-673">El código siguiente pasa una cadena a la función `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="bf764-673">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="bf764-674">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="bf764-674">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="bf764-675">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="bf764-675">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="bf764-676">Responder con un cuerpo y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="bf764-676">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="bf764-677">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="bf764-677">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="bf764-678">Obtiene las entidades que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="bf764-678">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="bf764-679">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="bf764-679">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-680">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-680">Requirements</span></span>

|<span data-ttu-id="bf764-681">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-681">Requirement</span></span>| <span data-ttu-id="bf764-682">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-682">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-683">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-683">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-684">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-684">1.0</span></span>|
|[<span data-ttu-id="bf764-685">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-685">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-686">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-686">ReadItem</span></span>|
|[<span data-ttu-id="bf764-687">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-687">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-688">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-688">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bf764-689">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="bf764-689">Returns:</span></span>

<span data-ttu-id="bf764-690">Tipo: [Entidades](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="bf764-690">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="bf764-691">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-691">Example</span></span>

<span data-ttu-id="bf764-692">En el ejemplo siguiente se obtiene acceso a las entidades de los contactos en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="bf764-692">The following example accesses the contacts entities in the current item's body.</span></span>

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="bf764-693">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="bf764-693">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="bf764-694">Obtiene una matriz de todas las entidades del tipo de entidad especificado que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="bf764-694">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="bf764-695">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="bf764-695">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bf764-696">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="bf764-696">Parameters:</span></span>

|<span data-ttu-id="bf764-697">Nombre</span><span class="sxs-lookup"><span data-stu-id="bf764-697">Name</span></span>| <span data-ttu-id="bf764-698">Tipo</span><span class="sxs-lookup"><span data-stu-id="bf764-698">Type</span></span>| <span data-ttu-id="bf764-699">Descripción</span><span class="sxs-lookup"><span data-stu-id="bf764-699">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="bf764-700">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="bf764-700">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="bf764-701">Uno de los valores de enumeración de EntityType.</span><span class="sxs-lookup"><span data-stu-id="bf764-701">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bf764-702">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-702">Requirements</span></span>

|<span data-ttu-id="bf764-703">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-703">Requirement</span></span>| <span data-ttu-id="bf764-704">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-704">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-705">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-705">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-706">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-706">1.0</span></span>|
|[<span data-ttu-id="bf764-707">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-707">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-708">Restringido</span><span class="sxs-lookup"><span data-stu-id="bf764-708">Restricted</span></span>|
|[<span data-ttu-id="bf764-709">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-709">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-710">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-710">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bf764-711">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="bf764-711">Returns:</span></span>

<span data-ttu-id="bf764-712">Si el valor que se pasa a `entityType` no es un miembro válido de la enumeración de `EntityType`, el método devuelve un valor NULL.</span><span class="sxs-lookup"><span data-stu-id="bf764-712">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="bf764-713">Si no existen entidades del tipo especificado están presentes en el cuerpo del elemento, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="bf764-713">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="bf764-714">De otro modo, el tipo de los objetos en la matriz devuelta depende del tipo de entidad solicitada en el parámetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="bf764-714">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="bf764-715">Aunque el nivel de permiso mínimo para usar este método es **Restricted**, algunos tipos de entidad requieren **ReadItem** para tener acceso, como se especifica en la tabla siguiente.</span><span class="sxs-lookup"><span data-stu-id="bf764-715">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="bf764-716">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="bf764-716">Value of `entityType`</span></span> | <span data-ttu-id="bf764-717">Tipo de objetos de la matriz devuelta</span><span class="sxs-lookup"><span data-stu-id="bf764-717">Type of objects in returned array</span></span> | <span data-ttu-id="bf764-718">Nivel de permiso requerido</span><span class="sxs-lookup"><span data-stu-id="bf764-718">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="bf764-719">Cadena</span><span class="sxs-lookup"><span data-stu-id="bf764-719">String</span></span> | <span data-ttu-id="bf764-720">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="bf764-720">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="bf764-721">Contacto</span><span class="sxs-lookup"><span data-stu-id="bf764-721">Contact</span></span> | <span data-ttu-id="bf764-722">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bf764-722">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="bf764-723">Cadena</span><span class="sxs-lookup"><span data-stu-id="bf764-723">String</span></span> | <span data-ttu-id="bf764-724">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bf764-724">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="bf764-725">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="bf764-725">MeetingSuggestion</span></span> | <span data-ttu-id="bf764-726">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bf764-726">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="bf764-727">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="bf764-727">PhoneNumber</span></span> | <span data-ttu-id="bf764-728">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="bf764-728">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="bf764-729">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="bf764-729">TaskSuggestion</span></span> | <span data-ttu-id="bf764-730">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bf764-730">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="bf764-731">Cadena</span><span class="sxs-lookup"><span data-stu-id="bf764-731">String</span></span> | <span data-ttu-id="bf764-732">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="bf764-732">**Restricted**</span></span> |

<span data-ttu-id="bf764-733">Tipo:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="bf764-733">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="bf764-734">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-734">Example</span></span>

<span data-ttu-id="bf764-735">En el ejemplo siguiente se muestra cómo obtener acceso a una matriz de cadenas que representan direcciones postales en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="bf764-735">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="bf764-736">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="bf764-736">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="bf764-737">Devuelve entidades conocidas en el elemento seleccionado que pasan el filtro con nombre definido en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="bf764-737">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="bf764-738">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="bf764-738">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bf764-739">El método `getFilteredEntitiesByName` devuelve las entidades que coinciden con la expresión regular definida en el elemento de regla [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) del archivo XML de manifiesto con el valor de elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="bf764-739">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bf764-740">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="bf764-740">Parameters:</span></span>

|<span data-ttu-id="bf764-741">Nombre</span><span class="sxs-lookup"><span data-stu-id="bf764-741">Name</span></span>| <span data-ttu-id="bf764-742">Tipo</span><span class="sxs-lookup"><span data-stu-id="bf764-742">Type</span></span>| <span data-ttu-id="bf764-743">Descripción</span><span class="sxs-lookup"><span data-stu-id="bf764-743">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="bf764-744">String</span><span class="sxs-lookup"><span data-stu-id="bf764-744">String</span></span>|<span data-ttu-id="bf764-745">El nombre del elemento de regla `ItemHasKnownEntity` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="bf764-745">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bf764-746">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-746">Requirements</span></span>

|<span data-ttu-id="bf764-747">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-747">Requirement</span></span>| <span data-ttu-id="bf764-748">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-748">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-749">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-749">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-750">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-750">1.0</span></span>|
|[<span data-ttu-id="bf764-751">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-751">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-752">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-752">ReadItem</span></span>|
|[<span data-ttu-id="bf764-753">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-753">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-754">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-754">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bf764-755">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="bf764-755">Returns:</span></span>

<span data-ttu-id="bf764-p146">Si no existe ningún elemento `ItemHasKnownEntity` en el manifiesto con un valor de elemento `FilterName` que coincida con el parámetro `name`, el método devuelve `null`. Si el parámetro `name` no coincide con un elemento `ItemHasKnownEntity` del manifiesto, pero no existen entidades en el elemento actual que coincidan, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="bf764-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="bf764-758">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="bf764-758">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="bf764-759">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="bf764-759">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="bf764-760">Devuelve valores de cadena en el elemento seleccionado que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="bf764-760">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="bf764-761">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="bf764-761">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bf764-p147">El método `getRegExMatches` devuelve las cadenas que coinciden con la expresión regular definida en cada elemento de regla `ItemHasRegularExpressionMatch` o `ItemHasKnownEntity` del archivo XML de manifiesto. Para una regla `ItemHasRegularExpressionMatch`, tiene que darse una cadena coincidente en la propiedad del elemento especificado por la regla. El tipo simple `PropertyName` define las propiedades admitidas.</span><span class="sxs-lookup"><span data-stu-id="bf764-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="bf764-765">Por ejemplo, considere un manifiesto de complemento que tenga el siguiente elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="bf764-765">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="bf764-766">El objeto devuelto desde `getRegExMatches` tendría dos propiedades: `fruits` y `veggies`.</span><span class="sxs-lookup"><span data-stu-id="bf764-766">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="bf764-p148">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="bf764-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf764-769">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-769">Requirements</span></span>

|<span data-ttu-id="bf764-770">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-770">Requirement</span></span>| <span data-ttu-id="bf764-771">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-772">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-772">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-773">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-773">1.0</span></span>|
|[<span data-ttu-id="bf764-774">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-774">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-775">ReadItem</span></span>|
|[<span data-ttu-id="bf764-776">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-776">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-777">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-777">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bf764-778">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="bf764-778">Returns:</span></span>

<span data-ttu-id="bf764-p149">Un objeto que contiene matrices de cadenas que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. El nombre de cada matriz es igual al valor correspondiente del atributo `RegExName` de la regla `ItemHasRegularExpressionMatch` coincidente o al atributo `FilterName` de la regla `ItemHasKnownEntity` coincidente.</span><span class="sxs-lookup"><span data-stu-id="bf764-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="bf764-781">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="bf764-781">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="bf764-782">Object</span><span class="sxs-lookup"><span data-stu-id="bf764-782">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="bf764-783">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-783">Example</span></span>

<span data-ttu-id="bf764-784">En el siguiente ejemplo, se muestra cómo tener acceso a la matriz de coincidencias de los <rule>elementos de expresión regular `fruits`y `veggies`, que se especifican en el manifiesto.</rule></span><span class="sxs-lookup"><span data-stu-id="bf764-784">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="bf764-785">→ de getregexmatchesbyname (Name) (nullable) {matriz. < cadena >}</span><span class="sxs-lookup"><span data-stu-id="bf764-785">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="bf764-786">Devuelve valores de cadena en el elemento seleccionado que coinciden con la expresión regular de nombre definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="bf764-786">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="bf764-787">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="bf764-787">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bf764-788">El método `getRegExMatchesByName` devuelve las cadenas que coinciden con la expresión regular definida en el elemento de regla `ItemHasRegularExpressionMatch` del archivo XML de manifiesto con el valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="bf764-788">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="bf764-p150">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="bf764-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bf764-791">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="bf764-791">Parameters:</span></span>

|<span data-ttu-id="bf764-792">Nombre</span><span class="sxs-lookup"><span data-stu-id="bf764-792">Name</span></span>| <span data-ttu-id="bf764-793">Tipo</span><span class="sxs-lookup"><span data-stu-id="bf764-793">Type</span></span>| <span data-ttu-id="bf764-794">Descripción</span><span class="sxs-lookup"><span data-stu-id="bf764-794">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="bf764-795">String</span><span class="sxs-lookup"><span data-stu-id="bf764-795">String</span></span>|<span data-ttu-id="bf764-796">El nombre del elemento de regla `ItemHasRegularExpressionMatch` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="bf764-796">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bf764-797">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-797">Requirements</span></span>

|<span data-ttu-id="bf764-798">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-798">Requirement</span></span>| <span data-ttu-id="bf764-799">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-799">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-800">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-800">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-801">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-801">1.0</span></span>|
|[<span data-ttu-id="bf764-802">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-802">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-803">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-803">ReadItem</span></span>|
|[<span data-ttu-id="bf764-804">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-804">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-805">Lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-805">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bf764-806">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="bf764-806">Returns:</span></span>

<span data-ttu-id="bf764-807">Una matriz que contiene las cadenas que coinciden con la expresión regular definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="bf764-807">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="bf764-808">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="bf764-808">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="bf764-809">< Cadena > de matriz.</span><span class="sxs-lookup"><span data-stu-id="bf764-809">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="bf764-810">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-810">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="bf764-811">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="bf764-811">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="bf764-812">Carga de forma asincrónica las propiedades personalizadas de este complemento en el elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="bf764-812">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="bf764-p151">Las propiedades personalizadas se almacenan como pares de clave/valor según la aplicación y el elemento. Este método devuelve un objeto `CustomProperties` en la devolución de llamada, que proporciona métodos para tener acceso a las propiedades personalizadas específicas del elemento y el complemento actual. Las propiedades personalizadas no están cifradas en el elemento, por lo que no debería usarse como almacenamiento seguro.</span><span class="sxs-lookup"><span data-stu-id="bf764-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bf764-816">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="bf764-816">Parameters:</span></span>

|<span data-ttu-id="bf764-817">Nombre</span><span class="sxs-lookup"><span data-stu-id="bf764-817">Name</span></span>| <span data-ttu-id="bf764-818">Tipo</span><span class="sxs-lookup"><span data-stu-id="bf764-818">Type</span></span>| <span data-ttu-id="bf764-819">Atributos</span><span class="sxs-lookup"><span data-stu-id="bf764-819">Attributes</span></span>| <span data-ttu-id="bf764-820">Descripción</span><span class="sxs-lookup"><span data-stu-id="bf764-820">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="bf764-821">función</span><span class="sxs-lookup"><span data-stu-id="bf764-821">function</span></span>||<span data-ttu-id="bf764-822">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bf764-822">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bf764-823">Las propiedades personalizadas se proporcionan como un objeto [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="bf764-823">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="bf764-824">Este objeto se puede utilizar para obtener, establecer y elimina las propiedades personalizadas del elemento y guarde los cambios en la propiedad personalizada que se vuelve a establecer en el servidor.</span><span class="sxs-lookup"><span data-stu-id="bf764-824">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="bf764-825">Objeto</span><span class="sxs-lookup"><span data-stu-id="bf764-825">Object</span></span>| <span data-ttu-id="bf764-826">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bf764-826">&lt;optional&gt;</span></span>|<span data-ttu-id="bf764-827">Los desarrolladores pueden proporcionar cualquier objeto que desean tener acceso en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="bf764-827">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="bf764-828">Se puede tener acceso a este objeto por el `asyncResult.asyncContext` (propiedad) en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="bf764-828">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bf764-829">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-829">Requirements</span></span>

|<span data-ttu-id="bf764-830">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-830">Requirement</span></span>| <span data-ttu-id="bf764-831">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-831">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-832">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-832">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-833">1.0</span><span class="sxs-lookup"><span data-stu-id="bf764-833">1.0</span></span>|
|[<span data-ttu-id="bf764-834">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-834">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-835">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bf764-835">ReadItem</span></span>|
|[<span data-ttu-id="bf764-836">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-836">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-837">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="bf764-837">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-838">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-838">Example</span></span>

<span data-ttu-id="bf764-p154">En el siguiente ejemplo de código se muestra cómo usar el método `loadCustomPropertiesAsync` para cargar de forma asincrónica propiedades personalizadas que son específicas del elemento actual. En el ejemplo también se muestra cómo usar el método `CustomProperties.saveAsync` para guardar estas propiedades en el servidor. Después de cargar las propiedades personalizadas, el ejemplo de código usa el método `CustomProperties.get` para leer la propiedad personalizada `myProp`, el método `CustomProperties.set` para escribir en la propiedad personalizada `otherProp` y finalmente se llama al método `saveAsync` para guardar las propiedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="bf764-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```JavaScript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="bf764-842">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bf764-842">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="bf764-843">Quita los datos adjuntos de un mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="bf764-843">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="bf764-p155">El método `removeAttachmentAsync` quita del elemento los datos adjuntos con el identificador especificado. Como práctica recomendada, debe usar el identificador de datos adjuntos para quitar datos adjuntos solo si la misma aplicación de correo ha agregado los datos adjuntos en la misma sesión. En Outlook Web App y OWA para los dispositivos, el identificador de datos adjuntos es válido solo en la misma sesión. Una sesión finaliza cuando el usuario cierra la aplicación, o si el usuario comienza a redactar en un formulario en línea y posteriormente se extrae el formulario en línea para continuar en una ventana independiente.</span><span class="sxs-lookup"><span data-stu-id="bf764-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bf764-848">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="bf764-848">Parameters:</span></span>

|<span data-ttu-id="bf764-849">Nombre</span><span class="sxs-lookup"><span data-stu-id="bf764-849">Name</span></span>| <span data-ttu-id="bf764-850">Tipo</span><span class="sxs-lookup"><span data-stu-id="bf764-850">Type</span></span>| <span data-ttu-id="bf764-851">Atributos</span><span class="sxs-lookup"><span data-stu-id="bf764-851">Attributes</span></span>| <span data-ttu-id="bf764-852">Descripción</span><span class="sxs-lookup"><span data-stu-id="bf764-852">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="bf764-853">String</span><span class="sxs-lookup"><span data-stu-id="bf764-853">String</span></span>||<span data-ttu-id="bf764-p156">El identificador de los datos adjuntos que se deben quitar. La longitud máxima de la cadena es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="bf764-p156">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="bf764-856">Object</span><span class="sxs-lookup"><span data-stu-id="bf764-856">Object</span></span>| <span data-ttu-id="bf764-857">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bf764-857">&lt;optional&gt;</span></span>|<span data-ttu-id="bf764-858">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="bf764-858">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bf764-859">Objeto</span><span class="sxs-lookup"><span data-stu-id="bf764-859">Object</span></span>| <span data-ttu-id="bf764-860">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bf764-860">&lt;optional&gt;</span></span>|<span data-ttu-id="bf764-861">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="bf764-861">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="bf764-862">función</span><span class="sxs-lookup"><span data-stu-id="bf764-862">function</span></span>| <span data-ttu-id="bf764-863">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bf764-863">&lt;optional&gt;</span></span>|<span data-ttu-id="bf764-864">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bf764-864">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="bf764-865">Si se produce un error en la eliminación de los datos adjuntos, la propiedad `asyncResult.error` contendrá un código de error con el motivo del error.</span><span class="sxs-lookup"><span data-stu-id="bf764-865">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bf764-866">Errores</span><span class="sxs-lookup"><span data-stu-id="bf764-866">Errors</span></span>

| <span data-ttu-id="bf764-867">Código de error</span><span class="sxs-lookup"><span data-stu-id="bf764-867">Error code</span></span> | <span data-ttu-id="bf764-868">Descripción</span><span class="sxs-lookup"><span data-stu-id="bf764-868">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="bf764-869">El identificador de datos adjuntos no existe.</span><span class="sxs-lookup"><span data-stu-id="bf764-869">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bf764-870">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bf764-870">Requirements</span></span>

|<span data-ttu-id="bf764-871">Requirement</span><span class="sxs-lookup"><span data-stu-id="bf764-871">Requirement</span></span>| <span data-ttu-id="bf764-872">Valor</span><span class="sxs-lookup"><span data-stu-id="bf764-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf764-873">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="bf764-873">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf764-874">1.1</span><span class="sxs-lookup"><span data-stu-id="bf764-874">1.1</span></span>|
|[<span data-ttu-id="bf764-875">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="bf764-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bf764-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bf764-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="bf764-877">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="bf764-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bf764-878">Redacción</span><span class="sxs-lookup"><span data-stu-id="bf764-878">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bf764-879">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="bf764-879">Example</span></span>

<span data-ttu-id="bf764-880">El código siguiente quita un archivo adjunto con un identificador de "0".</span><span class="sxs-lookup"><span data-stu-id="bf764-880">The following code removes an attachment with an identifier of '0'.</span></span>

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```