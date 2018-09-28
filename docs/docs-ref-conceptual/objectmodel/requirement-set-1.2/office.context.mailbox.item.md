
# <a name="item"></a><span data-ttu-id="2805a-101">elemento</span><span class="sxs-lookup"><span data-stu-id="2805a-101">item</span></span>

### <span data-ttu-id="2805a-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="2805a-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="2805a-p102">El espacio de nombres `item` se usa para acceder al mensaje seleccionado actualmente, a una convocatoria de reunión o a una cita. Puede determinar el tipo del `item` mediante la propiedad [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="2805a-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-106">Requirements</span></span>

|<span data-ttu-id="2805a-107">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-107">Requirement</span></span>| <span data-ttu-id="2805a-108">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-109">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-109">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-110">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-110">1.0</span></span>|
|[<span data-ttu-id="2805a-111">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-112">Restringido</span><span class="sxs-lookup"><span data-stu-id="2805a-112">Restricted</span></span>|
|[<span data-ttu-id="2805a-113">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-114">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="2805a-115">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-115">Example</span></span>

<span data-ttu-id="2805a-116">En el siguiente ejemplo de código de JavaScript, se muestra cómo tener acceso a la propiedad `subject` del elemento actual en Outlook.</span><span class="sxs-lookup"><span data-stu-id="2805a-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="2805a-117">Miembros</span><span class="sxs-lookup"><span data-stu-id="2805a-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="2805a-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="2805a-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="2805a-p103">Obtiene una matriz de datos adjuntos para el elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="2805a-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2805a-121">Ciertos tipos de archivos están bloqueados por Outlook debido a problemas de seguridad y, por tanto, no se devuelven.</span><span class="sxs-lookup"><span data-stu-id="2805a-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="2805a-122">Para obtener más información, vea [datos adjuntos bloqueados en Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="2805a-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-123">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-123">Type:</span></span>

*   <span data-ttu-id="2805a-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="2805a-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-125">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-125">Requirements</span></span>

|<span data-ttu-id="2805a-126">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-126">Requirement</span></span>| <span data-ttu-id="2805a-127">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-128">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-128">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-129">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-129">1.0</span></span>|
|[<span data-ttu-id="2805a-130">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-131">ReadItem</span></span>|
|[<span data-ttu-id="2805a-132">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-133">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-134">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-134">Example</span></span>

<span data-ttu-id="2805a-135">El siguiente código crea una cadena HTML con los detalles de todos los datos adjuntos en el elemento actual.</span><span class="sxs-lookup"><span data-stu-id="2805a-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="2805a-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2805a-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="2805a-137">Obtiene un objeto que proporciona métodos para obtener o actualizar a los destinatarios de la línea de CCO (con copia oculta) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="2805a-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="2805a-138">Solo modo Redacción.</span><span class="sxs-lookup"><span data-stu-id="2805a-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-139">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-139">Type:</span></span>

*   [<span data-ttu-id="2805a-140">Destinatarios</span><span class="sxs-lookup"><span data-stu-id="2805a-140">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="2805a-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-141">Requirements</span></span>

|<span data-ttu-id="2805a-142">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-142">Requirement</span></span>| <span data-ttu-id="2805a-143">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-144">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-145">1.1</span><span class="sxs-lookup"><span data-stu-id="2805a-145">1.1</span></span>|
|[<span data-ttu-id="2805a-146">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-147">ReadItem</span></span>|
|[<span data-ttu-id="2805a-148">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-149">Redacción</span><span class="sxs-lookup"><span data-stu-id="2805a-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-150">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="2805a-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="2805a-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="2805a-152">Obtiene un objeto que proporciona métodos para manipular el cuerpo de un elemento.</span><span class="sxs-lookup"><span data-stu-id="2805a-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-153">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-153">Type:</span></span>

*   [<span data-ttu-id="2805a-154">Cuerpo</span><span class="sxs-lookup"><span data-stu-id="2805a-154">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="2805a-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-155">Requirements</span></span>

|<span data-ttu-id="2805a-156">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-156">Requirement</span></span>| <span data-ttu-id="2805a-157">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-158">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-159">1.1</span><span class="sxs-lookup"><span data-stu-id="2805a-159">1.1</span></span>|
|[<span data-ttu-id="2805a-160">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-161">ReadItem</span></span>|
|[<span data-ttu-id="2805a-162">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-163">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="2805a-164">cc: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2805a-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="2805a-165">Proporciona acceso a los destinatarios de Cc (con copia) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="2805a-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="2805a-166">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="2805a-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2805a-167">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-167">Read mode</span></span>

<span data-ttu-id="2805a-p107">La propiedad `cc` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **CC** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="2805a-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2805a-170">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="2805a-170">Compose mode</span></span>

<span data-ttu-id="2805a-171">El `cc` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **Cc** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="2805a-171">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-172">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-172">Type:</span></span>

*   <span data-ttu-id="2805a-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2805a-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-174">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-174">Requirements</span></span>

|<span data-ttu-id="2805a-175">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-175">Requirement</span></span>| <span data-ttu-id="2805a-176">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-177">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-178">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-178">1.0</span></span>|
|[<span data-ttu-id="2805a-179">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-180">ReadItem</span></span>|
|[<span data-ttu-id="2805a-181">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-182">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-183">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="2805a-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="2805a-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="2805a-185">Obtiene un identificador de la conversación de correo electrónico que contiene un mensaje determinado.</span><span class="sxs-lookup"><span data-stu-id="2805a-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="2805a-p108">Puede obtener un entero para esta propiedad si la aplicación de correo se activa en respuestas o formularios de lectura o en formularios de redacción. Si, después, el usuario cambia el asunto del mensaje de respuesta, al enviar dicha respuesta, el Id. de conversación de ese mensaje cambiará y el valor que se obtuvo anteriormente ya no será de aplicación.</span><span class="sxs-lookup"><span data-stu-id="2805a-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="2805a-p109">Obtiene un valor NULL para esta propiedad para un nuevo elemento de un formulario de redacción. Si el usuario establece un asunto y guarda el elemento, la propiedad `conversationId` devolverá un valor.</span><span class="sxs-lookup"><span data-stu-id="2805a-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-190">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-190">Type:</span></span>

*   <span data-ttu-id="2805a-191">String</span><span class="sxs-lookup"><span data-stu-id="2805a-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-192">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-192">Requirements</span></span>

|<span data-ttu-id="2805a-193">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-193">Requirement</span></span>| <span data-ttu-id="2805a-194">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-195">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-195">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-196">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-196">1.0</span></span>|
|[<span data-ttu-id="2805a-197">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-198">ReadItem</span></span>|
|[<span data-ttu-id="2805a-199">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-200">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="2805a-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="2805a-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="2805a-p110">Obtiene la fecha y hora de creación de un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="2805a-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-204">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-204">Type:</span></span>

*   <span data-ttu-id="2805a-205">Fecha</span><span class="sxs-lookup"><span data-stu-id="2805a-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-206">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-206">Requirements</span></span>

|<span data-ttu-id="2805a-207">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-207">Requirement</span></span>| <span data-ttu-id="2805a-208">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-209">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-209">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-210">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-210">1.0</span></span>|
|[<span data-ttu-id="2805a-211">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-212">ReadItem</span></span>|
|[<span data-ttu-id="2805a-213">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-214">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-215">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="2805a-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="2805a-216">dateTimeModified :Date</span></span>

<span data-ttu-id="2805a-p111">Obtiene la fecha y hora en que se modificó por última vez un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="2805a-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2805a-219">Este miembro no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="2805a-219">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-220">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-220">Type:</span></span>

*   <span data-ttu-id="2805a-221">Fecha</span><span class="sxs-lookup"><span data-stu-id="2805a-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-222">Requirements</span></span>

|<span data-ttu-id="2805a-223">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-223">Requirement</span></span>| <span data-ttu-id="2805a-224">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-225">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-225">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-226">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-226">1.0</span></span>|
|[<span data-ttu-id="2805a-227">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-228">ReadItem</span></span>|
|[<span data-ttu-id="2805a-229">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-230">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-231">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="2805a-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="2805a-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="2805a-233">Obtiene o establece la fecha y la hora de finalización de la cita.</span><span class="sxs-lookup"><span data-stu-id="2805a-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="2805a-p112">La propiedad `end` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) para convertir el valor de la propiedad final a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="2805a-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2805a-236">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-236">Read mode</span></span>

<span data-ttu-id="2805a-237">La propiedad `end` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="2805a-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2805a-238">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="2805a-238">Compose mode</span></span>

<span data-ttu-id="2805a-239">La propiedad `end` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="2805a-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="2805a-240">Si usa el método [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) para establecer la hora de finalización, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="2805a-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-241">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-241">Type:</span></span>

*   <span data-ttu-id="2805a-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="2805a-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-243">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-243">Requirements</span></span>

|<span data-ttu-id="2805a-244">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-244">Requirement</span></span>| <span data-ttu-id="2805a-245">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-246">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-246">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-247">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-247">1.0</span></span>|
|[<span data-ttu-id="2805a-248">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-249">ReadItem</span></span>|
|[<span data-ttu-id="2805a-250">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-251">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-252">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-252">Example</span></span>

<span data-ttu-id="2805a-253">En el ejemplo siguiente, se establece la hora de finalización de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="2805a-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="2805a-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="2805a-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="2805a-p113">Obtiene la dirección de correo electrónico del remitente de un mensaje. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="2805a-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="2805a-p114">Las propiedades `from` y [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="2805a-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2805a-259">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `from` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="2805a-259">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-260">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-260">Type:</span></span>

*   [<span data-ttu-id="2805a-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2805a-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="2805a-262">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-262">Requirements</span></span>

|<span data-ttu-id="2805a-263">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-263">Requirement</span></span>| <span data-ttu-id="2805a-264">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-265">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-266">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-266">1.0</span></span>|
|[<span data-ttu-id="2805a-267">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-268">ReadItem</span></span>|
|[<span data-ttu-id="2805a-269">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-270">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="2805a-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="2805a-271">internetMessageId :String</span></span>

<span data-ttu-id="2805a-p115">Obtiene el identificador de mensaje de Internet de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="2805a-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-274">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-274">Type:</span></span>

*   <span data-ttu-id="2805a-275">String</span><span class="sxs-lookup"><span data-stu-id="2805a-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-276">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-276">Requirements</span></span>

|<span data-ttu-id="2805a-277">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-277">Requirement</span></span>| <span data-ttu-id="2805a-278">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-279">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-279">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-280">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-280">1.0</span></span>|
|[<span data-ttu-id="2805a-281">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-282">ReadItem</span></span>|
|[<span data-ttu-id="2805a-283">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-284">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-285">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="2805a-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="2805a-286">itemClass :String</span></span>

<span data-ttu-id="2805a-p116">Obtiene la clase de elemento Servicios Web Exchange del elemento seleccionado. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="2805a-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="2805a-p117">La propiedad `itemClass` especifica la clase de mensaje del elemento seleccionado. Las siguientes son las clases de mensaje predeterminadas para el elemento de mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="2805a-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="2805a-291">Tipo</span><span class="sxs-lookup"><span data-stu-id="2805a-291">Type</span></span> | <span data-ttu-id="2805a-292">Descripción</span><span class="sxs-lookup"><span data-stu-id="2805a-292">Description</span></span> | <span data-ttu-id="2805a-293">Clase de elemento</span><span class="sxs-lookup"><span data-stu-id="2805a-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="2805a-294">Elementos de cita</span><span class="sxs-lookup"><span data-stu-id="2805a-294">Appointment items</span></span> | <span data-ttu-id="2805a-295">Estos son los elementos de calendario de la clase de elemento `IPM.Appointment` o `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="2805a-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="2805a-296">Elementos de mensaje</span><span class="sxs-lookup"><span data-stu-id="2805a-296">Message items</span></span> | <span data-ttu-id="2805a-297">Entre estos se encuentran los mensajes de correo electrónico que tienen la clase de mensaje predeterminada `IPM.Note` y las convocatorias de reunión, respuestas y cancelaciones que usan `IPM.Schedule.Meeting` como clase de mensaje base.</span><span class="sxs-lookup"><span data-stu-id="2805a-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="2805a-298">Puede crear clases de mensaje personalizadas para ampliar una clase de mensaje predeterminada. Por ejemplo, puede crear una clase de mensaje de cita personalizada `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="2805a-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-299">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-299">Type:</span></span>

*   <span data-ttu-id="2805a-300">String</span><span class="sxs-lookup"><span data-stu-id="2805a-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-301">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-301">Requirements</span></span>

|<span data-ttu-id="2805a-302">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-302">Requirement</span></span>| <span data-ttu-id="2805a-303">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-304">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-304">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-305">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-305">1.0</span></span>|
|[<span data-ttu-id="2805a-306">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-307">ReadItem</span></span>|
|[<span data-ttu-id="2805a-308">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-309">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-310">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="2805a-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="2805a-311">(nullable) itemId :String</span></span>

<span data-ttu-id="2805a-p118">Obtiene el identificador de elemento de Servicios Web Exchange para el elemento actual. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="2805a-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2805a-314">El identificador que devuelve la propiedad `itemId` es el mismo que el identificador de elemento de Servicios Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="2805a-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="2805a-315">El `itemId` (propiedad) no es idéntico al identificador de entrada de Outlook o el identificador de la API de REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="2805a-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="2805a-316">Antes de realizar llamadas a API de REST utilizando este valor, se debe convertir mediante `Office.context.mailbox.convertToRestId`, que está disponible a partir de requisito establece 1.3.</span><span class="sxs-lookup"><span data-stu-id="2805a-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="2805a-317">Para obtener más información, vea [usar las API de REST de Outlook desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="2805a-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-318">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-318">Type:</span></span>

*   <span data-ttu-id="2805a-319">String</span><span class="sxs-lookup"><span data-stu-id="2805a-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-320">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-320">Requirements</span></span>

|<span data-ttu-id="2805a-321">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-321">Requirement</span></span>| <span data-ttu-id="2805a-322">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-323">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-323">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-324">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-324">1.0</span></span>|
|[<span data-ttu-id="2805a-325">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-326">ReadItem</span></span>|
|[<span data-ttu-id="2805a-327">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-328">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-329">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-329">Example</span></span>

<span data-ttu-id="2805a-p120">El siguiente código comprueba la presencia de un identificador de elemento. Si la propiedad `itemId` devuelve `null` o `undefined`, guarda el elemento en el servidor y obtiene el identificador de elemento desde el resultado asincrónico.</span><span class="sxs-lookup"><span data-stu-id="2805a-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="2805a-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="2805a-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="2805a-333">Obtiene el tipo de elemento que representa una instancia.</span><span class="sxs-lookup"><span data-stu-id="2805a-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="2805a-334">La propiedad `itemType` devuelve uno de los valores de enumeración de `ItemType`, lo que indica si la instancia del objeto `item` es un mensaje o una cita.</span><span class="sxs-lookup"><span data-stu-id="2805a-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-335">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-335">Type:</span></span>

*   [<span data-ttu-id="2805a-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="2805a-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="2805a-337">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-337">Requirements</span></span>

|<span data-ttu-id="2805a-338">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-338">Requirement</span></span>| <span data-ttu-id="2805a-339">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-340">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-341">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-341">1.0</span></span>|
|[<span data-ttu-id="2805a-342">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-343">ReadItem</span></span>|
|[<span data-ttu-id="2805a-344">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-345">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-346">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="2805a-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="2805a-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="2805a-348">Obtiene o establece la ubicación de una cita.</span><span class="sxs-lookup"><span data-stu-id="2805a-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2805a-349">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-349">Read mode</span></span>

<span data-ttu-id="2805a-350">La propiedad `location` devuelve una cadena que contiene la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="2805a-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2805a-351">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="2805a-351">Compose mode</span></span>

<span data-ttu-id="2805a-352">La propiedad `location` devuelve un objeto `Location` que proporciona métodos para obtener y establecer la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="2805a-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-353">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-353">Type:</span></span>

*   <span data-ttu-id="2805a-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="2805a-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-355">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-355">Requirements</span></span>

|<span data-ttu-id="2805a-356">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-356">Requirement</span></span>| <span data-ttu-id="2805a-357">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-358">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-358">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-359">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-359">1.0</span></span>|
|[<span data-ttu-id="2805a-360">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-361">ReadItem</span></span>|
|[<span data-ttu-id="2805a-362">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-363">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-364">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="2805a-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="2805a-365">normalizedSubject :String</span></span>

<span data-ttu-id="2805a-p121">Obtiene el asunto de un elemento, con todos los prefijos quitados (incluidos `RE:` y `FWD:`). Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="2805a-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="2805a-p122">La propiedad normalizedSubject obtiene el asunto del elemento, con los prefijos estándar (como `RE:` y `FW:`) agregados por los programas de correo electrónico. Para obtener el asunto del elemento con los prefijos intactos, use la propiedad [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject).</span><span class="sxs-lookup"><span data-stu-id="2805a-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-370">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-370">Type:</span></span>

*   <span data-ttu-id="2805a-371">String</span><span class="sxs-lookup"><span data-stu-id="2805a-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-372">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-372">Requirements</span></span>

|<span data-ttu-id="2805a-373">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-373">Requirement</span></span>| <span data-ttu-id="2805a-374">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-375">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-375">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-376">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-376">1.0</span></span>|
|[<span data-ttu-id="2805a-377">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-378">ReadItem</span></span>|
|[<span data-ttu-id="2805a-379">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-380">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-381">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="2805a-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2805a-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="2805a-383">Proporciona acceso a los asistentes opcionales de un evento.</span><span class="sxs-lookup"><span data-stu-id="2805a-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="2805a-384">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="2805a-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2805a-385">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-385">Read mode</span></span>

<span data-ttu-id="2805a-386">La propiedad `optionalAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente opcional a la reunión.</span><span class="sxs-lookup"><span data-stu-id="2805a-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2805a-387">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="2805a-387">Compose mode</span></span>

<span data-ttu-id="2805a-388">El `optionalAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes opcionales de una reunión.</span><span class="sxs-lookup"><span data-stu-id="2805a-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-389">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-389">Type:</span></span>

*   <span data-ttu-id="2805a-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2805a-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-391">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-391">Requirements</span></span>

|<span data-ttu-id="2805a-392">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-392">Requirement</span></span>| <span data-ttu-id="2805a-393">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-394">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-394">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-395">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-395">1.0</span></span>|
|[<span data-ttu-id="2805a-396">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-397">ReadItem</span></span>|
|[<span data-ttu-id="2805a-398">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-399">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-400">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="2805a-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="2805a-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="2805a-p124">Obtiene la dirección de correo electrónico del organizador de una reunión especificada. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="2805a-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-404">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-404">Type:</span></span>

*   [<span data-ttu-id="2805a-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2805a-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="2805a-406">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-406">Requirements</span></span>

|<span data-ttu-id="2805a-407">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-407">Requirement</span></span>| <span data-ttu-id="2805a-408">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-409">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-410">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-410">1.0</span></span>|
|[<span data-ttu-id="2805a-411">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-412">ReadItem</span></span>|
|[<span data-ttu-id="2805a-413">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-414">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-415">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="2805a-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2805a-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="2805a-417">Proporciona acceso a los asistentes necesarios de un evento.</span><span class="sxs-lookup"><span data-stu-id="2805a-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="2805a-418">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="2805a-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2805a-419">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-419">Read mode</span></span>

<span data-ttu-id="2805a-420">La propiedad `requiredAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente requerido a la reunión.</span><span class="sxs-lookup"><span data-stu-id="2805a-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2805a-421">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="2805a-421">Compose mode</span></span>

<span data-ttu-id="2805a-422">El `requiredAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes necesarios de una reunión.</span><span class="sxs-lookup"><span data-stu-id="2805a-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-423">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-423">Type:</span></span>

*   <span data-ttu-id="2805a-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2805a-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-425">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-425">Requirements</span></span>

|<span data-ttu-id="2805a-426">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-426">Requirement</span></span>| <span data-ttu-id="2805a-427">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-428">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-428">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-429">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-429">1.0</span></span>|
|[<span data-ttu-id="2805a-430">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-431">ReadItem</span></span>|
|[<span data-ttu-id="2805a-432">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-433">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-434">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="2805a-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="2805a-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="2805a-p126">Obtiene la dirección de correo electrónico del remitente de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="2805a-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="2805a-p127">Las propiedades [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) y `sender` representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="2805a-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2805a-440">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `sender` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="2805a-440">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-441">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-441">Type:</span></span>

*   [<span data-ttu-id="2805a-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2805a-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="2805a-443">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-443">Requirements</span></span>

|<span data-ttu-id="2805a-444">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-444">Requirement</span></span>| <span data-ttu-id="2805a-445">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-446">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-446">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-447">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-447">1.0</span></span>|
|[<span data-ttu-id="2805a-448">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-449">ReadItem</span></span>|
|[<span data-ttu-id="2805a-450">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-451">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-452">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="2805a-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="2805a-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="2805a-454">Obtiene o establece la fecha y la hora de inicio de la cita.</span><span class="sxs-lookup"><span data-stu-id="2805a-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="2805a-p128">La propiedad `start` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) para convertir el valor a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="2805a-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2805a-457">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-457">Read mode</span></span>

<span data-ttu-id="2805a-458">La propiedad `start` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="2805a-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2805a-459">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="2805a-459">Compose mode</span></span>

<span data-ttu-id="2805a-460">La propiedad `start` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="2805a-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="2805a-461">Si usa el método [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) para establecer la hora de inicio, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="2805a-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-462">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-462">Type:</span></span>

*   <span data-ttu-id="2805a-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="2805a-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-464">Requirements</span></span>

|<span data-ttu-id="2805a-465">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-465">Requirement</span></span>| <span data-ttu-id="2805a-466">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-467">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-467">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-468">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-468">1.0</span></span>|
|[<span data-ttu-id="2805a-469">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-470">ReadItem</span></span>|
|[<span data-ttu-id="2805a-471">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-472">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-473">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-473">Example</span></span>

<span data-ttu-id="2805a-474">En el ejemplo siguiente, se establece la hora de inicio de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="2805a-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="2805a-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="2805a-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="2805a-476">Obtiene o establece la descripción que se muestra en el campo de asunto de un elemento.</span><span class="sxs-lookup"><span data-stu-id="2805a-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="2805a-477">La propiedad `subject` obtiene o establece el asunto completo del elemento, como lo ha enviado el servidor de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="2805a-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2805a-478">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-478">Read mode</span></span>

<span data-ttu-id="2805a-p129">La propiedad `subject` devuelve una cadena. Use la propiedad [`normalizedSubject`](#normalizedsubject-string) para obtener el asunto menos cualquier prefijo inicial como `RE:` y `FW:`.</span><span class="sxs-lookup"><span data-stu-id="2805a-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="2805a-481">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="2805a-481">Compose mode</span></span>

<span data-ttu-id="2805a-482">La propiedad `subject` devuelve un objeto `Subject` que proporciona métodos para obtener y establecer el asunto.</span><span class="sxs-lookup"><span data-stu-id="2805a-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2805a-483">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-483">Type:</span></span>

*   <span data-ttu-id="2805a-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="2805a-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-485">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-485">Requirements</span></span>

|<span data-ttu-id="2805a-486">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-486">Requirement</span></span>| <span data-ttu-id="2805a-487">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-488">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-488">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-489">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-489">1.0</span></span>|
|[<span data-ttu-id="2805a-490">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-491">ReadItem</span></span>|
|[<span data-ttu-id="2805a-492">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-493">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="2805a-494">para: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2805a-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="2805a-495">Proporciona acceso a los destinatarios en la línea **para** de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="2805a-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="2805a-496">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="2805a-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2805a-497">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-497">Read mode</span></span>

<span data-ttu-id="2805a-p131">La propiedad `to` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **Para** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="2805a-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2805a-500">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="2805a-500">Compose mode</span></span>

<span data-ttu-id="2805a-501">El `to` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **para** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="2805a-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="2805a-502">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2805a-502">Type:</span></span>

*   <span data-ttu-id="2805a-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2805a-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-504">Requirements</span></span>

|<span data-ttu-id="2805a-505">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-505">Requirement</span></span>| <span data-ttu-id="2805a-506">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-507">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-508">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-508">1.0</span></span>|
|[<span data-ttu-id="2805a-509">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-510">ReadItem</span></span>|
|[<span data-ttu-id="2805a-511">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-512">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-513">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="2805a-514">Métodos</span><span class="sxs-lookup"><span data-stu-id="2805a-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="2805a-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2805a-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2805a-516">Agrega un archivo a un mensaje o cita como datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="2805a-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="2805a-517">El método `addFileAttachmentAsync` carga el archivo en el URI especificado y lo asocia al elemento en el formulario de redacción.</span><span class="sxs-lookup"><span data-stu-id="2805a-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="2805a-518">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="2805a-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2805a-519">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="2805a-519">Parameters:</span></span>

|<span data-ttu-id="2805a-520">Nombre</span><span class="sxs-lookup"><span data-stu-id="2805a-520">Name</span></span>| <span data-ttu-id="2805a-521">Tipo</span><span class="sxs-lookup"><span data-stu-id="2805a-521">Type</span></span>| <span data-ttu-id="2805a-522">Atributos</span><span class="sxs-lookup"><span data-stu-id="2805a-522">Attributes</span></span>| <span data-ttu-id="2805a-523">Descripción</span><span class="sxs-lookup"><span data-stu-id="2805a-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="2805a-524">String</span><span class="sxs-lookup"><span data-stu-id="2805a-524">String</span></span>||<span data-ttu-id="2805a-p132">El URI que proporciona la ubicación del archivo que se va a adjuntar al mensaje o a la cita. La longitud máxima es de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="2805a-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="2805a-527">String</span><span class="sxs-lookup"><span data-stu-id="2805a-527">String</span></span>||<span data-ttu-id="2805a-p133">El nombre de los datos adjuntos que se muestra mientras estos se cargan. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="2805a-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="2805a-530">Object</span><span class="sxs-lookup"><span data-stu-id="2805a-530">Object</span></span>| <span data-ttu-id="2805a-531">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-531">&lt;optional&gt;</span></span>|<span data-ttu-id="2805a-532">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="2805a-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2805a-533">Objeto</span><span class="sxs-lookup"><span data-stu-id="2805a-533">Object</span></span>| <span data-ttu-id="2805a-534">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-534">&lt;optional&gt;</span></span>|<span data-ttu-id="2805a-535">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="2805a-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2805a-536">función</span><span class="sxs-lookup"><span data-stu-id="2805a-536">function</span></span>| <span data-ttu-id="2805a-537">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-537">&lt;optional&gt;</span></span>|<span data-ttu-id="2805a-538">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2805a-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2805a-539">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2805a-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2805a-540">Si se produce un error en la carga de los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="2805a-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2805a-541">Errores</span><span class="sxs-lookup"><span data-stu-id="2805a-541">Errors</span></span>

| <span data-ttu-id="2805a-542">Código de error</span><span class="sxs-lookup"><span data-stu-id="2805a-542">Error code</span></span> | <span data-ttu-id="2805a-543">Descripción</span><span class="sxs-lookup"><span data-stu-id="2805a-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="2805a-544">Los datos adjuntos son más grandes de lo permitido.</span><span class="sxs-lookup"><span data-stu-id="2805a-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="2805a-545">Los datos adjuntos tienen una extensión que no está permitida.</span><span class="sxs-lookup"><span data-stu-id="2805a-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="2805a-546">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="2805a-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2805a-547">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-547">Requirements</span></span>

|<span data-ttu-id="2805a-548">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-548">Requirement</span></span>| <span data-ttu-id="2805a-549">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-550">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-550">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-551">1.1</span><span class="sxs-lookup"><span data-stu-id="2805a-551">1.1</span></span>|
|[<span data-ttu-id="2805a-552">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2805a-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="2805a-554">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-555">Redacción</span><span class="sxs-lookup"><span data-stu-id="2805a-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-556">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-556">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="2805a-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2805a-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2805a-558">Agrega un elemento de Exchange (por ejemplo, un mensaje) como datos adjuntos al mensaje o a la cita.</span><span class="sxs-lookup"><span data-stu-id="2805a-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="2805a-p134">El método `addItemAttachmentAsync` asocia el elemento con el identificador especificado de Exchange al elemento en el formulario de redacción. Si se especifica un método de devolución de llamada, se llama al método con un parámetro, `asyncResult`, que contiene el identificador de datos adjuntos o un código que indique cualquier error que se ha producido al adjuntar el elemento. Si es necesario, puede usar el parámetro `options` para pasar información de estado al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="2805a-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="2805a-562">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="2805a-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="2805a-563">Si el complemento de Office se está ejecutando en Outlook Web App, el `addItemAttachmentAsync` método puede adjuntar elementos a los elementos que no sea el elemento que se está editando; Sin embargo, esto no se admite y no se recomienda.</span><span class="sxs-lookup"><span data-stu-id="2805a-563">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2805a-564">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="2805a-564">Parameters:</span></span>

|<span data-ttu-id="2805a-565">Nombre</span><span class="sxs-lookup"><span data-stu-id="2805a-565">Name</span></span>| <span data-ttu-id="2805a-566">Tipo</span><span class="sxs-lookup"><span data-stu-id="2805a-566">Type</span></span>| <span data-ttu-id="2805a-567">Atributos</span><span class="sxs-lookup"><span data-stu-id="2805a-567">Attributes</span></span>| <span data-ttu-id="2805a-568">Descripción</span><span class="sxs-lookup"><span data-stu-id="2805a-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="2805a-569">String</span><span class="sxs-lookup"><span data-stu-id="2805a-569">String</span></span>||<span data-ttu-id="2805a-p135">El identificador de Exchange del elemento que debe adjuntarse. La longitud máxima es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="2805a-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="2805a-572">String</span><span class="sxs-lookup"><span data-stu-id="2805a-572">String</span></span>||<span data-ttu-id="2805a-p136">El asunto del elemento que debe adjuntarse. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="2805a-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="2805a-575">Object</span><span class="sxs-lookup"><span data-stu-id="2805a-575">Object</span></span>| <span data-ttu-id="2805a-576">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-576">&lt;optional&gt;</span></span>|<span data-ttu-id="2805a-577">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="2805a-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2805a-578">Objeto</span><span class="sxs-lookup"><span data-stu-id="2805a-578">Object</span></span>| <span data-ttu-id="2805a-579">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-579">&lt;optional&gt;</span></span>|<span data-ttu-id="2805a-580">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="2805a-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2805a-581">función</span><span class="sxs-lookup"><span data-stu-id="2805a-581">function</span></span>| <span data-ttu-id="2805a-582">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-582">&lt;optional&gt;</span></span>|<span data-ttu-id="2805a-583">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2805a-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2805a-584">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2805a-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2805a-585">Si se produce un error al agregar los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="2805a-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2805a-586">Errores</span><span class="sxs-lookup"><span data-stu-id="2805a-586">Errors</span></span>

| <span data-ttu-id="2805a-587">Código de error</span><span class="sxs-lookup"><span data-stu-id="2805a-587">Error code</span></span> | <span data-ttu-id="2805a-588">Descripción</span><span class="sxs-lookup"><span data-stu-id="2805a-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="2805a-589">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="2805a-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2805a-590">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-590">Requirements</span></span>

|<span data-ttu-id="2805a-591">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-591">Requirement</span></span>| <span data-ttu-id="2805a-592">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-593">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-593">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-594">1.1</span><span class="sxs-lookup"><span data-stu-id="2805a-594">1.1</span></span>|
|[<span data-ttu-id="2805a-595">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2805a-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="2805a-597">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-598">Redacción</span><span class="sxs-lookup"><span data-stu-id="2805a-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-599">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-599">Example</span></span>

<span data-ttu-id="2805a-600">En el ejemplo siguiente, se agrega un elemento existente de Outlook como un archivo adjunto con el nombre `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="2805a-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="2805a-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="2805a-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="2805a-602">Muestra un formulario de respuesta que incluye el remitente y todos los destinatarios del mensaje seleccionado o el organizador y todos los asistentes de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="2805a-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2805a-603">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="2805a-603">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2805a-604">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="2805a-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2805a-605">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyAllForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="2805a-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="2805a-p137">Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="2805a-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2805a-609">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="2805a-609">Parameters:</span></span>

|<span data-ttu-id="2805a-610">Nombre</span><span class="sxs-lookup"><span data-stu-id="2805a-610">Name</span></span>| <span data-ttu-id="2805a-611">Tipo</span><span class="sxs-lookup"><span data-stu-id="2805a-611">Type</span></span>| <span data-ttu-id="2805a-612">Descripción</span><span class="sxs-lookup"><span data-stu-id="2805a-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="2805a-613">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="2805a-613">String &#124; Object</span></span>| |<span data-ttu-id="2805a-p138">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="2805a-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2805a-616">**O**</span><span class="sxs-lookup"><span data-stu-id="2805a-616">**OR**</span></span><br/><span data-ttu-id="2805a-p139">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="2805a-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="2805a-619">String</span><span class="sxs-lookup"><span data-stu-id="2805a-619">String</span></span> | <span data-ttu-id="2805a-620">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-620">&lt;optional&gt;</span></span> | <span data-ttu-id="2805a-p140">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="2805a-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="2805a-623">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-623">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="2805a-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-624">&lt;optional&gt;</span></span> | <span data-ttu-id="2805a-625">Una matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="2805a-625">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="2805a-626">String</span><span class="sxs-lookup"><span data-stu-id="2805a-626">String</span></span> | | <span data-ttu-id="2805a-p141">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="2805a-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="2805a-629">String</span><span class="sxs-lookup"><span data-stu-id="2805a-629">String</span></span> | | <span data-ttu-id="2805a-630">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="2805a-630">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="2805a-631">String</span><span class="sxs-lookup"><span data-stu-id="2805a-631">String</span></span> | | <span data-ttu-id="2805a-p142">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="2805a-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="2805a-634">String</span><span class="sxs-lookup"><span data-stu-id="2805a-634">String</span></span> | | <span data-ttu-id="2805a-p143">Solo se usa si `type` se establece en `item`. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="2805a-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="2805a-638">función</span><span class="sxs-lookup"><span data-stu-id="2805a-638">function</span></span> | <span data-ttu-id="2805a-639">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-639">&lt;optional&gt;</span></span> | <span data-ttu-id="2805a-640">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2805a-640">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2805a-641">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-641">Requirements</span></span>

|<span data-ttu-id="2805a-642">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-642">Requirement</span></span>| <span data-ttu-id="2805a-643">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-644">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-644">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-645">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-645">1.0</span></span>|
|[<span data-ttu-id="2805a-646">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-647">ReadItem</span></span>|
|[<span data-ttu-id="2805a-648">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-649">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-649">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2805a-650">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="2805a-650">Examples</span></span>

<span data-ttu-id="2805a-651">El código siguiente pasa una cadena a la función `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="2805a-651">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="2805a-652">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="2805a-652">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="2805a-653">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="2805a-653">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2805a-654">Responder con un cuerpo y datos adjuntos de archivo.</span><span class="sxs-lookup"><span data-stu-id="2805a-654">Reply with a body and a file attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="2805a-655">Responder con un cuerpo y datos adjuntos de elemento.</span><span class="sxs-lookup"><span data-stu-id="2805a-655">Reply with a body and an item attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="2805a-656">Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="2805a-656">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="2805a-657">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="2805a-657">displayReplyForm(formData)</span></span>

<span data-ttu-id="2805a-658">Muestra un formulario de respuesta que incluye solo el remitente del mensaje seleccionado o el organizador de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="2805a-658">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2805a-659">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="2805a-659">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2805a-660">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="2805a-660">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2805a-661">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="2805a-661">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="2805a-p144">Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="2805a-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2805a-665">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="2805a-665">Parameters:</span></span>

|<span data-ttu-id="2805a-666">Nombre</span><span class="sxs-lookup"><span data-stu-id="2805a-666">Name</span></span>| <span data-ttu-id="2805a-667">Tipo</span><span class="sxs-lookup"><span data-stu-id="2805a-667">Type</span></span>| <span data-ttu-id="2805a-668">Descripción</span><span class="sxs-lookup"><span data-stu-id="2805a-668">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="2805a-669">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="2805a-669">String &#124; Object</span></span>| | <span data-ttu-id="2805a-p145">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="2805a-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2805a-672">**O**</span><span class="sxs-lookup"><span data-stu-id="2805a-672">**OR**</span></span><br/><span data-ttu-id="2805a-p146">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="2805a-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="2805a-675">String</span><span class="sxs-lookup"><span data-stu-id="2805a-675">String</span></span> | <span data-ttu-id="2805a-676">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-676">&lt;optional&gt;</span></span> | <span data-ttu-id="2805a-p147">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="2805a-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="2805a-679">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-679">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="2805a-680">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-680">&lt;optional&gt;</span></span> | <span data-ttu-id="2805a-681">Una matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="2805a-681">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="2805a-682">String</span><span class="sxs-lookup"><span data-stu-id="2805a-682">String</span></span> | | <span data-ttu-id="2805a-p148">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="2805a-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="2805a-685">String</span><span class="sxs-lookup"><span data-stu-id="2805a-685">String</span></span> | | <span data-ttu-id="2805a-686">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="2805a-686">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="2805a-687">String</span><span class="sxs-lookup"><span data-stu-id="2805a-687">String</span></span> | | <span data-ttu-id="2805a-p149">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="2805a-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="2805a-690">String</span><span class="sxs-lookup"><span data-stu-id="2805a-690">String</span></span> | | <span data-ttu-id="2805a-p150">Solo se usa si `type` se establece en `item`. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="2805a-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="2805a-694">función</span><span class="sxs-lookup"><span data-stu-id="2805a-694">function</span></span> | <span data-ttu-id="2805a-695">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-695">&lt;optional&gt;</span></span> | <span data-ttu-id="2805a-696">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2805a-696">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2805a-697">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-697">Requirements</span></span>

|<span data-ttu-id="2805a-698">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-698">Requirement</span></span>| <span data-ttu-id="2805a-699">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-699">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-700">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-700">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-701">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-701">1.0</span></span>|
|[<span data-ttu-id="2805a-702">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-702">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-703">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-703">ReadItem</span></span>|
|[<span data-ttu-id="2805a-704">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-704">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-705">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-705">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2805a-706">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="2805a-706">Examples</span></span>

<span data-ttu-id="2805a-707">El código siguiente pasa una cadena a la función `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="2805a-707">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="2805a-708">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="2805a-708">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="2805a-709">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="2805a-709">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2805a-710">Responder con un cuerpo y datos adjuntos de archivo.</span><span class="sxs-lookup"><span data-stu-id="2805a-710">Reply with a body and a file attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="2805a-711">Responder con un cuerpo y datos adjuntos de elemento.</span><span class="sxs-lookup"><span data-stu-id="2805a-711">Reply with a body and an item attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="2805a-712">Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="2805a-712">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="2805a-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="2805a-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="2805a-714">Obtiene las entidades que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="2805a-714">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2805a-715">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="2805a-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-716">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-716">Requirements</span></span>

|<span data-ttu-id="2805a-717">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-717">Requirement</span></span>| <span data-ttu-id="2805a-718">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-719">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-719">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-720">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-720">1.0</span></span>|
|[<span data-ttu-id="2805a-721">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-721">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-722">ReadItem</span></span>|
|[<span data-ttu-id="2805a-723">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-723">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-724">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-724">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2805a-725">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="2805a-725">Returns:</span></span>

<span data-ttu-id="2805a-726">Tipo: [Entidades](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="2805a-726">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="2805a-727">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-727">Example</span></span>

<span data-ttu-id="2805a-728">En el ejemplo siguiente se obtiene acceso a las entidades de los contactos en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="2805a-728">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="2805a-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="2805a-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="2805a-730">Obtiene una matriz de todas las entidades del tipo de entidad especificado que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="2805a-730">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2805a-731">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="2805a-731">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2805a-732">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="2805a-732">Parameters:</span></span>

|<span data-ttu-id="2805a-733">Nombre</span><span class="sxs-lookup"><span data-stu-id="2805a-733">Name</span></span>| <span data-ttu-id="2805a-734">Tipo</span><span class="sxs-lookup"><span data-stu-id="2805a-734">Type</span></span>| <span data-ttu-id="2805a-735">Descripción</span><span class="sxs-lookup"><span data-stu-id="2805a-735">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="2805a-736">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="2805a-736">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="2805a-737">Uno de los valores de enumeración de EntityType.</span><span class="sxs-lookup"><span data-stu-id="2805a-737">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2805a-738">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-738">Requirements</span></span>

|<span data-ttu-id="2805a-739">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-739">Requirement</span></span>| <span data-ttu-id="2805a-740">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-741">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-741">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-742">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-742">1.0</span></span>|
|[<span data-ttu-id="2805a-743">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-743">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-744">Restringido</span><span class="sxs-lookup"><span data-stu-id="2805a-744">Restricted</span></span>|
|[<span data-ttu-id="2805a-745">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-745">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-746">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-746">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2805a-747">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="2805a-747">Returns:</span></span>

<span data-ttu-id="2805a-748">Si el valor que se pasa a `entityType` no es un miembro válido de la enumeración de `EntityType`, el método devuelve un valor NULL.</span><span class="sxs-lookup"><span data-stu-id="2805a-748">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="2805a-749">Si no existen entidades del tipo especificado están presentes en el cuerpo del elemento, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="2805a-749">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="2805a-750">De otro modo, el tipo de los objetos en la matriz devuelta depende del tipo de entidad solicitada en el parámetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="2805a-750">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="2805a-751">Aunque el nivel de permiso mínimo para usar este método es **Restricted**, algunos tipos de entidad requieren **ReadItem** para tener acceso, como se especifica en la tabla siguiente.</span><span class="sxs-lookup"><span data-stu-id="2805a-751">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="2805a-752">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="2805a-752">Value of `entityType`</span></span> | <span data-ttu-id="2805a-753">Tipo de objetos de la matriz devuelta</span><span class="sxs-lookup"><span data-stu-id="2805a-753">Type of objects in returned array</span></span> | <span data-ttu-id="2805a-754">Nivel de permiso requerido</span><span class="sxs-lookup"><span data-stu-id="2805a-754">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="2805a-755">Cadena</span><span class="sxs-lookup"><span data-stu-id="2805a-755">String</span></span> | <span data-ttu-id="2805a-756">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="2805a-756">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="2805a-757">Contacto</span><span class="sxs-lookup"><span data-stu-id="2805a-757">Contact</span></span> | <span data-ttu-id="2805a-758">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2805a-758">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="2805a-759">Cadena</span><span class="sxs-lookup"><span data-stu-id="2805a-759">String</span></span> | <span data-ttu-id="2805a-760">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2805a-760">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="2805a-761">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="2805a-761">MeetingSuggestion</span></span> | <span data-ttu-id="2805a-762">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2805a-762">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="2805a-763">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="2805a-763">PhoneNumber</span></span> | <span data-ttu-id="2805a-764">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="2805a-764">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="2805a-765">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="2805a-765">TaskSuggestion</span></span> | <span data-ttu-id="2805a-766">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2805a-766">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="2805a-767">Cadena</span><span class="sxs-lookup"><span data-stu-id="2805a-767">String</span></span> | <span data-ttu-id="2805a-768">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="2805a-768">**Restricted**</span></span> |

<span data-ttu-id="2805a-769">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="2805a-769">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="2805a-770">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-770">Example</span></span>

<span data-ttu-id="2805a-771">En el ejemplo siguiente se muestra cómo obtener acceso a una matriz de cadenas que representan direcciones postales en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="2805a-771">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="2805a-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="2805a-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="2805a-773">Devuelve entidades conocidas en el elemento seleccionado que pasan el filtro con nombre definido en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="2805a-773">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2805a-774">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="2805a-774">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2805a-775">El método `getFilteredEntitiesByName` devuelve las entidades que coinciden con la expresión regular definida en el elemento de regla [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) del archivo XML de manifiesto con el valor de elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="2805a-775">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2805a-776">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="2805a-776">Parameters:</span></span>

|<span data-ttu-id="2805a-777">Nombre</span><span class="sxs-lookup"><span data-stu-id="2805a-777">Name</span></span>| <span data-ttu-id="2805a-778">Tipo</span><span class="sxs-lookup"><span data-stu-id="2805a-778">Type</span></span>| <span data-ttu-id="2805a-779">Descripción</span><span class="sxs-lookup"><span data-stu-id="2805a-779">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="2805a-780">String</span><span class="sxs-lookup"><span data-stu-id="2805a-780">String</span></span>|<span data-ttu-id="2805a-781">El nombre del elemento de regla `ItemHasKnownEntity` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="2805a-781">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2805a-782">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-782">Requirements</span></span>

|<span data-ttu-id="2805a-783">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-783">Requirement</span></span>| <span data-ttu-id="2805a-784">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-784">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-785">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-785">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-786">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-786">1.0</span></span>|
|[<span data-ttu-id="2805a-787">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-787">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-788">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-788">ReadItem</span></span>|
|[<span data-ttu-id="2805a-789">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-789">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-790">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-790">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2805a-791">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="2805a-791">Returns:</span></span>

<span data-ttu-id="2805a-p152">Si no existe ningún elemento `ItemHasKnownEntity` en el manifiesto con un valor de elemento `FilterName` que coincida con el parámetro `name`, el método devuelve `null`. Si el parámetro `name` no coincide con un elemento `ItemHasKnownEntity` del manifiesto, pero no existen entidades en el elemento actual que coincidan, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="2805a-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="2805a-794">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="2805a-794">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="2805a-795">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="2805a-795">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="2805a-796">Devuelve valores de cadena en el elemento seleccionado que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="2805a-796">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2805a-797">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="2805a-797">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2805a-p153">El método `getRegExMatches` devuelve las cadenas que coinciden con la expresión regular definida en cada elemento de regla `ItemHasRegularExpressionMatch` o `ItemHasKnownEntity` del archivo XML de manifiesto. Para una regla `ItemHasRegularExpressionMatch`, tiene que darse una cadena coincidente en la propiedad del elemento especificado por la regla. El tipo simple `PropertyName` define las propiedades admitidas.</span><span class="sxs-lookup"><span data-stu-id="2805a-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="2805a-801">Por ejemplo, considere un manifiesto de complemento que tenga el siguiente elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="2805a-801">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="2805a-802">El objeto devuelto desde `getRegExMatches` tendría dos propiedades: `fruits` y `veggies`.</span><span class="sxs-lookup"><span data-stu-id="2805a-802">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="2805a-p154">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="2805a-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2805a-805">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-805">Requirements</span></span>

|<span data-ttu-id="2805a-806">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-806">Requirement</span></span>| <span data-ttu-id="2805a-807">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-808">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-808">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-809">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-809">1.0</span></span>|
|[<span data-ttu-id="2805a-810">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-811">ReadItem</span></span>|
|[<span data-ttu-id="2805a-812">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-813">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-813">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2805a-814">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="2805a-814">Returns:</span></span>

<span data-ttu-id="2805a-p155">Un objeto que contiene matrices de cadenas que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. El nombre de cada matriz es igual al valor correspondiente del atributo `RegExName` de la regla `ItemHasRegularExpressionMatch` coincidente o al atributo `FilterName` de la regla `ItemHasKnownEntity` coincidente.</span><span class="sxs-lookup"><span data-stu-id="2805a-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="2805a-817">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="2805a-817">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2805a-818">Object</span><span class="sxs-lookup"><span data-stu-id="2805a-818">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="2805a-819">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-819">Example</span></span>

<span data-ttu-id="2805a-820">En el siguiente ejemplo, se muestra cómo tener acceso a la matriz de coincidencias de los <rule>elementos de expresión regular `fruits`y `veggies`, que se especifican en el manifiesto.</rule></span><span class="sxs-lookup"><span data-stu-id="2805a-820">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="2805a-821">→ de getregexmatchesbyname (Name) (nullable) {matriz. < cadena >}</span><span class="sxs-lookup"><span data-stu-id="2805a-821">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="2805a-822">Devuelve valores de cadena en el elemento seleccionado que coinciden con la expresión regular de nombre definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="2805a-822">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2805a-823">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="2805a-823">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2805a-824">El método `getRegExMatchesByName` devuelve las cadenas que coinciden con la expresión regular definida en el elemento de regla `ItemHasRegularExpressionMatch` del archivo XML de manifiesto con el valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="2805a-824">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="2805a-p156">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="2805a-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2805a-827">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="2805a-827">Parameters:</span></span>

|<span data-ttu-id="2805a-828">Nombre</span><span class="sxs-lookup"><span data-stu-id="2805a-828">Name</span></span>| <span data-ttu-id="2805a-829">Tipo</span><span class="sxs-lookup"><span data-stu-id="2805a-829">Type</span></span>| <span data-ttu-id="2805a-830">Descripción</span><span class="sxs-lookup"><span data-stu-id="2805a-830">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="2805a-831">String</span><span class="sxs-lookup"><span data-stu-id="2805a-831">String</span></span>|<span data-ttu-id="2805a-832">El nombre del elemento de regla `ItemHasRegularExpressionMatch` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="2805a-832">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2805a-833">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-833">Requirements</span></span>

|<span data-ttu-id="2805a-834">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-834">Requirement</span></span>| <span data-ttu-id="2805a-835">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-836">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-836">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-837">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-837">1.0</span></span>|
|[<span data-ttu-id="2805a-838">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-839">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-839">ReadItem</span></span>|
|[<span data-ttu-id="2805a-840">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-841">Lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-841">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2805a-842">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="2805a-842">Returns:</span></span>

<span data-ttu-id="2805a-843">Una matriz que contiene las cadenas que coinciden con la expresión regular definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="2805a-843">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="2805a-844">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="2805a-844">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2805a-845">< Cadena > de matriz.</span><span class="sxs-lookup"><span data-stu-id="2805a-845">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="2805a-846">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-846">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="2805a-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="2805a-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="2805a-848">Devuelve asincrónicamente datos seleccionados desde el asunto o el cuerpo de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="2805a-848">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="2805a-p157">Si no hay ninguna selección, pero el cursor está en el cuerpo o el asunto, el método devuelve null para los datos seleccionados. Si se selecciona un campo que no sea el cuerpo o el asunto, el método devuelve el error `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="2805a-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2805a-851">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="2805a-851">Parameters:</span></span>

|<span data-ttu-id="2805a-852">Nombre</span><span class="sxs-lookup"><span data-stu-id="2805a-852">Name</span></span>| <span data-ttu-id="2805a-853">Tipo</span><span class="sxs-lookup"><span data-stu-id="2805a-853">Type</span></span>| <span data-ttu-id="2805a-854">Atributos</span><span class="sxs-lookup"><span data-stu-id="2805a-854">Attributes</span></span>| <span data-ttu-id="2805a-855">Descripción</span><span class="sxs-lookup"><span data-stu-id="2805a-855">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="2805a-856">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="2805a-856">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="2805a-p158">Solicita un formato para los datos. Si es Texto, el método devuelve el texto sin formato como cadena, quitando toda etiqueta HTML que hubiera. Si es HTML, el método devuelve el texto seleccionado, ya sea texto sin formato o HTML.</span><span class="sxs-lookup"><span data-stu-id="2805a-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="2805a-860">Object</span><span class="sxs-lookup"><span data-stu-id="2805a-860">Object</span></span>| <span data-ttu-id="2805a-861">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-861">&lt;optional&gt;</span></span>|<span data-ttu-id="2805a-862">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="2805a-862">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2805a-863">Objeto</span><span class="sxs-lookup"><span data-stu-id="2805a-863">Object</span></span>| <span data-ttu-id="2805a-864">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-864">&lt;optional&gt;</span></span>|<span data-ttu-id="2805a-865">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="2805a-865">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2805a-866">función</span><span class="sxs-lookup"><span data-stu-id="2805a-866">function</span></span>||<span data-ttu-id="2805a-867">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2805a-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2805a-868">Para tener acceso a los datos seleccionados desde el método de devolución de llamada, llame a `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="2805a-868">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="2805a-869">Para obtener acceso a la propiedad de origen del que proviene la selección, llame a `asyncResult.value.sourceProperty`, que será cualquiera `body` o `subject`.</span><span class="sxs-lookup"><span data-stu-id="2805a-869">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2805a-870">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-870">Requirements</span></span>

|<span data-ttu-id="2805a-871">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-871">Requirement</span></span>| <span data-ttu-id="2805a-872">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-873">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-873">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-874">1.2</span><span class="sxs-lookup"><span data-stu-id="2805a-874">1.2</span></span>|
|[<span data-ttu-id="2805a-875">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2805a-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="2805a-877">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-878">Redacción</span><span class="sxs-lookup"><span data-stu-id="2805a-878">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="2805a-879">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="2805a-879">Returns:</span></span>

<span data-ttu-id="2805a-880">Los datos seleccionados como cadena con formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="2805a-880">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="2805a-881">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="2805a-881">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2805a-882">String</span><span class="sxs-lookup"><span data-stu-id="2805a-882">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="2805a-883">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-883">Example</span></span>

```JavaScript
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="2805a-884">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2805a-884">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="2805a-885">Carga de forma asincrónica las propiedades personalizadas de este complemento en el elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="2805a-885">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="2805a-p160">Las propiedades personalizadas se almacenan como pares de clave/valor según la aplicación y el elemento. Este método devuelve un objeto `CustomProperties` en la devolución de llamada, que proporciona métodos para tener acceso a las propiedades personalizadas específicas del elemento y el complemento actual. Las propiedades personalizadas no están cifradas en el elemento, por lo que no debería usarse como almacenamiento seguro.</span><span class="sxs-lookup"><span data-stu-id="2805a-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2805a-889">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="2805a-889">Parameters:</span></span>

|<span data-ttu-id="2805a-890">Nombre</span><span class="sxs-lookup"><span data-stu-id="2805a-890">Name</span></span>| <span data-ttu-id="2805a-891">Tipo</span><span class="sxs-lookup"><span data-stu-id="2805a-891">Type</span></span>| <span data-ttu-id="2805a-892">Atributos</span><span class="sxs-lookup"><span data-stu-id="2805a-892">Attributes</span></span>| <span data-ttu-id="2805a-893">Descripción</span><span class="sxs-lookup"><span data-stu-id="2805a-893">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="2805a-894">función</span><span class="sxs-lookup"><span data-stu-id="2805a-894">function</span></span>||<span data-ttu-id="2805a-895">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2805a-895">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2805a-896">Las propiedades personalizadas se proporcionan como un objeto [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2805a-896">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="2805a-897">Este objeto se puede utilizar para obtener, establecer y elimina las propiedades personalizadas del elemento y guarde los cambios en la propiedad personalizada que se vuelve a establecer en el servidor.</span><span class="sxs-lookup"><span data-stu-id="2805a-897">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="2805a-898">Objeto</span><span class="sxs-lookup"><span data-stu-id="2805a-898">Object</span></span>| <span data-ttu-id="2805a-899">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-899">&lt;optional&gt;</span></span>|<span data-ttu-id="2805a-900">Los desarrolladores pueden proporcionar cualquier objeto que desean tener acceso en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="2805a-900">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="2805a-901">Se puede tener acceso a este objeto por el `asyncResult.asyncContext` (propiedad) en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="2805a-901">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2805a-902">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-902">Requirements</span></span>

|<span data-ttu-id="2805a-903">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-903">Requirement</span></span>| <span data-ttu-id="2805a-904">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-905">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-906">1.0</span><span class="sxs-lookup"><span data-stu-id="2805a-906">1.0</span></span>|
|[<span data-ttu-id="2805a-907">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2805a-908">ReadItem</span></span>|
|[<span data-ttu-id="2805a-909">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-910">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="2805a-910">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-911">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-911">Example</span></span>

<span data-ttu-id="2805a-p163">En el siguiente ejemplo de código se muestra cómo usar el método `loadCustomPropertiesAsync` para cargar de forma asincrónica propiedades personalizadas que son específicas del elemento actual. En el ejemplo también se muestra cómo usar el método `CustomProperties.saveAsync` para guardar estas propiedades en el servidor. Después de cargar las propiedades personalizadas, el ejemplo de código usa el método `CustomProperties.get` para leer la propiedad personalizada `myProp`, el método `CustomProperties.set` para escribir en la propiedad personalizada `otherProp` y finalmente se llama al método `saveAsync` para guardar las propiedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="2805a-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="2805a-915">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2805a-915">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="2805a-916">Quita los datos adjuntos de un mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="2805a-916">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="2805a-p164">El método `removeAttachmentAsync` quita del elemento los datos adjuntos con el identificador especificado. Como práctica recomendada, debe usar el identificador de datos adjuntos para quitar datos adjuntos solo si la misma aplicación de correo ha agregado los datos adjuntos en la misma sesión. En Outlook Web App y OWA para los dispositivos, el identificador de datos adjuntos es válido solo en la misma sesión. Una sesión finaliza cuando el usuario cierra la aplicación, o si el usuario comienza a redactar en un formulario en línea y posteriormente se extrae el formulario en línea para continuar en una ventana independiente.</span><span class="sxs-lookup"><span data-stu-id="2805a-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2805a-921">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="2805a-921">Parameters:</span></span>

|<span data-ttu-id="2805a-922">Nombre</span><span class="sxs-lookup"><span data-stu-id="2805a-922">Name</span></span>| <span data-ttu-id="2805a-923">Tipo</span><span class="sxs-lookup"><span data-stu-id="2805a-923">Type</span></span>| <span data-ttu-id="2805a-924">Atributos</span><span class="sxs-lookup"><span data-stu-id="2805a-924">Attributes</span></span>| <span data-ttu-id="2805a-925">Descripción</span><span class="sxs-lookup"><span data-stu-id="2805a-925">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="2805a-926">String</span><span class="sxs-lookup"><span data-stu-id="2805a-926">String</span></span>||<span data-ttu-id="2805a-p165">El identificador de los datos adjuntos que se deben quitar. La longitud máxima de la cadena es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="2805a-p165">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="2805a-929">Object</span><span class="sxs-lookup"><span data-stu-id="2805a-929">Object</span></span>| <span data-ttu-id="2805a-930">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-930">&lt;optional&gt;</span></span>|<span data-ttu-id="2805a-931">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="2805a-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2805a-932">Objeto</span><span class="sxs-lookup"><span data-stu-id="2805a-932">Object</span></span>| <span data-ttu-id="2805a-933">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-933">&lt;optional&gt;</span></span>|<span data-ttu-id="2805a-934">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="2805a-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2805a-935">función</span><span class="sxs-lookup"><span data-stu-id="2805a-935">function</span></span>| <span data-ttu-id="2805a-936">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-936">&lt;optional&gt;</span></span>|<span data-ttu-id="2805a-937">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2805a-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2805a-938">Si se produce un error en la eliminación de los datos adjuntos, la propiedad `asyncResult.error` contendrá un código de error con el motivo del error.</span><span class="sxs-lookup"><span data-stu-id="2805a-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2805a-939">Errores</span><span class="sxs-lookup"><span data-stu-id="2805a-939">Errors</span></span>

| <span data-ttu-id="2805a-940">Código de error</span><span class="sxs-lookup"><span data-stu-id="2805a-940">Error code</span></span> | <span data-ttu-id="2805a-941">Descripción</span><span class="sxs-lookup"><span data-stu-id="2805a-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="2805a-942">El identificador de datos adjuntos no existe.</span><span class="sxs-lookup"><span data-stu-id="2805a-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2805a-943">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-943">Requirements</span></span>

|<span data-ttu-id="2805a-944">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-944">Requirement</span></span>| <span data-ttu-id="2805a-945">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-946">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-946">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-947">1.1</span><span class="sxs-lookup"><span data-stu-id="2805a-947">1.1</span></span>|
|[<span data-ttu-id="2805a-948">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2805a-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="2805a-950">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-951">Redacción</span><span class="sxs-lookup"><span data-stu-id="2805a-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-952">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-952">Example</span></span>

<span data-ttu-id="2805a-953">El código siguiente quita un archivo adjunto con un identificador de "0".</span><span class="sxs-lookup"><span data-stu-id="2805a-953">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="2805a-954">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="2805a-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="2805a-955">Inserta asincrónicamente datos en el cuerpo o el asunto de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="2805a-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="2805a-p166">El método `setSelectedDataAsync` inserta la cadena especificada en la posición del cursor en el asunto o el cuerpo del elemento, o si se selecciona texto en el editor, reemplaza el texto seleccionado. Si el cursor no está en el cuerpo o el campo del asunto, se devuelve un error. Después de la inserción, el cursor se coloca al final del contenido insertado.</span><span class="sxs-lookup"><span data-stu-id="2805a-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2805a-959">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="2805a-959">Parameters:</span></span>

|<span data-ttu-id="2805a-960">Nombre</span><span class="sxs-lookup"><span data-stu-id="2805a-960">Name</span></span>| <span data-ttu-id="2805a-961">Tipo</span><span class="sxs-lookup"><span data-stu-id="2805a-961">Type</span></span>| <span data-ttu-id="2805a-962">Atributos</span><span class="sxs-lookup"><span data-stu-id="2805a-962">Attributes</span></span>| <span data-ttu-id="2805a-963">Descripción</span><span class="sxs-lookup"><span data-stu-id="2805a-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="2805a-964">String</span><span class="sxs-lookup"><span data-stu-id="2805a-964">String</span></span>||<span data-ttu-id="2805a-p167">Datos que se van a insertar. Los datos no deben superar 1.000.000 de caracteres. Si se pasan más de 1.000.000 de caracteres, se produce una excepción `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="2805a-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="2805a-968">Object</span><span class="sxs-lookup"><span data-stu-id="2805a-968">Object</span></span>| <span data-ttu-id="2805a-969">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-969">&lt;optional&gt;</span></span>|<span data-ttu-id="2805a-970">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="2805a-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2805a-971">Objeto</span><span class="sxs-lookup"><span data-stu-id="2805a-971">Object</span></span>| <span data-ttu-id="2805a-972">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-972">&lt;optional&gt;</span></span>|<span data-ttu-id="2805a-973">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="2805a-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="2805a-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="2805a-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="2805a-975">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2805a-975">&lt;optional&gt;</span></span>|<span data-ttu-id="2805a-p168">Si `text`, se aplica el estilo actual en Outlook Web App y Outlook. Si el campo es un editor de HTML, se insertan solo los datos de texto, aunque los datos sean HTML.</span><span class="sxs-lookup"><span data-stu-id="2805a-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="2805a-p169">Si `html` y el campo admiten HTML (el asunto no), el estilo actual se aplica en Outlook Web App y el estilo predeterminado en Outlook. Si el campo es un campo de texto, se devuelve un error `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="2805a-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="2805a-980">Si `coercionType` no está establecido, el resultado depende del campo: si el campo es HTML, se usa HTML; si el campo es texto, se usa texto sin formato.</span><span class="sxs-lookup"><span data-stu-id="2805a-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="2805a-981">function</span><span class="sxs-lookup"><span data-stu-id="2805a-981">function</span></span>||<span data-ttu-id="2805a-982">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2805a-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2805a-983">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2805a-983">Requirements</span></span>

|<span data-ttu-id="2805a-984">Requirement</span><span class="sxs-lookup"><span data-stu-id="2805a-984">Requirement</span></span>| <span data-ttu-id="2805a-985">Valor</span><span class="sxs-lookup"><span data-stu-id="2805a-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="2805a-986">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="2805a-986">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2805a-987">1.2</span><span class="sxs-lookup"><span data-stu-id="2805a-987">1.2</span></span>|
|[<span data-ttu-id="2805a-988">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="2805a-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2805a-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2805a-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="2805a-990">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="2805a-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2805a-991">Redacción</span><span class="sxs-lookup"><span data-stu-id="2805a-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2805a-992">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="2805a-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```