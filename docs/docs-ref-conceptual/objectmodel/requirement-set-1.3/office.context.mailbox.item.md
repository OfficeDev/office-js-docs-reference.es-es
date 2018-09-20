
# <a name="item"></a><span data-ttu-id="d75cb-101">elemento</span><span class="sxs-lookup"><span data-stu-id="d75cb-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="d75cb-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="d75cb-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="d75cb-p101">El espacio de nombres `item` se usa para acceder al mensaje seleccionado actualmente, a una convocatoria de reunión o a una cita. Puede determinar el tipo del `item` mediante la propiedad [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="d75cb-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-105">Requirements</span></span>

|<span data-ttu-id="d75cb-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-106">Requirement</span></span>| <span data-ttu-id="d75cb-107">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-108">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-109">1.0</span></span>|
|[<span data-ttu-id="d75cb-110">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-111">Restringido</span><span class="sxs-lookup"><span data-stu-id="d75cb-111">Restricted</span></span>|
|[<span data-ttu-id="d75cb-112">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-113">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-113">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="d75cb-114">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-114">Example</span></span>

<span data-ttu-id="d75cb-115">En el siguiente ejemplo de código de JavaScript, se muestra cómo tener acceso a la propiedad `subject` del elemento actual en Outlook.</span><span class="sxs-lookup"><span data-stu-id="d75cb-115">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```
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

### <a name="members"></a><span data-ttu-id="d75cb-116">Miembros</span><span class="sxs-lookup"><span data-stu-id="d75cb-116">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook13officeattachmentdetails"></a><span data-ttu-id="d75cb-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d75cb-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

<span data-ttu-id="d75cb-p102">Obtiene una matriz de datos adjuntos para el elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d75cb-120">Ciertos tipos de archivos están bloqueados por Outlook debido a problemas de seguridad y, por tanto, no se devuelven.</span><span class="sxs-lookup"><span data-stu-id="d75cb-120">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="d75cb-121">Para obtener más información, vea [datos adjuntos bloqueados en Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="d75cb-121">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-122">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-122">Type:</span></span>

*   <span data-ttu-id="d75cb-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d75cb-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-124">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-124">Requirements</span></span>

|<span data-ttu-id="d75cb-125">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-125">Requirement</span></span>| <span data-ttu-id="d75cb-126">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-126">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-127">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-127">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-128">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-128">1.0</span></span>|
|[<span data-ttu-id="d75cb-129">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-129">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-130">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-130">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-131">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-131">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-132">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-132">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-133">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-133">Example</span></span>

<span data-ttu-id="d75cb-134">El siguiente código crea una cadena HTML con los detalles de todos los datos adjuntos en el elemento actual.</span><span class="sxs-lookup"><span data-stu-id="d75cb-134">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```
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

####  <a name="bcc-recipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="d75cb-135">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d75cb-135">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="d75cb-136">Obtiene un objeto que proporciona métodos para obtener o actualizar a los destinatarios de la línea de CCO (con copia oculta) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="d75cb-136">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="d75cb-137">Solo modo Redacción.</span><span class="sxs-lookup"><span data-stu-id="d75cb-137">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-138">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-138">Type:</span></span>

*   [<span data-ttu-id="d75cb-139">Destinatarios</span><span class="sxs-lookup"><span data-stu-id="d75cb-139">Recipients</span></span>](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="d75cb-140">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-140">Requirements</span></span>

|<span data-ttu-id="d75cb-141">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-141">Requirement</span></span>| <span data-ttu-id="d75cb-142">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-142">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-143">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-143">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-144">1.1</span><span class="sxs-lookup"><span data-stu-id="d75cb-144">1.1</span></span>|
|[<span data-ttu-id="d75cb-145">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-145">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-146">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-147">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-147">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-148">Redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-148">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-149">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-149">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook13officebody"></a><span data-ttu-id="d75cb-150">body :[Body](/javascript/api/outlook_1_3/office.body)</span><span class="sxs-lookup"><span data-stu-id="d75cb-150">body :[Body](/javascript/api/outlook_1_3/office.body)</span></span>

<span data-ttu-id="d75cb-151">Obtiene un objeto que proporciona métodos para manipular el cuerpo de un elemento.</span><span class="sxs-lookup"><span data-stu-id="d75cb-151">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-152">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-152">Type:</span></span>

*   [<span data-ttu-id="d75cb-153">Cuerpo</span><span class="sxs-lookup"><span data-stu-id="d75cb-153">Body</span></span>](/javascript/api/outlook_1_3/office.body)

##### <a name="requirements"></a><span data-ttu-id="d75cb-154">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-154">Requirements</span></span>

|<span data-ttu-id="d75cb-155">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-155">Requirement</span></span>| <span data-ttu-id="d75cb-156">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-157">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-157">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-158">1.1</span><span class="sxs-lookup"><span data-stu-id="d75cb-158">1.1</span></span>|
|[<span data-ttu-id="d75cb-159">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-159">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-160">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-160">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-161">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-162">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-162">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="d75cb-163">cc: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d75cb-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="d75cb-164">Proporciona acceso a los destinatarios de Cc (con copia) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="d75cb-164">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="d75cb-165">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="d75cb-165">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d75cb-166">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-166">Read mode</span></span>

<span data-ttu-id="d75cb-p106">La propiedad `cc` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **CC** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d75cb-169">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-169">Compose mode</span></span>

<span data-ttu-id="d75cb-170">El `cc` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **Cc** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="d75cb-170">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-171">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-171">Type:</span></span>

*   <span data-ttu-id="d75cb-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d75cb-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-173">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-173">Requirements</span></span>

|<span data-ttu-id="d75cb-174">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-174">Requirement</span></span>| <span data-ttu-id="d75cb-175">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-176">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-176">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-177">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-177">1.0</span></span>|
|[<span data-ttu-id="d75cb-178">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-178">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-179">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-180">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-181">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-181">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-182">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-182">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="d75cb-183">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="d75cb-183">(nullable) conversationId :String</span></span>

<span data-ttu-id="d75cb-184">Obtiene un identificador de la conversación de correo electrónico que contiene un mensaje determinado.</span><span class="sxs-lookup"><span data-stu-id="d75cb-184">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="d75cb-p107">Puede obtener un entero para esta propiedad si la aplicación de correo se activa en respuestas o formularios de lectura o en formularios de redacción. Si, después, el usuario cambia el asunto del mensaje de respuesta, al enviar dicha respuesta, el Id. de conversación de ese mensaje cambiará y el valor que se obtuvo anteriormente ya no será de aplicación.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="d75cb-p108">Obtiene un valor NULL para esta propiedad para un nuevo elemento de un formulario de redacción. Si el usuario establece un asunto y guarda el elemento, la propiedad `conversationId` devolverá un valor.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-189">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-189">Type:</span></span>

*   <span data-ttu-id="d75cb-190">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-190">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-191">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-191">Requirements</span></span>

|<span data-ttu-id="d75cb-192">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-192">Requirement</span></span>| <span data-ttu-id="d75cb-193">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-194">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-194">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-195">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-195">1.0</span></span>|
|[<span data-ttu-id="d75cb-196">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-197">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-198">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-199">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-199">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="d75cb-200">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="d75cb-200">dateTimeCreated :Date</span></span>

<span data-ttu-id="d75cb-p109">Obtiene la fecha y hora de creación de un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-203">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-203">Type:</span></span>

*   <span data-ttu-id="d75cb-204">Fecha</span><span class="sxs-lookup"><span data-stu-id="d75cb-204">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-205">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-205">Requirements</span></span>

|<span data-ttu-id="d75cb-206">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-206">Requirement</span></span>| <span data-ttu-id="d75cb-207">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-208">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-208">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-209">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-209">1.0</span></span>|
|[<span data-ttu-id="d75cb-210">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-210">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-211">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-211">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-212">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-213">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-213">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-214">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-214">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="d75cb-215">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="d75cb-215">dateTimeModified :Date</span></span>

<span data-ttu-id="d75cb-p110">Obtiene la fecha y hora en que se modificó por última vez un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d75cb-218">Este miembro no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="d75cb-218">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-219">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-219">Type:</span></span>

*   <span data-ttu-id="d75cb-220">Fecha</span><span class="sxs-lookup"><span data-stu-id="d75cb-220">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-221">Requirements</span></span>

|<span data-ttu-id="d75cb-222">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-222">Requirement</span></span>| <span data-ttu-id="d75cb-223">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-224">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-225">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-225">1.0</span></span>|
|[<span data-ttu-id="d75cb-226">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-227">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-228">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-229">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-230">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-230">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="d75cb-231">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="d75cb-231">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="d75cb-232">Obtiene o establece la fecha y la hora de finalización de la cita.</span><span class="sxs-lookup"><span data-stu-id="d75cb-232">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="d75cb-p111">La propiedad `end` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) para convertir el valor de la propiedad final a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d75cb-235">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-235">Read mode</span></span>

<span data-ttu-id="d75cb-236">La propiedad `end` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-236">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d75cb-237">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-237">Compose mode</span></span>

<span data-ttu-id="d75cb-238">La propiedad `end` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-238">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="d75cb-239">Si usa el método [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) para establecer la hora de finalización, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="d75cb-239">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-240">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-240">Type:</span></span>

*   <span data-ttu-id="d75cb-241">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="d75cb-241">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-242">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-242">Requirements</span></span>

|<span data-ttu-id="d75cb-243">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-243">Requirement</span></span>| <span data-ttu-id="d75cb-244">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-245">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-245">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-246">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-246">1.0</span></span>|
|[<span data-ttu-id="d75cb-247">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-247">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-248">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-249">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-249">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-250">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-250">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-251">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-251">Example</span></span>

<span data-ttu-id="d75cb-252">En el ejemplo siguiente, se establece la hora de finalización de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-252">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="d75cb-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="d75cb-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="d75cb-p112">Obtiene la dirección de correo electrónico del remitente de un mensaje. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="d75cb-p113">Las propiedades `from` y [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d75cb-258">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `from` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-258">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-259">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-259">Type:</span></span>

*   [<span data-ttu-id="d75cb-260">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d75cb-260">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="d75cb-261">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-261">Requirements</span></span>

|<span data-ttu-id="d75cb-262">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-262">Requirement</span></span>| <span data-ttu-id="d75cb-263">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-264">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-265">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-265">1.0</span></span>|
|[<span data-ttu-id="d75cb-266">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-267">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-268">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-269">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-269">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="d75cb-270">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="d75cb-270">internetMessageId :String</span></span>

<span data-ttu-id="d75cb-p114">Obtiene el identificador de mensaje de Internet de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-273">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-273">Type:</span></span>

*   <span data-ttu-id="d75cb-274">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-275">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-275">Requirements</span></span>

|<span data-ttu-id="d75cb-276">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-276">Requirement</span></span>| <span data-ttu-id="d75cb-277">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-278">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-278">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-279">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-279">1.0</span></span>|
|[<span data-ttu-id="d75cb-280">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-281">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-282">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-283">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-283">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-284">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-284">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="d75cb-285">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="d75cb-285">itemClass :String</span></span>

<span data-ttu-id="d75cb-p115">Obtiene la clase de elemento Servicios Web Exchange del elemento seleccionado. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="d75cb-p116">La propiedad `itemClass` especifica la clase de mensaje del elemento seleccionado. Las siguientes son las clases de mensaje predeterminadas para el elemento de mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="d75cb-290">Tipo</span><span class="sxs-lookup"><span data-stu-id="d75cb-290">Type</span></span> | <span data-ttu-id="d75cb-291">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-291">Description</span></span> | <span data-ttu-id="d75cb-292">Clase de elemento</span><span class="sxs-lookup"><span data-stu-id="d75cb-292">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="d75cb-293">Elementos de cita</span><span class="sxs-lookup"><span data-stu-id="d75cb-293">Appointment items</span></span> | <span data-ttu-id="d75cb-294">Estos son los elementos de calendario de la clase de elemento `IPM.Appointment` o `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-294">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="d75cb-295">Elementos de mensaje</span><span class="sxs-lookup"><span data-stu-id="d75cb-295">Message items</span></span> | <span data-ttu-id="d75cb-296">Entre estos se encuentran los mensajes de correo electrónico que tienen la clase de mensaje predeterminada `IPM.Note` y las convocatorias de reunión, respuestas y cancelaciones que usan `IPM.Schedule.Meeting` como clase de mensaje base.</span><span class="sxs-lookup"><span data-stu-id="d75cb-296">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="d75cb-297">Puede crear clases de mensaje personalizadas para ampliar una clase de mensaje predeterminada. Por ejemplo, puede crear una clase de mensaje de cita personalizada `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-297">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-298">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-298">Type:</span></span>

*   <span data-ttu-id="d75cb-299">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-299">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-300">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-300">Requirements</span></span>

|<span data-ttu-id="d75cb-301">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-301">Requirement</span></span>| <span data-ttu-id="d75cb-302">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-303">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-304">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-304">1.0</span></span>|
|[<span data-ttu-id="d75cb-305">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-306">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-307">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-308">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-309">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-309">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="d75cb-310">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="d75cb-310">(nullable) itemId :String</span></span>

<span data-ttu-id="d75cb-p117">Obtiene el identificador de elemento de Servicios Web Exchange para el elemento actual. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d75cb-313">El identificador que devuelve la propiedad `itemId` es el mismo que el identificador de elemento de Servicios Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="d75cb-313">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="d75cb-314">El `itemId` (propiedad) no es idéntico al identificador de entrada de Outlook o el identificador de la API de REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="d75cb-314">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="d75cb-315">Antes de realizar llamadas a la API de REST con este valor, se debe convertir mediante [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="d75cb-315">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d75cb-316">Para obtener más información, vea [usar las API de REST de Outlook desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="d75cb-316">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="d75cb-p119">La propiedad `itemId` no está disponible en el modo redacción. Si se requiere un identificador de elemento, el método [`saveAsync`](#saveasyncoptions-callback) puede usarse para guardar el elemento en el servidor, que devolverá el identificador de elemento en el parámetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) de la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-319">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-319">Type:</span></span>

*   <span data-ttu-id="d75cb-320">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-321">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-321">Requirements</span></span>

|<span data-ttu-id="d75cb-322">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-322">Requirement</span></span>| <span data-ttu-id="d75cb-323">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-324">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-324">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-325">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-325">1.0</span></span>|
|[<span data-ttu-id="d75cb-326">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-327">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-328">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-329">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-330">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-330">Example</span></span>

<span data-ttu-id="d75cb-p120">El siguiente código comprueba la presencia de un identificador de elemento. Si la propiedad `itemId` devuelve `null` o `undefined`, guarda el elemento en el servidor y obtiene el identificador de elemento desde el resultado asincrónico.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype"></a><span data-ttu-id="d75cb-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="d75cb-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="d75cb-334">Obtiene el tipo de elemento que representa una instancia.</span><span class="sxs-lookup"><span data-stu-id="d75cb-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="d75cb-335">La propiedad `itemType` devuelve uno de los valores de enumeración de `ItemType`, lo que indica si la instancia del objeto `item` es un mensaje o una cita.</span><span class="sxs-lookup"><span data-stu-id="d75cb-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-336">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-336">Type:</span></span>

*   [<span data-ttu-id="d75cb-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="d75cb-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="d75cb-338">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-338">Requirements</span></span>

|<span data-ttu-id="d75cb-339">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-339">Requirement</span></span>| <span data-ttu-id="d75cb-340">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-341">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-341">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-342">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-342">1.0</span></span>|
|[<span data-ttu-id="d75cb-343">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-344">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-345">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-346">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-347">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-347">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook13officelocation"></a><span data-ttu-id="d75cb-348">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="d75cb-348">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span></span>

<span data-ttu-id="d75cb-349">Obtiene o establece la ubicación de una cita.</span><span class="sxs-lookup"><span data-stu-id="d75cb-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d75cb-350">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-350">Read mode</span></span>

<span data-ttu-id="d75cb-351">La propiedad `location` devuelve una cadena que contiene la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="d75cb-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d75cb-352">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-352">Compose mode</span></span>

<span data-ttu-id="d75cb-353">La propiedad `location` devuelve un objeto `Location` que proporciona métodos para obtener y establecer la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="d75cb-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-354">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-354">Type:</span></span>

*   <span data-ttu-id="d75cb-355">String | [Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="d75cb-355">String | [Location](/javascript/api/outlook_1_3/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-356">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-356">Requirements</span></span>

|<span data-ttu-id="d75cb-357">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-357">Requirement</span></span>| <span data-ttu-id="d75cb-358">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-359">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-359">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-360">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-360">1.0</span></span>|
|[<span data-ttu-id="d75cb-361">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-362">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-363">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-364">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-365">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-365">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="d75cb-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="d75cb-366">normalizedSubject :String</span></span>

<span data-ttu-id="d75cb-p121">Obtiene el asunto de un elemento, con todos los prefijos quitados (incluidos `RE:` y `FWD:`). Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="d75cb-p122">La propiedad normalizedSubject obtiene el asunto del elemento, con los prefijos estándar (como `RE:` y `FW:`) agregados por los programas de correo electrónico. Para obtener el asunto del elemento con los prefijos intactos, use la propiedad [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject).</span><span class="sxs-lookup"><span data-stu-id="d75cb-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-371">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-371">Type:</span></span>

*   <span data-ttu-id="d75cb-372">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-373">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-373">Requirements</span></span>

|<span data-ttu-id="d75cb-374">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-374">Requirement</span></span>| <span data-ttu-id="d75cb-375">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-376">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-376">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-377">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-377">1.0</span></span>|
|[<span data-ttu-id="d75cb-378">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-379">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-380">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-381">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-382">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-382">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages"></a><span data-ttu-id="d75cb-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="d75cb-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span></span>

<span data-ttu-id="d75cb-384">Obtiene los mensajes de notificación de un elemento.</span><span class="sxs-lookup"><span data-stu-id="d75cb-384">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-385">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-385">Type:</span></span>

*   [<span data-ttu-id="d75cb-386">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="d75cb-386">NotificationMessages</span></span>](/javascript/api/outlook_1_3/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="d75cb-387">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-387">Requirements</span></span>

|<span data-ttu-id="d75cb-388">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-388">Requirement</span></span>| <span data-ttu-id="d75cb-389">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-390">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-390">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-391">1.3</span><span class="sxs-lookup"><span data-stu-id="d75cb-391">1.3</span></span>|
|[<span data-ttu-id="d75cb-392">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-393">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-394">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-395">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-395">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="d75cb-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d75cb-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="d75cb-397">Proporciona acceso a los asistentes opcionales de un evento.</span><span class="sxs-lookup"><span data-stu-id="d75cb-397">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="d75cb-398">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="d75cb-398">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d75cb-399">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-399">Read mode</span></span>

<span data-ttu-id="d75cb-400">La propiedad `optionalAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente opcional a la reunión.</span><span class="sxs-lookup"><span data-stu-id="d75cb-400">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d75cb-401">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-401">Compose mode</span></span>

<span data-ttu-id="d75cb-402">El `optionalAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes opcionales de una reunión.</span><span class="sxs-lookup"><span data-stu-id="d75cb-402">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-403">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-403">Type:</span></span>

*   <span data-ttu-id="d75cb-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d75cb-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-405">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-405">Requirements</span></span>

|<span data-ttu-id="d75cb-406">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-406">Requirement</span></span>| <span data-ttu-id="d75cb-407">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-408">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-408">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-409">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-409">1.0</span></span>|
|[<span data-ttu-id="d75cb-410">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-411">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-412">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-413">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-413">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-414">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-414">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="d75cb-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="d75cb-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="d75cb-p124">Obtiene la dirección de correo electrónico del organizador de una reunión especificada. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-418">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-418">Type:</span></span>

*   [<span data-ttu-id="d75cb-419">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d75cb-419">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="d75cb-420">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-420">Requirements</span></span>

|<span data-ttu-id="d75cb-421">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-421">Requirement</span></span>| <span data-ttu-id="d75cb-422">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-423">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-423">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-424">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-424">1.0</span></span>|
|[<span data-ttu-id="d75cb-425">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-426">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-427">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-428">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-429">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-429">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="d75cb-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d75cb-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="d75cb-431">Proporciona acceso a los asistentes necesarios de un evento.</span><span class="sxs-lookup"><span data-stu-id="d75cb-431">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="d75cb-432">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="d75cb-432">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d75cb-433">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-433">Read mode</span></span>

<span data-ttu-id="d75cb-434">La propiedad `requiredAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente requerido a la reunión.</span><span class="sxs-lookup"><span data-stu-id="d75cb-434">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d75cb-435">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-435">Compose mode</span></span>

<span data-ttu-id="d75cb-436">El `requiredAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes necesarios de una reunión.</span><span class="sxs-lookup"><span data-stu-id="d75cb-436">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-437">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-437">Type:</span></span>

*   <span data-ttu-id="d75cb-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d75cb-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-439">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-439">Requirements</span></span>

|<span data-ttu-id="d75cb-440">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-440">Requirement</span></span>| <span data-ttu-id="d75cb-441">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-442">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-442">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-443">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-443">1.0</span></span>|
|[<span data-ttu-id="d75cb-444">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-445">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-446">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-447">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-448">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-448">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="d75cb-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="d75cb-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="d75cb-p126">Obtiene la dirección de correo electrónico del remitente de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="d75cb-p127">Las propiedades [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) y `sender` representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d75cb-454">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `sender` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-454">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-455">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-455">Type:</span></span>

*   [<span data-ttu-id="d75cb-456">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d75cb-456">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="d75cb-457">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-457">Requirements</span></span>

|<span data-ttu-id="d75cb-458">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-458">Requirement</span></span>| <span data-ttu-id="d75cb-459">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-460">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-460">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-461">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-461">1.0</span></span>|
|[<span data-ttu-id="d75cb-462">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-463">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-464">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-465">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-466">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-466">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="d75cb-467">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="d75cb-467">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="d75cb-468">Obtiene o establece la fecha y la hora de inicio de la cita.</span><span class="sxs-lookup"><span data-stu-id="d75cb-468">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="d75cb-p128">La propiedad `start` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) para convertir el valor a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d75cb-471">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-471">Read mode</span></span>

<span data-ttu-id="d75cb-472">La propiedad `start` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-472">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d75cb-473">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-473">Compose mode</span></span>

<span data-ttu-id="d75cb-474">La propiedad `start` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-474">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="d75cb-475">Si usa el método [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) para establecer la hora de inicio, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="d75cb-475">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-476">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-476">Type:</span></span>

*   <span data-ttu-id="d75cb-477">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="d75cb-477">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-478">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-478">Requirements</span></span>

|<span data-ttu-id="d75cb-479">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-479">Requirement</span></span>| <span data-ttu-id="d75cb-480">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-481">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-481">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-482">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-482">1.0</span></span>|
|[<span data-ttu-id="d75cb-483">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-484">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-485">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-486">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-487">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-487">Example</span></span>

<span data-ttu-id="d75cb-488">En el ejemplo siguiente, se establece la hora de inicio de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-488">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

####  <a name="subject-stringsubjectjavascriptapioutlook13officesubject"></a><span data-ttu-id="d75cb-489">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="d75cb-489">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

<span data-ttu-id="d75cb-490">Obtiene o establece la descripción que se muestra en el campo de asunto de un elemento.</span><span class="sxs-lookup"><span data-stu-id="d75cb-490">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="d75cb-491">La propiedad `subject` obtiene o establece el asunto completo del elemento, como lo ha enviado el servidor de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="d75cb-491">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d75cb-492">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-492">Read mode</span></span>

<span data-ttu-id="d75cb-p129">La propiedad `subject` devuelve una cadena. Use la propiedad [`normalizedSubject`](#normalizedsubject-string) para obtener el asunto menos cualquier prefijo inicial como `RE:` y `FW:`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="d75cb-495">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-495">Compose mode</span></span>

<span data-ttu-id="d75cb-496">La propiedad `subject` devuelve un objeto `Subject` que proporciona métodos para obtener y establecer el asunto.</span><span class="sxs-lookup"><span data-stu-id="d75cb-496">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d75cb-497">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-497">Type:</span></span>

*   <span data-ttu-id="d75cb-498">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="d75cb-498">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-499">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-499">Requirements</span></span>

|<span data-ttu-id="d75cb-500">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-500">Requirement</span></span>| <span data-ttu-id="d75cb-501">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-502">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-502">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-503">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-503">1.0</span></span>|
|[<span data-ttu-id="d75cb-504">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-505">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-506">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-507">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-507">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="d75cb-508">para: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d75cb-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="d75cb-509">Proporciona acceso a los destinatarios en la línea **para** de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="d75cb-509">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="d75cb-510">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="d75cb-510">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d75cb-511">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-511">Read mode</span></span>

<span data-ttu-id="d75cb-p131">La propiedad `to` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **Para** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d75cb-514">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-514">Compose mode</span></span>

<span data-ttu-id="d75cb-515">El `to` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **para** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="d75cb-515">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="d75cb-516">Tipo:</span><span class="sxs-lookup"><span data-stu-id="d75cb-516">Type:</span></span>

*   <span data-ttu-id="d75cb-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d75cb-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-518">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-518">Requirements</span></span>

|<span data-ttu-id="d75cb-519">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-519">Requirement</span></span>| <span data-ttu-id="d75cb-520">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-521">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-522">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-522">1.0</span></span>|
|[<span data-ttu-id="d75cb-523">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-524">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-525">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-526">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-527">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-527">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="d75cb-528">Métodos</span><span class="sxs-lookup"><span data-stu-id="d75cb-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="d75cb-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d75cb-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d75cb-530">Agrega un archivo a un mensaje o cita como datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="d75cb-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="d75cb-531">El método `addFileAttachmentAsync` carga el archivo en el URI especificado y lo asocia al elemento en el formulario de redacción.</span><span class="sxs-lookup"><span data-stu-id="d75cb-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="d75cb-532">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="d75cb-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d75cb-533">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="d75cb-533">Parameters:</span></span>

|<span data-ttu-id="d75cb-534">Nombre</span><span class="sxs-lookup"><span data-stu-id="d75cb-534">Name</span></span>| <span data-ttu-id="d75cb-535">Tipo</span><span class="sxs-lookup"><span data-stu-id="d75cb-535">Type</span></span>| <span data-ttu-id="d75cb-536">Atributos</span><span class="sxs-lookup"><span data-stu-id="d75cb-536">Attributes</span></span>| <span data-ttu-id="d75cb-537">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="d75cb-538">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-538">String</span></span>||<span data-ttu-id="d75cb-p132">El URI que proporciona la ubicación del archivo que se va a adjuntar al mensaje o a la cita. La longitud máxima es de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="d75cb-541">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-541">String</span></span>||<span data-ttu-id="d75cb-p133">El nombre de los datos adjuntos que se muestra mientras estos se cargan. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="d75cb-544">Object</span><span class="sxs-lookup"><span data-stu-id="d75cb-544">Object</span></span>| <span data-ttu-id="d75cb-545">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-545">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-546">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="d75cb-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d75cb-547">Objeto</span><span class="sxs-lookup"><span data-stu-id="d75cb-547">Object</span></span>| <span data-ttu-id="d75cb-548">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-548">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-549">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="d75cb-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d75cb-550">función</span><span class="sxs-lookup"><span data-stu-id="d75cb-550">function</span></span>| <span data-ttu-id="d75cb-551">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-551">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-552">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d75cb-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d75cb-553">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d75cb-554">Si se produce un error en la carga de los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="d75cb-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d75cb-555">Errores</span><span class="sxs-lookup"><span data-stu-id="d75cb-555">Errors</span></span>

| <span data-ttu-id="d75cb-556">Código de error</span><span class="sxs-lookup"><span data-stu-id="d75cb-556">Error code</span></span> | <span data-ttu-id="d75cb-557">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="d75cb-558">Los datos adjuntos son más grandes de lo permitido.</span><span class="sxs-lookup"><span data-stu-id="d75cb-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="d75cb-559">Los datos adjuntos tienen una extensión que no está permitida.</span><span class="sxs-lookup"><span data-stu-id="d75cb-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="d75cb-560">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="d75cb-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d75cb-561">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-561">Requirements</span></span>

|<span data-ttu-id="d75cb-562">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-562">Requirement</span></span>| <span data-ttu-id="d75cb-563">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-564">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-564">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-565">1.1</span><span class="sxs-lookup"><span data-stu-id="d75cb-565">1.1</span></span>|
|[<span data-ttu-id="d75cb-566">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="d75cb-568">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-569">Redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-570">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-570">Example</span></span>

```
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="d75cb-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d75cb-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d75cb-572">Agrega un elemento de Exchange (por ejemplo, un mensaje) como datos adjuntos al mensaje o a la cita.</span><span class="sxs-lookup"><span data-stu-id="d75cb-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="d75cb-p134">El método `addItemAttachmentAsync` asocia el elemento con el identificador especificado de Exchange al elemento en el formulario de redacción. Si se especifica un método de devolución de llamada, se llama al método con un parámetro, `asyncResult`, que contiene el identificador de datos adjuntos o un código que indique cualquier error que se ha producido al adjuntar el elemento. Si es necesario, puede usar el parámetro `options` para pasar información de estado al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="d75cb-576">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="d75cb-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="d75cb-577">Si el complemento de Office se está ejecutando en Outlook Web App, el `addItemAttachmentAsync` método puede adjuntar elementos a los elementos que no sea el elemento que se está editando; Sin embargo, esto no se admite y no se recomienda.</span><span class="sxs-lookup"><span data-stu-id="d75cb-577">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d75cb-578">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="d75cb-578">Parameters:</span></span>

|<span data-ttu-id="d75cb-579">Nombre</span><span class="sxs-lookup"><span data-stu-id="d75cb-579">Name</span></span>| <span data-ttu-id="d75cb-580">Tipo</span><span class="sxs-lookup"><span data-stu-id="d75cb-580">Type</span></span>| <span data-ttu-id="d75cb-581">Atributos</span><span class="sxs-lookup"><span data-stu-id="d75cb-581">Attributes</span></span>| <span data-ttu-id="d75cb-582">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="d75cb-583">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-583">String</span></span>||<span data-ttu-id="d75cb-p135">El identificador de Exchange del elemento que debe adjuntarse. La longitud máxima es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="d75cb-586">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-586">String</span></span>||<span data-ttu-id="d75cb-p136">El asunto del elemento que debe adjuntarse. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="d75cb-589">Object</span><span class="sxs-lookup"><span data-stu-id="d75cb-589">Object</span></span>| <span data-ttu-id="d75cb-590">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-590">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-591">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="d75cb-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d75cb-592">Objeto</span><span class="sxs-lookup"><span data-stu-id="d75cb-592">Object</span></span>| <span data-ttu-id="d75cb-593">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-593">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-594">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="d75cb-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d75cb-595">función</span><span class="sxs-lookup"><span data-stu-id="d75cb-595">function</span></span>| <span data-ttu-id="d75cb-596">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-596">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-597">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d75cb-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d75cb-598">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d75cb-599">Si se produce un error al agregar los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="d75cb-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d75cb-600">Errores</span><span class="sxs-lookup"><span data-stu-id="d75cb-600">Errors</span></span>

| <span data-ttu-id="d75cb-601">Código de error</span><span class="sxs-lookup"><span data-stu-id="d75cb-601">Error code</span></span> | <span data-ttu-id="d75cb-602">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="d75cb-603">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="d75cb-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d75cb-604">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-604">Requirements</span></span>

|<span data-ttu-id="d75cb-605">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-605">Requirement</span></span>| <span data-ttu-id="d75cb-606">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-607">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-607">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-608">1.1</span><span class="sxs-lookup"><span data-stu-id="d75cb-608">1.1</span></span>|
|[<span data-ttu-id="d75cb-609">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="d75cb-611">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-612">Redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-613">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-613">Example</span></span>

<span data-ttu-id="d75cb-614">En el ejemplo siguiente, se agrega un elemento existente de Outlook como un archivo adjunto con el nombre `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```
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

####  <a name="close"></a><span data-ttu-id="d75cb-615">close()</span><span class="sxs-lookup"><span data-stu-id="d75cb-615">close()</span></span>

<span data-ttu-id="d75cb-616">Cierra el elemento actual que se está redactando.</span><span class="sxs-lookup"><span data-stu-id="d75cb-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="d75cb-p137">El comportamiento del método `close` depende del estado actual del elemento que se está redactando. Si el elemento tiene cambios sin guardar, el cliente solicita al usuario que guarde, descarte o cancele la acción Cerrar.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="d75cb-619">En Outlook, en el sitio web, si el elemento es una cita y anteriormente se ha guardado con `saveAsync`, se pide al usuario al guardar, descartar o cancelar incluso si no ha habido cambios desde la último el elemento guardado.</span><span class="sxs-lookup"><span data-stu-id="d75cb-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="d75cb-620">En el cliente de escritorio de Outlook, si el mensaje es una respuesta directa, el método `close` no tiene ningún efecto.</span><span class="sxs-lookup"><span data-stu-id="d75cb-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-621">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-621">Requirements</span></span>

|<span data-ttu-id="d75cb-622">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-622">Requirement</span></span>| <span data-ttu-id="d75cb-623">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-624">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-624">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-625">1.3</span><span class="sxs-lookup"><span data-stu-id="d75cb-625">1.3</span></span>|
|[<span data-ttu-id="d75cb-626">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-627">Restringido</span><span class="sxs-lookup"><span data-stu-id="d75cb-627">Restricted</span></span>|
|[<span data-ttu-id="d75cb-628">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-629">Redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-629">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="d75cb-630">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="d75cb-630">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="d75cb-631">Muestra un formulario de respuesta que incluye el remitente y todos los destinatarios del mensaje seleccionado o el organizador y todos los asistentes de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="d75cb-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d75cb-632">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="d75cb-632">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d75cb-633">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="d75cb-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d75cb-634">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyAllForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="d75cb-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="d75cb-p138">Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d75cb-638">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="d75cb-638">Parameters:</span></span>

|<span data-ttu-id="d75cb-639">Nombre</span><span class="sxs-lookup"><span data-stu-id="d75cb-639">Name</span></span>| <span data-ttu-id="d75cb-640">Tipo</span><span class="sxs-lookup"><span data-stu-id="d75cb-640">Type</span></span>| <span data-ttu-id="d75cb-641">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="d75cb-642">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d75cb-642">String &#124; Object</span></span>| |<span data-ttu-id="d75cb-p139">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d75cb-645">**O**</span><span class="sxs-lookup"><span data-stu-id="d75cb-645">**OR**</span></span><br/><span data-ttu-id="d75cb-p140">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="d75cb-648">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-648">String</span></span> | <span data-ttu-id="d75cb-649">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-649">&lt;optional&gt;</span></span> | <span data-ttu-id="d75cb-p141">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="d75cb-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d75cb-653">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-653">&lt;optional&gt;</span></span> | <span data-ttu-id="d75cb-654">Una matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="d75cb-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="d75cb-655">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-655">String</span></span> | | <span data-ttu-id="d75cb-p142">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="d75cb-658">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-658">String</span></span> | | <span data-ttu-id="d75cb-659">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="d75cb-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="d75cb-660">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-660">String</span></span> | | <span data-ttu-id="d75cb-p143">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="d75cb-663">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-663">String</span></span> | | <span data-ttu-id="d75cb-p144">Solo se usa si `type` se establece en `item`. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="d75cb-667">función</span><span class="sxs-lookup"><span data-stu-id="d75cb-667">function</span></span> | <span data-ttu-id="d75cb-668">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-668">&lt;optional&gt;</span></span> | <span data-ttu-id="d75cb-669">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d75cb-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d75cb-670">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-670">Requirements</span></span>

|<span data-ttu-id="d75cb-671">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-671">Requirement</span></span>| <span data-ttu-id="d75cb-672">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-673">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-673">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-674">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-674">1.0</span></span>|
|[<span data-ttu-id="d75cb-675">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-676">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-677">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-678">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d75cb-679">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="d75cb-679">Examples</span></span>

<span data-ttu-id="d75cb-680">El código siguiente pasa una cadena a la función `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="d75cb-681">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="d75cb-681">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="d75cb-682">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="d75cb-682">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d75cb-683">Responder con un cuerpo y datos adjuntos de archivo.</span><span class="sxs-lookup"><span data-stu-id="d75cb-683">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="d75cb-684">Responder con un cuerpo y datos adjuntos de elemento.</span><span class="sxs-lookup"><span data-stu-id="d75cb-684">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="d75cb-685">Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="d75cb-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="d75cb-686">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="d75cb-686">displayReplyForm(formData)</span></span>

<span data-ttu-id="d75cb-687">Muestra un formulario de respuesta que incluye solo el remitente del mensaje seleccionado o el organizador de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="d75cb-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d75cb-688">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="d75cb-688">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d75cb-689">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="d75cb-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d75cb-690">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="d75cb-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="d75cb-p145">Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d75cb-694">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="d75cb-694">Parameters:</span></span>

|<span data-ttu-id="d75cb-695">Nombre</span><span class="sxs-lookup"><span data-stu-id="d75cb-695">Name</span></span>| <span data-ttu-id="d75cb-696">Tipo</span><span class="sxs-lookup"><span data-stu-id="d75cb-696">Type</span></span>| <span data-ttu-id="d75cb-697">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="d75cb-698">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d75cb-698">String &#124; Object</span></span>| | <span data-ttu-id="d75cb-p146">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d75cb-701">**O**</span><span class="sxs-lookup"><span data-stu-id="d75cb-701">**OR**</span></span><br/><span data-ttu-id="d75cb-p147">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="d75cb-704">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-704">String</span></span> | <span data-ttu-id="d75cb-705">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-705">&lt;optional&gt;</span></span> | <span data-ttu-id="d75cb-p148">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="d75cb-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d75cb-709">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-709">&lt;optional&gt;</span></span> | <span data-ttu-id="d75cb-710">Una matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="d75cb-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="d75cb-711">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-711">String</span></span> | | <span data-ttu-id="d75cb-p149">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="d75cb-714">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-714">String</span></span> | | <span data-ttu-id="d75cb-715">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="d75cb-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="d75cb-716">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-716">String</span></span> | | <span data-ttu-id="d75cb-p150">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="d75cb-719">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-719">String</span></span> | | <span data-ttu-id="d75cb-p151">Solo se usa si `type` se establece en `item`. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="d75cb-723">función</span><span class="sxs-lookup"><span data-stu-id="d75cb-723">function</span></span> | <span data-ttu-id="d75cb-724">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-724">&lt;optional&gt;</span></span> | <span data-ttu-id="d75cb-725">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d75cb-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d75cb-726">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-726">Requirements</span></span>

|<span data-ttu-id="d75cb-727">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-727">Requirement</span></span>| <span data-ttu-id="d75cb-728">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-729">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-729">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-730">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-730">1.0</span></span>|
|[<span data-ttu-id="d75cb-731">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-732">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-733">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-734">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d75cb-735">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="d75cb-735">Examples</span></span>

<span data-ttu-id="d75cb-736">El código siguiente pasa una cadena a la función `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="d75cb-737">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="d75cb-737">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="d75cb-738">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="d75cb-738">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d75cb-739">Responder con un cuerpo y datos adjuntos de archivo.</span><span class="sxs-lookup"><span data-stu-id="d75cb-739">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="d75cb-740">Responder con un cuerpo y datos adjuntos de elemento.</span><span class="sxs-lookup"><span data-stu-id="d75cb-740">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="d75cb-741">Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="d75cb-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="getentities--entitiesjavascriptapioutlook13officeentities"></a><span data-ttu-id="d75cb-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="d75cb-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span></span>

<span data-ttu-id="d75cb-743">Obtiene las entidades que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="d75cb-743">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d75cb-744">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="d75cb-744">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-745">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-745">Requirements</span></span>

|<span data-ttu-id="d75cb-746">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-746">Requirement</span></span>| <span data-ttu-id="d75cb-747">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-748">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-748">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-749">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-749">1.0</span></span>|
|[<span data-ttu-id="d75cb-750">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-751">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-752">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-753">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d75cb-754">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="d75cb-754">Returns:</span></span>

<span data-ttu-id="d75cb-755">Tipo: [Entidades](/javascript/api/outlook_1_3/office.entities)</span><span class="sxs-lookup"><span data-stu-id="d75cb-755">Type: [Entities](/javascript/api/outlook_1_3/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="d75cb-756">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-756">Example</span></span>

<span data-ttu-id="d75cb-757">En el ejemplo siguiente se obtiene acceso a las entidades de los contactos en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="d75cb-757">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="d75cb-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="d75cb-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="d75cb-759">Obtiene una matriz de todas las entidades del tipo de entidad especificado que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="d75cb-759">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d75cb-760">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="d75cb-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d75cb-761">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="d75cb-761">Parameters:</span></span>

|<span data-ttu-id="d75cb-762">Nombre</span><span class="sxs-lookup"><span data-stu-id="d75cb-762">Name</span></span>| <span data-ttu-id="d75cb-763">Tipo</span><span class="sxs-lookup"><span data-stu-id="d75cb-763">Type</span></span>| <span data-ttu-id="d75cb-764">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="d75cb-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="d75cb-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.entitytype)|<span data-ttu-id="d75cb-766">Uno de los valores de enumeración de EntityType.</span><span class="sxs-lookup"><span data-stu-id="d75cb-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d75cb-767">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-767">Requirements</span></span>

|<span data-ttu-id="d75cb-768">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-768">Requirement</span></span>| <span data-ttu-id="d75cb-769">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-770">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-770">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-771">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-771">1.0</span></span>|
|[<span data-ttu-id="d75cb-772">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-773">Restringido</span><span class="sxs-lookup"><span data-stu-id="d75cb-773">Restricted</span></span>|
|[<span data-ttu-id="d75cb-774">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-775">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d75cb-776">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="d75cb-776">Returns:</span></span>

<span data-ttu-id="d75cb-777">Si el valor que se pasa a `entityType` no es un miembro válido de la enumeración de `EntityType`, el método devuelve un valor NULL.</span><span class="sxs-lookup"><span data-stu-id="d75cb-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="d75cb-778">Si no existen entidades del tipo especificado están presentes en el cuerpo del elemento, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="d75cb-778">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="d75cb-779">De otro modo, el tipo de los objetos en la matriz devuelta depende del tipo de entidad solicitada en el parámetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="d75cb-780">Aunque el nivel de permiso mínimo para usar este método es **Restricted**, algunos tipos de entidad requieren **ReadItem** para tener acceso, como se especifica en la tabla siguiente.</span><span class="sxs-lookup"><span data-stu-id="d75cb-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="d75cb-781">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="d75cb-781">Value of `entityType`</span></span> | <span data-ttu-id="d75cb-782">Tipo de objetos de la matriz devuelta</span><span class="sxs-lookup"><span data-stu-id="d75cb-782">Type of objects in returned array</span></span> | <span data-ttu-id="d75cb-783">Nivel de permiso requerido</span><span class="sxs-lookup"><span data-stu-id="d75cb-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="d75cb-784">Cadena</span><span class="sxs-lookup"><span data-stu-id="d75cb-784">String</span></span> | <span data-ttu-id="d75cb-785">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="d75cb-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="d75cb-786">Contacto</span><span class="sxs-lookup"><span data-stu-id="d75cb-786">Contact</span></span> | <span data-ttu-id="d75cb-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d75cb-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="d75cb-788">Cadena</span><span class="sxs-lookup"><span data-stu-id="d75cb-788">String</span></span> | <span data-ttu-id="d75cb-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d75cb-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="d75cb-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="d75cb-790">MeetingSuggestion</span></span> | <span data-ttu-id="d75cb-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d75cb-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="d75cb-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="d75cb-792">PhoneNumber</span></span> | <span data-ttu-id="d75cb-793">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="d75cb-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="d75cb-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="d75cb-794">TaskSuggestion</span></span> | <span data-ttu-id="d75cb-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d75cb-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="d75cb-796">Cadena</span><span class="sxs-lookup"><span data-stu-id="d75cb-796">String</span></span> | <span data-ttu-id="d75cb-797">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="d75cb-797">**Restricted**</span></span> |

<span data-ttu-id="d75cb-798">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="d75cb-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="d75cb-799">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-799">Example</span></span>

<span data-ttu-id="d75cb-800">En el ejemplo siguiente se muestra cómo obtener acceso a una matriz de cadenas que representan direcciones postales en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="d75cb-800">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="d75cb-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="d75cb-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="d75cb-802">Devuelve entidades conocidas en el elemento seleccionado que pasan el filtro con nombre definido en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="d75cb-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d75cb-803">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="d75cb-803">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d75cb-804">El método `getFilteredEntitiesByName` devuelve las entidades que coinciden con la expresión regular definida en el elemento de regla [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) del archivo XML de manifiesto con el valor de elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="d75cb-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d75cb-805">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="d75cb-805">Parameters:</span></span>

|<span data-ttu-id="d75cb-806">Nombre</span><span class="sxs-lookup"><span data-stu-id="d75cb-806">Name</span></span>| <span data-ttu-id="d75cb-807">Tipo</span><span class="sxs-lookup"><span data-stu-id="d75cb-807">Type</span></span>| <span data-ttu-id="d75cb-808">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="d75cb-809">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-809">String</span></span>|<span data-ttu-id="d75cb-810">El nombre del elemento de regla `ItemHasKnownEntity` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="d75cb-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d75cb-811">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-811">Requirements</span></span>

|<span data-ttu-id="d75cb-812">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-812">Requirement</span></span>| <span data-ttu-id="d75cb-813">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-814">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-814">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-815">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-815">1.0</span></span>|
|[<span data-ttu-id="d75cb-816">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-817">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-818">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-819">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d75cb-820">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="d75cb-820">Returns:</span></span>

<span data-ttu-id="d75cb-p153">Si no existe ningún elemento `ItemHasKnownEntity` en el manifiesto con un valor de elemento `FilterName` que coincida con el parámetro `name`, el método devuelve `null`. Si el parámetro `name` no coincide con un elemento `ItemHasKnownEntity` del manifiesto, pero no existen entidades en el elemento actual que coincidan, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="d75cb-823">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="d75cb-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="d75cb-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d75cb-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="d75cb-825">Devuelve valores de cadena en el elemento seleccionado que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="d75cb-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d75cb-826">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="d75cb-826">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d75cb-p154">El método `getRegExMatches` devuelve las cadenas que coinciden con la expresión regular definida en cada elemento de regla `ItemHasRegularExpressionMatch` o `ItemHasKnownEntity` del archivo XML de manifiesto. Para una regla `ItemHasRegularExpressionMatch`, tiene que darse una cadena coincidente en la propiedad del elemento especificado por la regla. El tipo simple `PropertyName` define las propiedades admitidas.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d75cb-830">Por ejemplo, considere un manifiesto de complemento que tenga el siguiente elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="d75cb-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d75cb-831">El objeto devuelto desde `getRegExMatches` tendría dos propiedades: `fruits` y `veggies`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d75cb-p155">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados. En su lugar, use el método [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) para recuperar todo el cuerpo.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d75cb-835">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-835">Requirements</span></span>

|<span data-ttu-id="d75cb-836">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-836">Requirement</span></span>| <span data-ttu-id="d75cb-837">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-838">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-838">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-839">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-839">1.0</span></span>|
|[<span data-ttu-id="d75cb-840">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-841">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-842">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-843">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d75cb-844">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="d75cb-844">Returns:</span></span>

<span data-ttu-id="d75cb-p156">Un objeto que contiene matrices de cadenas que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. El nombre de cada matriz es igual al valor correspondiente del atributo `RegExName` de la regla `ItemHasRegularExpressionMatch` coincidente o al atributo `FilterName` de la regla `ItemHasKnownEntity` coincidente.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="d75cb-847">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="d75cb-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d75cb-848">Object</span><span class="sxs-lookup"><span data-stu-id="d75cb-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d75cb-849">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-849">Example</span></span>

<span data-ttu-id="d75cb-850">En el siguiente ejemplo, se muestra cómo tener acceso a la matriz de coincidencias de los <rule>elementos de expresión regular `fruits`y `veggies`, que se especifican en el manifiesto.</rule></span><span class="sxs-lookup"><span data-stu-id="d75cb-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="d75cb-851">→ de getregexmatchesbyname (Name) (nullable) {matriz. < cadena >}</span><span class="sxs-lookup"><span data-stu-id="d75cb-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="d75cb-852">Devuelve valores de cadena en el elemento seleccionado que coinciden con la expresión regular de nombre definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="d75cb-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d75cb-853">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="d75cb-853">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d75cb-854">El método `getRegExMatchesByName` devuelve las cadenas que coinciden con la expresión regular definida en el elemento de regla `ItemHasRegularExpressionMatch` del archivo XML de manifiesto con el valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="d75cb-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="d75cb-p157">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d75cb-857">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="d75cb-857">Parameters:</span></span>

|<span data-ttu-id="d75cb-858">Nombre</span><span class="sxs-lookup"><span data-stu-id="d75cb-858">Name</span></span>| <span data-ttu-id="d75cb-859">Tipo</span><span class="sxs-lookup"><span data-stu-id="d75cb-859">Type</span></span>| <span data-ttu-id="d75cb-860">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="d75cb-861">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-861">String</span></span>|<span data-ttu-id="d75cb-862">El nombre del elemento de regla `ItemHasRegularExpressionMatch` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="d75cb-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d75cb-863">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-863">Requirements</span></span>

|<span data-ttu-id="d75cb-864">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-864">Requirement</span></span>| <span data-ttu-id="d75cb-865">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-866">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-866">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-867">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-867">1.0</span></span>|
|[<span data-ttu-id="d75cb-868">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-869">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-870">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-871">Lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d75cb-872">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="d75cb-872">Returns:</span></span>

<span data-ttu-id="d75cb-873">Una matriz que contiene las cadenas que coinciden con la expresión regular definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="d75cb-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="d75cb-874">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="d75cb-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d75cb-875">< Cadena > de matriz.</span><span class="sxs-lookup"><span data-stu-id="d75cb-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d75cb-876">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-876">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="d75cb-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="d75cb-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="d75cb-878">Devuelve asincrónicamente datos seleccionados desde el asunto o el cuerpo de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="d75cb-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="d75cb-p158">Si no hay ninguna selección, pero el cursor está en el cuerpo o el asunto, el método devuelve null para los datos seleccionados. Si se selecciona un campo que no sea el cuerpo o el asunto, el método devuelve el error `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d75cb-881">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="d75cb-881">Parameters:</span></span>

|<span data-ttu-id="d75cb-882">Nombre</span><span class="sxs-lookup"><span data-stu-id="d75cb-882">Name</span></span>| <span data-ttu-id="d75cb-883">Tipo</span><span class="sxs-lookup"><span data-stu-id="d75cb-883">Type</span></span>| <span data-ttu-id="d75cb-884">Atributos</span><span class="sxs-lookup"><span data-stu-id="d75cb-884">Attributes</span></span>| <span data-ttu-id="d75cb-885">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="d75cb-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d75cb-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="d75cb-p159">Solicita un formato para los datos. Si es Texto, el método devuelve el texto sin formato como cadena, quitando toda etiqueta HTML que hubiera. Si es HTML, el método devuelve el texto seleccionado, ya sea texto sin formato o HTML.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="d75cb-890">Object</span><span class="sxs-lookup"><span data-stu-id="d75cb-890">Object</span></span>| <span data-ttu-id="d75cb-891">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-891">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-892">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="d75cb-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d75cb-893">Objeto</span><span class="sxs-lookup"><span data-stu-id="d75cb-893">Object</span></span>| <span data-ttu-id="d75cb-894">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-894">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-895">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="d75cb-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d75cb-896">función</span><span class="sxs-lookup"><span data-stu-id="d75cb-896">function</span></span>||<span data-ttu-id="d75cb-897">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d75cb-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d75cb-898">Para tener acceso a los datos seleccionados desde el método de devolución de llamada, llame a `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="d75cb-899">Para obtener acceso a la propiedad de origen del que proviene la selección, llame a `asyncResult.value.sourceProperty`, que será cualquiera `body` o `subject`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-899">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d75cb-900">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-900">Requirements</span></span>

|<span data-ttu-id="d75cb-901">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-901">Requirement</span></span>| <span data-ttu-id="d75cb-902">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-903">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-903">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-904">1.2</span><span class="sxs-lookup"><span data-stu-id="d75cb-904">1.2</span></span>|
|[<span data-ttu-id="d75cb-905">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="d75cb-907">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-908">Redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="d75cb-909">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="d75cb-909">Returns:</span></span>

<span data-ttu-id="d75cb-910">Los datos seleccionados como cadena con formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="d75cb-911">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="d75cb-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d75cb-912">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d75cb-913">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-913">Example</span></span>

```
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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="d75cb-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d75cb-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="d75cb-915">Carga de forma asincrónica las propiedades personalizadas de este complemento en el elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="d75cb-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="d75cb-p161">Las propiedades personalizadas se almacenan como pares de clave/valor según la aplicación y el elemento. Este método devuelve un objeto `CustomProperties` en la devolución de llamada, que proporciona métodos para tener acceso a las propiedades personalizadas específicas del elemento y el complemento actual. Las propiedades personalizadas no están cifradas en el elemento, por lo que no debería usarse como almacenamiento seguro.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d75cb-919">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="d75cb-919">Parameters:</span></span>

|<span data-ttu-id="d75cb-920">Nombre</span><span class="sxs-lookup"><span data-stu-id="d75cb-920">Name</span></span>| <span data-ttu-id="d75cb-921">Tipo</span><span class="sxs-lookup"><span data-stu-id="d75cb-921">Type</span></span>| <span data-ttu-id="d75cb-922">Atributos</span><span class="sxs-lookup"><span data-stu-id="d75cb-922">Attributes</span></span>| <span data-ttu-id="d75cb-923">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d75cb-924">función</span><span class="sxs-lookup"><span data-stu-id="d75cb-924">function</span></span>||<span data-ttu-id="d75cb-925">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d75cb-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d75cb-926">Las propiedades personalizadas se proporcionan como un objeto [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="d75cb-927">Este objeto se puede utilizar para obtener, establecer y elimina las propiedades personalizadas del elemento y guarde los cambios en la propiedad personalizada que se vuelve a establecer en el servidor.</span><span class="sxs-lookup"><span data-stu-id="d75cb-927">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="d75cb-928">Objeto</span><span class="sxs-lookup"><span data-stu-id="d75cb-928">Object</span></span>| <span data-ttu-id="d75cb-929">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-929">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-930">Los desarrolladores pueden proporcionar cualquier objeto que desean tener acceso en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="d75cb-930">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="d75cb-931">Se puede tener acceso a este objeto por el `asyncResult.asyncContext` (propiedad) en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="d75cb-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d75cb-932">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-932">Requirements</span></span>

|<span data-ttu-id="d75cb-933">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-933">Requirement</span></span>| <span data-ttu-id="d75cb-934">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-935">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-935">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-936">1.0</span><span class="sxs-lookup"><span data-stu-id="d75cb-936">1.0</span></span>|
|[<span data-ttu-id="d75cb-937">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-938">ReadItem</span></span>|
|[<span data-ttu-id="d75cb-939">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-940">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="d75cb-940">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-941">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-941">Example</span></span>

<span data-ttu-id="d75cb-p164">En el siguiente ejemplo de código se muestra cómo usar el método `loadCustomPropertiesAsync` para cargar de forma asincrónica propiedades personalizadas que son específicas del elemento actual. En el ejemplo también se muestra cómo usar el método `CustomProperties.saveAsync` para guardar estas propiedades en el servidor. Después de cargar las propiedades personalizadas, el ejemplo de código usa el método `CustomProperties.get` para leer la propiedad personalizada `myProp`, el método `CustomProperties.set` para escribir en la propiedad personalizada `otherProp` y finalmente se llama al método `saveAsync` para guardar las propiedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="d75cb-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d75cb-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="d75cb-946">Quita los datos adjuntos de un mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="d75cb-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="d75cb-p165">El método `removeAttachmentAsync` quita del elemento los datos adjuntos con el identificador especificado. Como práctica recomendada, debe usar el identificador de datos adjuntos para quitar datos adjuntos solo si la misma aplicación de correo ha agregado los datos adjuntos en la misma sesión. En Outlook Web App y OWA para los dispositivos, el identificador de datos adjuntos es válido solo en la misma sesión. Una sesión finaliza cuando el usuario cierra la aplicación, o si el usuario comienza a redactar en un formulario en línea y posteriormente se extrae el formulario en línea para continuar en una ventana independiente.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d75cb-951">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="d75cb-951">Parameters:</span></span>

|<span data-ttu-id="d75cb-952">Nombre</span><span class="sxs-lookup"><span data-stu-id="d75cb-952">Name</span></span>| <span data-ttu-id="d75cb-953">Tipo</span><span class="sxs-lookup"><span data-stu-id="d75cb-953">Type</span></span>| <span data-ttu-id="d75cb-954">Atributos</span><span class="sxs-lookup"><span data-stu-id="d75cb-954">Attributes</span></span>| <span data-ttu-id="d75cb-955">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="d75cb-956">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-956">String</span></span>||<span data-ttu-id="d75cb-p166">El identificador de los datos adjuntos que se deben quitar. La longitud máxima de la cadena es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p166">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="d75cb-959">Object</span><span class="sxs-lookup"><span data-stu-id="d75cb-959">Object</span></span>| <span data-ttu-id="d75cb-960">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-960">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-961">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="d75cb-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d75cb-962">Objeto</span><span class="sxs-lookup"><span data-stu-id="d75cb-962">Object</span></span>| <span data-ttu-id="d75cb-963">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-963">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-964">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="d75cb-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d75cb-965">función</span><span class="sxs-lookup"><span data-stu-id="d75cb-965">function</span></span>| <span data-ttu-id="d75cb-966">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-966">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-967">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d75cb-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d75cb-968">Si se produce un error en la eliminación de los datos adjuntos, la propiedad `asyncResult.error` contendrá un código de error con el motivo del error.</span><span class="sxs-lookup"><span data-stu-id="d75cb-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d75cb-969">Errores</span><span class="sxs-lookup"><span data-stu-id="d75cb-969">Errors</span></span>

| <span data-ttu-id="d75cb-970">Código de error</span><span class="sxs-lookup"><span data-stu-id="d75cb-970">Error code</span></span> | <span data-ttu-id="d75cb-971">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="d75cb-972">El identificador de datos adjuntos no existe.</span><span class="sxs-lookup"><span data-stu-id="d75cb-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d75cb-973">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-973">Requirements</span></span>

|<span data-ttu-id="d75cb-974">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-974">Requirement</span></span>| <span data-ttu-id="d75cb-975">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-976">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-976">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-977">1.1</span><span class="sxs-lookup"><span data-stu-id="d75cb-977">1.1</span></span>|
|[<span data-ttu-id="d75cb-978">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="d75cb-980">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-981">Redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-982">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-982">Example</span></span>

<span data-ttu-id="d75cb-983">El código siguiente quita un archivo adjunto con un identificador de "0".</span><span class="sxs-lookup"><span data-stu-id="d75cb-983">The following code removes an attachment with an identifier of '0'.</span></span>

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="d75cb-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d75cb-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="d75cb-985">Guarda un elemento de forma asincrónica.</span><span class="sxs-lookup"><span data-stu-id="d75cb-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="d75cb-p167">Cuando se invoca, este método guarda el mensaje actual como un borrador y devuelve el identificador de elemento a través del método de devolución de llamada. En el modo en línea de Outlook Web App o Outlook, el elemento se guarda en el servidor. En modo en caché de Outlook, se guarda el elemento en la caché local.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p167">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="d75cb-989">Si el complemento llama a `saveAsync` en un elemento en modo de redacción con el fin de obtener una `itemId` para usar con EWS o la API de REST, tenga en cuenta que cuando Outlook está en modo en caché, puede tardar algún tiempo antes de que el elemento realmente se sincroniza con el servidor.</span><span class="sxs-lookup"><span data-stu-id="d75cb-989">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="d75cb-990">Hasta que se sincroniza el elemento, el uso de la `itemId` devolverá un error.</span><span class="sxs-lookup"><span data-stu-id="d75cb-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="d75cb-p169">Dado que las citas no tienen ningún estado de borrador, si se llama a `saveAsync` en una cita en modo de redacción, el elemento se guardará como una cita normal en el calendario del usuario. Para las citas nuevas que todavía no se han guardado, no se enviará ninguna invitación. Si se guarda una cita existente, se enviará una actualización a los asistentes agregados o eliminados.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p169">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="d75cb-994">Los siguientes clientes tienen un comportamiento diferente para `saveAsync` en citas en modo de redacción:</span><span class="sxs-lookup"><span data-stu-id="d75cb-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="d75cb-995">Mac Outlook no admite `saveAsync` en una reunión en el modo de redacción.</span><span class="sxs-lookup"><span data-stu-id="d75cb-995">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="d75cb-996">Al llamar a `saveAsync` en una reunión en Mac Outlook devolverá un error.</span><span class="sxs-lookup"><span data-stu-id="d75cb-996">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="d75cb-997">Outlook en el web siempre envía una invitación o actualizar cuándo `saveAsync` se llama en una cita en modo de redacción.</span><span class="sxs-lookup"><span data-stu-id="d75cb-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d75cb-998">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="d75cb-998">Parameters:</span></span>

|<span data-ttu-id="d75cb-999">Nombre</span><span class="sxs-lookup"><span data-stu-id="d75cb-999">Name</span></span>| <span data-ttu-id="d75cb-1000">Tipo</span><span class="sxs-lookup"><span data-stu-id="d75cb-1000">Type</span></span>| <span data-ttu-id="d75cb-1001">Atributos</span><span class="sxs-lookup"><span data-stu-id="d75cb-1001">Attributes</span></span>| <span data-ttu-id="d75cb-1002">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="d75cb-1003">Objeto</span><span class="sxs-lookup"><span data-stu-id="d75cb-1003">Object</span></span>| <span data-ttu-id="d75cb-1004">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-1005">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="d75cb-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d75cb-1006">Objeto</span><span class="sxs-lookup"><span data-stu-id="d75cb-1006">Object</span></span>| <span data-ttu-id="d75cb-1007">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-1008">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="d75cb-1008">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d75cb-1009">función</span><span class="sxs-lookup"><span data-stu-id="d75cb-1009">function</span></span>||<span data-ttu-id="d75cb-1010">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d75cb-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d75cb-1011">En caso de éxito, se proporciona el identificador de elemento en el `asyncResult.value` (propiedad).</span><span class="sxs-lookup"><span data-stu-id="d75cb-1011">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d75cb-1012">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-1012">Requirements</span></span>

|<span data-ttu-id="d75cb-1013">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-1013">Requirement</span></span>| <span data-ttu-id="d75cb-1014">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-1015">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-1015">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="d75cb-1016">1.3</span></span>|
|[<span data-ttu-id="d75cb-1017">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="d75cb-1019">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-1020">Redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d75cb-1021">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="d75cb-1021">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="d75cb-p171">A continuación, se muestra un ejemplo del parámetro `result` que se ha pasado a la función de devolución de llamada. La propiedad `value` contiene el identificador de elemento del elemento.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p171">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="d75cb-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="d75cb-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="d75cb-1025">Inserta asincrónicamente datos en el cuerpo o el asunto de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="d75cb-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="d75cb-p172">El método `setSelectedDataAsync` inserta la cadena especificada en la posición del cursor en el asunto o el cuerpo del elemento, o si se selecciona texto en el editor, reemplaza el texto seleccionado. Si el cursor no está en el cuerpo o el campo del asunto, se devuelve un error. Después de la inserción, el cursor se coloca al final del contenido insertado.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p172">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d75cb-1029">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="d75cb-1029">Parameters:</span></span>

|<span data-ttu-id="d75cb-1030">Nombre</span><span class="sxs-lookup"><span data-stu-id="d75cb-1030">Name</span></span>| <span data-ttu-id="d75cb-1031">Tipo</span><span class="sxs-lookup"><span data-stu-id="d75cb-1031">Type</span></span>| <span data-ttu-id="d75cb-1032">Atributos</span><span class="sxs-lookup"><span data-stu-id="d75cb-1032">Attributes</span></span>| <span data-ttu-id="d75cb-1033">Descripción</span><span class="sxs-lookup"><span data-stu-id="d75cb-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="d75cb-1034">String</span><span class="sxs-lookup"><span data-stu-id="d75cb-1034">String</span></span>||<span data-ttu-id="d75cb-p173">Datos que se van a insertar. Los datos no deben superar 1.000.000 de caracteres. Si se pasan más de 1.000.000 de caracteres, se produce una excepción `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p173">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="d75cb-1038">Object</span><span class="sxs-lookup"><span data-stu-id="d75cb-1038">Object</span></span>| <span data-ttu-id="d75cb-1039">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-1040">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="d75cb-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d75cb-1041">Objeto</span><span class="sxs-lookup"><span data-stu-id="d75cb-1041">Object</span></span>| <span data-ttu-id="d75cb-1042">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-1043">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="d75cb-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="d75cb-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d75cb-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="d75cb-1045">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d75cb-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="d75cb-p174">Si `text`, se aplica el estilo actual en Outlook Web App y Outlook. Si el campo es un editor de HTML, se insertan solo los datos de texto, aunque los datos sean HTML.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p174">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="d75cb-p175">Si `html` y el campo admiten HTML (el asunto no), el estilo actual se aplica en Outlook Web App y el estilo predeterminado en Outlook. Si el campo es un campo de texto, se devuelve un error `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="d75cb-p175">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="d75cb-1050">Si `coercionType` no está establecido, el resultado depende del campo: si el campo es HTML, se usa HTML; si el campo es texto, se usa texto sin formato.</span><span class="sxs-lookup"><span data-stu-id="d75cb-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="d75cb-1051">function</span><span class="sxs-lookup"><span data-stu-id="d75cb-1051">function</span></span>||<span data-ttu-id="d75cb-1052">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d75cb-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d75cb-1053">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d75cb-1053">Requirements</span></span>

|<span data-ttu-id="d75cb-1054">Requirement</span><span class="sxs-lookup"><span data-stu-id="d75cb-1054">Requirement</span></span>| <span data-ttu-id="d75cb-1055">Valor</span><span class="sxs-lookup"><span data-stu-id="d75cb-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="d75cb-1056">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="d75cb-1056">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d75cb-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="d75cb-1057">1.2</span></span>|
|[<span data-ttu-id="d75cb-1058">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="d75cb-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d75cb-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d75cb-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="d75cb-1060">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="d75cb-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d75cb-1061">Redacción</span><span class="sxs-lookup"><span data-stu-id="d75cb-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d75cb-1062">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d75cb-1062">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```