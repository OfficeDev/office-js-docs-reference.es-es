
# <a name="item"></a><span data-ttu-id="6cedd-101">elemento</span><span class="sxs-lookup"><span data-stu-id="6cedd-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="6cedd-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="6cedd-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="6cedd-p101">El espacio de nombres `item` se usa para acceder al mensaje seleccionado actualmente, a una convocatoria de reunión o a una cita. Puede determinar el tipo del `item` mediante la propiedad [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="6cedd-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-105">Requirements</span></span>

|<span data-ttu-id="6cedd-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-106">Requirement</span></span>| <span data-ttu-id="6cedd-107">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-108">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-109">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-109">1.0</span></span>|
|[<span data-ttu-id="6cedd-110">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-111">Restringido</span><span class="sxs-lookup"><span data-stu-id="6cedd-111">Restricted</span></span>|
|[<span data-ttu-id="6cedd-112">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-113">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6cedd-114">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="6cedd-114">Members and methods</span></span>

| <span data-ttu-id="6cedd-115">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-115">Member</span></span> | <span data-ttu-id="6cedd-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="6cedd-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6cedd-117">attachments</span><span class="sxs-lookup"><span data-stu-id="6cedd-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="6cedd-118">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-118">Member</span></span> |
| [<span data-ttu-id="6cedd-119">bcc</span><span class="sxs-lookup"><span data-stu-id="6cedd-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="6cedd-120">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-120">Member</span></span> |
| [<span data-ttu-id="6cedd-121">body</span><span class="sxs-lookup"><span data-stu-id="6cedd-121">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="6cedd-122">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-122">Member</span></span> |
| [<span data-ttu-id="6cedd-123">cc</span><span class="sxs-lookup"><span data-stu-id="6cedd-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="6cedd-124">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-124">Member</span></span> |
| [<span data-ttu-id="6cedd-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="6cedd-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="6cedd-126">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-126">Member</span></span> |
| [<span data-ttu-id="6cedd-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="6cedd-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="6cedd-128">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-128">Member</span></span> |
| [<span data-ttu-id="6cedd-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="6cedd-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="6cedd-130">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-130">Member</span></span> |
| [<span data-ttu-id="6cedd-131">end</span><span class="sxs-lookup"><span data-stu-id="6cedd-131">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="6cedd-132">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-132">Member</span></span> |
| [<span data-ttu-id="6cedd-133">from</span><span class="sxs-lookup"><span data-stu-id="6cedd-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="6cedd-134">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-134">Member</span></span> |
| [<span data-ttu-id="6cedd-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="6cedd-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="6cedd-136">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-136">Member</span></span> |
| [<span data-ttu-id="6cedd-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="6cedd-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="6cedd-138">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-138">Member</span></span> |
| [<span data-ttu-id="6cedd-139">itemId</span><span class="sxs-lookup"><span data-stu-id="6cedd-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="6cedd-140">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-140">Member</span></span> |
| [<span data-ttu-id="6cedd-141">itemType</span><span class="sxs-lookup"><span data-stu-id="6cedd-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="6cedd-142">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-142">Member</span></span> |
| [<span data-ttu-id="6cedd-143">location</span><span class="sxs-lookup"><span data-stu-id="6cedd-143">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="6cedd-144">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-144">Member</span></span> |
| [<span data-ttu-id="6cedd-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="6cedd-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="6cedd-146">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-146">Member</span></span> |
| [<span data-ttu-id="6cedd-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="6cedd-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="6cedd-148">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-148">Member</span></span> |
| [<span data-ttu-id="6cedd-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="6cedd-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="6cedd-150">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-150">Member</span></span> |
| [<span data-ttu-id="6cedd-151">organizer</span><span class="sxs-lookup"><span data-stu-id="6cedd-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="6cedd-152">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-152">Member</span></span> |
| [<span data-ttu-id="6cedd-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="6cedd-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="6cedd-154">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-154">Member</span></span> |
| [<span data-ttu-id="6cedd-155">sender</span><span class="sxs-lookup"><span data-stu-id="6cedd-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="6cedd-156">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-156">Member</span></span> |
| [<span data-ttu-id="6cedd-157">start</span><span class="sxs-lookup"><span data-stu-id="6cedd-157">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="6cedd-158">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-158">Member</span></span> |
| [<span data-ttu-id="6cedd-159">subject</span><span class="sxs-lookup"><span data-stu-id="6cedd-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="6cedd-160">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-160">Member</span></span> |
| [<span data-ttu-id="6cedd-161">to</span><span class="sxs-lookup"><span data-stu-id="6cedd-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="6cedd-162">Miembro	</span><span class="sxs-lookup"><span data-stu-id="6cedd-162">Member</span></span> |
| [<span data-ttu-id="6cedd-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="6cedd-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="6cedd-164">Método</span><span class="sxs-lookup"><span data-stu-id="6cedd-164">Method</span></span> |
| [<span data-ttu-id="6cedd-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="6cedd-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="6cedd-166">Método</span><span class="sxs-lookup"><span data-stu-id="6cedd-166">Method</span></span> |
| [<span data-ttu-id="6cedd-167">close</span><span class="sxs-lookup"><span data-stu-id="6cedd-167">close</span></span>](#close) | <span data-ttu-id="6cedd-168">Método</span><span class="sxs-lookup"><span data-stu-id="6cedd-168">Method</span></span> |
| [<span data-ttu-id="6cedd-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="6cedd-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="6cedd-170">Método</span><span class="sxs-lookup"><span data-stu-id="6cedd-170">Method</span></span> |
| [<span data-ttu-id="6cedd-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="6cedd-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="6cedd-172">Método</span><span class="sxs-lookup"><span data-stu-id="6cedd-172">Method</span></span> |
| [<span data-ttu-id="6cedd-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="6cedd-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="6cedd-174">Método</span><span class="sxs-lookup"><span data-stu-id="6cedd-174">Method</span></span> |
| [<span data-ttu-id="6cedd-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="6cedd-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="6cedd-176">Método</span><span class="sxs-lookup"><span data-stu-id="6cedd-176">Method</span></span> |
| [<span data-ttu-id="6cedd-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="6cedd-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="6cedd-178">Método</span><span class="sxs-lookup"><span data-stu-id="6cedd-178">Method</span></span> |
| [<span data-ttu-id="6cedd-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="6cedd-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="6cedd-180">Método</span><span class="sxs-lookup"><span data-stu-id="6cedd-180">Method</span></span> |
| [<span data-ttu-id="6cedd-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="6cedd-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="6cedd-182">Método</span><span class="sxs-lookup"><span data-stu-id="6cedd-182">Method</span></span> |
| [<span data-ttu-id="6cedd-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="6cedd-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="6cedd-184">Método</span><span class="sxs-lookup"><span data-stu-id="6cedd-184">Method</span></span> |
| [<span data-ttu-id="6cedd-185">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="6cedd-185">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="6cedd-186">Método</span><span class="sxs-lookup"><span data-stu-id="6cedd-186">Method</span></span> |
| [<span data-ttu-id="6cedd-187">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="6cedd-187">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="6cedd-188">Método</span><span class="sxs-lookup"><span data-stu-id="6cedd-188">Method</span></span> |
| [<span data-ttu-id="6cedd-189">saveAsync</span><span class="sxs-lookup"><span data-stu-id="6cedd-189">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="6cedd-190">Método</span><span class="sxs-lookup"><span data-stu-id="6cedd-190">Method</span></span> |
| [<span data-ttu-id="6cedd-191">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="6cedd-191">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="6cedd-192">Método</span><span class="sxs-lookup"><span data-stu-id="6cedd-192">Method</span></span> |

### <a name="example"></a><span data-ttu-id="6cedd-193">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-193">Example</span></span>

<span data-ttu-id="6cedd-194">En el siguiente ejemplo de código de JavaScript, se muestra cómo tener acceso a la propiedad `subject` del elemento actual en Outlook.</span><span class="sxs-lookup"><span data-stu-id="6cedd-194">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="6cedd-195">Miembros</span><span class="sxs-lookup"><span data-stu-id="6cedd-195">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="6cedd-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="6cedd-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="6cedd-p102">Obtiene una matriz de datos adjuntos para el elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6cedd-199">Ciertos tipos de archivos están bloqueados por Outlook debido a problemas de seguridad y, por tanto, no se devuelven.</span><span class="sxs-lookup"><span data-stu-id="6cedd-199">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="6cedd-200">Para obtener más información, vea [datos adjuntos bloqueados en Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="6cedd-200">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-201">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-201">Type:</span></span>

*   <span data-ttu-id="6cedd-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="6cedd-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-203">Requirements</span></span>

|<span data-ttu-id="6cedd-204">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-204">Requirement</span></span>| <span data-ttu-id="6cedd-205">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-206">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-206">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-207">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-207">1.0</span></span>|
|[<span data-ttu-id="6cedd-208">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-209">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-210">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-211">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-211">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-212">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-212">Example</span></span>

<span data-ttu-id="6cedd-213">El siguiente código crea una cadena HTML con los detalles de todos los datos adjuntos en el elemento actual.</span><span class="sxs-lookup"><span data-stu-id="6cedd-213">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="6cedd-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cedd-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="6cedd-215">Obtiene un objeto que proporciona métodos para obtener o actualizar a los destinatarios de la línea de CCO (con copia oculta) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="6cedd-215">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="6cedd-216">Solo modo Redacción.</span><span class="sxs-lookup"><span data-stu-id="6cedd-216">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-217">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-217">Type:</span></span>

*   [<span data-ttu-id="6cedd-218">Destinatarios</span><span class="sxs-lookup"><span data-stu-id="6cedd-218">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="6cedd-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-219">Requirements</span></span>

|<span data-ttu-id="6cedd-220">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-220">Requirement</span></span>| <span data-ttu-id="6cedd-221">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-222">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-222">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-223">1.1</span><span class="sxs-lookup"><span data-stu-id="6cedd-223">1.1</span></span>|
|[<span data-ttu-id="6cedd-224">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-225">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-226">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-227">Redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-228">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-228">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="6cedd-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="6cedd-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="6cedd-230">Obtiene un objeto que proporciona métodos para manipular el cuerpo de un elemento.</span><span class="sxs-lookup"><span data-stu-id="6cedd-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-231">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-231">Type:</span></span>

*   [<span data-ttu-id="6cedd-232">Cuerpo</span><span class="sxs-lookup"><span data-stu-id="6cedd-232">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="6cedd-233">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-233">Requirements</span></span>

|<span data-ttu-id="6cedd-234">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-234">Requirement</span></span>| <span data-ttu-id="6cedd-235">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-236">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-236">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-237">1.1</span><span class="sxs-lookup"><span data-stu-id="6cedd-237">1.1</span></span>|
|[<span data-ttu-id="6cedd-238">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-239">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-240">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-241">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-241">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="6cedd-242">cc: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cedd-242">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="6cedd-243">Proporciona acceso a los destinatarios de Cc (con copia) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="6cedd-243">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="6cedd-244">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="6cedd-244">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cedd-245">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-245">Read mode</span></span>

<span data-ttu-id="6cedd-p106">La propiedad `cc` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **CC** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6cedd-248">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-248">Compose mode</span></span>

<span data-ttu-id="6cedd-249">El `cc` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **Cc** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="6cedd-249">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-250">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-250">Type:</span></span>

*   <span data-ttu-id="6cedd-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cedd-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-252">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-252">Requirements</span></span>

|<span data-ttu-id="6cedd-253">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-253">Requirement</span></span>| <span data-ttu-id="6cedd-254">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-255">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-255">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-256">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-256">1.0</span></span>|
|[<span data-ttu-id="6cedd-257">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-257">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-258">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-259">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-259">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-260">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-260">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-261">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-261">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="6cedd-262">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="6cedd-262">(nullable) conversationId :String</span></span>

<span data-ttu-id="6cedd-263">Obtiene un identificador de la conversación de correo electrónico que contiene un mensaje determinado.</span><span class="sxs-lookup"><span data-stu-id="6cedd-263">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="6cedd-p107">Puede obtener un entero para esta propiedad si la aplicación de correo se activa en respuestas o formularios de lectura o en formularios de redacción. Si, después, el usuario cambia el asunto del mensaje de respuesta, al enviar dicha respuesta, el Id. de conversación de ese mensaje cambiará y el valor que se obtuvo anteriormente ya no será de aplicación.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="6cedd-p108">Obtiene un valor NULL para esta propiedad para un nuevo elemento de un formulario de redacción. Si el usuario establece un asunto y guarda el elemento, la propiedad `conversationId` devolverá un valor.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-268">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-268">Type:</span></span>

*   <span data-ttu-id="6cedd-269">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-269">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-270">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-270">Requirements</span></span>

|<span data-ttu-id="6cedd-271">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-271">Requirement</span></span>| <span data-ttu-id="6cedd-272">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-273">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-274">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-274">1.0</span></span>|
|[<span data-ttu-id="6cedd-275">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-276">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-277">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-278">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-278">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="6cedd-279">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="6cedd-279">dateTimeCreated :Date</span></span>

<span data-ttu-id="6cedd-p109">Obtiene la fecha y hora de creación de un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-282">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-282">Type:</span></span>

*   <span data-ttu-id="6cedd-283">Fecha</span><span class="sxs-lookup"><span data-stu-id="6cedd-283">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-284">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-284">Requirements</span></span>

|<span data-ttu-id="6cedd-285">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-285">Requirement</span></span>| <span data-ttu-id="6cedd-286">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-287">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-287">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-288">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-288">1.0</span></span>|
|[<span data-ttu-id="6cedd-289">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-289">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-290">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-290">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-291">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-291">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-292">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-292">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-293">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-293">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="6cedd-294">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="6cedd-294">dateTimeModified :Date</span></span>

<span data-ttu-id="6cedd-p110">Obtiene la fecha y hora en que se modificó por última vez un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6cedd-297">Este miembro no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="6cedd-297">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-298">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-298">Type:</span></span>

*   <span data-ttu-id="6cedd-299">Fecha</span><span class="sxs-lookup"><span data-stu-id="6cedd-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-300">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-300">Requirements</span></span>

|<span data-ttu-id="6cedd-301">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-301">Requirement</span></span>| <span data-ttu-id="6cedd-302">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-303">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-304">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-304">1.0</span></span>|
|[<span data-ttu-id="6cedd-305">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-306">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-307">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-308">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-309">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-309">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="6cedd-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="6cedd-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="6cedd-311">Obtiene o establece la fecha y la hora de finalización de la cita.</span><span class="sxs-lookup"><span data-stu-id="6cedd-311">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="6cedd-p111">La propiedad `end` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) para convertir el valor de la propiedad final a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cedd-314">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-314">Read mode</span></span>

<span data-ttu-id="6cedd-315">La propiedad `end` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-315">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6cedd-316">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-316">Compose mode</span></span>

<span data-ttu-id="6cedd-317">La propiedad `end` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-317">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="6cedd-318">Si usa el método [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) para establecer la hora de finalización, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="6cedd-318">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-319">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-319">Type:</span></span>

*   <span data-ttu-id="6cedd-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="6cedd-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-321">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-321">Requirements</span></span>

|<span data-ttu-id="6cedd-322">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-322">Requirement</span></span>| <span data-ttu-id="6cedd-323">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-324">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-324">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-325">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-325">1.0</span></span>|
|[<span data-ttu-id="6cedd-326">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-327">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-328">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-329">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-329">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-330">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-330">Example</span></span>

<span data-ttu-id="6cedd-331">En el ejemplo siguiente, se establece la hora de finalización de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-331">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="6cedd-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="6cedd-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="6cedd-p112">Obtiene la dirección de correo electrónico del remitente de un mensaje. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="6cedd-p113">Las propiedades `from` y [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="6cedd-337">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `from` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-337">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-338">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-338">Type:</span></span>

*   [<span data-ttu-id="6cedd-339">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6cedd-339">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="6cedd-340">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-340">Requirements</span></span>

|<span data-ttu-id="6cedd-341">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-341">Requirement</span></span>| <span data-ttu-id="6cedd-342">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-342">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-343">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-343">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-344">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-344">1.0</span></span>|
|[<span data-ttu-id="6cedd-345">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-345">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-346">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-346">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-347">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-347">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-348">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-348">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="6cedd-349">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="6cedd-349">internetMessageId :String</span></span>

<span data-ttu-id="6cedd-p114">Obtiene el identificador de mensaje de Internet de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-352">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-352">Type:</span></span>

*   <span data-ttu-id="6cedd-353">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-353">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-354">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-354">Requirements</span></span>

|<span data-ttu-id="6cedd-355">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-355">Requirement</span></span>| <span data-ttu-id="6cedd-356">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-357">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-357">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-358">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-358">1.0</span></span>|
|[<span data-ttu-id="6cedd-359">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-360">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-361">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-362">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-362">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-363">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-363">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="6cedd-364">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="6cedd-364">itemClass :String</span></span>

<span data-ttu-id="6cedd-p115">Obtiene la clase de elemento Servicios Web Exchange del elemento seleccionado. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="6cedd-p116">La propiedad `itemClass` especifica la clase de mensaje del elemento seleccionado. Las siguientes son las clases de mensaje predeterminadas para el elemento de mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="6cedd-369">Tipo</span><span class="sxs-lookup"><span data-stu-id="6cedd-369">Type</span></span> | <span data-ttu-id="6cedd-370">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-370">Description</span></span> | <span data-ttu-id="6cedd-371">Clase de elemento</span><span class="sxs-lookup"><span data-stu-id="6cedd-371">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="6cedd-372">Elementos de cita</span><span class="sxs-lookup"><span data-stu-id="6cedd-372">Appointment items</span></span> | <span data-ttu-id="6cedd-373">Estos son los elementos de calendario de la clase de elemento `IPM.Appointment` o `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-373">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="6cedd-374">Elementos de mensaje</span><span class="sxs-lookup"><span data-stu-id="6cedd-374">Message items</span></span> | <span data-ttu-id="6cedd-375">Entre estos se encuentran los mensajes de correo electrónico que tienen la clase de mensaje predeterminada `IPM.Note` y las convocatorias de reunión, respuestas y cancelaciones que usan `IPM.Schedule.Meeting` como clase de mensaje base.</span><span class="sxs-lookup"><span data-stu-id="6cedd-375">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="6cedd-376">Puede crear clases de mensaje personalizadas para ampliar una clase de mensaje predeterminada. Por ejemplo, puede crear una clase de mensaje de cita personalizada `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-376">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-377">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-377">Type:</span></span>

*   <span data-ttu-id="6cedd-378">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-378">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-379">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-379">Requirements</span></span>

|<span data-ttu-id="6cedd-380">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-380">Requirement</span></span>| <span data-ttu-id="6cedd-381">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-381">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-382">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-382">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-383">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-383">1.0</span></span>|
|[<span data-ttu-id="6cedd-384">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-384">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-385">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-386">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-386">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-387">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-387">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-388">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-388">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="6cedd-389">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="6cedd-389">(nullable) itemId :String</span></span>

<span data-ttu-id="6cedd-p117">Obtiene el identificador de elemento de Servicios Web Exchange para el elemento actual. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6cedd-392">El identificador que devuelve la propiedad `itemId` es el mismo que el identificador de elemento de Servicios Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="6cedd-392">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="6cedd-393">El `itemId` (propiedad) no es idéntico al identificador de entrada de Outlook o el identificador de la API de REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="6cedd-393">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="6cedd-394">Antes de realizar llamadas a la API de REST con este valor, se debe convertir mediante [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="6cedd-394">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="6cedd-395">Para obtener más información, vea [usar las API de REST de Outlook desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="6cedd-395">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="6cedd-p119">La propiedad `itemId` no está disponible en el modo redacción. Si se requiere un identificador de elemento, el método [`saveAsync`](#saveasyncoptions-callback) puede usarse para guardar el elemento en el servidor, que devolverá el identificador de elemento en el parámetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) de la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-398">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-398">Type:</span></span>

*   <span data-ttu-id="6cedd-399">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-400">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-400">Requirements</span></span>

|<span data-ttu-id="6cedd-401">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-401">Requirement</span></span>| <span data-ttu-id="6cedd-402">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-403">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-403">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-404">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-404">1.0</span></span>|
|[<span data-ttu-id="6cedd-405">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-406">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-407">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-408">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-409">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-409">Example</span></span>

<span data-ttu-id="6cedd-p120">El siguiente código comprueba la presencia de un identificador de elemento. Si la propiedad `itemId` devuelve `null` o `undefined`, guarda el elemento en el servidor y obtiene el identificador de elemento desde el resultado asincrónico.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="6cedd-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="6cedd-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="6cedd-413">Obtiene el tipo de elemento que representa una instancia.</span><span class="sxs-lookup"><span data-stu-id="6cedd-413">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="6cedd-414">La propiedad `itemType` devuelve uno de los valores de enumeración de `ItemType`, lo que indica si la instancia del objeto `item` es un mensaje o una cita.</span><span class="sxs-lookup"><span data-stu-id="6cedd-414">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-415">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-415">Type:</span></span>

*   [<span data-ttu-id="6cedd-416">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="6cedd-416">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="6cedd-417">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-417">Requirements</span></span>

|<span data-ttu-id="6cedd-418">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-418">Requirement</span></span>| <span data-ttu-id="6cedd-419">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-420">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-420">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-421">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-421">1.0</span></span>|
|[<span data-ttu-id="6cedd-422">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-423">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-424">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-425">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-425">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-426">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-426">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="6cedd-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="6cedd-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="6cedd-428">Obtiene o establece la ubicación de una cita.</span><span class="sxs-lookup"><span data-stu-id="6cedd-428">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cedd-429">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-429">Read mode</span></span>

<span data-ttu-id="6cedd-430">La propiedad `location` devuelve una cadena que contiene la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="6cedd-430">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6cedd-431">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-431">Compose mode</span></span>

<span data-ttu-id="6cedd-432">La propiedad `location` devuelve un objeto `Location` que proporciona métodos para obtener y establecer la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="6cedd-432">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-433">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-433">Type:</span></span>

*   <span data-ttu-id="6cedd-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="6cedd-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-435">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-435">Requirements</span></span>

|<span data-ttu-id="6cedd-436">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-436">Requirement</span></span>| <span data-ttu-id="6cedd-437">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-438">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-438">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-439">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-439">1.0</span></span>|
|[<span data-ttu-id="6cedd-440">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-441">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-442">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-443">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-444">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-444">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="6cedd-445">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="6cedd-445">normalizedSubject :String</span></span>

<span data-ttu-id="6cedd-p121">Obtiene el asunto de un elemento, con todos los prefijos quitados (incluidos `RE:` y `FWD:`). Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="6cedd-p122">La propiedad normalizedSubject obtiene el asunto del elemento, con los prefijos estándar (como `RE:` y `FW:`) agregados por los programas de correo electrónico. Para obtener el asunto del elemento con los prefijos intactos, use la propiedad [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject).</span><span class="sxs-lookup"><span data-stu-id="6cedd-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-450">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-450">Type:</span></span>

*   <span data-ttu-id="6cedd-451">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-451">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-452">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-452">Requirements</span></span>

|<span data-ttu-id="6cedd-453">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-453">Requirement</span></span>| <span data-ttu-id="6cedd-454">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-455">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-455">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-456">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-456">1.0</span></span>|
|[<span data-ttu-id="6cedd-457">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-458">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-459">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-460">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-460">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-461">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-461">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="6cedd-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="6cedd-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="6cedd-463">Obtiene los mensajes de notificación de un elemento.</span><span class="sxs-lookup"><span data-stu-id="6cedd-463">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-464">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-464">Type:</span></span>

*   [<span data-ttu-id="6cedd-465">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="6cedd-465">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="6cedd-466">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-466">Requirements</span></span>

|<span data-ttu-id="6cedd-467">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-467">Requirement</span></span>| <span data-ttu-id="6cedd-468">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-469">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-469">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-470">1.3</span><span class="sxs-lookup"><span data-stu-id="6cedd-470">1.3</span></span>|
|[<span data-ttu-id="6cedd-471">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-472">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-473">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-474">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-474">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="6cedd-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cedd-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="6cedd-476">Proporciona acceso a los asistentes opcionales de un evento.</span><span class="sxs-lookup"><span data-stu-id="6cedd-476">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="6cedd-477">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="6cedd-477">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cedd-478">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-478">Read mode</span></span>

<span data-ttu-id="6cedd-479">La propiedad `optionalAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente opcional a la reunión.</span><span class="sxs-lookup"><span data-stu-id="6cedd-479">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6cedd-480">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-480">Compose mode</span></span>

<span data-ttu-id="6cedd-481">El `optionalAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes opcionales de una reunión.</span><span class="sxs-lookup"><span data-stu-id="6cedd-481">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-482">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-482">Type:</span></span>

*   <span data-ttu-id="6cedd-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cedd-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-484">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-484">Requirements</span></span>

|<span data-ttu-id="6cedd-485">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-485">Requirement</span></span>| <span data-ttu-id="6cedd-486">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-487">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-487">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-488">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-488">1.0</span></span>|
|[<span data-ttu-id="6cedd-489">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-490">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-491">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-492">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-492">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-493">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-493">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="6cedd-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="6cedd-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="6cedd-p124">Obtiene la dirección de correo electrónico del organizador de una reunión especificada. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-497">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-497">Type:</span></span>

*   [<span data-ttu-id="6cedd-498">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6cedd-498">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="6cedd-499">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-499">Requirements</span></span>

|<span data-ttu-id="6cedd-500">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-500">Requirement</span></span>| <span data-ttu-id="6cedd-501">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-502">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-502">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-503">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-503">1.0</span></span>|
|[<span data-ttu-id="6cedd-504">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-505">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-506">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-507">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-508">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-508">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="6cedd-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cedd-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="6cedd-510">Proporciona acceso a los asistentes necesarios de un evento.</span><span class="sxs-lookup"><span data-stu-id="6cedd-510">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="6cedd-511">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="6cedd-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cedd-512">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-512">Read mode</span></span>

<span data-ttu-id="6cedd-513">La propiedad `requiredAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente requerido a la reunión.</span><span class="sxs-lookup"><span data-stu-id="6cedd-513">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6cedd-514">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-514">Compose mode</span></span>

<span data-ttu-id="6cedd-515">El `requiredAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes necesarios de una reunión.</span><span class="sxs-lookup"><span data-stu-id="6cedd-515">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-516">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-516">Type:</span></span>

*   <span data-ttu-id="6cedd-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cedd-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-518">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-518">Requirements</span></span>

|<span data-ttu-id="6cedd-519">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-519">Requirement</span></span>| <span data-ttu-id="6cedd-520">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-521">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-522">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-522">1.0</span></span>|
|[<span data-ttu-id="6cedd-523">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-524">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-525">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-526">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-527">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-527">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="6cedd-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="6cedd-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="6cedd-p126">Obtiene la dirección de correo electrónico del remitente de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="6cedd-p127">Las propiedades [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) y `sender` representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="6cedd-533">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `sender` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-533">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-534">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-534">Type:</span></span>

*   [<span data-ttu-id="6cedd-535">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6cedd-535">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="6cedd-536">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-536">Requirements</span></span>

|<span data-ttu-id="6cedd-537">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-537">Requirement</span></span>| <span data-ttu-id="6cedd-538">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-539">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-539">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-540">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-540">1.0</span></span>|
|[<span data-ttu-id="6cedd-541">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-541">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-542">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-543">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-543">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-544">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-545">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-545">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="6cedd-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="6cedd-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="6cedd-547">Obtiene o establece la fecha y la hora de inicio de la cita.</span><span class="sxs-lookup"><span data-stu-id="6cedd-547">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="6cedd-p128">La propiedad `start` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) para convertir el valor a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cedd-550">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-550">Read mode</span></span>

<span data-ttu-id="6cedd-551">La propiedad `start` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-551">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6cedd-552">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-552">Compose mode</span></span>

<span data-ttu-id="6cedd-553">La propiedad `start` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-553">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="6cedd-554">Si usa el método [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) para establecer la hora de inicio, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="6cedd-554">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-555">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-555">Type:</span></span>

*   <span data-ttu-id="6cedd-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="6cedd-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-557">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-557">Requirements</span></span>

|<span data-ttu-id="6cedd-558">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-558">Requirement</span></span>| <span data-ttu-id="6cedd-559">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-559">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-560">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-560">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-561">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-561">1.0</span></span>|
|[<span data-ttu-id="6cedd-562">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-562">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-563">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-563">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-564">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-564">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-565">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-565">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-566">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-566">Example</span></span>

<span data-ttu-id="6cedd-567">En el ejemplo siguiente, se establece la hora de inicio de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-567">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="6cedd-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="6cedd-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="6cedd-569">Obtiene o establece la descripción que se muestra en el campo de asunto de un elemento.</span><span class="sxs-lookup"><span data-stu-id="6cedd-569">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="6cedd-570">La propiedad `subject` obtiene o establece el asunto completo del elemento, como lo ha enviado el servidor de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="6cedd-570">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cedd-571">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-571">Read mode</span></span>

<span data-ttu-id="6cedd-p129">La propiedad `subject` devuelve una cadena. Use la propiedad [`normalizedSubject`](#normalizedsubject-string) para obtener el asunto menos cualquier prefijo inicial como `RE:` y `FW:`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="6cedd-574">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-574">Compose mode</span></span>

<span data-ttu-id="6cedd-575">La propiedad `subject` devuelve un objeto `Subject` que proporciona métodos para obtener y establecer el asunto.</span><span class="sxs-lookup"><span data-stu-id="6cedd-575">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="6cedd-576">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-576">Type:</span></span>

*   <span data-ttu-id="6cedd-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="6cedd-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-578">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-578">Requirements</span></span>

|<span data-ttu-id="6cedd-579">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-579">Requirement</span></span>| <span data-ttu-id="6cedd-580">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-581">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-581">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-582">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-582">1.0</span></span>|
|[<span data-ttu-id="6cedd-583">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-584">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-585">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-586">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-586">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="6cedd-587">para: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cedd-587">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="6cedd-588">Proporciona acceso a los destinatarios en la línea **para** de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="6cedd-588">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="6cedd-589">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="6cedd-589">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cedd-590">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-590">Read mode</span></span>

<span data-ttu-id="6cedd-p131">La propiedad `to` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **Para** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6cedd-593">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-593">Compose mode</span></span>

<span data-ttu-id="6cedd-594">El `to` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **para** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="6cedd-594">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="6cedd-595">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6cedd-595">Type:</span></span>

*   <span data-ttu-id="6cedd-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cedd-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-597">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-597">Requirements</span></span>

|<span data-ttu-id="6cedd-598">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-598">Requirement</span></span>| <span data-ttu-id="6cedd-599">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-599">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-600">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-600">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-601">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-601">1.0</span></span>|
|[<span data-ttu-id="6cedd-602">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-602">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-603">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-603">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-604">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-604">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-605">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-605">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-606">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-606">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="6cedd-607">Métodos</span><span class="sxs-lookup"><span data-stu-id="6cedd-607">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="6cedd-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6cedd-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="6cedd-609">Agrega un archivo a un mensaje o cita como datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="6cedd-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="6cedd-610">El método `addFileAttachmentAsync` carga el archivo en el URI especificado y lo asocia al elemento en el formulario de redacción.</span><span class="sxs-lookup"><span data-stu-id="6cedd-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="6cedd-611">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="6cedd-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cedd-612">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="6cedd-612">Parameters:</span></span>

|<span data-ttu-id="6cedd-613">Nombre</span><span class="sxs-lookup"><span data-stu-id="6cedd-613">Name</span></span>| <span data-ttu-id="6cedd-614">Tipo</span><span class="sxs-lookup"><span data-stu-id="6cedd-614">Type</span></span>| <span data-ttu-id="6cedd-615">Atributos</span><span class="sxs-lookup"><span data-stu-id="6cedd-615">Attributes</span></span>| <span data-ttu-id="6cedd-616">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="6cedd-617">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-617">String</span></span>||<span data-ttu-id="6cedd-p132">El URI que proporciona la ubicación del archivo que se va a adjuntar al mensaje o a la cita. La longitud máxima es de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="6cedd-620">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-620">String</span></span>||<span data-ttu-id="6cedd-p133">El nombre de los datos adjuntos que se muestra mientras estos se cargan. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="6cedd-623">Object</span><span class="sxs-lookup"><span data-stu-id="6cedd-623">Object</span></span>| <span data-ttu-id="6cedd-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-624">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-625">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="6cedd-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="6cedd-626">Objeto</span><span class="sxs-lookup"><span data-stu-id="6cedd-626">Object</span></span> | <span data-ttu-id="6cedd-627">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-627">&lt;optional&gt;</span></span> | <span data-ttu-id="6cedd-628">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="6cedd-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="6cedd-629">Booleano</span><span class="sxs-lookup"><span data-stu-id="6cedd-629">Boolean</span></span> | <span data-ttu-id="6cedd-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-630">&lt;optional&gt;</span></span> | <span data-ttu-id="6cedd-631">Si está establecido en `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="6cedd-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="6cedd-632">Función</span><span class="sxs-lookup"><span data-stu-id="6cedd-632">function</span></span>| <span data-ttu-id="6cedd-633">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-633">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-634">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cedd-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6cedd-635">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="6cedd-636">Si se produce un error en la carga de los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="6cedd-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6cedd-637">Errores</span><span class="sxs-lookup"><span data-stu-id="6cedd-637">Errors</span></span>

| <span data-ttu-id="6cedd-638">Código de error</span><span class="sxs-lookup"><span data-stu-id="6cedd-638">Error code</span></span> | <span data-ttu-id="6cedd-639">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="6cedd-640">Los datos adjuntos son más grandes de lo permitido.</span><span class="sxs-lookup"><span data-stu-id="6cedd-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="6cedd-641">Los datos adjuntos tienen una extensión que no está permitida.</span><span class="sxs-lookup"><span data-stu-id="6cedd-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="6cedd-642">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="6cedd-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6cedd-643">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-643">Requirements</span></span>

|<span data-ttu-id="6cedd-644">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-644">Requirement</span></span>| <span data-ttu-id="6cedd-645">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-646">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-646">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-647">1.1</span><span class="sxs-lookup"><span data-stu-id="6cedd-647">1.1</span></span>|
|[<span data-ttu-id="6cedd-648">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-648">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="6cedd-650">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-650">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-651">Redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="6cedd-652">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="6cedd-652">Examples</span></span>

```js
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

<span data-ttu-id="6cedd-653">En el siguiente ejemplo, se agrega un archivo de imagen como un dato adjunto en línea y hace referencia a los datos adjuntos del cuerpo del mensaje.</span><span class="sxs-lookup"><span data-stu-id="6cedd-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="6cedd-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6cedd-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="6cedd-655">Agrega un elemento de Exchange (por ejemplo, un mensaje) como datos adjuntos al mensaje o a la cita.</span><span class="sxs-lookup"><span data-stu-id="6cedd-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="6cedd-p134">El método `addItemAttachmentAsync` asocia el elemento con el identificador especificado de Exchange al elemento en el formulario de redacción. Si se especifica un método de devolución de llamada, se llama al método con un parámetro, `asyncResult`, que contiene el identificador de datos adjuntos o un código que indique cualquier error que se ha producido al adjuntar el elemento. Si es necesario, puede usar el parámetro `options` para pasar información de estado al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="6cedd-659">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="6cedd-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="6cedd-660">Si el complemento de Office se está ejecutando en Outlook Web App, el `addItemAttachmentAsync` método puede adjuntar elementos a los elementos que no sea el elemento que se está editando; Sin embargo, esto no se admite y no se recomienda.</span><span class="sxs-lookup"><span data-stu-id="6cedd-660">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cedd-661">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="6cedd-661">Parameters:</span></span>

|<span data-ttu-id="6cedd-662">Nombre</span><span class="sxs-lookup"><span data-stu-id="6cedd-662">Name</span></span>| <span data-ttu-id="6cedd-663">Tipo</span><span class="sxs-lookup"><span data-stu-id="6cedd-663">Type</span></span>| <span data-ttu-id="6cedd-664">Atributos</span><span class="sxs-lookup"><span data-stu-id="6cedd-664">Attributes</span></span>| <span data-ttu-id="6cedd-665">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="6cedd-666">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-666">String</span></span>||<span data-ttu-id="6cedd-p135">El identificador de Exchange del elemento que debe adjuntarse. La longitud máxima es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="6cedd-669">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-669">String</span></span>||<span data-ttu-id="6cedd-p136">El asunto del elemento que debe adjuntarse. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="6cedd-672">Object</span><span class="sxs-lookup"><span data-stu-id="6cedd-672">Object</span></span>| <span data-ttu-id="6cedd-673">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-673">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-674">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="6cedd-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6cedd-675">Objeto</span><span class="sxs-lookup"><span data-stu-id="6cedd-675">Object</span></span>| <span data-ttu-id="6cedd-676">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-676">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-677">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="6cedd-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6cedd-678">función</span><span class="sxs-lookup"><span data-stu-id="6cedd-678">function</span></span>| <span data-ttu-id="6cedd-679">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-679">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-680">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cedd-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6cedd-681">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="6cedd-682">Si se produce un error al agregar los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="6cedd-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6cedd-683">Errores</span><span class="sxs-lookup"><span data-stu-id="6cedd-683">Errors</span></span>

| <span data-ttu-id="6cedd-684">Código de error</span><span class="sxs-lookup"><span data-stu-id="6cedd-684">Error code</span></span> | <span data-ttu-id="6cedd-685">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="6cedd-686">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="6cedd-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6cedd-687">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-687">Requirements</span></span>

|<span data-ttu-id="6cedd-688">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-688">Requirement</span></span>| <span data-ttu-id="6cedd-689">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-690">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-690">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-691">1.1</span><span class="sxs-lookup"><span data-stu-id="6cedd-691">1.1</span></span>|
|[<span data-ttu-id="6cedd-692">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-692">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="6cedd-694">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-694">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-695">Redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-696">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-696">Example</span></span>

<span data-ttu-id="6cedd-697">En el ejemplo siguiente, se agrega un elemento existente de Outlook como un archivo adjunto con el nombre `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="6cedd-698">close()</span><span class="sxs-lookup"><span data-stu-id="6cedd-698">close()</span></span>

<span data-ttu-id="6cedd-699">Cierra el elemento actual que se está redactando.</span><span class="sxs-lookup"><span data-stu-id="6cedd-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="6cedd-p137">El comportamiento del método `close` depende del estado actual del elemento que se está redactando. Si el elemento tiene cambios sin guardar, el cliente solicita al usuario que guarde, descarte o cancele la acción Cerrar.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="6cedd-702">En Outlook, en el sitio web, si el elemento es una cita y anteriormente se ha guardado con `saveAsync`, se pide al usuario al guardar, descartar o cancelar incluso si no ha habido cambios desde la último el elemento guardado.</span><span class="sxs-lookup"><span data-stu-id="6cedd-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="6cedd-703">En el cliente de escritorio de Outlook, si el mensaje es una respuesta directa, el método `close` no tiene ningún efecto.</span><span class="sxs-lookup"><span data-stu-id="6cedd-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-704">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-704">Requirements</span></span>

|<span data-ttu-id="6cedd-705">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-705">Requirement</span></span>| <span data-ttu-id="6cedd-706">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-707">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-707">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-708">1.3</span><span class="sxs-lookup"><span data-stu-id="6cedd-708">1.3</span></span>|
|[<span data-ttu-id="6cedd-709">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-709">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-710">Restringido</span><span class="sxs-lookup"><span data-stu-id="6cedd-710">Restricted</span></span>|
|[<span data-ttu-id="6cedd-711">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-711">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-712">Redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-712">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="6cedd-713">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="6cedd-713">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="6cedd-714">Muestra un formulario de respuesta que incluye el remitente y todos los destinatarios del mensaje seleccionado o el organizador y todos los asistentes de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="6cedd-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6cedd-715">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="6cedd-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6cedd-716">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="6cedd-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="6cedd-717">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyAllForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="6cedd-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="6cedd-p138">Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cedd-721">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="6cedd-721">Parameters:</span></span>

| <span data-ttu-id="6cedd-722">Nombre</span><span class="sxs-lookup"><span data-stu-id="6cedd-722">Name</span></span> | <span data-ttu-id="6cedd-723">Tipo</span><span class="sxs-lookup"><span data-stu-id="6cedd-723">Type</span></span> | <span data-ttu-id="6cedd-724">Atributos</span><span class="sxs-lookup"><span data-stu-id="6cedd-724">Attributes</span></span> | <span data-ttu-id="6cedd-725">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="6cedd-726">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="6cedd-726">String &#124; Object</span></span>| |<span data-ttu-id="6cedd-p139">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="6cedd-729">**O**</span><span class="sxs-lookup"><span data-stu-id="6cedd-729">**OR**</span></span><br/><span data-ttu-id="6cedd-p140">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="6cedd-732">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-732">String</span></span> | <span data-ttu-id="6cedd-733">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-733">&lt;optional&gt;</span></span> | <span data-ttu-id="6cedd-p141">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="6cedd-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="6cedd-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-737">&lt;optional&gt;</span></span> | <span data-ttu-id="6cedd-738">Una matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="6cedd-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="6cedd-739">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-739">String</span></span> | | <span data-ttu-id="6cedd-p142">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="6cedd-742">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-742">String</span></span> | | <span data-ttu-id="6cedd-743">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="6cedd-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="6cedd-744">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-744">String</span></span> | | <span data-ttu-id="6cedd-p143">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="6cedd-747">Booleano</span><span class="sxs-lookup"><span data-stu-id="6cedd-747">Boolean</span></span> | | <span data-ttu-id="6cedd-p144">Solo se usa si `type` se establece en `file`. Si `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="6cedd-750">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-750">String</span></span> | | <span data-ttu-id="6cedd-p145">Solo se usa si `type` se establece en `item`. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="6cedd-754">función</span><span class="sxs-lookup"><span data-stu-id="6cedd-754">function</span></span> | <span data-ttu-id="6cedd-755">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-755">&lt;optional&gt;</span></span> | <span data-ttu-id="6cedd-756">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cedd-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6cedd-757">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-757">Requirements</span></span>

|<span data-ttu-id="6cedd-758">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-758">Requirement</span></span>| <span data-ttu-id="6cedd-759">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-760">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-760">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-761">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-761">1.0</span></span>|
|[<span data-ttu-id="6cedd-762">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-762">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-763">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-764">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-764">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-765">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="6cedd-766">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="6cedd-766">Examples</span></span>

<span data-ttu-id="6cedd-767">El código siguiente pasa una cadena a la función `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="6cedd-768">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="6cedd-768">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="6cedd-769">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="6cedd-769">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="6cedd-770">Responder con un cuerpo y datos adjuntos de archivo.</span><span class="sxs-lookup"><span data-stu-id="6cedd-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="6cedd-771">Responder con un cuerpo y datos adjuntos de elemento.</span><span class="sxs-lookup"><span data-stu-id="6cedd-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="6cedd-772">Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="6cedd-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="6cedd-773">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="6cedd-773">displayReplyForm(formData)</span></span>

<span data-ttu-id="6cedd-774">Muestra un formulario de respuesta que incluye solo el remitente del mensaje seleccionado o el organizador de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="6cedd-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6cedd-775">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="6cedd-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6cedd-776">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="6cedd-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="6cedd-777">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="6cedd-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="6cedd-p146">Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cedd-781">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="6cedd-781">Parameters:</span></span>

| <span data-ttu-id="6cedd-782">Nombre</span><span class="sxs-lookup"><span data-stu-id="6cedd-782">Name</span></span> | <span data-ttu-id="6cedd-783">Tipo</span><span class="sxs-lookup"><span data-stu-id="6cedd-783">Type</span></span> | <span data-ttu-id="6cedd-784">Atributos</span><span class="sxs-lookup"><span data-stu-id="6cedd-784">Attributes</span></span> | <span data-ttu-id="6cedd-785">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="6cedd-786">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="6cedd-786">String &#124; Object</span></span>| | <span data-ttu-id="6cedd-p147">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="6cedd-789">**O**</span><span class="sxs-lookup"><span data-stu-id="6cedd-789">**OR**</span></span><br/><span data-ttu-id="6cedd-p148">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="6cedd-792">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-792">String</span></span> | <span data-ttu-id="6cedd-793">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-793">&lt;optional&gt;</span></span> | <span data-ttu-id="6cedd-p149">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="6cedd-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="6cedd-797">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-797">&lt;optional&gt;</span></span> | <span data-ttu-id="6cedd-798">Una matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="6cedd-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="6cedd-799">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-799">String</span></span> | | <span data-ttu-id="6cedd-p150">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="6cedd-802">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-802">String</span></span> | | <span data-ttu-id="6cedd-803">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="6cedd-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="6cedd-804">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-804">String</span></span> | | <span data-ttu-id="6cedd-p151">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="6cedd-807">Booleano</span><span class="sxs-lookup"><span data-stu-id="6cedd-807">Boolean</span></span> | | <span data-ttu-id="6cedd-p152">Solo se usa si `type` se establece en `file`. Si `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="6cedd-810">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-810">String</span></span> | | <span data-ttu-id="6cedd-p153">Solo se usa si `type` se establece en `item`. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="6cedd-814">función</span><span class="sxs-lookup"><span data-stu-id="6cedd-814">function</span></span> | <span data-ttu-id="6cedd-815">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-815">&lt;optional&gt;</span></span> | <span data-ttu-id="6cedd-816">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cedd-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6cedd-817">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-817">Requirements</span></span>

|<span data-ttu-id="6cedd-818">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-818">Requirement</span></span>| <span data-ttu-id="6cedd-819">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-820">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-820">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-821">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-821">1.0</span></span>|
|[<span data-ttu-id="6cedd-822">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-822">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-823">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-824">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-824">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-825">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="6cedd-826">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="6cedd-826">Examples</span></span>

<span data-ttu-id="6cedd-827">El código siguiente pasa una cadena a la función `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="6cedd-828">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="6cedd-828">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="6cedd-829">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="6cedd-829">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="6cedd-830">Responder con un cuerpo y datos adjuntos de archivo.</span><span class="sxs-lookup"><span data-stu-id="6cedd-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="6cedd-831">Responder con un cuerpo y datos adjuntos de elemento.</span><span class="sxs-lookup"><span data-stu-id="6cedd-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="6cedd-832">Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="6cedd-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="6cedd-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="6cedd-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="6cedd-834">Obtiene las entidades que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="6cedd-834">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="6cedd-835">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="6cedd-835">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-836">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-836">Requirements</span></span>

|<span data-ttu-id="6cedd-837">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-837">Requirement</span></span>| <span data-ttu-id="6cedd-838">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-839">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-839">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-840">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-840">1.0</span></span>|
|[<span data-ttu-id="6cedd-841">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-842">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-843">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-844">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6cedd-845">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="6cedd-845">Returns:</span></span>

<span data-ttu-id="6cedd-846">Tipo: [Entidades](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="6cedd-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="6cedd-847">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-847">Example</span></span>

<span data-ttu-id="6cedd-848">En el ejemplo siguiente se obtiene acceso a las entidades de los contactos en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="6cedd-848">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="6cedd-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="6cedd-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="6cedd-850">Obtiene una matriz de todas las entidades del tipo de entidad especificado que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="6cedd-850">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="6cedd-851">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="6cedd-851">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cedd-852">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="6cedd-852">Parameters:</span></span>

|<span data-ttu-id="6cedd-853">Nombre</span><span class="sxs-lookup"><span data-stu-id="6cedd-853">Name</span></span>| <span data-ttu-id="6cedd-854">Tipo</span><span class="sxs-lookup"><span data-stu-id="6cedd-854">Type</span></span>| <span data-ttu-id="6cedd-855">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="6cedd-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="6cedd-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="6cedd-857">Uno de los valores de enumeración de EntityType.</span><span class="sxs-lookup"><span data-stu-id="6cedd-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cedd-858">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-858">Requirements</span></span>

|<span data-ttu-id="6cedd-859">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-859">Requirement</span></span>| <span data-ttu-id="6cedd-860">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-861">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-861">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-862">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-862">1.0</span></span>|
|[<span data-ttu-id="6cedd-863">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-864">Restringido</span><span class="sxs-lookup"><span data-stu-id="6cedd-864">Restricted</span></span>|
|[<span data-ttu-id="6cedd-865">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-866">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6cedd-867">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="6cedd-867">Returns:</span></span>

<span data-ttu-id="6cedd-868">Si el valor que se pasa a `entityType` no es un miembro válido de la enumeración de `EntityType`, el método devuelve un valor NULL.</span><span class="sxs-lookup"><span data-stu-id="6cedd-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="6cedd-869">Si no existen entidades del tipo especificado están presentes en el cuerpo del elemento, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="6cedd-869">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="6cedd-870">De otro modo, el tipo de los objetos en la matriz devuelta depende del tipo de entidad solicitada en el parámetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="6cedd-871">Aunque el nivel de permiso mínimo para usar este método es **Restricted**, algunos tipos de entidad requieren **ReadItem** para tener acceso, como se especifica en la tabla siguiente.</span><span class="sxs-lookup"><span data-stu-id="6cedd-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="6cedd-872">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="6cedd-872">Value of `entityType`</span></span> | <span data-ttu-id="6cedd-873">Tipo de objetos de la matriz devuelta</span><span class="sxs-lookup"><span data-stu-id="6cedd-873">Type of objects in returned array</span></span> | <span data-ttu-id="6cedd-874">Nivel de permiso requerido</span><span class="sxs-lookup"><span data-stu-id="6cedd-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="6cedd-875">Cadena</span><span class="sxs-lookup"><span data-stu-id="6cedd-875">String</span></span> | <span data-ttu-id="6cedd-876">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="6cedd-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="6cedd-877">Contacto</span><span class="sxs-lookup"><span data-stu-id="6cedd-877">Contact</span></span> | <span data-ttu-id="6cedd-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6cedd-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="6cedd-879">Cadena</span><span class="sxs-lookup"><span data-stu-id="6cedd-879">String</span></span> | <span data-ttu-id="6cedd-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6cedd-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="6cedd-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="6cedd-881">MeetingSuggestion</span></span> | <span data-ttu-id="6cedd-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6cedd-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="6cedd-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="6cedd-883">PhoneNumber</span></span> | <span data-ttu-id="6cedd-884">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="6cedd-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="6cedd-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="6cedd-885">TaskSuggestion</span></span> | <span data-ttu-id="6cedd-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6cedd-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="6cedd-887">Cadena</span><span class="sxs-lookup"><span data-stu-id="6cedd-887">String</span></span> | <span data-ttu-id="6cedd-888">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="6cedd-888">**Restricted**</span></span> |

<span data-ttu-id="6cedd-889">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="6cedd-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="6cedd-890">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-890">Example</span></span>

<span data-ttu-id="6cedd-891">En el ejemplo siguiente se muestra cómo obtener acceso a una matriz de cadenas que representan direcciones postales en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="6cedd-891">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="6cedd-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="6cedd-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="6cedd-893">Devuelve entidades conocidas en el elemento seleccionado que pasan el filtro con nombre definido en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="6cedd-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6cedd-894">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="6cedd-894">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6cedd-895">El método `getFilteredEntitiesByName` devuelve las entidades que coinciden con la expresión regular definida en el elemento de regla [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) del archivo XML de manifiesto con el valor de elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="6cedd-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cedd-896">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="6cedd-896">Parameters:</span></span>

|<span data-ttu-id="6cedd-897">Nombre</span><span class="sxs-lookup"><span data-stu-id="6cedd-897">Name</span></span>| <span data-ttu-id="6cedd-898">Tipo</span><span class="sxs-lookup"><span data-stu-id="6cedd-898">Type</span></span>| <span data-ttu-id="6cedd-899">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="6cedd-900">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-900">String</span></span>|<span data-ttu-id="6cedd-901">El nombre del elemento de regla `ItemHasKnownEntity` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="6cedd-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cedd-902">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-902">Requirements</span></span>

|<span data-ttu-id="6cedd-903">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-903">Requirement</span></span>| <span data-ttu-id="6cedd-904">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-905">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-906">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-906">1.0</span></span>|
|[<span data-ttu-id="6cedd-907">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-908">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-909">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-910">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6cedd-911">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="6cedd-911">Returns:</span></span>

<span data-ttu-id="6cedd-p155">Si no existe ningún elemento `ItemHasKnownEntity` en el manifiesto con un valor de elemento `FilterName` que coincida con el parámetro `name`, el método devuelve `null`. Si el parámetro `name` no coincide con un elemento `ItemHasKnownEntity` del manifiesto, pero no existen entidades en el elemento actual que coincidan, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="6cedd-914">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="6cedd-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="6cedd-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="6cedd-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="6cedd-916">Devuelve valores de cadena en el elemento seleccionado que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="6cedd-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6cedd-917">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="6cedd-917">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6cedd-p156">El método `getRegExMatches` devuelve las cadenas que coinciden con la expresión regular definida en cada elemento de regla `ItemHasRegularExpressionMatch` o `ItemHasKnownEntity` del archivo XML de manifiesto. Para una regla `ItemHasRegularExpressionMatch`, tiene que darse una cadena coincidente en la propiedad del elemento especificado por la regla. El tipo simple `PropertyName` define las propiedades admitidas.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="6cedd-921">Por ejemplo, considere un manifiesto de complemento que tenga el siguiente elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="6cedd-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="6cedd-922">El objeto devuelto desde `getRegExMatches` tendría dos propiedades: `fruits` y `veggies`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="6cedd-p157">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados. En su lugar, use el método [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) para recuperar todo el cuerpo.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cedd-926">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-926">Requirements</span></span>

|<span data-ttu-id="6cedd-927">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-927">Requirement</span></span>| <span data-ttu-id="6cedd-928">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-929">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-929">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-930">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-930">1.0</span></span>|
|[<span data-ttu-id="6cedd-931">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-931">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-932">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-933">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-933">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-934">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6cedd-935">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="6cedd-935">Returns:</span></span>

<span data-ttu-id="6cedd-p158">Un objeto que contiene matrices de cadenas que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. El nombre de cada matriz es igual al valor correspondiente del atributo `RegExName` de la regla `ItemHasRegularExpressionMatch` coincidente o al atributo `FilterName` de la regla `ItemHasKnownEntity` coincidente.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="6cedd-938">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="6cedd-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6cedd-939">Object</span><span class="sxs-lookup"><span data-stu-id="6cedd-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="6cedd-940">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-940">Example</span></span>

<span data-ttu-id="6cedd-941">En el siguiente ejemplo, se muestra cómo tener acceso a la matriz de coincidencias de los <rule>elementos de expresión regular `fruits`y `veggies`, que se especifican en el manifiesto.</rule></span><span class="sxs-lookup"><span data-stu-id="6cedd-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="6cedd-942">→ de getregexmatchesbyname (Name) (nullable) {matriz. < cadena >}</span><span class="sxs-lookup"><span data-stu-id="6cedd-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="6cedd-943">Devuelve valores de cadena en el elemento seleccionado que coinciden con la expresión regular de nombre definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="6cedd-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6cedd-944">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="6cedd-944">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6cedd-945">El método `getRegExMatchesByName` devuelve las cadenas que coinciden con la expresión regular definida en el elemento de regla `ItemHasRegularExpressionMatch` del archivo XML de manifiesto con el valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="6cedd-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="6cedd-p159">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cedd-948">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="6cedd-948">Parameters:</span></span>

|<span data-ttu-id="6cedd-949">Nombre</span><span class="sxs-lookup"><span data-stu-id="6cedd-949">Name</span></span>| <span data-ttu-id="6cedd-950">Tipo</span><span class="sxs-lookup"><span data-stu-id="6cedd-950">Type</span></span>| <span data-ttu-id="6cedd-951">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="6cedd-952">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-952">String</span></span>|<span data-ttu-id="6cedd-953">El nombre del elemento de regla `ItemHasRegularExpressionMatch` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="6cedd-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cedd-954">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-954">Requirements</span></span>

|<span data-ttu-id="6cedd-955">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-955">Requirement</span></span>| <span data-ttu-id="6cedd-956">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-957">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-957">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-958">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-958">1.0</span></span>|
|[<span data-ttu-id="6cedd-959">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-960">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-961">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-962">Lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6cedd-963">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="6cedd-963">Returns:</span></span>

<span data-ttu-id="6cedd-964">Una matriz que contiene las cadenas que coinciden con la expresión regular definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="6cedd-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="6cedd-965">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="6cedd-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6cedd-966">< Cadena > de matriz.</span><span class="sxs-lookup"><span data-stu-id="6cedd-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="6cedd-967">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-967">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="6cedd-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="6cedd-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="6cedd-969">Devuelve asincrónicamente datos seleccionados desde el asunto o el cuerpo de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="6cedd-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="6cedd-p160">Si no hay ninguna selección, pero el cursor está en el cuerpo o el asunto, el método devuelve null para los datos seleccionados. Si se selecciona un campo que no sea el cuerpo o el asunto, el método devuelve el error `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cedd-972">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="6cedd-972">Parameters:</span></span>

|<span data-ttu-id="6cedd-973">Nombre</span><span class="sxs-lookup"><span data-stu-id="6cedd-973">Name</span></span>| <span data-ttu-id="6cedd-974">Tipo</span><span class="sxs-lookup"><span data-stu-id="6cedd-974">Type</span></span>| <span data-ttu-id="6cedd-975">Atributos</span><span class="sxs-lookup"><span data-stu-id="6cedd-975">Attributes</span></span>| <span data-ttu-id="6cedd-976">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="6cedd-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="6cedd-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="6cedd-p161">Solicita un formato para los datos. Si es Texto, el método devuelve el texto sin formato como cadena, quitando toda etiqueta HTML que hubiera. Si es HTML, el método devuelve el texto seleccionado, ya sea texto sin formato o HTML.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="6cedd-981">Object</span><span class="sxs-lookup"><span data-stu-id="6cedd-981">Object</span></span>| <span data-ttu-id="6cedd-982">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-982">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-983">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="6cedd-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6cedd-984">Objeto</span><span class="sxs-lookup"><span data-stu-id="6cedd-984">Object</span></span>| <span data-ttu-id="6cedd-985">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-985">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-986">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="6cedd-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6cedd-987">función</span><span class="sxs-lookup"><span data-stu-id="6cedd-987">function</span></span>||<span data-ttu-id="6cedd-988">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cedd-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6cedd-989">Para tener acceso a los datos seleccionados desde el método de devolución de llamada, llame a `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="6cedd-990">Para obtener acceso a la propiedad de origen del que proviene la selección, llame a `asyncResult.value.sourceProperty`, que será cualquiera `body` o `subject`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-990">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cedd-991">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-991">Requirements</span></span>

|<span data-ttu-id="6cedd-992">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-992">Requirement</span></span>| <span data-ttu-id="6cedd-993">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-994">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-994">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-995">1.2</span><span class="sxs-lookup"><span data-stu-id="6cedd-995">1.2</span></span>|
|[<span data-ttu-id="6cedd-996">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-996">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="6cedd-998">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-998">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-999">Redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="6cedd-1000">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="6cedd-1000">Returns:</span></span>

<span data-ttu-id="6cedd-1001">Los datos seleccionados como cadena con formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="6cedd-1002">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="6cedd-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6cedd-1003">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="6cedd-1004">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-1004">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="6cedd-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="6cedd-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="6cedd-1006">Carga de forma asincrónica las propiedades personalizadas de este complemento en el elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="6cedd-p163">Las propiedades personalizadas se almacenan como pares de clave/valor según la aplicación y el elemento. Este método devuelve un objeto `CustomProperties` en la devolución de llamada, que proporciona métodos para tener acceso a las propiedades personalizadas específicas del elemento y el complemento actual. Las propiedades personalizadas no están cifradas en el elemento, por lo que no debería usarse como almacenamiento seguro.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cedd-1010">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="6cedd-1010">Parameters:</span></span>

|<span data-ttu-id="6cedd-1011">Nombre</span><span class="sxs-lookup"><span data-stu-id="6cedd-1011">Name</span></span>| <span data-ttu-id="6cedd-1012">Tipo</span><span class="sxs-lookup"><span data-stu-id="6cedd-1012">Type</span></span>| <span data-ttu-id="6cedd-1013">Atributos</span><span class="sxs-lookup"><span data-stu-id="6cedd-1013">Attributes</span></span>| <span data-ttu-id="6cedd-1014">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="6cedd-1015">función</span><span class="sxs-lookup"><span data-stu-id="6cedd-1015">function</span></span>||<span data-ttu-id="6cedd-1016">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cedd-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6cedd-1017">Las propiedades personalizadas se proporcionan como un objeto [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="6cedd-1018">Este objeto se puede utilizar para obtener, establecer y elimina las propiedades personalizadas del elemento y guarde los cambios en la propiedad personalizada que se vuelve a establecer en el servidor.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1018">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="6cedd-1019">Objeto</span><span class="sxs-lookup"><span data-stu-id="6cedd-1019">Object</span></span>| <span data-ttu-id="6cedd-1020">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-1021">Los desarrolladores pueden proporcionar cualquier objeto que desean tener acceso en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1021">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="6cedd-1022">Se puede tener acceso a este objeto por el `asyncResult.asyncContext` (propiedad) en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cedd-1023">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-1023">Requirements</span></span>

|<span data-ttu-id="6cedd-1024">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-1024">Requirement</span></span>| <span data-ttu-id="6cedd-1025">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-1026">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-1026">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="6cedd-1027">1.0</span></span>|
|[<span data-ttu-id="6cedd-1028">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-1029">ReadItem</span></span>|
|[<span data-ttu-id="6cedd-1030">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-1031">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="6cedd-1031">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-1032">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-1032">Example</span></span>

<span data-ttu-id="6cedd-p166">En el siguiente ejemplo de código se muestra cómo usar el método `loadCustomPropertiesAsync` para cargar de forma asincrónica propiedades personalizadas que son específicas del elemento actual. En el ejemplo también se muestra cómo usar el método `CustomProperties.saveAsync` para guardar estas propiedades en el servidor. Después de cargar las propiedades personalizadas, el ejemplo de código usa el método `CustomProperties.get` para leer la propiedad personalizada `myProp`, el método `CustomProperties.set` para escribir en la propiedad personalizada `otherProp` y finalmente se llama al método `saveAsync` para guardar las propiedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="6cedd-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6cedd-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="6cedd-1037">Quita los datos adjuntos de un mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="6cedd-p167">El método `removeAttachmentAsync` quita del elemento los datos adjuntos con el identificador especificado. Como práctica recomendada, debe usar el identificador de datos adjuntos para quitar datos adjuntos solo si la misma aplicación de correo ha agregado los datos adjuntos en la misma sesión. En Outlook Web App y OWA para los dispositivos, el identificador de datos adjuntos es válido solo en la misma sesión. Una sesión finaliza cuando el usuario cierra la aplicación, o si el usuario comienza a redactar en un formulario en línea y posteriormente se extrae el formulario en línea para continuar en una ventana independiente.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cedd-1042">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="6cedd-1042">Parameters:</span></span>

|<span data-ttu-id="6cedd-1043">Nombre</span><span class="sxs-lookup"><span data-stu-id="6cedd-1043">Name</span></span>| <span data-ttu-id="6cedd-1044">Tipo</span><span class="sxs-lookup"><span data-stu-id="6cedd-1044">Type</span></span>| <span data-ttu-id="6cedd-1045">Atributos</span><span class="sxs-lookup"><span data-stu-id="6cedd-1045">Attributes</span></span>| <span data-ttu-id="6cedd-1046">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="6cedd-1047">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-1047">String</span></span>||<span data-ttu-id="6cedd-p168">El identificador de los datos adjuntos que se deben quitar. La longitud máxima de la cadena es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p168">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="6cedd-1050">Object</span><span class="sxs-lookup"><span data-stu-id="6cedd-1050">Object</span></span>| <span data-ttu-id="6cedd-1051">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-1052">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6cedd-1053">Objeto</span><span class="sxs-lookup"><span data-stu-id="6cedd-1053">Object</span></span>| <span data-ttu-id="6cedd-1054">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-1055">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6cedd-1056">función</span><span class="sxs-lookup"><span data-stu-id="6cedd-1056">function</span></span>| <span data-ttu-id="6cedd-1057">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-1058">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cedd-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6cedd-1059">Si se produce un error en la eliminación de los datos adjuntos, la propiedad `asyncResult.error` contendrá un código de error con el motivo del error.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6cedd-1060">Errores</span><span class="sxs-lookup"><span data-stu-id="6cedd-1060">Errors</span></span>

| <span data-ttu-id="6cedd-1061">Código de error</span><span class="sxs-lookup"><span data-stu-id="6cedd-1061">Error code</span></span> | <span data-ttu-id="6cedd-1062">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="6cedd-1063">El identificador de datos adjuntos no existe.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6cedd-1064">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-1064">Requirements</span></span>

|<span data-ttu-id="6cedd-1065">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-1065">Requirement</span></span>| <span data-ttu-id="6cedd-1066">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-1067">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-1067">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="6cedd-1068">1.1</span></span>|
|[<span data-ttu-id="6cedd-1069">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="6cedd-1071">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-1072">Redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-1073">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-1073">Example</span></span>

<span data-ttu-id="6cedd-1074">El código siguiente quita un archivo adjunto con un identificador de "0".</span><span class="sxs-lookup"><span data-stu-id="6cedd-1074">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="6cedd-1075">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="6cedd-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="6cedd-1076">Guarda un elemento de forma asincrónica.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="6cedd-p169">Cuando se invoca, este método guarda el mensaje actual como un borrador y devuelve el identificador de elemento a través del método de devolución de llamada. En el modo en línea de Outlook Web App o Outlook, el elemento se guarda en el servidor. En modo en caché de Outlook, se guarda el elemento en la caché local.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p169">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="6cedd-1080">Si el complemento llama a `saveAsync` en un elemento en modo de redacción con el fin de obtener una `itemId` para usar con EWS o la API de REST, tenga en cuenta que cuando Outlook está en modo en caché, puede tardar algún tiempo antes de que el elemento realmente se sincroniza con el servidor.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1080">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="6cedd-1081">Hasta que se sincroniza el elemento, el uso de la `itemId` devolverá un error.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="6cedd-p171">Dado que las citas no tienen ningún estado de borrador, si se llama a `saveAsync` en una cita en modo de redacción, el elemento se guardará como una cita normal en el calendario del usuario. Para las citas nuevas que todavía no se han guardado, no se enviará ninguna invitación. Si se guarda una cita existente, se enviará una actualización a los asistentes agregados o eliminados.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p171">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="6cedd-1085">Los siguientes clientes tienen un comportamiento diferente para `saveAsync` en citas en modo de redacción:</span><span class="sxs-lookup"><span data-stu-id="6cedd-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="6cedd-1086">Mac Outlook no admite `saveAsync` en una reunión en el modo de redacción.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1086">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="6cedd-1087">Al llamar a `saveAsync` en una reunión en Mac Outlook devolverá un error.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1087">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="6cedd-1088">Outlook en el web siempre envía una invitación o actualizar cuándo `saveAsync` se llama en una cita en modo de redacción.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cedd-1089">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="6cedd-1089">Parameters:</span></span>

|<span data-ttu-id="6cedd-1090">Nombre</span><span class="sxs-lookup"><span data-stu-id="6cedd-1090">Name</span></span>| <span data-ttu-id="6cedd-1091">Tipo</span><span class="sxs-lookup"><span data-stu-id="6cedd-1091">Type</span></span>| <span data-ttu-id="6cedd-1092">Atributos</span><span class="sxs-lookup"><span data-stu-id="6cedd-1092">Attributes</span></span>| <span data-ttu-id="6cedd-1093">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="6cedd-1094">Objeto</span><span class="sxs-lookup"><span data-stu-id="6cedd-1094">Object</span></span>| <span data-ttu-id="6cedd-1095">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-1096">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6cedd-1097">Objeto</span><span class="sxs-lookup"><span data-stu-id="6cedd-1097">Object</span></span>| <span data-ttu-id="6cedd-1098">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-1099">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6cedd-1100">función</span><span class="sxs-lookup"><span data-stu-id="6cedd-1100">function</span></span>||<span data-ttu-id="6cedd-1101">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cedd-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6cedd-1102">En caso de éxito, se proporciona el identificador de elemento en el `asyncResult.value` (propiedad).</span><span class="sxs-lookup"><span data-stu-id="6cedd-1102">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cedd-1103">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-1103">Requirements</span></span>

|<span data-ttu-id="6cedd-1104">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-1104">Requirement</span></span>| <span data-ttu-id="6cedd-1105">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-1106">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-1106">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="6cedd-1107">1.3</span></span>|
|[<span data-ttu-id="6cedd-1108">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="6cedd-1110">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-1111">Redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="6cedd-1112">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="6cedd-1112">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="6cedd-p173">A continuación, se muestra un ejemplo del parámetro `result` que se ha pasado a la función de devolución de llamada. La propiedad `value` contiene el identificador de elemento del elemento.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p173">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="6cedd-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="6cedd-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="6cedd-1116">Inserta asincrónicamente datos en el cuerpo o el asunto de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="6cedd-p174">El método `setSelectedDataAsync` inserta la cadena especificada en la posición del cursor en el asunto o el cuerpo del elemento, o si se selecciona texto en el editor, reemplaza el texto seleccionado. Si el cursor no está en el cuerpo o el campo del asunto, se devuelve un error. Después de la inserción, el cursor se coloca al final del contenido insertado.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p174">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cedd-1120">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="6cedd-1120">Parameters:</span></span>

|<span data-ttu-id="6cedd-1121">Nombre</span><span class="sxs-lookup"><span data-stu-id="6cedd-1121">Name</span></span>| <span data-ttu-id="6cedd-1122">Tipo</span><span class="sxs-lookup"><span data-stu-id="6cedd-1122">Type</span></span>| <span data-ttu-id="6cedd-1123">Atributos</span><span class="sxs-lookup"><span data-stu-id="6cedd-1123">Attributes</span></span>| <span data-ttu-id="6cedd-1124">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cedd-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="6cedd-1125">String</span><span class="sxs-lookup"><span data-stu-id="6cedd-1125">String</span></span>||<span data-ttu-id="6cedd-p175">Datos que se van a insertar. Los datos no deben superar 1.000.000 de caracteres. Si se pasan más de 1.000.000 de caracteres, se produce una excepción `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p175">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="6cedd-1129">Object</span><span class="sxs-lookup"><span data-stu-id="6cedd-1129">Object</span></span>| <span data-ttu-id="6cedd-1130">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-1131">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6cedd-1132">Objeto</span><span class="sxs-lookup"><span data-stu-id="6cedd-1132">Object</span></span>| <span data-ttu-id="6cedd-1133">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-1134">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="6cedd-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="6cedd-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="6cedd-1136">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cedd-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="6cedd-p176">Si `text`, se aplica el estilo actual en Outlook Web App y Outlook. Si el campo es un editor de HTML, se insertan solo los datos de texto, aunque los datos sean HTML.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p176">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="6cedd-p177">Si `html` y el campo admiten HTML (el asunto no), el estilo actual se aplica en Outlook Web App y el estilo predeterminado en Outlook. Si el campo es un campo de texto, se devuelve un error `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="6cedd-p177">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="6cedd-1141">Si `coercionType` no está establecido, el resultado depende del campo: si el campo es HTML, se usa HTML; si el campo es texto, se usa texto sin formato.</span><span class="sxs-lookup"><span data-stu-id="6cedd-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="6cedd-1142">function</span><span class="sxs-lookup"><span data-stu-id="6cedd-1142">function</span></span>||<span data-ttu-id="6cedd-1143">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cedd-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6cedd-1144">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6cedd-1144">Requirements</span></span>

|<span data-ttu-id="6cedd-1145">Requirement</span><span class="sxs-lookup"><span data-stu-id="6cedd-1145">Requirement</span></span>| <span data-ttu-id="6cedd-1146">Valor</span><span class="sxs-lookup"><span data-stu-id="6cedd-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cedd-1147">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="6cedd-1147">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cedd-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="6cedd-1148">1.2</span></span>|
|[<span data-ttu-id="6cedd-1149">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="6cedd-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cedd-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6cedd-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="6cedd-1151">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="6cedd-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cedd-1152">Redacción</span><span class="sxs-lookup"><span data-stu-id="6cedd-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6cedd-1153">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="6cedd-1153">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```