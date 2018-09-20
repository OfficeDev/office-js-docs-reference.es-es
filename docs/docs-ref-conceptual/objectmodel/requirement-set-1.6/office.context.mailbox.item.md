
# <a name="item"></a><span data-ttu-id="61a0b-101">elemento</span><span class="sxs-lookup"><span data-stu-id="61a0b-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="61a0b-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="61a0b-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="61a0b-p101">El espacio de nombres `item` se usa para acceder al mensaje seleccionado actualmente, a una convocatoria de reunión o a una cita. Puede determinar el tipo del `item` mediante la propiedad [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="61a0b-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-105">Requirements</span></span>

|<span data-ttu-id="61a0b-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-106">Requirement</span></span>| <span data-ttu-id="61a0b-107">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-108">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-109">1.0</span></span>|
|[<span data-ttu-id="61a0b-110">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-111">Restringido</span><span class="sxs-lookup"><span data-stu-id="61a0b-111">Restricted</span></span>|
|[<span data-ttu-id="61a0b-112">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-113">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="61a0b-114">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="61a0b-114">Members and methods</span></span>

| <span data-ttu-id="61a0b-115">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-115">Member</span></span> | <span data-ttu-id="61a0b-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="61a0b-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="61a0b-117">attachments</span><span class="sxs-lookup"><span data-stu-id="61a0b-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="61a0b-118">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-118">Member</span></span> |
| [<span data-ttu-id="61a0b-119">bcc</span><span class="sxs-lookup"><span data-stu-id="61a0b-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="61a0b-120">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-120">Member</span></span> |
| [<span data-ttu-id="61a0b-121">body</span><span class="sxs-lookup"><span data-stu-id="61a0b-121">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="61a0b-122">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-122">Member</span></span> |
| [<span data-ttu-id="61a0b-123">cc</span><span class="sxs-lookup"><span data-stu-id="61a0b-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="61a0b-124">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-124">Member</span></span> |
| [<span data-ttu-id="61a0b-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="61a0b-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="61a0b-126">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-126">Member</span></span> |
| [<span data-ttu-id="61a0b-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="61a0b-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="61a0b-128">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-128">Member</span></span> |
| [<span data-ttu-id="61a0b-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="61a0b-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="61a0b-130">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-130">Member</span></span> |
| [<span data-ttu-id="61a0b-131">end</span><span class="sxs-lookup"><span data-stu-id="61a0b-131">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="61a0b-132">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-132">Member</span></span> |
| [<span data-ttu-id="61a0b-133">from</span><span class="sxs-lookup"><span data-stu-id="61a0b-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="61a0b-134">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-134">Member</span></span> |
| [<span data-ttu-id="61a0b-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="61a0b-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="61a0b-136">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-136">Member</span></span> |
| [<span data-ttu-id="61a0b-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="61a0b-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="61a0b-138">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-138">Member</span></span> |
| [<span data-ttu-id="61a0b-139">itemId</span><span class="sxs-lookup"><span data-stu-id="61a0b-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="61a0b-140">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-140">Member</span></span> |
| [<span data-ttu-id="61a0b-141">itemType</span><span class="sxs-lookup"><span data-stu-id="61a0b-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="61a0b-142">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-142">Member</span></span> |
| [<span data-ttu-id="61a0b-143">location</span><span class="sxs-lookup"><span data-stu-id="61a0b-143">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="61a0b-144">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-144">Member</span></span> |
| [<span data-ttu-id="61a0b-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="61a0b-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="61a0b-146">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-146">Member</span></span> |
| [<span data-ttu-id="61a0b-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="61a0b-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="61a0b-148">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-148">Member</span></span> |
| [<span data-ttu-id="61a0b-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="61a0b-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="61a0b-150">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-150">Member</span></span> |
| [<span data-ttu-id="61a0b-151">organizer</span><span class="sxs-lookup"><span data-stu-id="61a0b-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="61a0b-152">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-152">Member</span></span> |
| [<span data-ttu-id="61a0b-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="61a0b-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="61a0b-154">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-154">Member</span></span> |
| [<span data-ttu-id="61a0b-155">sender</span><span class="sxs-lookup"><span data-stu-id="61a0b-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="61a0b-156">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-156">Member</span></span> |
| [<span data-ttu-id="61a0b-157">start</span><span class="sxs-lookup"><span data-stu-id="61a0b-157">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="61a0b-158">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-158">Member</span></span> |
| [<span data-ttu-id="61a0b-159">subject</span><span class="sxs-lookup"><span data-stu-id="61a0b-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="61a0b-160">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-160">Member</span></span> |
| [<span data-ttu-id="61a0b-161">to</span><span class="sxs-lookup"><span data-stu-id="61a0b-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="61a0b-162">Miembro	</span><span class="sxs-lookup"><span data-stu-id="61a0b-162">Member</span></span> |
| [<span data-ttu-id="61a0b-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="61a0b-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="61a0b-164">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-164">Method</span></span> |
| [<span data-ttu-id="61a0b-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="61a0b-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="61a0b-166">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-166">Method</span></span> |
| [<span data-ttu-id="61a0b-167">close</span><span class="sxs-lookup"><span data-stu-id="61a0b-167">close</span></span>](#close) | <span data-ttu-id="61a0b-168">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-168">Method</span></span> |
| [<span data-ttu-id="61a0b-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="61a0b-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="61a0b-170">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-170">Method</span></span> |
| [<span data-ttu-id="61a0b-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="61a0b-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="61a0b-172">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-172">Method</span></span> |
| [<span data-ttu-id="61a0b-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="61a0b-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="61a0b-174">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-174">Method</span></span> |
| [<span data-ttu-id="61a0b-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="61a0b-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="61a0b-176">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-176">Method</span></span> |
| [<span data-ttu-id="61a0b-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="61a0b-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="61a0b-178">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-178">Method</span></span> |
| [<span data-ttu-id="61a0b-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="61a0b-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="61a0b-180">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-180">Method</span></span> |
| [<span data-ttu-id="61a0b-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="61a0b-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="61a0b-182">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-182">Method</span></span> |
| [<span data-ttu-id="61a0b-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="61a0b-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="61a0b-184">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-184">Method</span></span> |
| [<span data-ttu-id="61a0b-185">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="61a0b-185">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="61a0b-186">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-186">Method</span></span> |
| [<span data-ttu-id="61a0b-187">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="61a0b-187">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="61a0b-188">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-188">Method</span></span> |
| [<span data-ttu-id="61a0b-189">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="61a0b-189">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="61a0b-190">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-190">Method</span></span> |
| [<span data-ttu-id="61a0b-191">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="61a0b-191">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="61a0b-192">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-192">Method</span></span> |
| [<span data-ttu-id="61a0b-193">saveAsync</span><span class="sxs-lookup"><span data-stu-id="61a0b-193">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="61a0b-194">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-194">Method</span></span> |
| [<span data-ttu-id="61a0b-195">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="61a0b-195">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="61a0b-196">Método</span><span class="sxs-lookup"><span data-stu-id="61a0b-196">Method</span></span> |

### <a name="example"></a><span data-ttu-id="61a0b-197">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-197">Example</span></span>

<span data-ttu-id="61a0b-198">En el siguiente ejemplo de código de JavaScript, se muestra cómo tener acceso a la propiedad `subject` del elemento actual en Outlook.</span><span class="sxs-lookup"><span data-stu-id="61a0b-198">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="61a0b-199">Miembros</span><span class="sxs-lookup"><span data-stu-id="61a0b-199">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="61a0b-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="61a0b-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="61a0b-p102">Obtiene una matriz de datos adjuntos para el elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-203">Ciertos tipos de archivos están bloqueados por Outlook debido a problemas de seguridad y, por tanto, no se devuelven.</span><span class="sxs-lookup"><span data-stu-id="61a0b-203">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="61a0b-204">Para obtener más información, vea [datos adjuntos bloqueados en Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="61a0b-204">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-205">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-205">Type:</span></span>

*   <span data-ttu-id="61a0b-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="61a0b-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-207">Requirements</span></span>

|<span data-ttu-id="61a0b-208">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-208">Requirement</span></span>| <span data-ttu-id="61a0b-209">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-210">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-211">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-211">1.0</span></span>|
|[<span data-ttu-id="61a0b-212">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-213">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-214">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-215">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-216">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-216">Example</span></span>

<span data-ttu-id="61a0b-217">El siguiente código crea una cadena HTML con los detalles de todos los datos adjuntos en el elemento actual.</span><span class="sxs-lookup"><span data-stu-id="61a0b-217">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="61a0b-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="61a0b-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="61a0b-219">Obtiene un objeto que proporciona métodos para obtener o actualizar a los destinatarios de la línea de CCO (con copia oculta) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="61a0b-219">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="61a0b-220">Solo modo Redacción.</span><span class="sxs-lookup"><span data-stu-id="61a0b-220">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-221">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-221">Type:</span></span>

*   [<span data-ttu-id="61a0b-222">Destinatarios</span><span class="sxs-lookup"><span data-stu-id="61a0b-222">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="61a0b-223">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-223">Requirements</span></span>

|<span data-ttu-id="61a0b-224">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-224">Requirement</span></span>| <span data-ttu-id="61a0b-225">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-226">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-226">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-227">1.1</span><span class="sxs-lookup"><span data-stu-id="61a0b-227">1.1</span></span>|
|[<span data-ttu-id="61a0b-228">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-229">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-230">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-231">Redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-231">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-232">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-232">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="61a0b-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="61a0b-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="61a0b-234">Obtiene un objeto que proporciona métodos para manipular el cuerpo de un elemento.</span><span class="sxs-lookup"><span data-stu-id="61a0b-234">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-235">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-235">Type:</span></span>

*   [<span data-ttu-id="61a0b-236">Cuerpo</span><span class="sxs-lookup"><span data-stu-id="61a0b-236">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="61a0b-237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-237">Requirements</span></span>

|<span data-ttu-id="61a0b-238">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-238">Requirement</span></span>| <span data-ttu-id="61a0b-239">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-240">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-241">1.1</span><span class="sxs-lookup"><span data-stu-id="61a0b-241">1.1</span></span>|
|[<span data-ttu-id="61a0b-242">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-243">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-244">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-245">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-245">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="61a0b-246">cc: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="61a0b-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="61a0b-247">Proporciona acceso a los destinatarios de Cc (con copia) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="61a0b-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="61a0b-248">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="61a0b-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="61a0b-249">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-249">Read mode</span></span>

<span data-ttu-id="61a0b-p106">La propiedad `cc` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **CC** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="61a0b-252">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-252">Compose mode</span></span>

<span data-ttu-id="61a0b-253">El `cc` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **Cc** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="61a0b-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-254">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-254">Type:</span></span>

*   <span data-ttu-id="61a0b-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="61a0b-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-256">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-256">Requirements</span></span>

|<span data-ttu-id="61a0b-257">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-257">Requirement</span></span>| <span data-ttu-id="61a0b-258">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-259">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-259">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-260">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-260">1.0</span></span>|
|[<span data-ttu-id="61a0b-261">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-261">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-262">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-263">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-263">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-264">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-264">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-265">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-265">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="61a0b-266">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="61a0b-266">(nullable) conversationId :String</span></span>

<span data-ttu-id="61a0b-267">Obtiene un identificador de la conversación de correo electrónico que contiene un mensaje determinado.</span><span class="sxs-lookup"><span data-stu-id="61a0b-267">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="61a0b-p107">Puede obtener un entero para esta propiedad si la aplicación de correo se activa en respuestas o formularios de lectura o en formularios de redacción. Si, después, el usuario cambia el asunto del mensaje de respuesta, al enviar dicha respuesta, el Id. de conversación de ese mensaje cambiará y el valor que se obtuvo anteriormente ya no será de aplicación.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="61a0b-p108">Obtiene un valor NULL para esta propiedad para un nuevo elemento de un formulario de redacción. Si el usuario establece un asunto y guarda el elemento, la propiedad `conversationId` devolverá un valor.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-272">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-272">Type:</span></span>

*   <span data-ttu-id="61a0b-273">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-273">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-274">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-274">Requirements</span></span>

|<span data-ttu-id="61a0b-275">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-275">Requirement</span></span>| <span data-ttu-id="61a0b-276">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-277">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-277">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-278">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-278">1.0</span></span>|
|[<span data-ttu-id="61a0b-279">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-280">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-281">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-282">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-282">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="61a0b-283">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="61a0b-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="61a0b-p109">Obtiene la fecha y hora de creación de un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-286">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-286">Type:</span></span>

*   <span data-ttu-id="61a0b-287">Fecha</span><span class="sxs-lookup"><span data-stu-id="61a0b-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-288">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-288">Requirements</span></span>

|<span data-ttu-id="61a0b-289">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-289">Requirement</span></span>| <span data-ttu-id="61a0b-290">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-291">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-291">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-292">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-292">1.0</span></span>|
|[<span data-ttu-id="61a0b-293">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-294">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-295">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-296">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-297">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-297">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="61a0b-298">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="61a0b-298">dateTimeModified :Date</span></span>

<span data-ttu-id="61a0b-p110">Obtiene la fecha y hora en que se modificó por última vez un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-301">Este miembro no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="61a0b-301">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-302">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-302">Type:</span></span>

*   <span data-ttu-id="61a0b-303">Fecha</span><span class="sxs-lookup"><span data-stu-id="61a0b-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-304">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-304">Requirements</span></span>

|<span data-ttu-id="61a0b-305">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-305">Requirement</span></span>| <span data-ttu-id="61a0b-306">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-307">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-307">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-308">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-308">1.0</span></span>|
|[<span data-ttu-id="61a0b-309">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-310">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-311">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-312">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-313">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-313">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="61a0b-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="61a0b-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="61a0b-315">Obtiene o establece la fecha y la hora de finalización de la cita.</span><span class="sxs-lookup"><span data-stu-id="61a0b-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="61a0b-p111">La propiedad `end` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) para convertir el valor de la propiedad final a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="61a0b-318">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-318">Read mode</span></span>

<span data-ttu-id="61a0b-319">La propiedad `end` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-319">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="61a0b-320">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-320">Compose mode</span></span>

<span data-ttu-id="61a0b-321">La propiedad `end` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="61a0b-322">Si usa el método [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) para establecer la hora de finalización, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="61a0b-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-323">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-323">Type:</span></span>

*   <span data-ttu-id="61a0b-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="61a0b-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-325">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-325">Requirements</span></span>

|<span data-ttu-id="61a0b-326">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-326">Requirement</span></span>| <span data-ttu-id="61a0b-327">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-328">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-328">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-329">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-329">1.0</span></span>|
|[<span data-ttu-id="61a0b-330">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-331">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-332">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-333">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-333">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-334">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-334">Example</span></span>

<span data-ttu-id="61a0b-335">En el ejemplo siguiente, se establece la hora de finalización de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-335">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="61a0b-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="61a0b-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="61a0b-p112">Obtiene la dirección de correo electrónico del remitente de un mensaje. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="61a0b-p113">Las propiedades `from` y [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-341">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `from` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-341">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-342">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-342">Type:</span></span>

*   [<span data-ttu-id="61a0b-343">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="61a0b-343">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="61a0b-344">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-344">Requirements</span></span>

|<span data-ttu-id="61a0b-345">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-345">Requirement</span></span>| <span data-ttu-id="61a0b-346">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-346">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-347">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-347">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-348">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-348">1.0</span></span>|
|[<span data-ttu-id="61a0b-349">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-349">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-350">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-350">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-351">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-351">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-352">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-352">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="61a0b-353">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="61a0b-353">internetMessageId :String</span></span>

<span data-ttu-id="61a0b-p114">Obtiene el identificador de mensaje de Internet de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-356">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-356">Type:</span></span>

*   <span data-ttu-id="61a0b-357">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-358">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-358">Requirements</span></span>

|<span data-ttu-id="61a0b-359">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-359">Requirement</span></span>| <span data-ttu-id="61a0b-360">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-361">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-362">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-362">1.0</span></span>|
|[<span data-ttu-id="61a0b-363">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-364">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-365">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-366">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-367">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-367">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="61a0b-368">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="61a0b-368">itemClass :String</span></span>

<span data-ttu-id="61a0b-p115">Obtiene la clase de elemento Servicios Web Exchange del elemento seleccionado. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="61a0b-p116">La propiedad `itemClass` especifica la clase de mensaje del elemento seleccionado. Las siguientes son las clases de mensaje predeterminadas para el elemento de mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="61a0b-373">Tipo</span><span class="sxs-lookup"><span data-stu-id="61a0b-373">Type</span></span> | <span data-ttu-id="61a0b-374">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-374">Description</span></span> | <span data-ttu-id="61a0b-375">Clase de elemento</span><span class="sxs-lookup"><span data-stu-id="61a0b-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="61a0b-376">Elementos de cita</span><span class="sxs-lookup"><span data-stu-id="61a0b-376">Appointment items</span></span> | <span data-ttu-id="61a0b-377">Estos son los elementos de calendario de la clase de elemento `IPM.Appointment` o `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="61a0b-378">Elementos de mensaje</span><span class="sxs-lookup"><span data-stu-id="61a0b-378">Message items</span></span> | <span data-ttu-id="61a0b-379">Entre estos se encuentran los mensajes de correo electrónico que tienen la clase de mensaje predeterminada `IPM.Note` y las convocatorias de reunión, respuestas y cancelaciones que usan `IPM.Schedule.Meeting` como clase de mensaje base.</span><span class="sxs-lookup"><span data-stu-id="61a0b-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="61a0b-380">Puede crear clases de mensaje personalizadas para ampliar una clase de mensaje predeterminada. Por ejemplo, puede crear una clase de mensaje de cita personalizada `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-381">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-381">Type:</span></span>

*   <span data-ttu-id="61a0b-382">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-383">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-383">Requirements</span></span>

|<span data-ttu-id="61a0b-384">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-384">Requirement</span></span>| <span data-ttu-id="61a0b-385">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-386">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-386">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-387">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-387">1.0</span></span>|
|[<span data-ttu-id="61a0b-388">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-389">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-390">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-391">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-392">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-392">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="61a0b-393">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="61a0b-393">(nullable) itemId :String</span></span>

<span data-ttu-id="61a0b-p117">Obtiene el identificador de elemento de Servicios Web Exchange para el elemento actual. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-396">El identificador que devuelve la propiedad `itemId` es el mismo que el identificador de elemento de Servicios Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="61a0b-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="61a0b-397">El `itemId` (propiedad) no es idéntico al identificador de entrada de Outlook o el identificador de la API de REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="61a0b-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="61a0b-398">Antes de realizar llamadas a la API de REST con este valor, se debe convertir mediante [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="61a0b-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="61a0b-399">Para obtener más información, vea [usar las API de REST de Outlook desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="61a0b-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="61a0b-p119">La propiedad `itemId` no está disponible en el modo redacción. Si se requiere un identificador de elemento, el método [`saveAsync`](#saveasyncoptions-callback) puede usarse para guardar el elemento en el servidor, que devolverá el identificador de elemento en el parámetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) de la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-402">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-402">Type:</span></span>

*   <span data-ttu-id="61a0b-403">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-404">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-404">Requirements</span></span>

|<span data-ttu-id="61a0b-405">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-405">Requirement</span></span>| <span data-ttu-id="61a0b-406">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-407">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-407">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-408">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-408">1.0</span></span>|
|[<span data-ttu-id="61a0b-409">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-410">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-411">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-412">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-413">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-413">Example</span></span>

<span data-ttu-id="61a0b-p120">El siguiente código comprueba la presencia de un identificador de elemento. Si la propiedad `itemId` devuelve `null` o `undefined`, guarda el elemento en el servidor y obtiene el identificador de elemento desde el resultado asincrónico.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="61a0b-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="61a0b-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="61a0b-417">Obtiene el tipo de elemento que representa una instancia.</span><span class="sxs-lookup"><span data-stu-id="61a0b-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="61a0b-418">La propiedad `itemType` devuelve uno de los valores de enumeración de `ItemType`, lo que indica si la instancia del objeto `item` es un mensaje o una cita.</span><span class="sxs-lookup"><span data-stu-id="61a0b-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-419">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-419">Type:</span></span>

*   [<span data-ttu-id="61a0b-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="61a0b-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="61a0b-421">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-421">Requirements</span></span>

|<span data-ttu-id="61a0b-422">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-422">Requirement</span></span>| <span data-ttu-id="61a0b-423">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-424">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-425">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-425">1.0</span></span>|
|[<span data-ttu-id="61a0b-426">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-427">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-428">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-429">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-429">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-430">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-430">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="61a0b-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="61a0b-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="61a0b-432">Obtiene o establece la ubicación de una cita.</span><span class="sxs-lookup"><span data-stu-id="61a0b-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="61a0b-433">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-433">Read mode</span></span>

<span data-ttu-id="61a0b-434">La propiedad `location` devuelve una cadena que contiene la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="61a0b-434">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="61a0b-435">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-435">Compose mode</span></span>

<span data-ttu-id="61a0b-436">La propiedad `location` devuelve un objeto `Location` que proporciona métodos para obtener y establecer la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="61a0b-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-437">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-437">Type:</span></span>

*   <span data-ttu-id="61a0b-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="61a0b-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-439">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-439">Requirements</span></span>

|<span data-ttu-id="61a0b-440">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-440">Requirement</span></span>| <span data-ttu-id="61a0b-441">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-442">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-442">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-443">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-443">1.0</span></span>|
|[<span data-ttu-id="61a0b-444">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-445">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-446">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-447">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-448">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-448">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="61a0b-449">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="61a0b-449">normalizedSubject :String</span></span>

<span data-ttu-id="61a0b-p121">Obtiene el asunto de un elemento, con todos los prefijos quitados (incluidos `RE:` y `FWD:`). Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="61a0b-p122">La propiedad normalizedSubject obtiene el asunto del elemento, con los prefijos estándar (como `RE:` y `FW:`) agregados por los programas de correo electrónico. Para obtener el asunto del elemento con los prefijos intactos, use la propiedad [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject).</span><span class="sxs-lookup"><span data-stu-id="61a0b-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-454">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-454">Type:</span></span>

*   <span data-ttu-id="61a0b-455">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-455">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-456">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-456">Requirements</span></span>

|<span data-ttu-id="61a0b-457">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-457">Requirement</span></span>| <span data-ttu-id="61a0b-458">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-459">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-459">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-460">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-460">1.0</span></span>|
|[<span data-ttu-id="61a0b-461">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-462">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-463">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-464">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-464">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-465">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-465">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="61a0b-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="61a0b-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="61a0b-467">Obtiene los mensajes de notificación de un elemento.</span><span class="sxs-lookup"><span data-stu-id="61a0b-467">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-468">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-468">Type:</span></span>

*   [<span data-ttu-id="61a0b-469">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="61a0b-469">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="61a0b-470">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-470">Requirements</span></span>

|<span data-ttu-id="61a0b-471">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-471">Requirement</span></span>| <span data-ttu-id="61a0b-472">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-473">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-473">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-474">1.3</span><span class="sxs-lookup"><span data-stu-id="61a0b-474">1.3</span></span>|
|[<span data-ttu-id="61a0b-475">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-476">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-477">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-478">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-478">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="61a0b-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="61a0b-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="61a0b-480">Proporciona acceso a los asistentes opcionales de un evento.</span><span class="sxs-lookup"><span data-stu-id="61a0b-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="61a0b-481">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="61a0b-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="61a0b-482">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-482">Read mode</span></span>

<span data-ttu-id="61a0b-483">La propiedad `optionalAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente opcional a la reunión.</span><span class="sxs-lookup"><span data-stu-id="61a0b-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="61a0b-484">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-484">Compose mode</span></span>

<span data-ttu-id="61a0b-485">El `optionalAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes opcionales de una reunión.</span><span class="sxs-lookup"><span data-stu-id="61a0b-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-486">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-486">Type:</span></span>

*   <span data-ttu-id="61a0b-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="61a0b-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-488">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-488">Requirements</span></span>

|<span data-ttu-id="61a0b-489">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-489">Requirement</span></span>| <span data-ttu-id="61a0b-490">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-491">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-491">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-492">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-492">1.0</span></span>|
|[<span data-ttu-id="61a0b-493">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-493">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-494">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-495">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-495">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-496">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-496">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-497">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-497">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="61a0b-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="61a0b-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="61a0b-p124">Obtiene la dirección de correo electrónico del organizador de una reunión especificada. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-501">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-501">Type:</span></span>

*   [<span data-ttu-id="61a0b-502">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="61a0b-502">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="61a0b-503">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-503">Requirements</span></span>

|<span data-ttu-id="61a0b-504">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-504">Requirement</span></span>| <span data-ttu-id="61a0b-505">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-506">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-507">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-507">1.0</span></span>|
|[<span data-ttu-id="61a0b-508">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-509">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-510">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-511">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-511">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-512">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-512">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="61a0b-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="61a0b-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="61a0b-514">Proporciona acceso a los asistentes necesarios de un evento.</span><span class="sxs-lookup"><span data-stu-id="61a0b-514">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="61a0b-515">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="61a0b-515">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="61a0b-516">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-516">Read mode</span></span>

<span data-ttu-id="61a0b-517">La propiedad `requiredAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente requerido a la reunión.</span><span class="sxs-lookup"><span data-stu-id="61a0b-517">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="61a0b-518">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-518">Compose mode</span></span>

<span data-ttu-id="61a0b-519">El `requiredAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes necesarios de una reunión.</span><span class="sxs-lookup"><span data-stu-id="61a0b-519">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-520">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-520">Type:</span></span>

*   <span data-ttu-id="61a0b-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="61a0b-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-522">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-522">Requirements</span></span>

|<span data-ttu-id="61a0b-523">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-523">Requirement</span></span>| <span data-ttu-id="61a0b-524">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-525">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-525">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-526">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-526">1.0</span></span>|
|[<span data-ttu-id="61a0b-527">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-527">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-528">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-529">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-530">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-530">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-531">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-531">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="61a0b-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="61a0b-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="61a0b-p126">Obtiene la dirección de correo electrónico del remitente de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="61a0b-p127">Las propiedades [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) y `sender` representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-537">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `sender` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-537">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-538">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-538">Type:</span></span>

*   [<span data-ttu-id="61a0b-539">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="61a0b-539">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="61a0b-540">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-540">Requirements</span></span>

|<span data-ttu-id="61a0b-541">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-541">Requirement</span></span>| <span data-ttu-id="61a0b-542">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-543">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-543">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-544">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-544">1.0</span></span>|
|[<span data-ttu-id="61a0b-545">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-545">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-546">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-547">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-547">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-548">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-548">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-549">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-549">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="61a0b-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="61a0b-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="61a0b-551">Obtiene o establece la fecha y la hora de inicio de la cita.</span><span class="sxs-lookup"><span data-stu-id="61a0b-551">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="61a0b-p128">La propiedad `start` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) para convertir el valor a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="61a0b-554">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-554">Read mode</span></span>

<span data-ttu-id="61a0b-555">La propiedad `start` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-555">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="61a0b-556">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-556">Compose mode</span></span>

<span data-ttu-id="61a0b-557">La propiedad `start` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-557">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="61a0b-558">Si usa el método [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) para establecer la hora de inicio, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="61a0b-558">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-559">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-559">Type:</span></span>

*   <span data-ttu-id="61a0b-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="61a0b-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-561">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-561">Requirements</span></span>

|<span data-ttu-id="61a0b-562">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-562">Requirement</span></span>| <span data-ttu-id="61a0b-563">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-564">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-564">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-565">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-565">1.0</span></span>|
|[<span data-ttu-id="61a0b-566">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-567">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-568">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-569">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-570">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-570">Example</span></span>

<span data-ttu-id="61a0b-571">En el ejemplo siguiente, se establece la hora de inicio de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-571">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="61a0b-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="61a0b-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="61a0b-573">Obtiene o establece la descripción que se muestra en el campo de asunto de un elemento.</span><span class="sxs-lookup"><span data-stu-id="61a0b-573">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="61a0b-574">La propiedad `subject` obtiene o establece el asunto completo del elemento, como lo ha enviado el servidor de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="61a0b-574">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="61a0b-575">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-575">Read mode</span></span>

<span data-ttu-id="61a0b-p129">La propiedad `subject` devuelve una cadena. Use la propiedad [`normalizedSubject`](#normalizedsubject-string) para obtener el asunto menos cualquier prefijo inicial como `RE:` y `FW:`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="61a0b-578">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-578">Compose mode</span></span>

<span data-ttu-id="61a0b-579">La propiedad `subject` devuelve un objeto `Subject` que proporciona métodos para obtener y establecer el asunto.</span><span class="sxs-lookup"><span data-stu-id="61a0b-579">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="61a0b-580">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-580">Type:</span></span>

*   <span data-ttu-id="61a0b-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="61a0b-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-582">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-582">Requirements</span></span>

|<span data-ttu-id="61a0b-583">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-583">Requirement</span></span>| <span data-ttu-id="61a0b-584">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-585">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-585">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-586">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-586">1.0</span></span>|
|[<span data-ttu-id="61a0b-587">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-588">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-589">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-590">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-590">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="61a0b-591">para: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="61a0b-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="61a0b-592">Proporciona acceso a los destinatarios en la línea **para** de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="61a0b-592">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="61a0b-593">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="61a0b-593">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="61a0b-594">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-594">Read mode</span></span>

<span data-ttu-id="61a0b-p131">La propiedad `to` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **Para** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="61a0b-597">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-597">Compose mode</span></span>

<span data-ttu-id="61a0b-598">El `to` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **para** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="61a0b-598">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="61a0b-599">Tipo:</span><span class="sxs-lookup"><span data-stu-id="61a0b-599">Type:</span></span>

*   <span data-ttu-id="61a0b-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="61a0b-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-601">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-601">Requirements</span></span>

|<span data-ttu-id="61a0b-602">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-602">Requirement</span></span>| <span data-ttu-id="61a0b-603">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-604">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-604">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-605">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-605">1.0</span></span>|
|[<span data-ttu-id="61a0b-606">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-607">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-607">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-608">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-609">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-610">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-610">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="61a0b-611">Métodos</span><span class="sxs-lookup"><span data-stu-id="61a0b-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="61a0b-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="61a0b-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="61a0b-613">Agrega un archivo a un mensaje o cita como datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="61a0b-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="61a0b-614">El método `addFileAttachmentAsync` carga el archivo en el URI especificado y lo asocia al elemento en el formulario de redacción.</span><span class="sxs-lookup"><span data-stu-id="61a0b-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="61a0b-615">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="61a0b-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="61a0b-616">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="61a0b-616">Parameters:</span></span>

|<span data-ttu-id="61a0b-617">Nombre</span><span class="sxs-lookup"><span data-stu-id="61a0b-617">Name</span></span>| <span data-ttu-id="61a0b-618">Tipo</span><span class="sxs-lookup"><span data-stu-id="61a0b-618">Type</span></span>| <span data-ttu-id="61a0b-619">Atributos</span><span class="sxs-lookup"><span data-stu-id="61a0b-619">Attributes</span></span>| <span data-ttu-id="61a0b-620">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="61a0b-621">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-621">String</span></span>||<span data-ttu-id="61a0b-p132">El URI que proporciona la ubicación del archivo que se va a adjuntar al mensaje o a la cita. La longitud máxima es de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="61a0b-624">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-624">String</span></span>||<span data-ttu-id="61a0b-p133">El nombre de los datos adjuntos que se muestra mientras estos se cargan. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="61a0b-627">Object</span><span class="sxs-lookup"><span data-stu-id="61a0b-627">Object</span></span>| <span data-ttu-id="61a0b-628">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-628">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-629">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="61a0b-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="61a0b-630">Objeto</span><span class="sxs-lookup"><span data-stu-id="61a0b-630">Object</span></span> | <span data-ttu-id="61a0b-631">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-631">&lt;optional&gt;</span></span> | <span data-ttu-id="61a0b-632">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="61a0b-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="61a0b-633">Booleano</span><span class="sxs-lookup"><span data-stu-id="61a0b-633">Boolean</span></span> | <span data-ttu-id="61a0b-634">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-634">&lt;optional&gt;</span></span> | <span data-ttu-id="61a0b-635">Si está establecido en `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="61a0b-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="61a0b-636">Función</span><span class="sxs-lookup"><span data-stu-id="61a0b-636">function</span></span>| <span data-ttu-id="61a0b-637">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-637">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-638">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="61a0b-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="61a0b-639">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="61a0b-640">Si se produce un error en la carga de los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="61a0b-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="61a0b-641">Errores</span><span class="sxs-lookup"><span data-stu-id="61a0b-641">Errors</span></span>

| <span data-ttu-id="61a0b-642">Código de error</span><span class="sxs-lookup"><span data-stu-id="61a0b-642">Error code</span></span> | <span data-ttu-id="61a0b-643">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="61a0b-644">Los datos adjuntos son más grandes de lo permitido.</span><span class="sxs-lookup"><span data-stu-id="61a0b-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="61a0b-645">Los datos adjuntos tienen una extensión que no está permitida.</span><span class="sxs-lookup"><span data-stu-id="61a0b-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="61a0b-646">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="61a0b-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="61a0b-647">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-647">Requirements</span></span>

|<span data-ttu-id="61a0b-648">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-648">Requirement</span></span>| <span data-ttu-id="61a0b-649">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-650">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-650">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-651">1.1</span><span class="sxs-lookup"><span data-stu-id="61a0b-651">1.1</span></span>|
|[<span data-ttu-id="61a0b-652">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="61a0b-654">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-655">Redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="61a0b-656">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="61a0b-656">Examples</span></span>

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

<span data-ttu-id="61a0b-657">En el siguiente ejemplo, se agrega un archivo de imagen como un dato adjunto en línea y hace referencia a los datos adjuntos del cuerpo del mensaje.</span><span class="sxs-lookup"><span data-stu-id="61a0b-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="61a0b-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="61a0b-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="61a0b-659">Agrega un elemento de Exchange (por ejemplo, un mensaje) como datos adjuntos al mensaje o a la cita.</span><span class="sxs-lookup"><span data-stu-id="61a0b-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="61a0b-p134">El método `addItemAttachmentAsync` asocia el elemento con el identificador especificado de Exchange al elemento en el formulario de redacción. Si se especifica un método de devolución de llamada, se llama al método con un parámetro, `asyncResult`, que contiene el identificador de datos adjuntos o un código que indique cualquier error que se ha producido al adjuntar el elemento. Si es necesario, puede usar el parámetro `options` para pasar información de estado al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="61a0b-663">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="61a0b-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="61a0b-664">Si el complemento de Office se está ejecutando en Outlook Web App, el `addItemAttachmentAsync` método puede adjuntar elementos a los elementos que no sea el elemento que se está editando; Sin embargo, esto no se admite y no se recomienda.</span><span class="sxs-lookup"><span data-stu-id="61a0b-664">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="61a0b-665">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="61a0b-665">Parameters:</span></span>

|<span data-ttu-id="61a0b-666">Nombre</span><span class="sxs-lookup"><span data-stu-id="61a0b-666">Name</span></span>| <span data-ttu-id="61a0b-667">Tipo</span><span class="sxs-lookup"><span data-stu-id="61a0b-667">Type</span></span>| <span data-ttu-id="61a0b-668">Atributos</span><span class="sxs-lookup"><span data-stu-id="61a0b-668">Attributes</span></span>| <span data-ttu-id="61a0b-669">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="61a0b-670">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-670">String</span></span>||<span data-ttu-id="61a0b-p135">El identificador de Exchange del elemento que debe adjuntarse. La longitud máxima es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="61a0b-673">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-673">String</span></span>||<span data-ttu-id="61a0b-p136">El asunto del elemento que debe adjuntarse. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="61a0b-676">Object</span><span class="sxs-lookup"><span data-stu-id="61a0b-676">Object</span></span>| <span data-ttu-id="61a0b-677">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-677">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-678">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="61a0b-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="61a0b-679">Objeto</span><span class="sxs-lookup"><span data-stu-id="61a0b-679">Object</span></span>| <span data-ttu-id="61a0b-680">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-680">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-681">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="61a0b-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="61a0b-682">función</span><span class="sxs-lookup"><span data-stu-id="61a0b-682">function</span></span>| <span data-ttu-id="61a0b-683">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-683">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-684">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="61a0b-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="61a0b-685">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="61a0b-686">Si se produce un error al agregar los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="61a0b-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="61a0b-687">Errores</span><span class="sxs-lookup"><span data-stu-id="61a0b-687">Errors</span></span>

| <span data-ttu-id="61a0b-688">Código de error</span><span class="sxs-lookup"><span data-stu-id="61a0b-688">Error code</span></span> | <span data-ttu-id="61a0b-689">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="61a0b-690">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="61a0b-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="61a0b-691">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-691">Requirements</span></span>

|<span data-ttu-id="61a0b-692">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-692">Requirement</span></span>| <span data-ttu-id="61a0b-693">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-694">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-694">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-695">1.1</span><span class="sxs-lookup"><span data-stu-id="61a0b-695">1.1</span></span>|
|[<span data-ttu-id="61a0b-696">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-696">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="61a0b-698">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-698">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-699">Redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-700">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-700">Example</span></span>

<span data-ttu-id="61a0b-701">En el ejemplo siguiente, se agrega un elemento existente de Outlook como un archivo adjunto con el nombre `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="61a0b-702">close()</span><span class="sxs-lookup"><span data-stu-id="61a0b-702">close()</span></span>

<span data-ttu-id="61a0b-703">Cierra el elemento actual que se está redactando.</span><span class="sxs-lookup"><span data-stu-id="61a0b-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="61a0b-p137">El comportamiento del método `close` depende del estado actual del elemento que se está redactando. Si el elemento tiene cambios sin guardar, el cliente solicita al usuario que guarde, descarte o cancele la acción Cerrar.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-706">En Outlook, en el sitio web, si el elemento es una cita y anteriormente se ha guardado con `saveAsync`, se pide al usuario al guardar, descartar o cancelar incluso si no ha habido cambios desde la último el elemento guardado.</span><span class="sxs-lookup"><span data-stu-id="61a0b-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="61a0b-707">En el cliente de escritorio de Outlook, si el mensaje es una respuesta directa, el método `close` no tiene ningún efecto.</span><span class="sxs-lookup"><span data-stu-id="61a0b-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-708">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-708">Requirements</span></span>

|<span data-ttu-id="61a0b-709">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-709">Requirement</span></span>| <span data-ttu-id="61a0b-710">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-711">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-711">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-712">1.3</span><span class="sxs-lookup"><span data-stu-id="61a0b-712">1.3</span></span>|
|[<span data-ttu-id="61a0b-713">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-713">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-714">Restringido</span><span class="sxs-lookup"><span data-stu-id="61a0b-714">Restricted</span></span>|
|[<span data-ttu-id="61a0b-715">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-715">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-716">Redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-716">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="61a0b-717">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="61a0b-717">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="61a0b-718">Muestra un formulario de respuesta que incluye el remitente y todos los destinatarios del mensaje seleccionado o el organizador y todos los asistentes de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="61a0b-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-719">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="61a0b-719">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="61a0b-720">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="61a0b-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="61a0b-721">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyAllForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="61a0b-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="61a0b-p138">Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="61a0b-725">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="61a0b-725">Parameters:</span></span>

| <span data-ttu-id="61a0b-726">Nombre</span><span class="sxs-lookup"><span data-stu-id="61a0b-726">Name</span></span> | <span data-ttu-id="61a0b-727">Tipo</span><span class="sxs-lookup"><span data-stu-id="61a0b-727">Type</span></span> | <span data-ttu-id="61a0b-728">Atributos</span><span class="sxs-lookup"><span data-stu-id="61a0b-728">Attributes</span></span> | <span data-ttu-id="61a0b-729">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="61a0b-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="61a0b-730">String &#124; Object</span></span>| |<span data-ttu-id="61a0b-p139">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="61a0b-733">**O**</span><span class="sxs-lookup"><span data-stu-id="61a0b-733">**OR**</span></span><br/><span data-ttu-id="61a0b-p140">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="61a0b-736">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-736">String</span></span> | <span data-ttu-id="61a0b-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-737">&lt;optional&gt;</span></span> | <span data-ttu-id="61a0b-p141">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="61a0b-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="61a0b-741">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-741">&lt;optional&gt;</span></span> | <span data-ttu-id="61a0b-742">Una matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="61a0b-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="61a0b-743">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-743">String</span></span> | | <span data-ttu-id="61a0b-p142">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="61a0b-746">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-746">String</span></span> | | <span data-ttu-id="61a0b-747">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="61a0b-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="61a0b-748">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-748">String</span></span> | | <span data-ttu-id="61a0b-p143">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="61a0b-751">Booleano</span><span class="sxs-lookup"><span data-stu-id="61a0b-751">Boolean</span></span> | | <span data-ttu-id="61a0b-p144">Solo se usa si `type` se establece en `file`. Si `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="61a0b-754">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-754">String</span></span> | | <span data-ttu-id="61a0b-p145">Solo se usa si `type` se establece en `item`. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="61a0b-758">función</span><span class="sxs-lookup"><span data-stu-id="61a0b-758">function</span></span> | <span data-ttu-id="61a0b-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-759">&lt;optional&gt;</span></span> | <span data-ttu-id="61a0b-760">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="61a0b-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="61a0b-761">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-761">Requirements</span></span>

|<span data-ttu-id="61a0b-762">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-762">Requirement</span></span>| <span data-ttu-id="61a0b-763">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-764">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-764">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-765">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-765">1.0</span></span>|
|[<span data-ttu-id="61a0b-766">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-766">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-767">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-768">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-768">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-769">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="61a0b-770">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="61a0b-770">Examples</span></span>

<span data-ttu-id="61a0b-771">El código siguiente pasa una cadena a la función `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="61a0b-772">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="61a0b-772">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="61a0b-773">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="61a0b-773">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="61a0b-774">Responder con un cuerpo y datos adjuntos de archivo.</span><span class="sxs-lookup"><span data-stu-id="61a0b-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="61a0b-775">Responder con un cuerpo y datos adjuntos de elemento.</span><span class="sxs-lookup"><span data-stu-id="61a0b-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="61a0b-776">Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="61a0b-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="61a0b-777">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="61a0b-777">displayReplyForm(formData)</span></span>

<span data-ttu-id="61a0b-778">Muestra un formulario de respuesta que incluye solo el remitente del mensaje seleccionado o el organizador de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="61a0b-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-779">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="61a0b-779">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="61a0b-780">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="61a0b-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="61a0b-781">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="61a0b-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="61a0b-p146">Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="61a0b-785">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="61a0b-785">Parameters:</span></span>

| <span data-ttu-id="61a0b-786">Nombre</span><span class="sxs-lookup"><span data-stu-id="61a0b-786">Name</span></span> | <span data-ttu-id="61a0b-787">Tipo</span><span class="sxs-lookup"><span data-stu-id="61a0b-787">Type</span></span> | <span data-ttu-id="61a0b-788">Atributos</span><span class="sxs-lookup"><span data-stu-id="61a0b-788">Attributes</span></span> | <span data-ttu-id="61a0b-789">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="61a0b-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="61a0b-790">String &#124; Object</span></span>| | <span data-ttu-id="61a0b-p147">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="61a0b-793">**O**</span><span class="sxs-lookup"><span data-stu-id="61a0b-793">**OR**</span></span><br/><span data-ttu-id="61a0b-p148">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="61a0b-796">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-796">String</span></span> | <span data-ttu-id="61a0b-797">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-797">&lt;optional&gt;</span></span> | <span data-ttu-id="61a0b-p149">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="61a0b-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="61a0b-801">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-801">&lt;optional&gt;</span></span> | <span data-ttu-id="61a0b-802">Una matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="61a0b-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="61a0b-803">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-803">String</span></span> | | <span data-ttu-id="61a0b-p150">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="61a0b-806">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-806">String</span></span> | | <span data-ttu-id="61a0b-807">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="61a0b-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="61a0b-808">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-808">String</span></span> | | <span data-ttu-id="61a0b-p151">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="61a0b-811">Booleano</span><span class="sxs-lookup"><span data-stu-id="61a0b-811">Boolean</span></span> | | <span data-ttu-id="61a0b-p152">Solo se usa si `type` se establece en `file`. Si `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="61a0b-814">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-814">String</span></span> | | <span data-ttu-id="61a0b-p153">Solo se usa si `type` se establece en `item`. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="61a0b-818">función</span><span class="sxs-lookup"><span data-stu-id="61a0b-818">function</span></span> | <span data-ttu-id="61a0b-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-819">&lt;optional&gt;</span></span> | <span data-ttu-id="61a0b-820">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="61a0b-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="61a0b-821">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-821">Requirements</span></span>

|<span data-ttu-id="61a0b-822">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-822">Requirement</span></span>| <span data-ttu-id="61a0b-823">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-824">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-824">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-825">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-825">1.0</span></span>|
|[<span data-ttu-id="61a0b-826">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-826">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-827">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-828">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-828">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-829">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="61a0b-830">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="61a0b-830">Examples</span></span>

<span data-ttu-id="61a0b-831">El código siguiente pasa una cadena a la función `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="61a0b-832">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="61a0b-832">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="61a0b-833">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="61a0b-833">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="61a0b-834">Responder con un cuerpo y datos adjuntos de archivo.</span><span class="sxs-lookup"><span data-stu-id="61a0b-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="61a0b-835">Responder con un cuerpo y datos adjuntos de elemento.</span><span class="sxs-lookup"><span data-stu-id="61a0b-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="61a0b-836">Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="61a0b-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="61a0b-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="61a0b-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="61a0b-838">Obtiene las entidades que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="61a0b-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-839">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="61a0b-839">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-840">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-840">Requirements</span></span>

|<span data-ttu-id="61a0b-841">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-841">Requirement</span></span>| <span data-ttu-id="61a0b-842">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-843">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-843">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-844">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-844">1.0</span></span>|
|[<span data-ttu-id="61a0b-845">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-845">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-846">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-847">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-847">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-848">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="61a0b-849">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="61a0b-849">Returns:</span></span>

<span data-ttu-id="61a0b-850">Tipo: [Entidades](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="61a0b-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="61a0b-851">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-851">Example</span></span>

<span data-ttu-id="61a0b-852">En el ejemplo siguiente se obtiene acceso a las entidades de los contactos en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="61a0b-852">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="61a0b-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="61a0b-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="61a0b-854">Obtiene una matriz de todas las entidades del tipo de entidad especificado que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="61a0b-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-855">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="61a0b-855">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="61a0b-856">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="61a0b-856">Parameters:</span></span>

|<span data-ttu-id="61a0b-857">Nombre</span><span class="sxs-lookup"><span data-stu-id="61a0b-857">Name</span></span>| <span data-ttu-id="61a0b-858">Tipo</span><span class="sxs-lookup"><span data-stu-id="61a0b-858">Type</span></span>| <span data-ttu-id="61a0b-859">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="61a0b-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="61a0b-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="61a0b-861">Uno de los valores de enumeración de EntityType.</span><span class="sxs-lookup"><span data-stu-id="61a0b-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="61a0b-862">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-862">Requirements</span></span>

|<span data-ttu-id="61a0b-863">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-863">Requirement</span></span>| <span data-ttu-id="61a0b-864">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-865">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-865">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-866">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-866">1.0</span></span>|
|[<span data-ttu-id="61a0b-867">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-867">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-868">Restringido</span><span class="sxs-lookup"><span data-stu-id="61a0b-868">Restricted</span></span>|
|[<span data-ttu-id="61a0b-869">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-869">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-870">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="61a0b-871">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="61a0b-871">Returns:</span></span>

<span data-ttu-id="61a0b-872">Si el valor que se pasa a `entityType` no es un miembro válido de la enumeración de `EntityType`, el método devuelve un valor NULL.</span><span class="sxs-lookup"><span data-stu-id="61a0b-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="61a0b-873">Si no existen entidades del tipo especificado están presentes en el cuerpo del elemento, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="61a0b-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="61a0b-874">De otro modo, el tipo de los objetos en la matriz devuelta depende del tipo de entidad solicitada en el parámetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="61a0b-875">Aunque el nivel de permiso mínimo para usar este método es **Restricted**, algunos tipos de entidad requieren **ReadItem** para tener acceso, como se especifica en la tabla siguiente.</span><span class="sxs-lookup"><span data-stu-id="61a0b-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="61a0b-876">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="61a0b-876">Value of `entityType`</span></span> | <span data-ttu-id="61a0b-877">Tipo de objetos de la matriz devuelta</span><span class="sxs-lookup"><span data-stu-id="61a0b-877">Type of objects in returned array</span></span> | <span data-ttu-id="61a0b-878">Nivel de permiso requerido</span><span class="sxs-lookup"><span data-stu-id="61a0b-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="61a0b-879">Cadena</span><span class="sxs-lookup"><span data-stu-id="61a0b-879">String</span></span> | <span data-ttu-id="61a0b-880">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="61a0b-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="61a0b-881">Contacto</span><span class="sxs-lookup"><span data-stu-id="61a0b-881">Contact</span></span> | <span data-ttu-id="61a0b-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="61a0b-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="61a0b-883">Cadena</span><span class="sxs-lookup"><span data-stu-id="61a0b-883">String</span></span> | <span data-ttu-id="61a0b-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="61a0b-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="61a0b-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="61a0b-885">MeetingSuggestion</span></span> | <span data-ttu-id="61a0b-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="61a0b-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="61a0b-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="61a0b-887">PhoneNumber</span></span> | <span data-ttu-id="61a0b-888">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="61a0b-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="61a0b-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="61a0b-889">TaskSuggestion</span></span> | <span data-ttu-id="61a0b-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="61a0b-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="61a0b-891">Cadena</span><span class="sxs-lookup"><span data-stu-id="61a0b-891">String</span></span> | <span data-ttu-id="61a0b-892">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="61a0b-892">**Restricted**</span></span> |

<span data-ttu-id="61a0b-893">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="61a0b-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="61a0b-894">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-894">Example</span></span>

<span data-ttu-id="61a0b-895">En el ejemplo siguiente se muestra cómo obtener acceso a una matriz de cadenas que representan direcciones postales en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="61a0b-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="61a0b-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="61a0b-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="61a0b-897">Devuelve entidades conocidas en el elemento seleccionado que pasan el filtro con nombre definido en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="61a0b-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-898">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="61a0b-898">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="61a0b-899">El método `getFilteredEntitiesByName` devuelve las entidades que coinciden con la expresión regular definida en el elemento de regla [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) del archivo XML de manifiesto con el valor de elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="61a0b-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="61a0b-900">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="61a0b-900">Parameters:</span></span>

|<span data-ttu-id="61a0b-901">Nombre</span><span class="sxs-lookup"><span data-stu-id="61a0b-901">Name</span></span>| <span data-ttu-id="61a0b-902">Tipo</span><span class="sxs-lookup"><span data-stu-id="61a0b-902">Type</span></span>| <span data-ttu-id="61a0b-903">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="61a0b-904">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-904">String</span></span>|<span data-ttu-id="61a0b-905">El nombre del elemento de regla `ItemHasKnownEntity` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="61a0b-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="61a0b-906">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-906">Requirements</span></span>

|<span data-ttu-id="61a0b-907">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-907">Requirement</span></span>| <span data-ttu-id="61a0b-908">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-909">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-909">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-910">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-910">1.0</span></span>|
|[<span data-ttu-id="61a0b-911">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-911">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-912">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-913">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-913">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-914">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="61a0b-915">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="61a0b-915">Returns:</span></span>

<span data-ttu-id="61a0b-p155">Si no existe ningún elemento `ItemHasKnownEntity` en el manifiesto con un valor de elemento `FilterName` que coincida con el parámetro `name`, el método devuelve `null`. Si el parámetro `name` no coincide con un elemento `ItemHasKnownEntity` del manifiesto, pero no existen entidades en el elemento actual que coincidan, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="61a0b-918">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="61a0b-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="61a0b-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="61a0b-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="61a0b-920">Devuelve valores de cadena en el elemento seleccionado que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="61a0b-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-921">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="61a0b-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="61a0b-p156">El método `getRegExMatches` devuelve las cadenas que coinciden con la expresión regular definida en cada elemento de regla `ItemHasRegularExpressionMatch` o `ItemHasKnownEntity` del archivo XML de manifiesto. Para una regla `ItemHasRegularExpressionMatch`, tiene que darse una cadena coincidente en la propiedad del elemento especificado por la regla. El tipo simple `PropertyName` define las propiedades admitidas.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="61a0b-925">Por ejemplo, considere un manifiesto de complemento que tenga el siguiente elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="61a0b-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="61a0b-926">El objeto devuelto desde `getRegExMatches` tendría dos propiedades: `fruits` y `veggies`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="61a0b-p157">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados. En su lugar, use el método [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) para recuperar todo el cuerpo.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-930">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-930">Requirements</span></span>

|<span data-ttu-id="61a0b-931">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-931">Requirement</span></span>| <span data-ttu-id="61a0b-932">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-933">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-933">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-934">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-934">1.0</span></span>|
|[<span data-ttu-id="61a0b-935">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-935">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-936">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-937">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-937">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-938">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="61a0b-939">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="61a0b-939">Returns:</span></span>

<span data-ttu-id="61a0b-p158">Un objeto que contiene matrices de cadenas que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. El nombre de cada matriz es igual al valor correspondiente del atributo `RegExName` de la regla `ItemHasRegularExpressionMatch` coincidente o al atributo `FilterName` de la regla `ItemHasKnownEntity` coincidente.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="61a0b-942">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="61a0b-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="61a0b-943">Object</span><span class="sxs-lookup"><span data-stu-id="61a0b-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="61a0b-944">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-944">Example</span></span>

<span data-ttu-id="61a0b-945">En el siguiente ejemplo, se muestra cómo tener acceso a la matriz de coincidencias de los elementos de regla de expresión regular `fruits` y `veggies`, que se especifican en el manifiesto.</span><span class="sxs-lookup"><span data-stu-id="61a0b-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="61a0b-946">→ de getregexmatchesbyname (Name) (nullable) {matriz. < cadena >}</span><span class="sxs-lookup"><span data-stu-id="61a0b-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="61a0b-947">Devuelve valores de cadena en el elemento seleccionado que coinciden con la expresión regular de nombre definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="61a0b-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-948">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="61a0b-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="61a0b-949">El método `getRegExMatchesByName` devuelve las cadenas que coinciden con la expresión regular definida en el elemento de regla `ItemHasRegularExpressionMatch` del archivo XML de manifiesto con el valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="61a0b-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="61a0b-p159">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="61a0b-952">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="61a0b-952">Parameters:</span></span>

|<span data-ttu-id="61a0b-953">Nombre</span><span class="sxs-lookup"><span data-stu-id="61a0b-953">Name</span></span>| <span data-ttu-id="61a0b-954">Tipo</span><span class="sxs-lookup"><span data-stu-id="61a0b-954">Type</span></span>| <span data-ttu-id="61a0b-955">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="61a0b-956">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-956">String</span></span>|<span data-ttu-id="61a0b-957">El nombre del elemento de regla `ItemHasRegularExpressionMatch` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="61a0b-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="61a0b-958">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-958">Requirements</span></span>

|<span data-ttu-id="61a0b-959">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-959">Requirement</span></span>| <span data-ttu-id="61a0b-960">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-961">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-961">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-962">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-962">1.0</span></span>|
|[<span data-ttu-id="61a0b-963">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-963">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-964">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-965">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-965">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-966">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="61a0b-967">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="61a0b-967">Returns:</span></span>

<span data-ttu-id="61a0b-968">Una matriz que contiene las cadenas que coinciden con la expresión regular definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="61a0b-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="61a0b-969">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="61a0b-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="61a0b-970">< Cadena > de matriz.</span><span class="sxs-lookup"><span data-stu-id="61a0b-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="61a0b-971">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-971">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="61a0b-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="61a0b-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="61a0b-973">Devuelve asincrónicamente datos seleccionados desde el asunto o el cuerpo de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="61a0b-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="61a0b-p160">Si no hay ninguna selección, pero el cursor está en el cuerpo o el asunto, el método devuelve null para los datos seleccionados. Si se selecciona un campo que no sea el cuerpo o el asunto, el método devuelve el error `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="61a0b-976">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="61a0b-976">Parameters:</span></span>

|<span data-ttu-id="61a0b-977">Nombre</span><span class="sxs-lookup"><span data-stu-id="61a0b-977">Name</span></span>| <span data-ttu-id="61a0b-978">Tipo</span><span class="sxs-lookup"><span data-stu-id="61a0b-978">Type</span></span>| <span data-ttu-id="61a0b-979">Atributos</span><span class="sxs-lookup"><span data-stu-id="61a0b-979">Attributes</span></span>| <span data-ttu-id="61a0b-980">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="61a0b-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="61a0b-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="61a0b-p161">Solicita un formato para los datos. Si es Texto, el método devuelve el texto sin formato como cadena, quitando toda etiqueta HTML que hubiera. Si es HTML, el método devuelve el texto seleccionado, ya sea texto sin formato o HTML.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="61a0b-985">Object</span><span class="sxs-lookup"><span data-stu-id="61a0b-985">Object</span></span>| <span data-ttu-id="61a0b-986">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-986">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-987">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="61a0b-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="61a0b-988">Objeto</span><span class="sxs-lookup"><span data-stu-id="61a0b-988">Object</span></span>| <span data-ttu-id="61a0b-989">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-989">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-990">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="61a0b-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="61a0b-991">función</span><span class="sxs-lookup"><span data-stu-id="61a0b-991">function</span></span>||<span data-ttu-id="61a0b-992">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="61a0b-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="61a0b-993">Para tener acceso a los datos seleccionados desde el método de devolución de llamada, llame a `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="61a0b-994">Para obtener acceso a la propiedad de origen del que proviene la selección, llame a `asyncResult.value.sourceProperty`, que será cualquiera `body` o `subject`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-994">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="61a0b-995">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-995">Requirements</span></span>

|<span data-ttu-id="61a0b-996">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-996">Requirement</span></span>| <span data-ttu-id="61a0b-997">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-998">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-998">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-999">1.2</span><span class="sxs-lookup"><span data-stu-id="61a0b-999">1.2</span></span>|
|[<span data-ttu-id="61a0b-1000">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1000">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="61a0b-1002">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-1002">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-1003">Redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="61a0b-1004">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="61a0b-1004">Returns:</span></span>

<span data-ttu-id="61a0b-1005">Los datos seleccionados como cadena con formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="61a0b-1006">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="61a0b-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="61a0b-1007">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="61a0b-1008">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="61a0b-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="61a0b-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="61a0b-p163">Obtiene las entidades que se han detectado en una coincidencia resaltada que un usuario ha seleccionado. Las coincidencias resaltadas se aplican a [complementos contextuales](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="61a0b-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-1012">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1012">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-1013">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-1013">Requirements</span></span>

|<span data-ttu-id="61a0b-1014">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-1014">Requirement</span></span>| <span data-ttu-id="61a0b-1015">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-1016">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-1016">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="61a0b-1017">1.6</span></span> |
|[<span data-ttu-id="61a0b-1018">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1018">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-1019">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-1020">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-1020">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-1021">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="61a0b-1022">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="61a0b-1022">Returns:</span></span>

<span data-ttu-id="61a0b-1023">Tipo: [Entidades](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="61a0b-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="61a0b-1024">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1024">Example</span></span>

<span data-ttu-id="61a0b-1025">En el siguiente ejemplo se accede a las entidades de direcciones en la coincidencia resaltada seleccionada por el usuario.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="61a0b-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="61a0b-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="61a0b-p164">Devuelve valores de cadena en una coincidencia resaltada que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. Las coincidencias resaltadas se aplican a [complementos contextuales](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="61a0b-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-1029">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1029">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="61a0b-p165">El método `getSelectedRegExMatches` devuelve las cadenas que coinciden con la expresión regular definida en cada elemento de regla `ItemHasRegularExpressionMatch` o `ItemHasKnownEntity` del archivo XML de manifiesto. Para una regla `ItemHasRegularExpressionMatch`, tiene que darse una cadena coincidente en la propiedad del elemento especificado por la regla. El tipo simple `PropertyName` define las propiedades admitidas.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="61a0b-1033">Por ejemplo, considere un manifiesto de complemento que tenga el siguiente elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="61a0b-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="61a0b-1034">El objeto devuelto desde `getRegExMatches` tendría dos propiedades: `fruits` y `veggies`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="61a0b-p166">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados. En su lugar, use el método [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) para recuperar todo el cuerpo.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="61a0b-1038">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-1038">Requirements</span></span>

|<span data-ttu-id="61a0b-1039">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-1039">Requirement</span></span>| <span data-ttu-id="61a0b-1040">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-1041">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-1041">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="61a0b-1042">1.6</span></span> |
|[<span data-ttu-id="61a0b-1043">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1043">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-1044">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-1045">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-1045">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-1046">Lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="61a0b-1047">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="61a0b-1047">Returns:</span></span>

<span data-ttu-id="61a0b-p167">Un objeto que contiene matrices de cadenas que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. El nombre de cada matriz es igual al valor correspondiente del atributo `RegExName` de la regla `ItemHasRegularExpressionMatch` coincidente o al atributo `FilterName` de la regla `ItemHasKnownEntity` coincidente.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="61a0b-1050">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1050">Example</span></span>

<span data-ttu-id="61a0b-1051">En el siguiente ejemplo, se muestra cómo tener acceso a la matriz de coincidencias de los elementos de regla de expresión regular `fruits` y `veggies`, que se especifican en el manifiesto.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="61a0b-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="61a0b-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="61a0b-1053">Carga de forma asincrónica las propiedades personalizadas de este complemento en el elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="61a0b-p168">Las propiedades personalizadas se almacenan como pares de clave/valor según la aplicación y el elemento. Este método devuelve un objeto `CustomProperties` en la devolución de llamada, que proporciona métodos para tener acceso a las propiedades personalizadas específicas del elemento y el complemento actual. Las propiedades personalizadas no están cifradas en el elemento, por lo que no debería usarse como almacenamiento seguro.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="61a0b-1057">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="61a0b-1057">Parameters:</span></span>

|<span data-ttu-id="61a0b-1058">Nombre</span><span class="sxs-lookup"><span data-stu-id="61a0b-1058">Name</span></span>| <span data-ttu-id="61a0b-1059">Tipo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1059">Type</span></span>| <span data-ttu-id="61a0b-1060">Atributos</span><span class="sxs-lookup"><span data-stu-id="61a0b-1060">Attributes</span></span>| <span data-ttu-id="61a0b-1061">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="61a0b-1062">función</span><span class="sxs-lookup"><span data-stu-id="61a0b-1062">function</span></span>||<span data-ttu-id="61a0b-1063">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="61a0b-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="61a0b-1064">Las propiedades personalizadas se proporcionan como un objeto [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="61a0b-1065">Este objeto se puede utilizar para obtener, establecer y elimina las propiedades personalizadas del elemento y guarde los cambios en la propiedad personalizada que se vuelve a establecer en el servidor.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1065">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="61a0b-1066">Objeto</span><span class="sxs-lookup"><span data-stu-id="61a0b-1066">Object</span></span>| <span data-ttu-id="61a0b-1067">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-1068">Los desarrolladores pueden proporcionar cualquier objeto que desean tener acceso en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1068">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="61a0b-1069">Se puede tener acceso a este objeto por el `asyncResult.asyncContext` (propiedad) en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="61a0b-1070">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-1070">Requirements</span></span>

|<span data-ttu-id="61a0b-1071">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-1071">Requirement</span></span>| <span data-ttu-id="61a0b-1072">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-1073">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-1073">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="61a0b-1074">1.0</span></span>|
|[<span data-ttu-id="61a0b-1075">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1075">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-1076">ReadItem</span></span>|
|[<span data-ttu-id="61a0b-1077">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-1077">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-1078">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="61a0b-1078">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-1079">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1079">Example</span></span>

<span data-ttu-id="61a0b-p171">En el siguiente ejemplo de código se muestra cómo usar el método `loadCustomPropertiesAsync` para cargar de forma asincrónica propiedades personalizadas que son específicas del elemento actual. En el ejemplo también se muestra cómo usar el método `CustomProperties.saveAsync` para guardar estas propiedades en el servidor. Después de cargar las propiedades personalizadas, el ejemplo de código usa el método `CustomProperties.get` para leer la propiedad personalizada `myProp`, el método `CustomProperties.set` para escribir en la propiedad personalizada `otherProp` y finalmente se llama al método `saveAsync` para guardar las propiedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="61a0b-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="61a0b-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="61a0b-1084">Quita los datos adjuntos de un mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="61a0b-p172">El método `removeAttachmentAsync` quita del elemento los datos adjuntos con el identificador especificado. Como práctica recomendada, debe usar el identificador de datos adjuntos para quitar datos adjuntos solo si la misma aplicación de correo ha agregado los datos adjuntos en la misma sesión. En Outlook Web App y OWA para los dispositivos, el identificador de datos adjuntos es válido solo en la misma sesión. Una sesión finaliza cuando el usuario cierra la aplicación, o si el usuario comienza a redactar en un formulario en línea y posteriormente se extrae el formulario en línea para continuar en una ventana independiente.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="61a0b-1089">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="61a0b-1089">Parameters:</span></span>

|<span data-ttu-id="61a0b-1090">Nombre</span><span class="sxs-lookup"><span data-stu-id="61a0b-1090">Name</span></span>| <span data-ttu-id="61a0b-1091">Tipo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1091">Type</span></span>| <span data-ttu-id="61a0b-1092">Atributos</span><span class="sxs-lookup"><span data-stu-id="61a0b-1092">Attributes</span></span>| <span data-ttu-id="61a0b-1093">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="61a0b-1094">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-1094">String</span></span>||<span data-ttu-id="61a0b-p173">El identificador de los datos adjuntos que se deben quitar. La longitud máxima de la cadena es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p173">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="61a0b-1097">Object</span><span class="sxs-lookup"><span data-stu-id="61a0b-1097">Object</span></span>| <span data-ttu-id="61a0b-1098">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-1099">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="61a0b-1100">Objeto</span><span class="sxs-lookup"><span data-stu-id="61a0b-1100">Object</span></span>| <span data-ttu-id="61a0b-1101">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-1102">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="61a0b-1103">función</span><span class="sxs-lookup"><span data-stu-id="61a0b-1103">function</span></span>| <span data-ttu-id="61a0b-1104">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-1105">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="61a0b-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="61a0b-1106">Si se produce un error en la eliminación de los datos adjuntos, la propiedad `asyncResult.error` contendrá un código de error con el motivo del error.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="61a0b-1107">Errores</span><span class="sxs-lookup"><span data-stu-id="61a0b-1107">Errors</span></span>

| <span data-ttu-id="61a0b-1108">Código de error</span><span class="sxs-lookup"><span data-stu-id="61a0b-1108">Error code</span></span> | <span data-ttu-id="61a0b-1109">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="61a0b-1110">El identificador de datos adjuntos no existe.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="61a0b-1111">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-1111">Requirements</span></span>

|<span data-ttu-id="61a0b-1112">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-1112">Requirement</span></span>| <span data-ttu-id="61a0b-1113">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-1114">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-1114">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="61a0b-1115">1.1</span></span>|
|[<span data-ttu-id="61a0b-1116">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="61a0b-1118">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-1119">Redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-1120">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1120">Example</span></span>

<span data-ttu-id="61a0b-1121">El código siguiente quita un archivo adjunto con un identificador de "0".</span><span class="sxs-lookup"><span data-stu-id="61a0b-1121">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="61a0b-1122">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="61a0b-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="61a0b-1123">Guarda un elemento de forma asincrónica.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="61a0b-p174">Cuando se invoca, este método guarda el mensaje actual como un borrador y devuelve el identificador de elemento a través del método de devolución de llamada. En el modo en línea de Outlook Web App o Outlook, el elemento se guarda en el servidor. En modo en caché de Outlook, se guarda el elemento en la caché local.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p174">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-1127">Si el complemento llama a `saveAsync` en un elemento en modo de redacción con el fin de obtener una `itemId` para usar con EWS o la API de REST, tenga en cuenta que cuando Outlook está en modo en caché, puede tardar algún tiempo antes de que el elemento realmente se sincroniza con el servidor.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1127">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="61a0b-1128">Hasta que se sincroniza el elemento, el uso de la `itemId` devolverá un error.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="61a0b-p176">Dado que las citas no tienen ningún estado de borrador, si se llama a `saveAsync` en una cita en modo de redacción, el elemento se guardará como una cita normal en el calendario del usuario. Para las citas nuevas que todavía no se han guardado, no se enviará ninguna invitación. Si se guarda una cita existente, se enviará una actualización a los asistentes agregados o eliminados.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p176">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="61a0b-1132">Los siguientes clientes tienen un comportamiento diferente para `saveAsync` en citas en modo de redacción:</span><span class="sxs-lookup"><span data-stu-id="61a0b-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="61a0b-1133">Mac Outlook no admite `saveAsync` en una reunión en el modo de redacción.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1133">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="61a0b-1134">Al llamar a `saveAsync` en una reunión en Mac Outlook devolverá un error.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1134">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="61a0b-1135">Outlook en el web siempre envía una invitación o actualizar cuándo `saveAsync` se llama en una cita en modo de redacción.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="61a0b-1136">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="61a0b-1136">Parameters:</span></span>

|<span data-ttu-id="61a0b-1137">Nombre</span><span class="sxs-lookup"><span data-stu-id="61a0b-1137">Name</span></span>| <span data-ttu-id="61a0b-1138">Tipo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1138">Type</span></span>| <span data-ttu-id="61a0b-1139">Atributos</span><span class="sxs-lookup"><span data-stu-id="61a0b-1139">Attributes</span></span>| <span data-ttu-id="61a0b-1140">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="61a0b-1141">Objeto</span><span class="sxs-lookup"><span data-stu-id="61a0b-1141">Object</span></span>| <span data-ttu-id="61a0b-1142">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-1143">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="61a0b-1144">Objeto</span><span class="sxs-lookup"><span data-stu-id="61a0b-1144">Object</span></span>| <span data-ttu-id="61a0b-1145">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-1146">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="61a0b-1147">función</span><span class="sxs-lookup"><span data-stu-id="61a0b-1147">function</span></span>||<span data-ttu-id="61a0b-1148">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="61a0b-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="61a0b-1149">En caso de éxito, se proporciona el identificador de elemento en el `asyncResult.value` (propiedad).</span><span class="sxs-lookup"><span data-stu-id="61a0b-1149">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="61a0b-1150">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-1150">Requirements</span></span>

|<span data-ttu-id="61a0b-1151">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-1151">Requirement</span></span>| <span data-ttu-id="61a0b-1152">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-1153">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-1153">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="61a0b-1154">1.3</span></span>|
|[<span data-ttu-id="61a0b-1155">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="61a0b-1157">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-1158">Redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="61a0b-1159">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="61a0b-1159">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="61a0b-p178">A continuación, se muestra un ejemplo del parámetro `result` que se ha pasado a la función de devolución de llamada. La propiedad `value` contiene el identificador de elemento del elemento.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p178">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="61a0b-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="61a0b-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="61a0b-1163">Inserta asincrónicamente datos en el cuerpo o el asunto de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="61a0b-p179">El método `setSelectedDataAsync` inserta la cadena especificada en la posición del cursor en el asunto o el cuerpo del elemento, o si se selecciona texto en el editor, reemplaza el texto seleccionado. Si el cursor no está en el cuerpo o el campo del asunto, se devuelve un error. Después de la inserción, el cursor se coloca al final del contenido insertado.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p179">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="61a0b-1167">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="61a0b-1167">Parameters:</span></span>

|<span data-ttu-id="61a0b-1168">Nombre</span><span class="sxs-lookup"><span data-stu-id="61a0b-1168">Name</span></span>| <span data-ttu-id="61a0b-1169">Tipo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1169">Type</span></span>| <span data-ttu-id="61a0b-1170">Atributos</span><span class="sxs-lookup"><span data-stu-id="61a0b-1170">Attributes</span></span>| <span data-ttu-id="61a0b-1171">Descripción</span><span class="sxs-lookup"><span data-stu-id="61a0b-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="61a0b-1172">String</span><span class="sxs-lookup"><span data-stu-id="61a0b-1172">String</span></span>||<span data-ttu-id="61a0b-p180">Datos que se van a insertar. Los datos no deben superar 1.000.000 de caracteres. Si se pasan más de 1.000.000 de caracteres, se produce una excepción `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p180">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="61a0b-1176">Object</span><span class="sxs-lookup"><span data-stu-id="61a0b-1176">Object</span></span>| <span data-ttu-id="61a0b-1177">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-1178">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="61a0b-1179">Objeto</span><span class="sxs-lookup"><span data-stu-id="61a0b-1179">Object</span></span>| <span data-ttu-id="61a0b-1180">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-1181">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="61a0b-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="61a0b-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="61a0b-1183">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="61a0b-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="61a0b-p181">Si `text`, se aplica el estilo actual en Outlook Web App y Outlook. Si el campo es un editor de HTML, se insertan solo los datos de texto, aunque los datos sean HTML.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p181">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="61a0b-p182">Si `html` y el campo admiten HTML (el asunto no), el estilo actual se aplica en Outlook Web App y el estilo predeterminado en Outlook. Si el campo es un campo de texto, se devuelve un error `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="61a0b-p182">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="61a0b-1188">Si `coercionType` no está establecido, el resultado depende del campo: si el campo es HTML, se usa HTML; si el campo es texto, se usa texto sin formato.</span><span class="sxs-lookup"><span data-stu-id="61a0b-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="61a0b-1189">function</span><span class="sxs-lookup"><span data-stu-id="61a0b-1189">function</span></span>||<span data-ttu-id="61a0b-1190">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="61a0b-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="61a0b-1191">Requisitos</span><span class="sxs-lookup"><span data-stu-id="61a0b-1191">Requirements</span></span>

|<span data-ttu-id="61a0b-1192">Requirement</span><span class="sxs-lookup"><span data-stu-id="61a0b-1192">Requirement</span></span>| <span data-ttu-id="61a0b-1193">Valor</span><span class="sxs-lookup"><span data-stu-id="61a0b-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="61a0b-1194">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="61a0b-1194">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61a0b-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="61a0b-1195">1.2</span></span>|
|[<span data-ttu-id="61a0b-1196">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="61a0b-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="61a0b-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="61a0b-1198">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="61a0b-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61a0b-1199">Redacción</span><span class="sxs-lookup"><span data-stu-id="61a0b-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="61a0b-1200">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="61a0b-1200">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```