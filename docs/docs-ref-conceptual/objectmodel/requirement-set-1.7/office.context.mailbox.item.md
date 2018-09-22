
# <a name="item"></a><span data-ttu-id="09953-101">elemento</span><span class="sxs-lookup"><span data-stu-id="09953-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="09953-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="09953-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="09953-p101">El espacio de nombres `item` se usa para acceder al mensaje seleccionado actualmente, a una convocatoria de reunión o a una cita. Puede determinar el tipo del `item` mediante la propiedad [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="09953-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-105">Requirements</span></span>

|<span data-ttu-id="09953-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-106">Requirement</span></span>|<span data-ttu-id="09953-107">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-108">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-109">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-109">1.0</span></span>|
|[<span data-ttu-id="09953-110">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-111">Restringido</span><span class="sxs-lookup"><span data-stu-id="09953-111">Restricted</span></span>|
|[<span data-ttu-id="09953-112">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-113">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="09953-114">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="09953-114">Members and methods</span></span>

| <span data-ttu-id="09953-115">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-115">Member</span></span> | <span data-ttu-id="09953-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="09953-117">attachments</span><span class="sxs-lookup"><span data-stu-id="09953-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="09953-118">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-118">Member</span></span> |
| [<span data-ttu-id="09953-119">bcc</span><span class="sxs-lookup"><span data-stu-id="09953-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="09953-120">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-120">Member</span></span> |
| [<span data-ttu-id="09953-121">body</span><span class="sxs-lookup"><span data-stu-id="09953-121">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="09953-122">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-122">Member</span></span> |
| [<span data-ttu-id="09953-123">cc</span><span class="sxs-lookup"><span data-stu-id="09953-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="09953-124">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-124">Member</span></span> |
| [<span data-ttu-id="09953-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="09953-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="09953-126">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-126">Member</span></span> |
| [<span data-ttu-id="09953-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="09953-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="09953-128">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-128">Member</span></span> |
| [<span data-ttu-id="09953-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="09953-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="09953-130">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-130">Member</span></span> |
| [<span data-ttu-id="09953-131">end</span><span class="sxs-lookup"><span data-stu-id="09953-131">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="09953-132">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-132">Member</span></span> |
| [<span data-ttu-id="09953-133">from</span><span class="sxs-lookup"><span data-stu-id="09953-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="09953-134">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-134">Member</span></span> |
| [<span data-ttu-id="09953-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="09953-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="09953-136">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-136">Member</span></span> |
| [<span data-ttu-id="09953-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="09953-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="09953-138">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-138">Member</span></span> |
| [<span data-ttu-id="09953-139">itemId</span><span class="sxs-lookup"><span data-stu-id="09953-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="09953-140">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-140">Member</span></span> |
| [<span data-ttu-id="09953-141">itemType</span><span class="sxs-lookup"><span data-stu-id="09953-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="09953-142">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-142">Member</span></span> |
| [<span data-ttu-id="09953-143">location</span><span class="sxs-lookup"><span data-stu-id="09953-143">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="09953-144">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-144">Member</span></span> |
| [<span data-ttu-id="09953-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="09953-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="09953-146">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-146">Member</span></span> |
| [<span data-ttu-id="09953-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="09953-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="09953-148">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-148">Member</span></span> |
| [<span data-ttu-id="09953-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="09953-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="09953-150">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-150">Member</span></span> |
| [<span data-ttu-id="09953-151">organizer</span><span class="sxs-lookup"><span data-stu-id="09953-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="09953-152">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-152">Member</span></span> |
| [<span data-ttu-id="09953-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="09953-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="09953-154">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-154">Member</span></span> |
| [<span data-ttu-id="09953-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="09953-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="09953-156">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-156">Member</span></span> |
| [<span data-ttu-id="09953-157">sender</span><span class="sxs-lookup"><span data-stu-id="09953-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="09953-158">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-158">Member</span></span> |
| [<span data-ttu-id="09953-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="09953-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="09953-160">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-160">Member</span></span> |
| [<span data-ttu-id="09953-161">start</span><span class="sxs-lookup"><span data-stu-id="09953-161">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="09953-162">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-162">Member</span></span> |
| [<span data-ttu-id="09953-163">subject</span><span class="sxs-lookup"><span data-stu-id="09953-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="09953-164">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-164">Member</span></span> |
| [<span data-ttu-id="09953-165">to</span><span class="sxs-lookup"><span data-stu-id="09953-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="09953-166">Miembro	</span><span class="sxs-lookup"><span data-stu-id="09953-166">Member</span></span> |
| [<span data-ttu-id="09953-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="09953-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="09953-168">Método</span><span class="sxs-lookup"><span data-stu-id="09953-168">Method</span></span> |
| [<span data-ttu-id="09953-169">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="09953-169">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="09953-170">Método</span><span class="sxs-lookup"><span data-stu-id="09953-170">Method</span></span> |
| [<span data-ttu-id="09953-171">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="09953-171">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="09953-172">Método</span><span class="sxs-lookup"><span data-stu-id="09953-172">Method</span></span> |
| [<span data-ttu-id="09953-173">close</span><span class="sxs-lookup"><span data-stu-id="09953-173">close</span></span>](#close) | <span data-ttu-id="09953-174">Método</span><span class="sxs-lookup"><span data-stu-id="09953-174">Method</span></span> |
| [<span data-ttu-id="09953-175">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="09953-175">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="09953-176">Método</span><span class="sxs-lookup"><span data-stu-id="09953-176">Method</span></span> |
| [<span data-ttu-id="09953-177">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="09953-177">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="09953-178">Método</span><span class="sxs-lookup"><span data-stu-id="09953-178">Method</span></span> |
| [<span data-ttu-id="09953-179">getEntities</span><span class="sxs-lookup"><span data-stu-id="09953-179">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="09953-180">Método</span><span class="sxs-lookup"><span data-stu-id="09953-180">Method</span></span> |
| [<span data-ttu-id="09953-181">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="09953-181">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="09953-182">Método</span><span class="sxs-lookup"><span data-stu-id="09953-182">Method</span></span> |
| [<span data-ttu-id="09953-183">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="09953-183">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="09953-184">Método</span><span class="sxs-lookup"><span data-stu-id="09953-184">Method</span></span> |
| [<span data-ttu-id="09953-185">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="09953-185">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="09953-186">Método</span><span class="sxs-lookup"><span data-stu-id="09953-186">Method</span></span> |
| [<span data-ttu-id="09953-187">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="09953-187">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="09953-188">Método</span><span class="sxs-lookup"><span data-stu-id="09953-188">Method</span></span> |
| [<span data-ttu-id="09953-189">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="09953-189">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="09953-190">Método</span><span class="sxs-lookup"><span data-stu-id="09953-190">Method</span></span> |
| [<span data-ttu-id="09953-191">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="09953-191">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="09953-192">Método</span><span class="sxs-lookup"><span data-stu-id="09953-192">Method</span></span> |
| [<span data-ttu-id="09953-193">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="09953-193">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="09953-194">Método</span><span class="sxs-lookup"><span data-stu-id="09953-194">Method</span></span> |
| [<span data-ttu-id="09953-195">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="09953-195">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="09953-196">Método</span><span class="sxs-lookup"><span data-stu-id="09953-196">Method</span></span> |
| [<span data-ttu-id="09953-197">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="09953-197">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="09953-198">Método</span><span class="sxs-lookup"><span data-stu-id="09953-198">Method</span></span> |
| [<span data-ttu-id="09953-199">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="09953-199">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="09953-200">Método</span><span class="sxs-lookup"><span data-stu-id="09953-200">Method</span></span> |
| [<span data-ttu-id="09953-201">saveAsync</span><span class="sxs-lookup"><span data-stu-id="09953-201">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="09953-202">Método</span><span class="sxs-lookup"><span data-stu-id="09953-202">Method</span></span> |
| [<span data-ttu-id="09953-203">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="09953-203">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="09953-204">Método</span><span class="sxs-lookup"><span data-stu-id="09953-204">Method</span></span> |

### <a name="example"></a><span data-ttu-id="09953-205">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-205">Example</span></span>

<span data-ttu-id="09953-206">En el siguiente ejemplo de código de JavaScript, se muestra cómo tener acceso a la propiedad `subject` del elemento actual en Outlook.</span><span class="sxs-lookup"><span data-stu-id="09953-206">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="09953-207">Miembros</span><span class="sxs-lookup"><span data-stu-id="09953-207">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="09953-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="09953-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="09953-p102">Obtiene una matriz de datos adjuntos para el elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="09953-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-211">Ciertos tipos de archivos están bloqueados por Outlook debido a problemas de seguridad y, por tanto, no se devuelven.</span><span class="sxs-lookup"><span data-stu-id="09953-211">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="09953-212">Para obtener más información, vea [datos adjuntos bloqueados en Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="09953-212">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="09953-213">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-213">Type:</span></span>

*   <span data-ttu-id="09953-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="09953-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-215">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-215">Requirements</span></span>

|<span data-ttu-id="09953-216">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-216">Requirement</span></span>|<span data-ttu-id="09953-217">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-218">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-218">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-219">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-219">1.0</span></span>|
|[<span data-ttu-id="09953-220">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-221">ReadItem</span></span>|
|[<span data-ttu-id="09953-222">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-223">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-223">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-224">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-224">Example</span></span>

<span data-ttu-id="09953-225">El siguiente código crea una cadena HTML con los detalles de todos los datos adjuntos en el elemento actual.</span><span class="sxs-lookup"><span data-stu-id="09953-225">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="09953-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09953-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="09953-227">Obtiene un objeto que proporciona métodos para obtener o actualizar a los destinatarios de la línea de CCO (con copia oculta) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="09953-227">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="09953-228">Solo modo Redacción.</span><span class="sxs-lookup"><span data-stu-id="09953-228">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-229">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-229">Type:</span></span>

*   [<span data-ttu-id="09953-230">Destinatarios</span><span class="sxs-lookup"><span data-stu-id="09953-230">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="09953-231">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-231">Requirements</span></span>

|<span data-ttu-id="09953-232">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-232">Requirement</span></span>|<span data-ttu-id="09953-233">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-234">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-234">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-235">1.1</span><span class="sxs-lookup"><span data-stu-id="09953-235">1.1</span></span>|
|[<span data-ttu-id="09953-236">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-236">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-237">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-237">ReadItem</span></span>|
|[<span data-ttu-id="09953-238">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-238">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-239">Redacción</span><span class="sxs-lookup"><span data-stu-id="09953-239">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-240">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-240">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="09953-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="09953-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="09953-242">Obtiene un objeto que proporciona métodos para manipular el cuerpo de un elemento.</span><span class="sxs-lookup"><span data-stu-id="09953-242">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-243">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-243">Type:</span></span>

*   [<span data-ttu-id="09953-244">Cuerpo</span><span class="sxs-lookup"><span data-stu-id="09953-244">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="09953-245">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-245">Requirements</span></span>

|<span data-ttu-id="09953-246">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-246">Requirement</span></span>|<span data-ttu-id="09953-247">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-248">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-248">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-249">1.1</span><span class="sxs-lookup"><span data-stu-id="09953-249">1.1</span></span>|
|[<span data-ttu-id="09953-250">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-250">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-251">ReadItem</span></span>|
|[<span data-ttu-id="09953-252">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-252">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-253">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-253">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="09953-254">cc: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09953-254">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="09953-255">Proporciona acceso a los destinatarios de Cc (con copia) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="09953-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="09953-256">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="09953-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09953-257">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="09953-257">Read mode</span></span>

<span data-ttu-id="09953-p106">La propiedad `cc` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **CC** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="09953-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="09953-260">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="09953-260">Compose mode</span></span>

<span data-ttu-id="09953-261">El `cc` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **Cc** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="09953-261">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-262">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-262">Type:</span></span>

*   <span data-ttu-id="09953-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09953-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-264">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-264">Requirements</span></span>

|<span data-ttu-id="09953-265">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-265">Requirement</span></span>|<span data-ttu-id="09953-266">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-267">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-267">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-268">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-268">1.0</span></span>|
|[<span data-ttu-id="09953-269">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-270">ReadItem</span></span>|
|[<span data-ttu-id="09953-271">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-272">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-272">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-273">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-273">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="09953-274">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="09953-274">(nullable) conversationId :String</span></span>

<span data-ttu-id="09953-275">Obtiene un identificador de la conversación de correo electrónico que contiene un mensaje determinado.</span><span class="sxs-lookup"><span data-stu-id="09953-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="09953-p107">Puede obtener un entero para esta propiedad si la aplicación de correo se activa en respuestas o formularios de lectura o en formularios de redacción. Si, después, el usuario cambia el asunto del mensaje de respuesta, al enviar dicha respuesta, el Id. de conversación de ese mensaje cambiará y el valor que se obtuvo anteriormente ya no será de aplicación.</span><span class="sxs-lookup"><span data-stu-id="09953-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="09953-p108">Obtiene un valor NULL para esta propiedad para un nuevo elemento de un formulario de redacción. Si el usuario establece un asunto y guarda el elemento, la propiedad `conversationId` devolverá un valor.</span><span class="sxs-lookup"><span data-stu-id="09953-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-280">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-280">Type:</span></span>

*   <span data-ttu-id="09953-281">String</span><span class="sxs-lookup"><span data-stu-id="09953-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-282">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-282">Requirements</span></span>

|<span data-ttu-id="09953-283">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-283">Requirement</span></span>|<span data-ttu-id="09953-284">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-285">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-285">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-286">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-286">1.0</span></span>|
|[<span data-ttu-id="09953-287">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-287">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-288">ReadItem</span></span>|
|[<span data-ttu-id="09953-289">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-289">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-290">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-290">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="09953-291">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="09953-291">dateTimeCreated :Date</span></span>

<span data-ttu-id="09953-p109">Obtiene la fecha y hora de creación de un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="09953-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-294">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-294">Type:</span></span>

*   <span data-ttu-id="09953-295">Fecha</span><span class="sxs-lookup"><span data-stu-id="09953-295">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-296">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-296">Requirements</span></span>

|<span data-ttu-id="09953-297">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-297">Requirement</span></span>|<span data-ttu-id="09953-298">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-298">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-299">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-299">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-300">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-300">1.0</span></span>|
|[<span data-ttu-id="09953-301">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-301">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-302">ReadItem</span></span>|
|[<span data-ttu-id="09953-303">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-303">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-304">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-304">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-305">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-305">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="09953-306">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="09953-306">dateTimeModified :Date</span></span>

<span data-ttu-id="09953-p110">Obtiene la fecha y hora en que se modificó por última vez un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="09953-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-309">Este miembro no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="09953-309">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-310">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-310">Type:</span></span>

*   <span data-ttu-id="09953-311">Fecha</span><span class="sxs-lookup"><span data-stu-id="09953-311">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-312">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-312">Requirements</span></span>

|<span data-ttu-id="09953-313">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-313">Requirement</span></span>|<span data-ttu-id="09953-314">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-314">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-315">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-315">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-316">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-316">1.0</span></span>|
|[<span data-ttu-id="09953-317">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-317">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-318">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-318">ReadItem</span></span>|
|[<span data-ttu-id="09953-319">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-319">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-320">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-320">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-321">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-321">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="09953-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="09953-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="09953-323">Obtiene o establece la fecha y la hora de finalización de la cita.</span><span class="sxs-lookup"><span data-stu-id="09953-323">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="09953-p111">La propiedad `end` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) para convertir el valor de la propiedad final a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="09953-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09953-326">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="09953-326">Read mode</span></span>

<span data-ttu-id="09953-327">La propiedad `end` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="09953-327">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="09953-328">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="09953-328">Compose mode</span></span>

<span data-ttu-id="09953-329">La propiedad `end` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="09953-329">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="09953-330">Si usa el método [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) para establecer la hora de finalización, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="09953-330">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-331">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-331">Type:</span></span>

*   <span data-ttu-id="09953-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="09953-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-333">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-333">Requirements</span></span>

|<span data-ttu-id="09953-334">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-334">Requirement</span></span>|<span data-ttu-id="09953-335">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-336">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-337">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-337">1.0</span></span>|
|[<span data-ttu-id="09953-338">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-339">ReadItem</span></span>|
|[<span data-ttu-id="09953-340">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-341">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-342">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-342">Example</span></span>

<span data-ttu-id="09953-343">En el ejemplo siguiente, se establece la hora de finalización de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="09953-343">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="09953-344">de:[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[desde](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="09953-344">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="09953-345">Obtiene la dirección de correo electrónico del remitente de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="09953-345">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="09953-p112">Las propiedades `from` y [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="09953-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-348">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `from` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="09953-348">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09953-349">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="09953-349">Read mode</span></span>

<span data-ttu-id="09953-350">El `from` (propiedad) devuelve un `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="09953-350">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="09953-351">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="09953-351">Compose mode</span></span>

<span data-ttu-id="09953-352">El `from` (propiedad) devuelve un `From` objeto que proporciona un método para obtener el valor de.</span><span class="sxs-lookup"><span data-stu-id="09953-352">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="09953-353">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-353">Type:</span></span>

*   <span data-ttu-id="09953-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [desde](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="09953-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-355">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-355">Requirements</span></span>

|<span data-ttu-id="09953-356">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-356">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="09953-357">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-357">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-358">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-358">1.0</span></span>|<span data-ttu-id="09953-359">1.7</span><span class="sxs-lookup"><span data-stu-id="09953-359">1.7</span></span>|
|[<span data-ttu-id="09953-360">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-361">ReadItem</span></span>|<span data-ttu-id="09953-362">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="09953-362">ReadWriteItem</span></span>|
|[<span data-ttu-id="09953-363">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-364">Read</span><span class="sxs-lookup"><span data-stu-id="09953-364">Read</span></span>|<span data-ttu-id="09953-365">Redacción</span><span class="sxs-lookup"><span data-stu-id="09953-365">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="09953-366">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="09953-366">internetMessageId :String</span></span>

<span data-ttu-id="09953-p113">Obtiene el identificador de mensaje de Internet de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="09953-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-369">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-369">Type:</span></span>

*   <span data-ttu-id="09953-370">String</span><span class="sxs-lookup"><span data-stu-id="09953-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-371">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-371">Requirements</span></span>

|<span data-ttu-id="09953-372">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-372">Requirement</span></span>|<span data-ttu-id="09953-373">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-374">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-374">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-375">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-375">1.0</span></span>|
|[<span data-ttu-id="09953-376">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-376">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-377">ReadItem</span></span>|
|[<span data-ttu-id="09953-378">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-378">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-379">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-380">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-380">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="09953-381">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="09953-381">itemClass :String</span></span>

<span data-ttu-id="09953-p114">Obtiene la clase de elemento Servicios Web Exchange del elemento seleccionado. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="09953-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="09953-p115">La propiedad `itemClass` especifica la clase de mensaje del elemento seleccionado. Las siguientes son las clases de mensaje predeterminadas para el elemento de mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="09953-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="09953-386">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-386">Type</span></span>|<span data-ttu-id="09953-387">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-387">Description</span></span>|<span data-ttu-id="09953-388">Clase de elemento</span><span class="sxs-lookup"><span data-stu-id="09953-388">item class</span></span>|
|---|---|---|
|<span data-ttu-id="09953-389">Elementos de cita</span><span class="sxs-lookup"><span data-stu-id="09953-389">Appointment items</span></span>|<span data-ttu-id="09953-390">Estos son los elementos de calendario de la clase de elemento `IPM.Appointment` o `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="09953-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="09953-391">Elementos de mensaje</span><span class="sxs-lookup"><span data-stu-id="09953-391">Message items</span></span>|<span data-ttu-id="09953-392">Entre estos se encuentran los mensajes de correo electrónico que tienen la clase de mensaje predeterminada `IPM.Note` y las convocatorias de reunión, respuestas y cancelaciones que usan `IPM.Schedule.Meeting` como clase de mensaje base.</span><span class="sxs-lookup"><span data-stu-id="09953-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="09953-393">Puede crear clases de mensaje personalizadas para ampliar una clase de mensaje predeterminada. Por ejemplo, puede crear una clase de mensaje de cita personalizada `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="09953-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-394">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-394">Type:</span></span>

*   <span data-ttu-id="09953-395">String</span><span class="sxs-lookup"><span data-stu-id="09953-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-396">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-396">Requirements</span></span>

|<span data-ttu-id="09953-397">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-397">Requirement</span></span>|<span data-ttu-id="09953-398">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-399">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-399">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-400">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-400">1.0</span></span>|
|[<span data-ttu-id="09953-401">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-401">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-402">ReadItem</span></span>|
|[<span data-ttu-id="09953-403">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-403">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-404">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-405">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-405">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="09953-406">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="09953-406">(nullable) itemId :String</span></span>

<span data-ttu-id="09953-p116">Obtiene el identificador de elemento de Servicios Web Exchange para el elemento actual. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="09953-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-409">El identificador que devuelve la propiedad `itemId` es el mismo que el identificador de elemento de Servicios Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="09953-409">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="09953-410">El `itemId` (propiedad) no es idéntico al identificador de entrada de Outlook o el identificador de la API de REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="09953-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="09953-411">Antes de realizar llamadas a la API de REST con este valor, se debe convertir mediante [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="09953-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="09953-412">Para obtener más información, vea [usar las API de REST de Outlook desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="09953-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="09953-p118">La propiedad `itemId` no está disponible en el modo redacción. Si se requiere un identificador de elemento, el método [`saveAsync`](#saveasyncoptions-callback) puede usarse para guardar el elemento en el servidor, que devolverá el identificador de elemento en el parámetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) de la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="09953-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-415">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-415">Type:</span></span>

*   <span data-ttu-id="09953-416">String</span><span class="sxs-lookup"><span data-stu-id="09953-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-417">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-417">Requirements</span></span>

|<span data-ttu-id="09953-418">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-418">Requirement</span></span>|<span data-ttu-id="09953-419">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-420">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-420">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-421">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-421">1.0</span></span>|
|[<span data-ttu-id="09953-422">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-423">ReadItem</span></span>|
|[<span data-ttu-id="09953-424">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-425">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-426">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-426">Example</span></span>

<span data-ttu-id="09953-p119">El siguiente código comprueba la presencia de un identificador de elemento. Si la propiedad `itemId` devuelve `null` o `undefined`, guarda el elemento en el servidor y obtiene el identificador de elemento desde el resultado asincrónico.</span><span class="sxs-lookup"><span data-stu-id="09953-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="09953-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="09953-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="09953-430">Obtiene el tipo de elemento que representa una instancia.</span><span class="sxs-lookup"><span data-stu-id="09953-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="09953-431">La propiedad `itemType` devuelve uno de los valores de enumeración de `ItemType`, lo que indica si la instancia del objeto `item` es un mensaje o una cita.</span><span class="sxs-lookup"><span data-stu-id="09953-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-432">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-432">Type:</span></span>

*   [<span data-ttu-id="09953-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="09953-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="09953-434">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-434">Requirements</span></span>

|<span data-ttu-id="09953-435">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-435">Requirement</span></span>|<span data-ttu-id="09953-436">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-437">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-437">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-438">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-438">1.0</span></span>|
|[<span data-ttu-id="09953-439">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-439">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-440">ReadItem</span></span>|
|[<span data-ttu-id="09953-441">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-441">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-442">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-442">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-443">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-443">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="09953-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="09953-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="09953-445">Obtiene o establece la ubicación de una cita.</span><span class="sxs-lookup"><span data-stu-id="09953-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09953-446">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="09953-446">Read mode</span></span>

<span data-ttu-id="09953-447">La propiedad `location` devuelve una cadena que contiene la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="09953-447">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="09953-448">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="09953-448">Compose mode</span></span>

<span data-ttu-id="09953-449">La propiedad `location` devuelve un objeto `Location` que proporciona métodos para obtener y establecer la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="09953-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-450">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-450">Type:</span></span>

*   <span data-ttu-id="09953-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="09953-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-452">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-452">Requirements</span></span>

|<span data-ttu-id="09953-453">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-453">Requirement</span></span>|<span data-ttu-id="09953-454">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-455">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-455">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-456">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-456">1.0</span></span>|
|[<span data-ttu-id="09953-457">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-458">ReadItem</span></span>|
|[<span data-ttu-id="09953-459">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-460">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-460">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-461">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-461">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="09953-462">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="09953-462">normalizedSubject :String</span></span>

<span data-ttu-id="09953-p120">Obtiene el asunto de un elemento, con todos los prefijos quitados (incluidos `RE:` y `FWD:`). Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="09953-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="09953-p121">La propiedad normalizedSubject obtiene el asunto del elemento, con los prefijos estándar (como `RE:` y `FW:`) agregados por los programas de correo electrónico. Para obtener el asunto del elemento con los prefijos intactos, use la propiedad [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject).</span><span class="sxs-lookup"><span data-stu-id="09953-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-467">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-467">Type:</span></span>

*   <span data-ttu-id="09953-468">String</span><span class="sxs-lookup"><span data-stu-id="09953-468">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-469">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-469">Requirements</span></span>

|<span data-ttu-id="09953-470">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-470">Requirement</span></span>|<span data-ttu-id="09953-471">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-472">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-472">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-473">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-473">1.0</span></span>|
|[<span data-ttu-id="09953-474">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-474">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-475">ReadItem</span></span>|
|[<span data-ttu-id="09953-476">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-476">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-477">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-477">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-478">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-478">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="09953-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="09953-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="09953-480">Obtiene los mensajes de notificación de un elemento.</span><span class="sxs-lookup"><span data-stu-id="09953-480">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-481">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-481">Type:</span></span>

*   [<span data-ttu-id="09953-482">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="09953-482">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="09953-483">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-483">Requirements</span></span>

|<span data-ttu-id="09953-484">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-484">Requirement</span></span>|<span data-ttu-id="09953-485">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-485">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-486">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-486">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-487">1.3</span><span class="sxs-lookup"><span data-stu-id="09953-487">1.3</span></span>|
|[<span data-ttu-id="09953-488">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-488">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-489">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-489">ReadItem</span></span>|
|[<span data-ttu-id="09953-490">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-490">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-491">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-491">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="09953-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09953-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="09953-493">Proporciona acceso a los asistentes opcionales de un evento.</span><span class="sxs-lookup"><span data-stu-id="09953-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="09953-494">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="09953-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09953-495">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="09953-495">Read mode</span></span>

<span data-ttu-id="09953-496">La propiedad `optionalAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente opcional a la reunión.</span><span class="sxs-lookup"><span data-stu-id="09953-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="09953-497">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="09953-497">Compose mode</span></span>

<span data-ttu-id="09953-498">El `optionalAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes opcionales de una reunión.</span><span class="sxs-lookup"><span data-stu-id="09953-498">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-499">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-499">Type:</span></span>

*   <span data-ttu-id="09953-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09953-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-501">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-501">Requirements</span></span>

|<span data-ttu-id="09953-502">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-502">Requirement</span></span>|<span data-ttu-id="09953-503">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-503">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-504">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-504">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-505">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-505">1.0</span></span>|
|[<span data-ttu-id="09953-506">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-506">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-507">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-507">ReadItem</span></span>|
|[<span data-ttu-id="09953-508">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-508">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-509">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-509">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-510">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-510">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="09953-511">Organizador:[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizador](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="09953-511">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="09953-512">Obtiene la dirección de correo electrónico del organizador de una reunión especificada.</span><span class="sxs-lookup"><span data-stu-id="09953-512">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09953-513">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="09953-513">Read mode</span></span>

<span data-ttu-id="09953-514">El `organizer` (propiedad) devuelve un objeto [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) que representa el organizador de la reunión.</span><span class="sxs-lookup"><span data-stu-id="09953-514">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="09953-515">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="09953-515">Compose mode</span></span>

<span data-ttu-id="09953-516">El `organizer` (propiedad) devuelve un objeto de [Organizador](/javascript/api/outlook_1_7/office.organizer) que proporciona un método para obtener el valor del organizador.</span><span class="sxs-lookup"><span data-stu-id="09953-516">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-517">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-517">Type:</span></span>

*   <span data-ttu-id="09953-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizador](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="09953-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-519">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-519">Requirements</span></span>

|<span data-ttu-id="09953-520">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-520">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="09953-521">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-522">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-522">1.0</span></span>|<span data-ttu-id="09953-523">1.7</span><span class="sxs-lookup"><span data-stu-id="09953-523">1.7</span></span>|
|[<span data-ttu-id="09953-524">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-525">ReadItem</span></span>|<span data-ttu-id="09953-526">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="09953-526">ReadWriteItem</span></span>|
|[<span data-ttu-id="09953-527">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-527">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-528">Read</span><span class="sxs-lookup"><span data-stu-id="09953-528">Read</span></span>|<span data-ttu-id="09953-529">Redacción</span><span class="sxs-lookup"><span data-stu-id="09953-529">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-530">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-530">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="09953-531">periodicidad (nullable):[Periodicidad](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="09953-531">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="09953-532">Obtiene o establece el patrón de periodicidad de una cita.</span><span class="sxs-lookup"><span data-stu-id="09953-532">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="09953-533">Obtiene el patrón de periodicidad de una convocatoria de reunión.</span><span class="sxs-lookup"><span data-stu-id="09953-533">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="09953-534">Lectura y redacción modos para elementos de cita.</span><span class="sxs-lookup"><span data-stu-id="09953-534">Read and compose modes for appointment items.</span></span> <span data-ttu-id="09953-535">Modo de lectura para los elementos de solicitud de reunión.</span><span class="sxs-lookup"><span data-stu-id="09953-535">Read mode for meeting request items.</span></span>

<span data-ttu-id="09953-536">El `recurrence` (propiedad) devuelve un objeto de [Periodicidad](/javascript/api/outlook_1_7/office.recurrence) para solicitudes de reuniones o citas periódicas si un elemento es una serie o una instancia de una serie.</span><span class="sxs-lookup"><span data-stu-id="09953-536">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="09953-537">`null`se devuelve solo citas y convocatorias de reunión de citas únicas.</span><span class="sxs-lookup"><span data-stu-id="09953-537">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="09953-538">`undefined`se devuelve para los mensajes que no son las convocatorias de reunión.</span><span class="sxs-lookup"><span data-stu-id="09953-538">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="09953-539">Nota: Las convocatorias de reunión tienen un `itemClass` valor de IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="09953-539">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="09953-540">Nota: Si el objeto de periodicidad es `null`, esto indica que el objeto es una cita única o una convocatoria de reunión de una cita única y no forma parte de una serie.</span><span class="sxs-lookup"><span data-stu-id="09953-540">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-541">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-541">Type:</span></span>

* [<span data-ttu-id="09953-542">Periodicidad</span><span class="sxs-lookup"><span data-stu-id="09953-542">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="09953-543">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-543">Requirement</span></span>|<span data-ttu-id="09953-544">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-545">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-545">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-546">1.7</span><span class="sxs-lookup"><span data-stu-id="09953-546">1.7</span></span>|
|[<span data-ttu-id="09953-547">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-547">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-548">ReadItem</span></span>|
|[<span data-ttu-id="09953-549">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-549">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-550">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-550">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="09953-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09953-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="09953-552">Proporciona acceso a los asistentes necesarios de un evento.</span><span class="sxs-lookup"><span data-stu-id="09953-552">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="09953-553">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="09953-553">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09953-554">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="09953-554">Read mode</span></span>

<span data-ttu-id="09953-555">La propiedad `requiredAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente requerido a la reunión.</span><span class="sxs-lookup"><span data-stu-id="09953-555">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="09953-556">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="09953-556">Compose mode</span></span>

<span data-ttu-id="09953-557">El `requiredAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes necesarios de una reunión.</span><span class="sxs-lookup"><span data-stu-id="09953-557">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-558">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-558">Type:</span></span>

*   <span data-ttu-id="09953-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09953-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-560">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-560">Requirements</span></span>

|<span data-ttu-id="09953-561">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-561">Requirement</span></span>|<span data-ttu-id="09953-562">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-563">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-563">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-564">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-564">1.0</span></span>|
|[<span data-ttu-id="09953-565">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-566">ReadItem</span></span>|
|[<span data-ttu-id="09953-567">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-568">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-568">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-569">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-569">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="09953-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="09953-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="09953-p126">Obtiene la dirección de correo electrónico del remitente de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="09953-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="09953-p127">Las propiedades [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) y `sender` representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="09953-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-575">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `sender` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="09953-575">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-576">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-576">Type:</span></span>

*   [<span data-ttu-id="09953-577">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="09953-577">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="09953-578">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-578">Requirements</span></span>

|<span data-ttu-id="09953-579">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-579">Requirement</span></span>|<span data-ttu-id="09953-580">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-581">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-581">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-582">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-582">1.0</span></span>|
|[<span data-ttu-id="09953-583">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-584">ReadItem</span></span>|
|[<span data-ttu-id="09953-585">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-586">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-586">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-587">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-587">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="09953-588">seriesId (nullable): cadena</span><span class="sxs-lookup"><span data-stu-id="09953-588">(nullable) seriesId :String</span></span>

<span data-ttu-id="09953-589">Obtiene el identificador de la serie a la que pertenece una instancia.</span><span class="sxs-lookup"><span data-stu-id="09953-589">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="09953-590">En OWA y Outlook, el `seriesId` devuelve el identificador de Exchange Web Services (EWS) del elemento primario (serie) al que pertenece este producto.</span><span class="sxs-lookup"><span data-stu-id="09953-590">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="09953-591">Sin embargo, en iOS y Android, el `seriesId` devuelve el identificador del resto del elemento primario.</span><span class="sxs-lookup"><span data-stu-id="09953-591">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-592">El identificador que devuelve la propiedad `seriesId` es el mismo que el identificador de elemento de Servicios Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="09953-592">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="09953-593">El `seriesId` propiedad no es idéntica a los identificadores de Outlook utilizadas por la API de REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="09953-593">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="09953-594">Antes de realizar llamadas a la API de REST con este valor, se debe convertir mediante [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="09953-594">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="09953-595">Para obtener más información, vea [usar las API de REST de Outlook desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="09953-595">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="09953-596">El `seriesId` propiedad devuelve `null` para los elementos que no tienen elementos primarios, como único citas, elementos de serie, o solicitudes de reunión y devuelve `undefined` para todos los elementos que no son las convocatorias de reunión.</span><span class="sxs-lookup"><span data-stu-id="09953-596">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-597">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-597">Type:</span></span>

* <span data-ttu-id="09953-598">String</span><span class="sxs-lookup"><span data-stu-id="09953-598">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-599">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-599">Requirements</span></span>

|<span data-ttu-id="09953-600">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-600">Requirement</span></span>|<span data-ttu-id="09953-601">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-602">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-602">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-603">1.7</span><span class="sxs-lookup"><span data-stu-id="09953-603">1.7</span></span>|
|[<span data-ttu-id="09953-604">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-604">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-605">ReadItem</span></span>|
|[<span data-ttu-id="09953-606">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-606">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-607">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-607">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-608">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-608">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId; 
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="09953-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="09953-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="09953-610">Obtiene o establece la fecha y la hora de inicio de la cita.</span><span class="sxs-lookup"><span data-stu-id="09953-610">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="09953-p130">La propiedad `start` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) para convertir el valor a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="09953-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09953-613">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="09953-613">Read mode</span></span>

<span data-ttu-id="09953-614">La propiedad `start` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="09953-614">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="09953-615">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="09953-615">Compose mode</span></span>

<span data-ttu-id="09953-616">La propiedad `start` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="09953-616">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="09953-617">Si usa el método [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) para establecer la hora de inicio, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="09953-617">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-618">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-618">Type:</span></span>

*   <span data-ttu-id="09953-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="09953-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-620">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-620">Requirements</span></span>

|<span data-ttu-id="09953-621">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-621">Requirement</span></span>|<span data-ttu-id="09953-622">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-622">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-623">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-623">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-624">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-624">1.0</span></span>|
|[<span data-ttu-id="09953-625">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-625">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-626">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-626">ReadItem</span></span>|
|[<span data-ttu-id="09953-627">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-627">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-628">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-628">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-629">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-629">Example</span></span>

<span data-ttu-id="09953-630">En el ejemplo siguiente, se establece la hora de inicio de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="09953-630">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="09953-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="09953-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="09953-632">Obtiene o establece la descripción que se muestra en el campo de asunto de un elemento.</span><span class="sxs-lookup"><span data-stu-id="09953-632">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="09953-633">La propiedad `subject` obtiene o establece el asunto completo del elemento, como lo ha enviado el servidor de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="09953-633">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09953-634">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="09953-634">Read mode</span></span>

<span data-ttu-id="09953-p131">La propiedad `subject` devuelve una cadena. Use la propiedad [`normalizedSubject`](#normalizedsubject-string) para obtener el asunto menos cualquier prefijo inicial como `RE:` y `FW:`.</span><span class="sxs-lookup"><span data-stu-id="09953-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="09953-637">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="09953-637">Compose mode</span></span>

<span data-ttu-id="09953-638">La propiedad `subject` devuelve un objeto `Subject` que proporciona métodos para obtener y establecer el asunto.</span><span class="sxs-lookup"><span data-stu-id="09953-638">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="09953-639">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-639">Type:</span></span>

*   <span data-ttu-id="09953-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="09953-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-641">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-641">Requirements</span></span>

|<span data-ttu-id="09953-642">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-642">Requirement</span></span>|<span data-ttu-id="09953-643">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-644">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-644">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-645">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-645">1.0</span></span>|
|[<span data-ttu-id="09953-646">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-647">ReadItem</span></span>|
|[<span data-ttu-id="09953-648">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-649">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-649">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="09953-650">para: matriz. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09953-650">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="09953-651">Proporciona acceso a los destinatarios en la línea **para** de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="09953-651">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="09953-652">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="09953-652">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09953-653">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="09953-653">Read mode</span></span>

<span data-ttu-id="09953-p133">La propiedad `to` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **Para** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="09953-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="09953-656">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="09953-656">Compose mode</span></span>

<span data-ttu-id="09953-657">El `to` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **para** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="09953-657">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="09953-658">Tipo:</span><span class="sxs-lookup"><span data-stu-id="09953-658">Type:</span></span>

*   <span data-ttu-id="09953-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09953-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-660">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-660">Requirements</span></span>

|<span data-ttu-id="09953-661">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-661">Requirement</span></span>|<span data-ttu-id="09953-662">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-662">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-663">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-663">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-664">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-664">1.0</span></span>|
|[<span data-ttu-id="09953-665">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-665">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-666">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-666">ReadItem</span></span>|
|[<span data-ttu-id="09953-667">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-667">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-668">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-668">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-669">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-669">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="09953-670">Métodos</span><span class="sxs-lookup"><span data-stu-id="09953-670">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="09953-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="09953-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="09953-672">Agrega un archivo a un mensaje o cita como datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="09953-672">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="09953-673">El método `addFileAttachmentAsync` carga el archivo en el URI especificado y lo asocia al elemento en el formulario de redacción.</span><span class="sxs-lookup"><span data-stu-id="09953-673">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="09953-674">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="09953-674">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09953-675">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="09953-675">Parameters:</span></span>
|<span data-ttu-id="09953-676">Nombre</span><span class="sxs-lookup"><span data-stu-id="09953-676">Name</span></span>|<span data-ttu-id="09953-677">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-677">Type</span></span>|<span data-ttu-id="09953-678">Atributos</span><span class="sxs-lookup"><span data-stu-id="09953-678">Attributes</span></span>|<span data-ttu-id="09953-679">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-679">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="09953-680">String</span><span class="sxs-lookup"><span data-stu-id="09953-680">String</span></span>||<span data-ttu-id="09953-p134">El URI que proporciona la ubicación del archivo que se va a adjuntar al mensaje o a la cita. La longitud máxima es de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="09953-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="09953-683">String</span><span class="sxs-lookup"><span data-stu-id="09953-683">String</span></span>||<span data-ttu-id="09953-p135">El nombre de los datos adjuntos que se muestra mientras estos se cargan. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="09953-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="09953-686">Object</span><span class="sxs-lookup"><span data-stu-id="09953-686">Object</span></span>|<span data-ttu-id="09953-687">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-687">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-688">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="09953-688">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="09953-689">Objeto</span><span class="sxs-lookup"><span data-stu-id="09953-689">Object</span></span>|<span data-ttu-id="09953-690">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-690">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-691">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="09953-691">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="09953-692">Booleano</span><span class="sxs-lookup"><span data-stu-id="09953-692">Boolean</span></span>|<span data-ttu-id="09953-693">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-693">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-694">Si está establecido en `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="09953-694">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="09953-695">Función</span><span class="sxs-lookup"><span data-stu-id="09953-695">function</span></span>|<span data-ttu-id="09953-696">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-696">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-697">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="09953-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="09953-698">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="09953-698">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="09953-699">Si se produce un error en la carga de los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="09953-699">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="09953-700">Errores</span><span class="sxs-lookup"><span data-stu-id="09953-700">Errors</span></span>

|<span data-ttu-id="09953-701">Código de error</span><span class="sxs-lookup"><span data-stu-id="09953-701">Error code</span></span>|<span data-ttu-id="09953-702">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-702">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="09953-703">Los datos adjuntos son más grandes de lo permitido.</span><span class="sxs-lookup"><span data-stu-id="09953-703">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="09953-704">Los datos adjuntos tienen una extensión que no está permitida.</span><span class="sxs-lookup"><span data-stu-id="09953-704">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="09953-705">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="09953-705">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09953-706">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-706">Requirements</span></span>

|<span data-ttu-id="09953-707">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-707">Requirement</span></span>|<span data-ttu-id="09953-708">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-708">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-709">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-709">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-710">1.1</span><span class="sxs-lookup"><span data-stu-id="09953-710">1.1</span></span>|
|[<span data-ttu-id="09953-711">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-711">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-712">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="09953-712">ReadWriteItem</span></span>|
|[<span data-ttu-id="09953-713">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-713">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-714">Redacción</span><span class="sxs-lookup"><span data-stu-id="09953-714">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="09953-715">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="09953-715">Examples</span></span>

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

<span data-ttu-id="09953-716">En el siguiente ejemplo, se agrega un archivo de imagen como un dato adjunto en línea y hace referencia a los datos adjuntos del cuerpo del mensaje.</span><span class="sxs-lookup"><span data-stu-id="09953-716">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="09953-717">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="09953-717">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="09953-718">Agrega un controlador de eventos para un evento admitido.</span><span class="sxs-lookup"><span data-stu-id="09953-718">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="09953-719">Actualmente son los tipos de eventos compatibles `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, y`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="09953-719">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="09953-720">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="09953-720">Parameters:</span></span>

| <span data-ttu-id="09953-721">Nombre</span><span class="sxs-lookup"><span data-stu-id="09953-721">Name</span></span> | <span data-ttu-id="09953-722">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-722">Type</span></span> | <span data-ttu-id="09953-723">Atributos</span><span class="sxs-lookup"><span data-stu-id="09953-723">Attributes</span></span> | <span data-ttu-id="09953-724">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-724">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="09953-725">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="09953-725">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="09953-726">El evento que debe invocar el controlador.</span><span class="sxs-lookup"><span data-stu-id="09953-726">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="09953-727">Función</span><span class="sxs-lookup"><span data-stu-id="09953-727">Function</span></span> || <span data-ttu-id="09953-p136">La función que va a controlar el evento. La función debe aceptar un único parámetro, que es un literal de objeto. La propiedad `type` del parámetro coincidirá con el parámetro `eventType` que se ha pasado a `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="09953-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="09953-731">Objeto</span><span class="sxs-lookup"><span data-stu-id="09953-731">Object</span></span> | <span data-ttu-id="09953-732">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-732">&lt;optional&gt;</span></span> | <span data-ttu-id="09953-733">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="09953-733">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="09953-734">Objeto</span><span class="sxs-lookup"><span data-stu-id="09953-734">Object</span></span> | <span data-ttu-id="09953-735">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-735">&lt;optional&gt;</span></span> | <span data-ttu-id="09953-736">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="09953-736">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="09953-737">función</span><span class="sxs-lookup"><span data-stu-id="09953-737">function</span></span>| <span data-ttu-id="09953-738">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-738">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-739">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="09953-739">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09953-740">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-740">Requirements</span></span>

|<span data-ttu-id="09953-741">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-741">Requirement</span></span>| <span data-ttu-id="09953-742">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-742">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-743">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-743">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09953-744">1.7</span><span class="sxs-lookup"><span data-stu-id="09953-744">1.7</span></span> |
|[<span data-ttu-id="09953-745">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-745">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09953-746">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-746">ReadItem</span></span> |
|[<span data-ttu-id="09953-747">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-747">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09953-748">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-748">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="09953-749">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-749">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="09953-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="09953-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="09953-751">Agrega un elemento de Exchange (por ejemplo, un mensaje) como datos adjuntos al mensaje o a la cita.</span><span class="sxs-lookup"><span data-stu-id="09953-751">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="09953-p137">El método `addItemAttachmentAsync` asocia el elemento con el identificador especificado de Exchange al elemento en el formulario de redacción. Si se especifica un método de devolución de llamada, se llama al método con un parámetro, `asyncResult`, que contiene el identificador de datos adjuntos o un código que indique cualquier error que se ha producido al adjuntar el elemento. Si es necesario, puede usar el parámetro `options` para pasar información de estado al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="09953-p137">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="09953-755">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="09953-755">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="09953-756">Si el complemento de Office se está ejecutando en Outlook Web App, el `addItemAttachmentAsync` método puede adjuntar elementos a los elementos que no sea el elemento que se está editando; Sin embargo, esto no se admite y no se recomienda.</span><span class="sxs-lookup"><span data-stu-id="09953-756">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09953-757">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="09953-757">Parameters:</span></span>

|<span data-ttu-id="09953-758">Nombre</span><span class="sxs-lookup"><span data-stu-id="09953-758">Name</span></span>|<span data-ttu-id="09953-759">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-759">Type</span></span>|<span data-ttu-id="09953-760">Atributos</span><span class="sxs-lookup"><span data-stu-id="09953-760">Attributes</span></span>|<span data-ttu-id="09953-761">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-761">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="09953-762">String</span><span class="sxs-lookup"><span data-stu-id="09953-762">String</span></span>||<span data-ttu-id="09953-p138">El identificador de Exchange del elemento que debe adjuntarse. La longitud máxima es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="09953-p138">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="09953-765">String</span><span class="sxs-lookup"><span data-stu-id="09953-765">String</span></span>||<span data-ttu-id="09953-p139">El asunto del elemento que debe adjuntarse. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="09953-p139">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="09953-768">Object</span><span class="sxs-lookup"><span data-stu-id="09953-768">Object</span></span>|<span data-ttu-id="09953-769">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-769">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-770">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="09953-770">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="09953-771">Objeto</span><span class="sxs-lookup"><span data-stu-id="09953-771">Object</span></span>|<span data-ttu-id="09953-772">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-772">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-773">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="09953-773">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="09953-774">función</span><span class="sxs-lookup"><span data-stu-id="09953-774">function</span></span>|<span data-ttu-id="09953-775">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-775">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-776">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="09953-776">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="09953-777">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="09953-777">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="09953-778">Si se produce un error al agregar los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="09953-778">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="09953-779">Errores</span><span class="sxs-lookup"><span data-stu-id="09953-779">Errors</span></span>

|<span data-ttu-id="09953-780">Código de error</span><span class="sxs-lookup"><span data-stu-id="09953-780">Error code</span></span>|<span data-ttu-id="09953-781">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-781">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="09953-782">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="09953-782">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09953-783">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-783">Requirements</span></span>

|<span data-ttu-id="09953-784">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-784">Requirement</span></span>|<span data-ttu-id="09953-785">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-786">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-786">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-787">1.1</span><span class="sxs-lookup"><span data-stu-id="09953-787">1.1</span></span>|
|[<span data-ttu-id="09953-788">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-789">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="09953-789">ReadWriteItem</span></span>|
|[<span data-ttu-id="09953-790">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-791">Redacción</span><span class="sxs-lookup"><span data-stu-id="09953-791">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-792">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-792">Example</span></span>

<span data-ttu-id="09953-793">En el ejemplo siguiente, se agrega un elemento existente de Outlook como un archivo adjunto con el nombre `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="09953-793">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="09953-794">close()</span><span class="sxs-lookup"><span data-stu-id="09953-794">close()</span></span>

<span data-ttu-id="09953-795">Cierra el elemento actual que se está redactando.</span><span class="sxs-lookup"><span data-stu-id="09953-795">Closes the current item that is being composed.</span></span>

<span data-ttu-id="09953-p140">El comportamiento del método `close` depende del estado actual del elemento que se está redactando. Si el elemento tiene cambios sin guardar, el cliente solicita al usuario que guarde, descarte o cancele la acción Cerrar.</span><span class="sxs-lookup"><span data-stu-id="09953-p140">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-798">En Outlook, en el sitio web, si el elemento es una cita y anteriormente se ha guardado con `saveAsync`, se pide al usuario al guardar, descartar o cancelar incluso si no ha habido cambios desde la último el elemento guardado.</span><span class="sxs-lookup"><span data-stu-id="09953-798">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="09953-799">En el cliente de escritorio de Outlook, si el mensaje es una respuesta directa, el método `close` no tiene ningún efecto.</span><span class="sxs-lookup"><span data-stu-id="09953-799">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-800">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-800">Requirements</span></span>

|<span data-ttu-id="09953-801">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-801">Requirement</span></span>|<span data-ttu-id="09953-802">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-802">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-803">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-803">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-804">1.3</span><span class="sxs-lookup"><span data-stu-id="09953-804">1.3</span></span>|
|[<span data-ttu-id="09953-805">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-805">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-806">Restringido</span><span class="sxs-lookup"><span data-stu-id="09953-806">Restricted</span></span>|
|[<span data-ttu-id="09953-807">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-807">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-808">Redacción</span><span class="sxs-lookup"><span data-stu-id="09953-808">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="09953-809">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="09953-809">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="09953-810">Muestra un formulario de respuesta que incluye el remitente y todos los destinatarios del mensaje seleccionado o el organizador y todos los asistentes de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="09953-810">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-811">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="09953-811">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="09953-812">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="09953-812">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="09953-813">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyAllForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="09953-813">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="09953-p141">Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="09953-p141">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09953-817">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="09953-817">Parameters:</span></span>

|<span data-ttu-id="09953-818">Nombre</span><span class="sxs-lookup"><span data-stu-id="09953-818">Name</span></span>|<span data-ttu-id="09953-819">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-819">Type</span></span>|<span data-ttu-id="09953-820">Atributos</span><span class="sxs-lookup"><span data-stu-id="09953-820">Attributes</span></span>|<span data-ttu-id="09953-821">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-821">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="09953-822">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="09953-822">String &#124; Object</span></span>||<span data-ttu-id="09953-p142">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="09953-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="09953-825">**O**</span><span class="sxs-lookup"><span data-stu-id="09953-825">**OR**</span></span><br/><span data-ttu-id="09953-p143">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="09953-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="09953-828">String</span><span class="sxs-lookup"><span data-stu-id="09953-828">String</span></span>|<span data-ttu-id="09953-829">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-829">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-p144">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="09953-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="09953-832">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-832">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="09953-833">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-833">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-834">Una matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="09953-834">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="09953-835">String</span><span class="sxs-lookup"><span data-stu-id="09953-835">String</span></span>||<span data-ttu-id="09953-p145">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="09953-p145">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="09953-838">String</span><span class="sxs-lookup"><span data-stu-id="09953-838">String</span></span>||<span data-ttu-id="09953-839">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="09953-839">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="09953-840">String</span><span class="sxs-lookup"><span data-stu-id="09953-840">String</span></span>||<span data-ttu-id="09953-p146">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="09953-p146">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="09953-843">Booleano</span><span class="sxs-lookup"><span data-stu-id="09953-843">Boolean</span></span>||<span data-ttu-id="09953-p147">Solo se usa si `type` se establece en `file`. Si `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="09953-p147">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="09953-846">String</span><span class="sxs-lookup"><span data-stu-id="09953-846">String</span></span>||<span data-ttu-id="09953-p148">Solo se usa si `type` se establece en `item`. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="09953-p148">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="09953-850">función</span><span class="sxs-lookup"><span data-stu-id="09953-850">function</span></span>|<span data-ttu-id="09953-851">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-851">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-852">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="09953-852">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09953-853">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-853">Requirements</span></span>

|<span data-ttu-id="09953-854">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-854">Requirement</span></span>|<span data-ttu-id="09953-855">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-856">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-856">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-857">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-857">1.0</span></span>|
|[<span data-ttu-id="09953-858">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-858">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-859">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-859">ReadItem</span></span>|
|[<span data-ttu-id="09953-860">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-860">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-861">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-861">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="09953-862">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="09953-862">Examples</span></span>

<span data-ttu-id="09953-863">El código siguiente pasa una cadena a la función `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="09953-863">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="09953-864">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="09953-864">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="09953-865">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="09953-865">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="09953-866">Responder con un cuerpo y datos adjuntos de archivo.</span><span class="sxs-lookup"><span data-stu-id="09953-866">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="09953-867">Responder con un cuerpo y datos adjuntos de elemento.</span><span class="sxs-lookup"><span data-stu-id="09953-867">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="09953-868">Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="09953-868">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="09953-869">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="09953-869">displayReplyForm(formData)</span></span>

<span data-ttu-id="09953-870">Muestra un formulario de respuesta que incluye solo el remitente del mensaje seleccionado o el organizador de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="09953-870">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-871">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="09953-871">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="09953-872">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="09953-872">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="09953-873">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="09953-873">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="09953-p149">Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="09953-p149">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09953-877">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="09953-877">Parameters:</span></span>

|<span data-ttu-id="09953-878">Nombre</span><span class="sxs-lookup"><span data-stu-id="09953-878">Name</span></span>|<span data-ttu-id="09953-879">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-879">Type</span></span>|<span data-ttu-id="09953-880">Atributos</span><span class="sxs-lookup"><span data-stu-id="09953-880">Attributes</span></span>|<span data-ttu-id="09953-881">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-881">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="09953-882">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="09953-882">String &#124; Object</span></span>||<span data-ttu-id="09953-p150">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="09953-p150">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="09953-885">**O**</span><span class="sxs-lookup"><span data-stu-id="09953-885">**OR**</span></span><br/><span data-ttu-id="09953-p151">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="09953-p151">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="09953-888">String</span><span class="sxs-lookup"><span data-stu-id="09953-888">String</span></span>|<span data-ttu-id="09953-889">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-889">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-p152">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="09953-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="09953-892">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-892">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="09953-893">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-893">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-894">Una matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="09953-894">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="09953-895">String</span><span class="sxs-lookup"><span data-stu-id="09953-895">String</span></span>||<span data-ttu-id="09953-p153">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="09953-p153">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="09953-898">String</span><span class="sxs-lookup"><span data-stu-id="09953-898">String</span></span>||<span data-ttu-id="09953-899">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="09953-899">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="09953-900">String</span><span class="sxs-lookup"><span data-stu-id="09953-900">String</span></span>||<span data-ttu-id="09953-p154">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="09953-p154">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="09953-903">Booleano</span><span class="sxs-lookup"><span data-stu-id="09953-903">Boolean</span></span>||<span data-ttu-id="09953-p155">Solo se usa si `type` se establece en `file`. Si `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="09953-p155">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="09953-906">String</span><span class="sxs-lookup"><span data-stu-id="09953-906">String</span></span>||<span data-ttu-id="09953-p156">Solo se usa si `type` se establece en `item`. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="09953-p156">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="09953-910">función</span><span class="sxs-lookup"><span data-stu-id="09953-910">function</span></span>|<span data-ttu-id="09953-911">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-911">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-912">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="09953-912">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09953-913">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-913">Requirements</span></span>

|<span data-ttu-id="09953-914">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-914">Requirement</span></span>|<span data-ttu-id="09953-915">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-915">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-916">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-916">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-917">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-917">1.0</span></span>|
|[<span data-ttu-id="09953-918">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-918">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-919">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-919">ReadItem</span></span>|
|[<span data-ttu-id="09953-920">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-920">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-921">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-921">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="09953-922">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="09953-922">Examples</span></span>

<span data-ttu-id="09953-923">El código siguiente pasa una cadena a la función `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="09953-923">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="09953-924">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="09953-924">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="09953-925">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="09953-925">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="09953-926">Responder con un cuerpo y datos adjuntos de archivo.</span><span class="sxs-lookup"><span data-stu-id="09953-926">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="09953-927">Responder con un cuerpo y datos adjuntos de elemento.</span><span class="sxs-lookup"><span data-stu-id="09953-927">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="09953-928">Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="09953-928">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="09953-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="09953-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="09953-930">Obtiene las entidades que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="09953-930">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-931">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="09953-931">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-932">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-932">Requirements</span></span>

|<span data-ttu-id="09953-933">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-933">Requirement</span></span>|<span data-ttu-id="09953-934">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-935">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-935">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-936">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-936">1.0</span></span>|
|[<span data-ttu-id="09953-937">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-938">ReadItem</span></span>|
|[<span data-ttu-id="09953-939">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-940">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="09953-941">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="09953-941">Returns:</span></span>

<span data-ttu-id="09953-942">Tipo: [Entidades](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="09953-942">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="09953-943">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-943">Example</span></span>

<span data-ttu-id="09953-944">En el ejemplo siguiente se obtiene acceso a las entidades de los contactos en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="09953-944">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="09953-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="09953-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="09953-946">Obtiene una matriz de todas las entidades del tipo de entidad especificado que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="09953-946">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-947">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="09953-947">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09953-948">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="09953-948">Parameters:</span></span>

|<span data-ttu-id="09953-949">Nombre</span><span class="sxs-lookup"><span data-stu-id="09953-949">Name</span></span>|<span data-ttu-id="09953-950">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-950">Type</span></span>|<span data-ttu-id="09953-951">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-951">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="09953-952">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="09953-952">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="09953-953">Uno de los valores de enumeración de EntityType.</span><span class="sxs-lookup"><span data-stu-id="09953-953">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09953-954">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-954">Requirements</span></span>

|<span data-ttu-id="09953-955">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-955">Requirement</span></span>|<span data-ttu-id="09953-956">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-957">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-957">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-958">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-958">1.0</span></span>|
|[<span data-ttu-id="09953-959">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-960">Restringido</span><span class="sxs-lookup"><span data-stu-id="09953-960">Restricted</span></span>|
|[<span data-ttu-id="09953-961">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-962">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="09953-963">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="09953-963">Returns:</span></span>

<span data-ttu-id="09953-964">Si el valor que se pasa a `entityType` no es un miembro válido de la enumeración de `EntityType`, el método devuelve un valor NULL.</span><span class="sxs-lookup"><span data-stu-id="09953-964">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="09953-965">Si no existen entidades del tipo especificado están presentes en el cuerpo del elemento, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="09953-965">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="09953-966">De otro modo, el tipo de los objetos en la matriz devuelta depende del tipo de entidad solicitada en el parámetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="09953-966">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="09953-967">Aunque el nivel de permiso mínimo para usar este método es **Restricted**, algunos tipos de entidad requieren **ReadItem** para tener acceso, como se especifica en la tabla siguiente.</span><span class="sxs-lookup"><span data-stu-id="09953-967">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="09953-968">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="09953-968">Value of `entityType`</span></span>|<span data-ttu-id="09953-969">Tipo de objetos de la matriz devuelta</span><span class="sxs-lookup"><span data-stu-id="09953-969">Type of objects in returned array</span></span>|<span data-ttu-id="09953-970">Nivel de permiso requerido</span><span class="sxs-lookup"><span data-stu-id="09953-970">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="09953-971">Cadena</span><span class="sxs-lookup"><span data-stu-id="09953-971">String</span></span>|<span data-ttu-id="09953-972">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="09953-972">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="09953-973">Contacto</span><span class="sxs-lookup"><span data-stu-id="09953-973">Contact</span></span>|<span data-ttu-id="09953-974">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="09953-974">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="09953-975">Cadena</span><span class="sxs-lookup"><span data-stu-id="09953-975">String</span></span>|<span data-ttu-id="09953-976">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="09953-976">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="09953-977">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="09953-977">MeetingSuggestion</span></span>|<span data-ttu-id="09953-978">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="09953-978">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="09953-979">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="09953-979">PhoneNumber</span></span>|<span data-ttu-id="09953-980">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="09953-980">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="09953-981">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="09953-981">TaskSuggestion</span></span>|<span data-ttu-id="09953-982">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="09953-982">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="09953-983">Cadena</span><span class="sxs-lookup"><span data-stu-id="09953-983">String</span></span>|<span data-ttu-id="09953-984">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="09953-984">**Restricted**</span></span>|

<span data-ttu-id="09953-985">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="09953-985">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="09953-986">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-986">Example</span></span>

<span data-ttu-id="09953-987">En el ejemplo siguiente se muestra cómo obtener acceso a una matriz de cadenas que representan direcciones postales en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="09953-987">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="09953-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="09953-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="09953-989">Devuelve entidades conocidas en el elemento seleccionado que pasan el filtro con nombre definido en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="09953-989">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-990">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="09953-990">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="09953-991">El método `getFilteredEntitiesByName` devuelve las entidades que coinciden con la expresión regular definida en el elemento de regla [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) del archivo XML de manifiesto con el valor de elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="09953-991">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09953-992">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="09953-992">Parameters:</span></span>

|<span data-ttu-id="09953-993">Nombre</span><span class="sxs-lookup"><span data-stu-id="09953-993">Name</span></span>|<span data-ttu-id="09953-994">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-994">Type</span></span>|<span data-ttu-id="09953-995">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-995">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="09953-996">String</span><span class="sxs-lookup"><span data-stu-id="09953-996">String</span></span>|<span data-ttu-id="09953-997">El nombre del elemento de regla `ItemHasKnownEntity` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="09953-997">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09953-998">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-998">Requirements</span></span>

|<span data-ttu-id="09953-999">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-999">Requirement</span></span>|<span data-ttu-id="09953-1000">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-1000">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-1001">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-1001">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-1002">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-1002">1.0</span></span>|
|[<span data-ttu-id="09953-1003">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-1003">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-1004">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-1004">ReadItem</span></span>|
|[<span data-ttu-id="09953-1005">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-1005">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-1006">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-1006">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="09953-1007">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="09953-1007">Returns:</span></span>

<span data-ttu-id="09953-p158">Si no existe ningún elemento `ItemHasKnownEntity` en el manifiesto con un valor de elemento `FilterName` que coincida con el parámetro `name`, el método devuelve `null`. Si el parámetro `name` no coincide con un elemento `ItemHasKnownEntity` del manifiesto, pero no existen entidades en el elemento actual que coincidan, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="09953-p158">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="09953-1010">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="09953-1010">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="09953-1011">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="09953-1011">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="09953-1012">Devuelve valores de cadena en el elemento seleccionado que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="09953-1012">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-1013">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="09953-1013">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="09953-p159">El método `getRegExMatches` devuelve las cadenas que coinciden con la expresión regular definida en cada elemento de regla `ItemHasRegularExpressionMatch` o `ItemHasKnownEntity` del archivo XML de manifiesto. Para una regla `ItemHasRegularExpressionMatch`, tiene que darse una cadena coincidente en la propiedad del elemento especificado por la regla. El tipo simple `PropertyName` define las propiedades admitidas.</span><span class="sxs-lookup"><span data-stu-id="09953-p159">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="09953-1017">Por ejemplo, considere un manifiesto de complemento que tenga el siguiente elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="09953-1017">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="09953-1018">El objeto devuelto desde `getRegExMatches` tendría dos propiedades: `fruits` y `veggies`.</span><span class="sxs-lookup"><span data-stu-id="09953-1018">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="09953-p160">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados. En su lugar, use el método [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) para recuperar todo el cuerpo.</span><span class="sxs-lookup"><span data-stu-id="09953-p160">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-1022">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-1022">Requirements</span></span>

|<span data-ttu-id="09953-1023">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-1023">Requirement</span></span>|<span data-ttu-id="09953-1024">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-1024">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-1025">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-1025">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-1026">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-1026">1.0</span></span>|
|[<span data-ttu-id="09953-1027">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-1027">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-1028">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-1028">ReadItem</span></span>|
|[<span data-ttu-id="09953-1029">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-1029">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-1030">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-1030">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="09953-1031">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="09953-1031">Returns:</span></span>

<span data-ttu-id="09953-p161">Un objeto que contiene matrices de cadenas que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. El nombre de cada matriz es igual al valor correspondiente del atributo `RegExName` de la regla `ItemHasRegularExpressionMatch` coincidente o al atributo `FilterName` de la regla `ItemHasKnownEntity` coincidente.</span><span class="sxs-lookup"><span data-stu-id="09953-p161">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="09953-1034">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="09953-1034">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="09953-1035">Object</span><span class="sxs-lookup"><span data-stu-id="09953-1035">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="09953-1036">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-1036">Example</span></span>

<span data-ttu-id="09953-1037">En el siguiente ejemplo, se muestra cómo tener acceso a la matriz de coincidencias de los elementos de regla de expresión regular `fruits` y `veggies`, que se especifican en el manifiesto.</span><span class="sxs-lookup"><span data-stu-id="09953-1037">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="09953-1038">→ de getregexmatchesbyname (Name) (nullable) {matriz. < cadena >}</span><span class="sxs-lookup"><span data-stu-id="09953-1038">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="09953-1039">Devuelve valores de cadena en el elemento seleccionado que coinciden con la expresión regular de nombre definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="09953-1039">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-1040">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="09953-1040">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="09953-1041">El método `getRegExMatchesByName` devuelve las cadenas que coinciden con la expresión regular definida en el elemento de regla `ItemHasRegularExpressionMatch` del archivo XML de manifiesto con el valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="09953-1041">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="09953-p162">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="09953-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09953-1044">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="09953-1044">Parameters:</span></span>

|<span data-ttu-id="09953-1045">Nombre</span><span class="sxs-lookup"><span data-stu-id="09953-1045">Name</span></span>|<span data-ttu-id="09953-1046">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-1046">Type</span></span>|<span data-ttu-id="09953-1047">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-1047">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="09953-1048">String</span><span class="sxs-lookup"><span data-stu-id="09953-1048">String</span></span>|<span data-ttu-id="09953-1049">El nombre del elemento de regla `ItemHasRegularExpressionMatch` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="09953-1049">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09953-1050">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-1050">Requirements</span></span>

|<span data-ttu-id="09953-1051">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-1051">Requirement</span></span>|<span data-ttu-id="09953-1052">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-1053">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-1053">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-1054">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-1054">1.0</span></span>|
|[<span data-ttu-id="09953-1055">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-1056">ReadItem</span></span>|
|[<span data-ttu-id="09953-1057">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-1058">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-1058">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="09953-1059">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="09953-1059">Returns:</span></span>

<span data-ttu-id="09953-1060">Una matriz que contiene las cadenas que coinciden con la expresión regular definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="09953-1060">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="09953-1061">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="09953-1061">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="09953-1062">< Cadena > de matriz.</span><span class="sxs-lookup"><span data-stu-id="09953-1062">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="09953-1063">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-1063">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="09953-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="09953-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="09953-1065">Devuelve asincrónicamente datos seleccionados desde el asunto o el cuerpo de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="09953-1065">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="09953-p163">Si no hay ninguna selección, pero el cursor está en el cuerpo o el asunto, el método devuelve null para los datos seleccionados. Si se selecciona un campo que no sea el cuerpo o el asunto, el método devuelve el error `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="09953-p163">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09953-1068">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="09953-1068">Parameters:</span></span>

|<span data-ttu-id="09953-1069">Nombre</span><span class="sxs-lookup"><span data-stu-id="09953-1069">Name</span></span>|<span data-ttu-id="09953-1070">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-1070">Type</span></span>|<span data-ttu-id="09953-1071">Atributos</span><span class="sxs-lookup"><span data-stu-id="09953-1071">Attributes</span></span>|<span data-ttu-id="09953-1072">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-1072">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="09953-1073">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="09953-1073">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="09953-p164">Solicita un formato para los datos. Si es Texto, el método devuelve el texto sin formato como cadena, quitando toda etiqueta HTML que hubiera. Si es HTML, el método devuelve el texto seleccionado, ya sea texto sin formato o HTML.</span><span class="sxs-lookup"><span data-stu-id="09953-p164">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="09953-1077">Object</span><span class="sxs-lookup"><span data-stu-id="09953-1077">Object</span></span>|<span data-ttu-id="09953-1078">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-1078">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-1079">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="09953-1079">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="09953-1080">Objeto</span><span class="sxs-lookup"><span data-stu-id="09953-1080">Object</span></span>|<span data-ttu-id="09953-1081">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-1081">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-1082">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="09953-1082">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="09953-1083">función</span><span class="sxs-lookup"><span data-stu-id="09953-1083">function</span></span>||<span data-ttu-id="09953-1084">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="09953-1084">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="09953-1085">Para tener acceso a los datos seleccionados desde el método de devolución de llamada, llame a `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="09953-1085">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="09953-1086">Para obtener acceso a la propiedad de origen del que proviene la selección, llame a `asyncResult.value.sourceProperty`, que será cualquiera `body` o `subject`.</span><span class="sxs-lookup"><span data-stu-id="09953-1086">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09953-1087">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-1087">Requirements</span></span>

|<span data-ttu-id="09953-1088">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-1088">Requirement</span></span>|<span data-ttu-id="09953-1089">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-1089">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-1090">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-1090">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-1091">1.2</span><span class="sxs-lookup"><span data-stu-id="09953-1091">1.2</span></span>|
|[<span data-ttu-id="09953-1092">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-1092">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-1093">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="09953-1093">ReadWriteItem</span></span>|
|[<span data-ttu-id="09953-1094">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-1094">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-1095">Redacción</span><span class="sxs-lookup"><span data-stu-id="09953-1095">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="09953-1096">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="09953-1096">Returns:</span></span>

<span data-ttu-id="09953-1097">Los datos seleccionados como cadena con formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="09953-1097">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="09953-1098">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="09953-1098">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="09953-1099">String</span><span class="sxs-lookup"><span data-stu-id="09953-1099">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="09953-1100">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-1100">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="09953-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="09953-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="09953-p166">Obtiene las entidades que se han detectado en una coincidencia resaltada que un usuario ha seleccionado. Las coincidencias resaltadas se aplican a [complementos contextuales](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="09953-p166">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="09953-1104">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="09953-1104">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-1105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-1105">Requirements</span></span>

|<span data-ttu-id="09953-1106">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-1106">Requirement</span></span>|<span data-ttu-id="09953-1107">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-1107">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-1108">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-1108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-1109">1.6</span><span class="sxs-lookup"><span data-stu-id="09953-1109">1.6</span></span>|
|[<span data-ttu-id="09953-1110">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-1110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-1111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-1111">ReadItem</span></span>|
|[<span data-ttu-id="09953-1112">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-1112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-1113">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-1113">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="09953-1114">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="09953-1114">Returns:</span></span>

<span data-ttu-id="09953-1115">Tipo: [Entidades](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="09953-1115">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="09953-1116">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-1116">Example</span></span>

<span data-ttu-id="09953-1117">En el siguiente ejemplo se accede a las entidades de direcciones en la coincidencia resaltada seleccionada por el usuario.</span><span class="sxs-lookup"><span data-stu-id="09953-1117">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="09953-1118">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="09953-1118">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="09953-p167">Devuelve valores de cadena en una coincidencia resaltada que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. Las coincidencias resaltadas se aplican a [complementos contextuales](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="09953-p167">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="09953-1121">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="09953-1121">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="09953-p168">El método `getSelectedRegExMatches` devuelve las cadenas que coinciden con la expresión regular definida en cada elemento de regla `ItemHasRegularExpressionMatch` o `ItemHasKnownEntity` del archivo XML de manifiesto. Para una regla `ItemHasRegularExpressionMatch`, tiene que darse una cadena coincidente en la propiedad del elemento especificado por la regla. El tipo simple `PropertyName` define las propiedades admitidas.</span><span class="sxs-lookup"><span data-stu-id="09953-p168">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="09953-1125">Por ejemplo, considere un manifiesto de complemento que tenga el siguiente elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="09953-1125">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="09953-1126">El objeto devuelto desde `getRegExMatches` tendría dos propiedades: `fruits` y `veggies`.</span><span class="sxs-lookup"><span data-stu-id="09953-1126">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="09953-p169">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados. En su lugar, use el método [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) para recuperar todo el cuerpo.</span><span class="sxs-lookup"><span data-stu-id="09953-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="09953-1130">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-1130">Requirements</span></span>

|<span data-ttu-id="09953-1131">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-1131">Requirement</span></span>|<span data-ttu-id="09953-1132">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-1133">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-1133">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-1134">1.6</span><span class="sxs-lookup"><span data-stu-id="09953-1134">1.6</span></span>|
|[<span data-ttu-id="09953-1135">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-1135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-1136">ReadItem</span></span>|
|[<span data-ttu-id="09953-1137">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-1137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-1138">Lectura</span><span class="sxs-lookup"><span data-stu-id="09953-1138">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="09953-1139">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="09953-1139">Returns:</span></span>

<span data-ttu-id="09953-p170">Un objeto que contiene matrices de cadenas que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. El nombre de cada matriz es igual al valor correspondiente del atributo `RegExName` de la regla `ItemHasRegularExpressionMatch` coincidente o al atributo `FilterName` de la regla `ItemHasKnownEntity` coincidente.</span><span class="sxs-lookup"><span data-stu-id="09953-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="09953-1142">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-1142">Example</span></span>

<span data-ttu-id="09953-1143">En el siguiente ejemplo, se muestra cómo tener acceso a la matriz de coincidencias de los elementos de regla de expresión regular `fruits` y `veggies`, que se especifican en el manifiesto.</span><span class="sxs-lookup"><span data-stu-id="09953-1143">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="09953-1144">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="09953-1144">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="09953-1145">Carga de forma asincrónica las propiedades personalizadas de este complemento en el elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="09953-1145">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="09953-p171">Las propiedades personalizadas se almacenan como pares de clave/valor según la aplicación y el elemento. Este método devuelve un objeto `CustomProperties` en la devolución de llamada, que proporciona métodos para tener acceso a las propiedades personalizadas específicas del elemento y el complemento actual. Las propiedades personalizadas no están cifradas en el elemento, por lo que no debería usarse como almacenamiento seguro.</span><span class="sxs-lookup"><span data-stu-id="09953-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09953-1149">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="09953-1149">Parameters:</span></span>

|<span data-ttu-id="09953-1150">Nombre</span><span class="sxs-lookup"><span data-stu-id="09953-1150">Name</span></span>|<span data-ttu-id="09953-1151">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-1151">Type</span></span>|<span data-ttu-id="09953-1152">Atributos</span><span class="sxs-lookup"><span data-stu-id="09953-1152">Attributes</span></span>|<span data-ttu-id="09953-1153">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-1153">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="09953-1154">función</span><span class="sxs-lookup"><span data-stu-id="09953-1154">function</span></span>||<span data-ttu-id="09953-1155">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="09953-1155">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="09953-1156">Las propiedades personalizadas se proporcionan como un objeto [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="09953-1156">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="09953-1157">Este objeto se puede utilizar para obtener, establecer y elimina las propiedades personalizadas del elemento y guarde los cambios en la propiedad personalizada que se vuelve a establecer en el servidor.</span><span class="sxs-lookup"><span data-stu-id="09953-1157">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="09953-1158">Objeto</span><span class="sxs-lookup"><span data-stu-id="09953-1158">Object</span></span>|<span data-ttu-id="09953-1159">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-1160">Los desarrolladores pueden proporcionar cualquier objeto que desean tener acceso en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="09953-1160">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="09953-1161">Se puede tener acceso a este objeto por el `asyncResult.asyncContext` (propiedad) en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="09953-1161">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09953-1162">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-1162">Requirements</span></span>

|<span data-ttu-id="09953-1163">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-1163">Requirement</span></span>|<span data-ttu-id="09953-1164">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-1164">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-1165">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-1165">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-1166">1.0</span><span class="sxs-lookup"><span data-stu-id="09953-1166">1.0</span></span>|
|[<span data-ttu-id="09953-1167">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-1167">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-1168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-1168">ReadItem</span></span>|
|[<span data-ttu-id="09953-1169">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-1169">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-1170">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-1170">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-1171">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-1171">Example</span></span>

<span data-ttu-id="09953-p174">En el siguiente ejemplo de código se muestra cómo usar el método `loadCustomPropertiesAsync` para cargar de forma asincrónica propiedades personalizadas que son específicas del elemento actual. En el ejemplo también se muestra cómo usar el método `CustomProperties.saveAsync` para guardar estas propiedades en el servidor. Después de cargar las propiedades personalizadas, el ejemplo de código usa el método `CustomProperties.get` para leer la propiedad personalizada `myProp`, el método `CustomProperties.set` para escribir en la propiedad personalizada `otherProp` y finalmente se llama al método `saveAsync` para guardar las propiedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="09953-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="09953-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="09953-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="09953-1176">Quita los datos adjuntos de un mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="09953-1176">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="09953-p175">El método `removeAttachmentAsync` quita del elemento los datos adjuntos con el identificador especificado. Como práctica recomendada, debe usar el identificador de datos adjuntos para quitar datos adjuntos solo si la misma aplicación de correo ha agregado los datos adjuntos en la misma sesión. En Outlook Web App y OWA para los dispositivos, el identificador de datos adjuntos es válido solo en la misma sesión. Una sesión finaliza cuando el usuario cierra la aplicación, o si el usuario comienza a redactar en un formulario en línea y posteriormente se extrae el formulario en línea para continuar en una ventana independiente.</span><span class="sxs-lookup"><span data-stu-id="09953-p175">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09953-1181">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="09953-1181">Parameters:</span></span>

|<span data-ttu-id="09953-1182">Nombre</span><span class="sxs-lookup"><span data-stu-id="09953-1182">Name</span></span>|<span data-ttu-id="09953-1183">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-1183">Type</span></span>|<span data-ttu-id="09953-1184">Atributos</span><span class="sxs-lookup"><span data-stu-id="09953-1184">Attributes</span></span>|<span data-ttu-id="09953-1185">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-1185">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="09953-1186">String</span><span class="sxs-lookup"><span data-stu-id="09953-1186">String</span></span>||<span data-ttu-id="09953-p176">El identificador de los datos adjuntos que se deben quitar. La longitud máxima de la cadena es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="09953-p176">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="09953-1189">Object</span><span class="sxs-lookup"><span data-stu-id="09953-1189">Object</span></span>|<span data-ttu-id="09953-1190">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-1190">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-1191">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="09953-1191">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="09953-1192">Objeto</span><span class="sxs-lookup"><span data-stu-id="09953-1192">Object</span></span>|<span data-ttu-id="09953-1193">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-1193">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-1194">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="09953-1194">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="09953-1195">función</span><span class="sxs-lookup"><span data-stu-id="09953-1195">function</span></span>|<span data-ttu-id="09953-1196">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-1196">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-1197">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="09953-1197">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="09953-1198">Si se produce un error en la eliminación de los datos adjuntos, la propiedad `asyncResult.error` contendrá un código de error con el motivo del error.</span><span class="sxs-lookup"><span data-stu-id="09953-1198">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="09953-1199">Errores</span><span class="sxs-lookup"><span data-stu-id="09953-1199">Errors</span></span>

|<span data-ttu-id="09953-1200">Código de error</span><span class="sxs-lookup"><span data-stu-id="09953-1200">Error code</span></span>|<span data-ttu-id="09953-1201">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-1201">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="09953-1202">El identificador de datos adjuntos no existe.</span><span class="sxs-lookup"><span data-stu-id="09953-1202">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09953-1203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-1203">Requirements</span></span>

|<span data-ttu-id="09953-1204">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-1204">Requirement</span></span>|<span data-ttu-id="09953-1205">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-1205">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-1206">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-1206">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-1207">1.1</span><span class="sxs-lookup"><span data-stu-id="09953-1207">1.1</span></span>|
|[<span data-ttu-id="09953-1208">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-1208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-1209">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="09953-1209">ReadWriteItem</span></span>|
|[<span data-ttu-id="09953-1210">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-1210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-1211">Redacción</span><span class="sxs-lookup"><span data-stu-id="09953-1211">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-1212">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-1212">Example</span></span>

<span data-ttu-id="09953-1213">El código siguiente quita un archivo adjunto con un identificador de "0".</span><span class="sxs-lookup"><span data-stu-id="09953-1213">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="09953-1214">removeHandlerAsync (eventType, controlador, [opciones], [devolución de llamada])</span><span class="sxs-lookup"><span data-stu-id="09953-1214">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="09953-1215">Quita un controlador de eventos para un evento compatible.</span><span class="sxs-lookup"><span data-stu-id="09953-1215">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="09953-1216">Actualmente son los tipos de eventos compatibles `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, y`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="09953-1216">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="09953-1217">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="09953-1217">Parameters:</span></span>

| <span data-ttu-id="09953-1218">Nombre</span><span class="sxs-lookup"><span data-stu-id="09953-1218">Name</span></span> | <span data-ttu-id="09953-1219">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-1219">Type</span></span> | <span data-ttu-id="09953-1220">Atributos</span><span class="sxs-lookup"><span data-stu-id="09953-1220">Attributes</span></span> | <span data-ttu-id="09953-1221">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-1221">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="09953-1222">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="09953-1222">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="09953-1223">El evento que debe invocar el controlador.</span><span class="sxs-lookup"><span data-stu-id="09953-1223">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="09953-1224">Función</span><span class="sxs-lookup"><span data-stu-id="09953-1224">Function</span></span> || <span data-ttu-id="09953-p177">La función que va a controlar el evento. La función debe aceptar un único parámetro, que es un literal de objeto. La propiedad `type` del parámetro coincidirá con el parámetro `eventType` que se ha pasado a `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="09953-p177">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="09953-1228">Objeto</span><span class="sxs-lookup"><span data-stu-id="09953-1228">Object</span></span> | <span data-ttu-id="09953-1229">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="09953-1230">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="09953-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="09953-1231">Objeto</span><span class="sxs-lookup"><span data-stu-id="09953-1231">Object</span></span> | <span data-ttu-id="09953-1232">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="09953-1233">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="09953-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="09953-1234">función</span><span class="sxs-lookup"><span data-stu-id="09953-1234">function</span></span>| <span data-ttu-id="09953-1235">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-1236">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="09953-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09953-1237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-1237">Requirements</span></span>

|<span data-ttu-id="09953-1238">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-1238">Requirement</span></span>| <span data-ttu-id="09953-1239">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-1240">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-1240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09953-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="09953-1241">1.7</span></span> |
|[<span data-ttu-id="09953-1242">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-1242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09953-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09953-1243">ReadItem</span></span> |
|[<span data-ttu-id="09953-1244">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-1244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09953-1245">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="09953-1245">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="09953-1246">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-1246">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="09953-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="09953-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="09953-1248">Guarda un elemento de forma asincrónica.</span><span class="sxs-lookup"><span data-stu-id="09953-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="09953-p178">Cuando se invoca, este método guarda el mensaje actual como un borrador y devuelve el identificador de elemento a través del método de devolución de llamada. En el modo en línea de Outlook Web App o Outlook, el elemento se guarda en el servidor. En modo en caché de Outlook, se guarda el elemento en la caché local.</span><span class="sxs-lookup"><span data-stu-id="09953-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-1252">Si el complemento llama a `saveAsync` en un elemento en modo de redacción con el fin de obtener una `itemId` para usar con EWS o la API de REST, tenga en cuenta que cuando Outlook está en modo en caché, puede tardar algún tiempo antes de que el elemento realmente se sincroniza con el servidor.</span><span class="sxs-lookup"><span data-stu-id="09953-1252">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="09953-1253">Hasta que se sincroniza el elemento, el uso de la `itemId` devolverá un error.</span><span class="sxs-lookup"><span data-stu-id="09953-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="09953-p180">Dado que las citas no tienen ningún estado de borrador, si se llama a `saveAsync` en una cita en modo de redacción, el elemento se guardará como una cita normal en el calendario del usuario. Para las citas nuevas que todavía no se han guardado, no se enviará ninguna invitación. Si se guarda una cita existente, se enviará una actualización a los asistentes agregados o eliminados.</span><span class="sxs-lookup"><span data-stu-id="09953-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="09953-1257">Los siguientes clientes tienen un comportamiento diferente para `saveAsync` en citas en modo de redacción:</span><span class="sxs-lookup"><span data-stu-id="09953-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="09953-1258">Mac Outlook no admite `saveAsync` en una reunión en el modo de redacción.</span><span class="sxs-lookup"><span data-stu-id="09953-1258">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="09953-1259">Al llamar a `saveAsync` en una reunión en Mac Outlook devolverá un error.</span><span class="sxs-lookup"><span data-stu-id="09953-1259">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="09953-1260">Outlook en el web siempre envía una invitación o actualizar cuándo `saveAsync` se llama en una cita en modo de redacción.</span><span class="sxs-lookup"><span data-stu-id="09953-1260">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09953-1261">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="09953-1261">Parameters:</span></span>

|<span data-ttu-id="09953-1262">Nombre</span><span class="sxs-lookup"><span data-stu-id="09953-1262">Name</span></span>|<span data-ttu-id="09953-1263">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-1263">Type</span></span>|<span data-ttu-id="09953-1264">Atributos</span><span class="sxs-lookup"><span data-stu-id="09953-1264">Attributes</span></span>|<span data-ttu-id="09953-1265">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-1265">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="09953-1266">Objeto</span><span class="sxs-lookup"><span data-stu-id="09953-1266">Object</span></span>|<span data-ttu-id="09953-1267">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-1268">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="09953-1268">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="09953-1269">Objeto</span><span class="sxs-lookup"><span data-stu-id="09953-1269">Object</span></span>|<span data-ttu-id="09953-1270">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-1271">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="09953-1271">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="09953-1272">función</span><span class="sxs-lookup"><span data-stu-id="09953-1272">function</span></span>||<span data-ttu-id="09953-1273">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="09953-1273">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="09953-1274">En caso de éxito, se proporciona el identificador de elemento en el `asyncResult.value` (propiedad).</span><span class="sxs-lookup"><span data-stu-id="09953-1274">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09953-1275">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-1275">Requirements</span></span>

|<span data-ttu-id="09953-1276">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-1276">Requirement</span></span>|<span data-ttu-id="09953-1277">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-1277">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-1278">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-1278">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-1279">1.3</span><span class="sxs-lookup"><span data-stu-id="09953-1279">1.3</span></span>|
|[<span data-ttu-id="09953-1280">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-1280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-1281">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="09953-1281">ReadWriteItem</span></span>|
|[<span data-ttu-id="09953-1282">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-1282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-1283">Redacción</span><span class="sxs-lookup"><span data-stu-id="09953-1283">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="09953-1284">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="09953-1284">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="09953-p182">A continuación, se muestra un ejemplo del parámetro `result` que se ha pasado a la función de devolución de llamada. La propiedad `value` contiene el identificador de elemento del elemento.</span><span class="sxs-lookup"><span data-stu-id="09953-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="09953-1287">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="09953-1287">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="09953-1288">Inserta asincrónicamente datos en el cuerpo o el asunto de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="09953-1288">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="09953-p183">El método `setSelectedDataAsync` inserta la cadena especificada en la posición del cursor en el asunto o el cuerpo del elemento, o si se selecciona texto en el editor, reemplaza el texto seleccionado. Si el cursor no está en el cuerpo o el campo del asunto, se devuelve un error. Después de la inserción, el cursor se coloca al final del contenido insertado.</span><span class="sxs-lookup"><span data-stu-id="09953-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09953-1292">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="09953-1292">Parameters:</span></span>

|<span data-ttu-id="09953-1293">Nombre</span><span class="sxs-lookup"><span data-stu-id="09953-1293">Name</span></span>|<span data-ttu-id="09953-1294">Tipo</span><span class="sxs-lookup"><span data-stu-id="09953-1294">Type</span></span>|<span data-ttu-id="09953-1295">Atributos</span><span class="sxs-lookup"><span data-stu-id="09953-1295">Attributes</span></span>|<span data-ttu-id="09953-1296">Descripción</span><span class="sxs-lookup"><span data-stu-id="09953-1296">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="09953-1297">String</span><span class="sxs-lookup"><span data-stu-id="09953-1297">String</span></span>||<span data-ttu-id="09953-p184">Datos que se van a insertar. Los datos no deben superar 1.000.000 de caracteres. Si se pasan más de 1.000.000 de caracteres, se produce una excepción `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="09953-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="09953-1301">Object</span><span class="sxs-lookup"><span data-stu-id="09953-1301">Object</span></span>|<span data-ttu-id="09953-1302">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-1303">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="09953-1303">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="09953-1304">Objeto</span><span class="sxs-lookup"><span data-stu-id="09953-1304">Object</span></span>|<span data-ttu-id="09953-1305">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-1305">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-1306">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="09953-1306">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="09953-1307">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="09953-1307">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="09953-1308">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="09953-1308">&lt;optional&gt;</span></span>|<span data-ttu-id="09953-p185">Si `text`, se aplica el estilo actual en Outlook Web App y Outlook. Si el campo es un editor de HTML, se insertan solo los datos de texto, aunque los datos sean HTML.</span><span class="sxs-lookup"><span data-stu-id="09953-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="09953-p186">Si `html` y el campo admiten HTML (el asunto no), el estilo actual se aplica en Outlook Web App y el estilo predeterminado en Outlook. Si el campo es un campo de texto, se devuelve un error `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="09953-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="09953-1313">Si `coercionType` no está establecido, el resultado depende del campo: si el campo es HTML, se usa HTML; si el campo es texto, se usa texto sin formato.</span><span class="sxs-lookup"><span data-stu-id="09953-1313">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="09953-1314">function</span><span class="sxs-lookup"><span data-stu-id="09953-1314">function</span></span>||<span data-ttu-id="09953-1315">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="09953-1315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09953-1316">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09953-1316">Requirements</span></span>

|<span data-ttu-id="09953-1317">Requirement</span><span class="sxs-lookup"><span data-stu-id="09953-1317">Requirement</span></span>|<span data-ttu-id="09953-1318">Valor</span><span class="sxs-lookup"><span data-stu-id="09953-1318">Value</span></span>|
|---|---|
|[<span data-ttu-id="09953-1319">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="09953-1319">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="09953-1320">1.2</span><span class="sxs-lookup"><span data-stu-id="09953-1320">1.2</span></span>|
|[<span data-ttu-id="09953-1321">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="09953-1321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="09953-1322">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="09953-1322">ReadWriteItem</span></span>|
|[<span data-ttu-id="09953-1323">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="09953-1323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="09953-1324">Redacción</span><span class="sxs-lookup"><span data-stu-id="09953-1324">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="09953-1325">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="09953-1325">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```