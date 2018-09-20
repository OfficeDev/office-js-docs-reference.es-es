
# <a name="item"></a><span data-ttu-id="ff4ce-101">elemento</span><span class="sxs-lookup"><span data-stu-id="ff4ce-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="ff4ce-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="ff4ce-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="ff4ce-p101">El espacio de nombres `item` se usa para acceder al mensaje seleccionado actualmente, a una convocatoria de reunión o a una cita. Puede determinar el tipo del `item` mediante la propiedad [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-105">Requirements</span></span>

|<span data-ttu-id="ff4ce-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-106">Requirement</span></span>|<span data-ttu-id="ff4ce-107">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-108">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-109">1.0</span></span>|
|[<span data-ttu-id="ff4ce-110">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-111">Restringido</span><span class="sxs-lookup"><span data-stu-id="ff4ce-111">Restricted</span></span>|
|[<span data-ttu-id="ff4ce-112">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-113">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ff4ce-114">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-114">Members and methods</span></span>

| <span data-ttu-id="ff4ce-115">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-115">Member</span></span> | <span data-ttu-id="ff4ce-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ff4ce-117">attachments</span><span class="sxs-lookup"><span data-stu-id="ff4ce-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="ff4ce-118">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-118">Member</span></span> |
| [<span data-ttu-id="ff4ce-119">bcc</span><span class="sxs-lookup"><span data-stu-id="ff4ce-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="ff4ce-120">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-120">Member</span></span> |
| [<span data-ttu-id="ff4ce-121">body</span><span class="sxs-lookup"><span data-stu-id="ff4ce-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="ff4ce-122">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-122">Member</span></span> |
| [<span data-ttu-id="ff4ce-123">cc</span><span class="sxs-lookup"><span data-stu-id="ff4ce-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="ff4ce-124">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-124">Member</span></span> |
| [<span data-ttu-id="ff4ce-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="ff4ce-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="ff4ce-126">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-126">Member</span></span> |
| [<span data-ttu-id="ff4ce-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="ff4ce-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="ff4ce-128">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-128">Member</span></span> |
| [<span data-ttu-id="ff4ce-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="ff4ce-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="ff4ce-130">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-130">Member</span></span> |
| [<span data-ttu-id="ff4ce-131">end</span><span class="sxs-lookup"><span data-stu-id="ff4ce-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="ff4ce-132">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-132">Member</span></span> |
| [<span data-ttu-id="ff4ce-133">from</span><span class="sxs-lookup"><span data-stu-id="ff4ce-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="ff4ce-134">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-134">Member</span></span> |
| [<span data-ttu-id="ff4ce-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="ff4ce-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="ff4ce-136">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-136">Member</span></span> |
| [<span data-ttu-id="ff4ce-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="ff4ce-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="ff4ce-138">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-138">Member</span></span> |
| [<span data-ttu-id="ff4ce-139">itemId</span><span class="sxs-lookup"><span data-stu-id="ff4ce-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="ff4ce-140">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-140">Member</span></span> |
| [<span data-ttu-id="ff4ce-141">itemType</span><span class="sxs-lookup"><span data-stu-id="ff4ce-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="ff4ce-142">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-142">Member</span></span> |
| [<span data-ttu-id="ff4ce-143">location</span><span class="sxs-lookup"><span data-stu-id="ff4ce-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="ff4ce-144">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-144">Member</span></span> |
| [<span data-ttu-id="ff4ce-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="ff4ce-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="ff4ce-146">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-146">Member</span></span> |
| [<span data-ttu-id="ff4ce-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="ff4ce-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="ff4ce-148">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-148">Member</span></span> |
| [<span data-ttu-id="ff4ce-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="ff4ce-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="ff4ce-150">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-150">Member</span></span> |
| [<span data-ttu-id="ff4ce-151">organizer</span><span class="sxs-lookup"><span data-stu-id="ff4ce-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="ff4ce-152">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-152">Member</span></span> |
| [<span data-ttu-id="ff4ce-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="ff4ce-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="ff4ce-154">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-154">Member</span></span> |
| [<span data-ttu-id="ff4ce-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="ff4ce-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="ff4ce-156">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-156">Member</span></span> |
| [<span data-ttu-id="ff4ce-157">sender</span><span class="sxs-lookup"><span data-stu-id="ff4ce-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="ff4ce-158">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-158">Member</span></span> |
| [<span data-ttu-id="ff4ce-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="ff4ce-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="ff4ce-160">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-160">Member</span></span> |
| [<span data-ttu-id="ff4ce-161">start</span><span class="sxs-lookup"><span data-stu-id="ff4ce-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="ff4ce-162">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-162">Member</span></span> |
| [<span data-ttu-id="ff4ce-163">subject</span><span class="sxs-lookup"><span data-stu-id="ff4ce-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="ff4ce-164">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-164">Member</span></span> |
| [<span data-ttu-id="ff4ce-165">to</span><span class="sxs-lookup"><span data-stu-id="ff4ce-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="ff4ce-166">Miembro	</span><span class="sxs-lookup"><span data-stu-id="ff4ce-166">Member</span></span> |
| [<span data-ttu-id="ff4ce-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ff4ce-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="ff4ce-168">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-168">Method</span></span> |
| [<span data-ttu-id="ff4ce-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="ff4ce-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="ff4ce-170">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-170">Method</span></span> |
| [<span data-ttu-id="ff4ce-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="ff4ce-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="ff4ce-172">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-172">Method</span></span> |
| [<span data-ttu-id="ff4ce-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ff4ce-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="ff4ce-174">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-174">Method</span></span> |
| [<span data-ttu-id="ff4ce-175">close</span><span class="sxs-lookup"><span data-stu-id="ff4ce-175">close</span></span>](#close) | <span data-ttu-id="ff4ce-176">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-176">Method</span></span> |
| [<span data-ttu-id="ff4ce-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="ff4ce-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="ff4ce-178">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-178">Method</span></span> |
| [<span data-ttu-id="ff4ce-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="ff4ce-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="ff4ce-180">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-180">Method</span></span> |
| [<span data-ttu-id="ff4ce-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="ff4ce-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="ff4ce-182">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-182">Method</span></span> |
| [<span data-ttu-id="ff4ce-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="ff4ce-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="ff4ce-184">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-184">Method</span></span> |
| [<span data-ttu-id="ff4ce-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="ff4ce-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="ff4ce-186">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-186">Method</span></span> |
| [<span data-ttu-id="ff4ce-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="ff4ce-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="ff4ce-188">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-188">Method</span></span> |
| [<span data-ttu-id="ff4ce-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="ff4ce-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="ff4ce-190">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-190">Method</span></span> |
| [<span data-ttu-id="ff4ce-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="ff4ce-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="ff4ce-192">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-192">Method</span></span> |
| [<span data-ttu-id="ff4ce-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ff4ce-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="ff4ce-194">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-194">Method</span></span> |
| [<span data-ttu-id="ff4ce-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="ff4ce-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="ff4ce-196">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-196">Method</span></span> |
| [<span data-ttu-id="ff4ce-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="ff4ce-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="ff4ce-198">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-198">Method</span></span> |
| [<span data-ttu-id="ff4ce-199">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="ff4ce-199">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="ff4ce-200">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-200">Method</span></span> |
| [<span data-ttu-id="ff4ce-201">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ff4ce-201">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="ff4ce-202">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-202">Method</span></span> |
| [<span data-ttu-id="ff4ce-203">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="ff4ce-203">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="ff4ce-204">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-204">Method</span></span> |
| [<span data-ttu-id="ff4ce-205">saveAsync</span><span class="sxs-lookup"><span data-stu-id="ff4ce-205">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="ff4ce-206">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-206">Method</span></span> |
| [<span data-ttu-id="ff4ce-207">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ff4ce-207">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="ff4ce-208">Método</span><span class="sxs-lookup"><span data-stu-id="ff4ce-208">Method</span></span> |

### <a name="example"></a><span data-ttu-id="ff4ce-209">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-209">Example</span></span>

<span data-ttu-id="ff4ce-210">En el siguiente ejemplo de código de JavaScript, se muestra cómo tener acceso a la propiedad `subject` del elemento actual en Outlook.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-210">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="ff4ce-211">Miembros</span><span class="sxs-lookup"><span data-stu-id="ff4ce-211">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="ff4ce-212">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ff4ce-212">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="ff4ce-p102">Obtiene una matriz de datos adjuntos para el elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-215">Ciertos tipos de archivos están bloqueados por Outlook debido a problemas de seguridad y, por tanto, no se devuelven.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-215">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="ff4ce-216">Para obtener más información, vea [datos adjuntos bloqueados en Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-216">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-217">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-217">Type:</span></span>

*   <span data-ttu-id="ff4ce-218">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ff4ce-218">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-219">Requirements</span></span>

|<span data-ttu-id="ff4ce-220">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-220">Requirement</span></span>|<span data-ttu-id="ff4ce-221">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-222">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-222">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-223">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-223">1.0</span></span>|
|[<span data-ttu-id="ff4ce-224">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-225">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-226">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-227">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-227">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-228">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-228">Example</span></span>

<span data-ttu-id="ff4ce-229">El siguiente código crea una cadena HTML con los detalles de todos los datos adjuntos en el elemento actual.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-229">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="ff4ce-230">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-230">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="ff4ce-231">Obtiene un objeto que proporciona métodos para obtener o actualizar a los destinatarios de la línea de CCO (con copia oculta) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-231">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="ff4ce-232">Solo modo Redacción.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-232">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-233">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-233">Type:</span></span>

*   [<span data-ttu-id="ff4ce-234">Destinatarios</span><span class="sxs-lookup"><span data-stu-id="ff4ce-234">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="ff4ce-235">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-235">Requirements</span></span>

|<span data-ttu-id="ff4ce-236">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-236">Requirement</span></span>|<span data-ttu-id="ff4ce-237">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-237">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-238">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-238">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-239">1.1</span><span class="sxs-lookup"><span data-stu-id="ff4ce-239">1.1</span></span>|
|[<span data-ttu-id="ff4ce-240">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-240">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-241">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-241">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-242">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-242">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-243">Redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-243">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-244">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-244">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="ff4ce-245">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-245">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="ff4ce-246">Obtiene un objeto que proporciona métodos para manipular el cuerpo de un elemento.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-246">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-247">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-247">Type:</span></span>

*   [<span data-ttu-id="ff4ce-248">Cuerpo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-248">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="ff4ce-249">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-249">Requirements</span></span>

|<span data-ttu-id="ff4ce-250">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-250">Requirement</span></span>|<span data-ttu-id="ff4ce-251">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-251">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-252">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-252">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-253">1.1</span><span class="sxs-lookup"><span data-stu-id="ff4ce-253">1.1</span></span>|
|[<span data-ttu-id="ff4ce-254">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-254">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-255">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-255">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-256">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-256">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-257">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-257">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="ff4ce-258">cc: matriz. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="ff4ce-259">Proporciona acceso a los destinatarios de Cc (con copia) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="ff4ce-260">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ff4ce-261">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-261">Read mode</span></span>

<span data-ttu-id="ff4ce-p106">La propiedad `cc` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **CC** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ff4ce-264">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-264">Compose mode</span></span>

<span data-ttu-id="ff4ce-265">El `cc` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **Cc** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-266">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-266">Type:</span></span>

*   <span data-ttu-id="ff4ce-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-268">Requirements</span></span>

|<span data-ttu-id="ff4ce-269">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-269">Requirement</span></span>|<span data-ttu-id="ff4ce-270">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-271">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-271">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-272">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-272">1.0</span></span>|
|[<span data-ttu-id="ff4ce-273">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-274">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-275">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-276">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-277">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-277">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="ff4ce-278">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-278">(nullable) conversationId :String</span></span>

<span data-ttu-id="ff4ce-279">Obtiene un identificador de la conversación de correo electrónico que contiene un mensaje determinado.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-279">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="ff4ce-p107">Puede obtener un entero para esta propiedad si la aplicación de correo se activa en respuestas o formularios de lectura o en formularios de redacción. Si, después, el usuario cambia el asunto del mensaje de respuesta, al enviar dicha respuesta, el Id. de conversación de ese mensaje cambiará y el valor que se obtuvo anteriormente ya no será de aplicación.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="ff4ce-p108">Obtiene un valor NULL para esta propiedad para un nuevo elemento de un formulario de redacción. Si el usuario establece un asunto y guarda el elemento, la propiedad `conversationId` devolverá un valor.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-284">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-284">Type:</span></span>

*   <span data-ttu-id="ff4ce-285">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-285">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-286">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-286">Requirements</span></span>

|<span data-ttu-id="ff4ce-287">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-287">Requirement</span></span>|<span data-ttu-id="ff4ce-288">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-289">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-289">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-290">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-290">1.0</span></span>|
|[<span data-ttu-id="ff4ce-291">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-291">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-292">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-292">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-293">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-293">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-294">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-294">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="ff4ce-295">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="ff4ce-295">dateTimeCreated :Date</span></span>

<span data-ttu-id="ff4ce-p109">Obtiene la fecha y hora de creación de un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-298">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-298">Type:</span></span>

*   <span data-ttu-id="ff4ce-299">Fecha</span><span class="sxs-lookup"><span data-stu-id="ff4ce-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-300">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-300">Requirements</span></span>

|<span data-ttu-id="ff4ce-301">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-301">Requirement</span></span>|<span data-ttu-id="ff4ce-302">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-303">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-304">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-304">1.0</span></span>|
|[<span data-ttu-id="ff4ce-305">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-306">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-307">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-308">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-309">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-309">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="ff4ce-310">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="ff4ce-310">dateTimeModified :Date</span></span>

<span data-ttu-id="ff4ce-p110">Obtiene la fecha y hora en que se modificó por última vez un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-313">Este miembro no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-313">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-314">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-314">Type:</span></span>

*   <span data-ttu-id="ff4ce-315">Fecha</span><span class="sxs-lookup"><span data-stu-id="ff4ce-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-316">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-316">Requirements</span></span>

|<span data-ttu-id="ff4ce-317">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-317">Requirement</span></span>|<span data-ttu-id="ff4ce-318">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-319">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-319">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-320">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-320">1.0</span></span>|
|[<span data-ttu-id="ff4ce-321">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-322">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-323">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-324">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-325">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-325">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="ff4ce-326">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-326">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="ff4ce-327">Obtiene o establece la fecha y la hora de finalización de la cita.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="ff4ce-p111">La propiedad `end` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para convertir el valor de la propiedad final a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ff4ce-330">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-330">Read mode</span></span>

<span data-ttu-id="ff4ce-331">La propiedad `end` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-331">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ff4ce-332">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-332">Compose mode</span></span>

<span data-ttu-id="ff4ce-333">La propiedad `end` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="ff4ce-334">Si usa el método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para establecer la hora de finalización, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-334">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-335">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-335">Type:</span></span>

*   <span data-ttu-id="ff4ce-336">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-336">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-337">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-337">Requirements</span></span>

|<span data-ttu-id="ff4ce-338">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-338">Requirement</span></span>|<span data-ttu-id="ff4ce-339">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-340">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-341">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-341">1.0</span></span>|
|[<span data-ttu-id="ff4ce-342">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-343">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-344">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-345">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-346">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-346">Example</span></span>

<span data-ttu-id="ff4ce-347">En el ejemplo siguiente, se establece la hora de finalización de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-347">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="ff4ce-348">de:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[desde](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-348">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="ff4ce-349">Obtiene la dirección de correo electrónico del remitente de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-349">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="ff4ce-p112">Las propiedades `from` y [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-352">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `from` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-352">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ff4ce-353">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-353">Read mode</span></span>

<span data-ttu-id="ff4ce-354">El `from` (propiedad) devuelve un `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-354">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="ff4ce-355">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-355">Compose mode</span></span>

<span data-ttu-id="ff4ce-356">El `from` (propiedad) devuelve un `From` objeto que proporciona un método para obtener el valor de.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-356">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ff4ce-357">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-357">Type:</span></span>

*   <span data-ttu-id="ff4ce-358">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [desde](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-358">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-359">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-359">Requirements</span></span>

|<span data-ttu-id="ff4ce-360">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-360">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="ff4ce-361">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-362">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-362">1.0</span></span>|<span data-ttu-id="ff4ce-363">Vista previa</span><span class="sxs-lookup"><span data-stu-id="ff4ce-363">Preview</span></span>|
|[<span data-ttu-id="ff4ce-364">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-365">ReadItem</span></span>|<span data-ttu-id="ff4ce-366">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-366">ReadWriteItem</span></span>|
|[<span data-ttu-id="ff4ce-367">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-367">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-368">Read</span><span class="sxs-lookup"><span data-stu-id="ff4ce-368">Read</span></span>|<span data-ttu-id="ff4ce-369">Redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-369">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="ff4ce-370">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-370">internetMessageId :String</span></span>

<span data-ttu-id="ff4ce-p113">Obtiene el identificador de mensaje de Internet de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-373">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-373">Type:</span></span>

*   <span data-ttu-id="ff4ce-374">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-375">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-375">Requirements</span></span>

|<span data-ttu-id="ff4ce-376">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-376">Requirement</span></span>|<span data-ttu-id="ff4ce-377">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-378">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-378">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-379">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-379">1.0</span></span>|
|[<span data-ttu-id="ff4ce-380">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-380">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-381">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-382">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-382">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-383">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-384">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-384">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="ff4ce-385">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-385">itemClass :String</span></span>

<span data-ttu-id="ff4ce-p114">Obtiene la clase de elemento Servicios Web Exchange del elemento seleccionado. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="ff4ce-p115">La propiedad `itemClass` especifica la clase de mensaje del elemento seleccionado. Las siguientes son las clases de mensaje predeterminadas para el elemento de mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="ff4ce-390">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-390">Type</span></span>|<span data-ttu-id="ff4ce-391">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-391">Description</span></span>|<span data-ttu-id="ff4ce-392">Clase de elemento</span><span class="sxs-lookup"><span data-stu-id="ff4ce-392">item class</span></span>|
|---|---|---|
|<span data-ttu-id="ff4ce-393">Elementos de cita</span><span class="sxs-lookup"><span data-stu-id="ff4ce-393">Appointment items</span></span>|<span data-ttu-id="ff4ce-394">Estos son los elementos de calendario de la clase de elemento `IPM.Appointment` o `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-394">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="ff4ce-395">Elementos de mensaje</span><span class="sxs-lookup"><span data-stu-id="ff4ce-395">Message items</span></span>|<span data-ttu-id="ff4ce-396">Entre estos se encuentran los mensajes de correo electrónico que tienen la clase de mensaje predeterminada `IPM.Note` y las convocatorias de reunión, respuestas y cancelaciones que usan `IPM.Schedule.Meeting` como clase de mensaje base.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-396">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="ff4ce-397">Puede crear clases de mensaje personalizadas para ampliar una clase de mensaje predeterminada. Por ejemplo, puede crear una clase de mensaje de cita personalizada `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-397">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-398">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-398">Type:</span></span>

*   <span data-ttu-id="ff4ce-399">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-400">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-400">Requirements</span></span>

|<span data-ttu-id="ff4ce-401">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-401">Requirement</span></span>|<span data-ttu-id="ff4ce-402">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-403">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-403">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-404">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-404">1.0</span></span>|
|[<span data-ttu-id="ff4ce-405">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-406">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-407">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-408">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-409">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-409">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="ff4ce-410">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-410">(nullable) itemId :String</span></span>

<span data-ttu-id="ff4ce-p116">Obtiene el identificador de elemento de Servicios Web Exchange para el elemento actual. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-413">El identificador que devuelve la propiedad `itemId` es el mismo que el identificador de elemento de Servicios Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-413">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="ff4ce-414">El `itemId` (propiedad) no es idéntico al identificador de entrada de Outlook o el identificador de la API de REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-414">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="ff4ce-415">Antes de realizar llamadas a la API de REST con este valor, se debe convertir mediante [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-415">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="ff4ce-416">Para obtener más información, vea [usar las API de REST de Outlook desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-416">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="ff4ce-p118">La propiedad `itemId` no está disponible en el modo redacción. Si se requiere un identificador de elemento, el método [`saveAsync`](#saveasyncoptions-callback) puede usarse para guardar el elemento en el servidor, que devolverá el identificador de elemento en el parámetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) de la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-419">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-419">Type:</span></span>

*   <span data-ttu-id="ff4ce-420">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-421">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-421">Requirements</span></span>

|<span data-ttu-id="ff4ce-422">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-422">Requirement</span></span>|<span data-ttu-id="ff4ce-423">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-424">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-425">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-425">1.0</span></span>|
|[<span data-ttu-id="ff4ce-426">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-427">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-428">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-429">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-430">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-430">Example</span></span>

<span data-ttu-id="ff4ce-p119">El siguiente código comprueba la presencia de un identificador de elemento. Si la propiedad `itemId` devuelve `null` o `undefined`, guarda el elemento en el servidor y obtiene el identificador de elemento desde el resultado asincrónico.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="ff4ce-433">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-433">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="ff4ce-434">Obtiene el tipo de elemento que representa una instancia.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-434">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="ff4ce-435">La propiedad `itemType` devuelve uno de los valores de enumeración de `ItemType`, lo que indica si la instancia del objeto `item` es un mensaje o una cita.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-435">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-436">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-436">Type:</span></span>

*   [<span data-ttu-id="ff4ce-437">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="ff4ce-437">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="ff4ce-438">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-438">Requirements</span></span>

|<span data-ttu-id="ff4ce-439">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-439">Requirement</span></span>|<span data-ttu-id="ff4ce-440">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-440">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-441">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-441">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-442">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-442">1.0</span></span>|
|[<span data-ttu-id="ff4ce-443">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-443">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-444">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-445">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-445">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-446">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-446">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-447">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-447">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="ff4ce-448">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-448">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="ff4ce-449">Obtiene o establece la ubicación de una cita.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-449">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ff4ce-450">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-450">Read mode</span></span>

<span data-ttu-id="ff4ce-451">La propiedad `location` devuelve una cadena que contiene la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-451">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ff4ce-452">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-452">Compose mode</span></span>

<span data-ttu-id="ff4ce-453">La propiedad `location` devuelve un objeto `Location` que proporciona métodos para obtener y establecer la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-453">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-454">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-454">Type:</span></span>

*   <span data-ttu-id="ff4ce-455">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-455">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-456">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-456">Requirements</span></span>

|<span data-ttu-id="ff4ce-457">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-457">Requirement</span></span>|<span data-ttu-id="ff4ce-458">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-459">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-459">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-460">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-460">1.0</span></span>|
|[<span data-ttu-id="ff4ce-461">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-462">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-463">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-464">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-464">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-465">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-465">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="ff4ce-466">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-466">normalizedSubject :String</span></span>

<span data-ttu-id="ff4ce-p120">Obtiene el asunto de un elemento, con todos los prefijos quitados (incluidos `RE:` y `FWD:`). Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="ff4ce-p121">La propiedad normalizedSubject obtiene el asunto del elemento, con los prefijos estándar (como `RE:` y `FW:`) agregados por los programas de correo electrónico. Para obtener el asunto del elemento con los prefijos intactos, use la propiedad [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-471">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-471">Type:</span></span>

*   <span data-ttu-id="ff4ce-472">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-472">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-473">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-473">Requirements</span></span>

|<span data-ttu-id="ff4ce-474">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-474">Requirement</span></span>|<span data-ttu-id="ff4ce-475">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-476">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-476">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-477">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-477">1.0</span></span>|
|[<span data-ttu-id="ff4ce-478">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-478">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-479">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-480">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-480">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-481">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-481">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-482">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-482">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="ff4ce-483">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-483">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="ff4ce-484">Obtiene los mensajes de notificación de un elemento.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-484">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-485">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-485">Type:</span></span>

*   [<span data-ttu-id="ff4ce-486">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="ff4ce-486">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="ff4ce-487">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-487">Requirements</span></span>

|<span data-ttu-id="ff4ce-488">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-488">Requirement</span></span>|<span data-ttu-id="ff4ce-489">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-489">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-490">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-490">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-491">1.3</span><span class="sxs-lookup"><span data-stu-id="ff4ce-491">1.3</span></span>|
|[<span data-ttu-id="ff4ce-492">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-492">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-493">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-493">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-494">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-494">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-495">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-495">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="ff4ce-496">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-496">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="ff4ce-497">Proporciona acceso a los asistentes opcionales de un evento.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-497">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="ff4ce-498">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-498">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ff4ce-499">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-499">Read mode</span></span>

<span data-ttu-id="ff4ce-500">La propiedad `optionalAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente opcional a la reunión.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-500">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ff4ce-501">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-501">Compose mode</span></span>

<span data-ttu-id="ff4ce-502">El `optionalAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes opcionales de una reunión.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-502">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-503">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-503">Type:</span></span>

*   <span data-ttu-id="ff4ce-504">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-504">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-505">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-505">Requirements</span></span>

|<span data-ttu-id="ff4ce-506">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-506">Requirement</span></span>|<span data-ttu-id="ff4ce-507">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-508">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-508">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-509">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-509">1.0</span></span>|
|[<span data-ttu-id="ff4ce-510">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-511">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-512">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-513">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-514">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-514">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="ff4ce-515">Organizador:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizador](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-515">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="ff4ce-516">Obtiene la dirección de correo electrónico del organizador de una reunión especificada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-516">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ff4ce-517">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-517">Read mode</span></span>

<span data-ttu-id="ff4ce-518">El `organizer` (propiedad) devuelve un objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa el organizador de la reunión.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-518">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ff4ce-519">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-519">Compose mode</span></span>

<span data-ttu-id="ff4ce-520">El `organizer` (propiedad) devuelve un objeto de [Organizador](/javascript/api/outlook/office.organizer) que proporciona un método para obtener el valor del organizador.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-520">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-521">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-521">Type:</span></span>

*   <span data-ttu-id="ff4ce-522">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizador](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-522">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-523">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-523">Requirements</span></span>

|<span data-ttu-id="ff4ce-524">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-524">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="ff4ce-525">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-525">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-526">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-526">1.0</span></span>|<span data-ttu-id="ff4ce-527">Vista previa</span><span class="sxs-lookup"><span data-stu-id="ff4ce-527">Preview</span></span>|
|[<span data-ttu-id="ff4ce-528">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-528">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-529">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-529">ReadItem</span></span>|<span data-ttu-id="ff4ce-530">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-530">ReadWriteItem</span></span>|
|[<span data-ttu-id="ff4ce-531">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-531">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-532">Read</span><span class="sxs-lookup"><span data-stu-id="ff4ce-532">Read</span></span>|<span data-ttu-id="ff4ce-533">Redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-533">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-534">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-534">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="ff4ce-535">periodicidad (nullable):[Periodicidad](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-535">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="ff4ce-536">Obtiene o establece el patrón de periodicidad de una cita.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-536">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="ff4ce-537">Obtiene el patrón de periodicidad de una convocatoria de reunión.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-537">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="ff4ce-538">Lectura y redacción modos para elementos de cita.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-538">Read and compose modes for appointment items.</span></span> <span data-ttu-id="ff4ce-539">Modo de lectura para los elementos de solicitud de reunión.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-539">Read mode for meeting request items.</span></span>

<span data-ttu-id="ff4ce-540">El `recurrence` (propiedad) devuelve un objeto de [Periodicidad](/javascript/api/outlook/office.recurrence) para solicitudes de reuniones o citas periódicas si un elemento es una serie o una instancia de una serie.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-540">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="ff4ce-541">`null`se devuelve solo citas y convocatorias de reunión de citas únicas.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-541">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="ff4ce-542">`undefined`se devuelve para los mensajes que no son las convocatorias de reunión.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-542">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="ff4ce-543">Nota: Las convocatorias de reunión tienen un `itemClass` valor de IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-543">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="ff4ce-544">Nota: Si el objeto de periodicidad es `null`, esto indica que el objeto es una cita única o una convocatoria de reunión de una cita única y no forma parte de una serie.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-544">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-545">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-545">Type:</span></span>

* [<span data-ttu-id="ff4ce-546">Periodicidad</span><span class="sxs-lookup"><span data-stu-id="ff4ce-546">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="ff4ce-547">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-547">Requirement</span></span>|<span data-ttu-id="ff4ce-548">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-549">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-549">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-550">Vista previa</span><span class="sxs-lookup"><span data-stu-id="ff4ce-550">Preview</span></span>|
|[<span data-ttu-id="ff4ce-551">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-551">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-552">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-552">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-553">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-554">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-554">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="ff4ce-555">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-555">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="ff4ce-556">Proporciona acceso a los asistentes necesarios de un evento.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-556">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="ff4ce-557">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-557">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ff4ce-558">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-558">Read mode</span></span>

<span data-ttu-id="ff4ce-559">La propiedad `requiredAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente requerido a la reunión.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-559">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ff4ce-560">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-560">Compose mode</span></span>

<span data-ttu-id="ff4ce-561">El `requiredAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes necesarios de una reunión.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-561">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-562">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-562">Type:</span></span>

*   <span data-ttu-id="ff4ce-563">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-563">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-564">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-564">Requirements</span></span>

|<span data-ttu-id="ff4ce-565">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-565">Requirement</span></span>|<span data-ttu-id="ff4ce-566">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-567">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-567">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-568">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-568">1.0</span></span>|
|[<span data-ttu-id="ff4ce-569">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-569">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-570">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-571">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-571">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-572">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-572">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-573">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-573">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="ff4ce-574">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-574">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="ff4ce-p126">Obtiene la dirección de correo electrónico del remitente de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="ff4ce-p127">Las propiedades [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) y `sender` representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-579">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `sender` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-579">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-580">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-580">Type:</span></span>

*   [<span data-ttu-id="ff4ce-581">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ff4ce-581">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ff4ce-582">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-582">Requirements</span></span>

|<span data-ttu-id="ff4ce-583">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-583">Requirement</span></span>|<span data-ttu-id="ff4ce-584">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-585">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-585">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-586">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-586">1.0</span></span>|
|[<span data-ttu-id="ff4ce-587">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-588">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-589">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-590">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-590">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-591">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-591">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="ff4ce-592">seriesId (nullable): cadena</span><span class="sxs-lookup"><span data-stu-id="ff4ce-592">(nullable) seriesId :String</span></span>

<span data-ttu-id="ff4ce-593">Obtiene el identificador de la serie a la que pertenece una instancia.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-593">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="ff4ce-594">En OWA y Outlook, el `seriesId` devuelve el identificador de Exchange Web Services (EWS) del elemento primario (serie) al que pertenece este producto.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-594">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="ff4ce-595">Sin embargo, en iOS y Android, el `seriesId` devuelve el identificador del resto del elemento primario.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-595">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-596">El identificador que devuelve la propiedad `seriesId` es el mismo que el identificador de elemento de Servicios Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-596">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="ff4ce-597">El `seriesId` propiedad no es idéntica a los identificadores de Outlook utilizadas por la API de REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-597">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="ff4ce-598">Antes de realizar llamadas a la API de REST con este valor, se debe convertir mediante [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-598">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="ff4ce-599">Para obtener más información, vea [usar las API de REST de Outlook desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-599">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="ff4ce-600">El `seriesId` propiedad devuelve `null` para los elementos que no tienen elementos primarios, como único citas, elementos de serie, o solicitudes de reunión y devuelve `undefined` para todos los elementos que no son las convocatorias de reunión.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-600">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-601">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-601">Type:</span></span>

* <span data-ttu-id="ff4ce-602">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-602">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-603">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-603">Requirements</span></span>

|<span data-ttu-id="ff4ce-604">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-604">Requirement</span></span>|<span data-ttu-id="ff4ce-605">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-606">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-606">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-607">Vista previa</span><span class="sxs-lookup"><span data-stu-id="ff4ce-607">Preview</span></span>|
|[<span data-ttu-id="ff4ce-608">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-609">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-610">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-611">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-611">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-612">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-612">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId; 
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="ff4ce-613">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-613">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="ff4ce-614">Obtiene o establece la fecha y la hora de inicio de la cita.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-614">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="ff4ce-p130">La propiedad `start` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para convertir el valor a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ff4ce-617">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-617">Read mode</span></span>

<span data-ttu-id="ff4ce-618">La propiedad `start` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-618">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ff4ce-619">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-619">Compose mode</span></span>

<span data-ttu-id="ff4ce-620">La propiedad `start` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-620">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="ff4ce-621">Si usa el método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para establecer la hora de inicio, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-621">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-622">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-622">Type:</span></span>

*   <span data-ttu-id="ff4ce-623">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-623">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-624">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-624">Requirements</span></span>

|<span data-ttu-id="ff4ce-625">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-625">Requirement</span></span>|<span data-ttu-id="ff4ce-626">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-627">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-627">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-628">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-628">1.0</span></span>|
|[<span data-ttu-id="ff4ce-629">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-629">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-630">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-631">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-631">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-632">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-632">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-633">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-633">Example</span></span>

<span data-ttu-id="ff4ce-634">En el ejemplo siguiente, se establece la hora de inicio de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-634">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="ff4ce-635">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-635">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="ff4ce-636">Obtiene o establece la descripción que se muestra en el campo de asunto de un elemento.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-636">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="ff4ce-637">La propiedad `subject` obtiene o establece el asunto completo del elemento, como lo ha enviado el servidor de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-637">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ff4ce-638">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-638">Read mode</span></span>

<span data-ttu-id="ff4ce-p131">La propiedad `subject` devuelve una cadena. Use la propiedad [`normalizedSubject`](#normalizedsubject-string) para obtener el asunto menos cualquier prefijo inicial como `RE:` y `FW:`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="ff4ce-641">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-641">Compose mode</span></span>

<span data-ttu-id="ff4ce-642">La propiedad `subject` devuelve un objeto `Subject` que proporciona métodos para obtener y establecer el asunto.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-642">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ff4ce-643">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-643">Type:</span></span>

*   <span data-ttu-id="ff4ce-644">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-644">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-645">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-645">Requirements</span></span>

|<span data-ttu-id="ff4ce-646">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-646">Requirement</span></span>|<span data-ttu-id="ff4ce-647">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-648">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-648">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-649">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-649">1.0</span></span>|
|[<span data-ttu-id="ff4ce-650">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-650">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-651">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-651">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-652">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-652">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-653">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-653">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="ff4ce-654">para: matriz. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-654">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="ff4ce-655">Proporciona acceso a los destinatarios en la línea **para** de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-655">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="ff4ce-656">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-656">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ff4ce-657">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-657">Read mode</span></span>

<span data-ttu-id="ff4ce-p133">La propiedad `to` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **Para** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ff4ce-660">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-660">Compose mode</span></span>

<span data-ttu-id="ff4ce-661">El `to` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **para** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-661">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="ff4ce-662">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-662">Type:</span></span>

*   <span data-ttu-id="ff4ce-663">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-663">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-664">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-664">Requirements</span></span>

|<span data-ttu-id="ff4ce-665">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-665">Requirement</span></span>|<span data-ttu-id="ff4ce-666">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-666">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-667">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-667">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-668">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-668">1.0</span></span>|
|[<span data-ttu-id="ff4ce-669">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-669">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-670">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-670">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-671">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-671">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-672">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-672">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-673">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-673">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="ff4ce-674">Métodos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-674">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="ff4ce-675">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ff4ce-675">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ff4ce-676">Agrega un archivo a un mensaje o cita como datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-676">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="ff4ce-677">El método `addFileAttachmentAsync` carga el archivo en el URI especificado y lo asocia al elemento en el formulario de redacción.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-677">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="ff4ce-678">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-678">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-679">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-679">Parameters:</span></span>
|<span data-ttu-id="ff4ce-680">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-680">Name</span></span>|<span data-ttu-id="ff4ce-681">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-681">Type</span></span>|<span data-ttu-id="ff4ce-682">Atributos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-682">Attributes</span></span>|<span data-ttu-id="ff4ce-683">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-683">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="ff4ce-684">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-684">String</span></span>||<span data-ttu-id="ff4ce-p134">El URI que proporciona la ubicación del archivo que se va a adjuntar al mensaje o a la cita. La longitud máxima es de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="ff4ce-687">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-687">String</span></span>||<span data-ttu-id="ff4ce-p135">El nombre de los datos adjuntos que se muestra mientras estos se cargan. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="ff4ce-690">Object</span><span class="sxs-lookup"><span data-stu-id="ff4ce-690">Object</span></span>|<span data-ttu-id="ff4ce-691">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-691">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-692">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-692">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ff4ce-693">Objeto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-693">Object</span></span>|<span data-ttu-id="ff4ce-694">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-694">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-695">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-695">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="ff4ce-696">Booleano</span><span class="sxs-lookup"><span data-stu-id="ff4ce-696">Boolean</span></span>|<span data-ttu-id="ff4ce-697">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-697">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-698">Si está establecido en `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-698">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="ff4ce-699">Función</span><span class="sxs-lookup"><span data-stu-id="ff4ce-699">function</span></span>|<span data-ttu-id="ff4ce-700">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-700">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-701">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-701">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ff4ce-702">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-702">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ff4ce-703">Si se produce un error en la carga de los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-703">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ff4ce-704">Errores</span><span class="sxs-lookup"><span data-stu-id="ff4ce-704">Errors</span></span>

|<span data-ttu-id="ff4ce-705">Código de error</span><span class="sxs-lookup"><span data-stu-id="ff4ce-705">Error code</span></span>|<span data-ttu-id="ff4ce-706">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-706">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="ff4ce-707">Los datos adjuntos son más grandes de lo permitido.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-707">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="ff4ce-708">Los datos adjuntos tienen una extensión que no está permitida.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-708">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="ff4ce-709">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-709">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-710">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-710">Requirements</span></span>

|<span data-ttu-id="ff4ce-711">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-711">Requirement</span></span>|<span data-ttu-id="ff4ce-712">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-712">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-713">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-713">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-714">1.1</span><span class="sxs-lookup"><span data-stu-id="ff4ce-714">1.1</span></span>|
|[<span data-ttu-id="ff4ce-715">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-715">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-716">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-716">ReadWriteItem</span></span>|
|[<span data-ttu-id="ff4ce-717">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-717">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-718">Redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-718">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ff4ce-719">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-719">Examples</span></span>

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

<span data-ttu-id="ff4ce-720">En el siguiente ejemplo, se agrega un archivo de imagen como un dato adjunto en línea y hace referencia a los datos adjuntos del cuerpo del mensaje.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-720">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="ff4ce-721">addFileAttachmentFromBase64Async (base64File, attachmentName, [opciones], [devolución de llamada])</span><span class="sxs-lookup"><span data-stu-id="ff4ce-721">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ff4ce-722">Agrega un archivo desde el base64 de codificación para un mensaje o cita como datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-722">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="ff4ce-723">El `addFileAttachmentFromBase64Async` método carga el archivo de la codificación base64 y se adjunta al elemento en el formulario de redacción.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-723">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="ff4ce-724">Este método devuelve el identificador de datos adjuntos en el objeto AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-724">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="ff4ce-725">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-725">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-726">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-726">Parameters:</span></span>
|<span data-ttu-id="ff4ce-727">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-727">Name</span></span>|<span data-ttu-id="ff4ce-728">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-728">Type</span></span>|<span data-ttu-id="ff4ce-729">Atributos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-729">Attributes</span></span>|<span data-ttu-id="ff4ce-730">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-730">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="ff4ce-731">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-731">String</span></span>||<span data-ttu-id="ff4ce-732">La base64 codificado contenido de una imagen o un archivo que se agregarán a un correo electrónico o un evento.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-732">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="ff4ce-733">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-733">String</span></span>||<span data-ttu-id="ff4ce-p137">El nombre de los datos adjuntos que se muestra mientras estos se cargan. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="ff4ce-736">Object</span><span class="sxs-lookup"><span data-stu-id="ff4ce-736">Object</span></span>|<span data-ttu-id="ff4ce-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-737">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-738">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-738">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ff4ce-739">Objeto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-739">Object</span></span>|<span data-ttu-id="ff4ce-740">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-740">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-741">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-741">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="ff4ce-742">Booleano</span><span class="sxs-lookup"><span data-stu-id="ff4ce-742">Boolean</span></span>|<span data-ttu-id="ff4ce-743">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-743">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-744">Si está establecido en `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-744">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="ff4ce-745">Función</span><span class="sxs-lookup"><span data-stu-id="ff4ce-745">function</span></span>|<span data-ttu-id="ff4ce-746">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-746">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-747">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-747">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ff4ce-748">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-748">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ff4ce-749">Si se produce un error en la carga de los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-749">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ff4ce-750">Errores</span><span class="sxs-lookup"><span data-stu-id="ff4ce-750">Errors</span></span>

|<span data-ttu-id="ff4ce-751">Código de error</span><span class="sxs-lookup"><span data-stu-id="ff4ce-751">Error code</span></span>|<span data-ttu-id="ff4ce-752">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-752">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="ff4ce-753">Los datos adjuntos son más grandes de lo permitido.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-753">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="ff4ce-754">Los datos adjuntos tienen una extensión que no está permitida.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-754">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="ff4ce-755">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-755">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-756">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-756">Requirements</span></span>

|<span data-ttu-id="ff4ce-757">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-757">Requirement</span></span>|<span data-ttu-id="ff4ce-758">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-758">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-759">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-759">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-760">Vista previa</span><span class="sxs-lookup"><span data-stu-id="ff4ce-760">Preview</span></span>|
|[<span data-ttu-id="ff4ce-761">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-761">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-762">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-762">ReadWriteItem</span></span>|
|[<span data-ttu-id="ff4ce-763">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-763">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-764">Redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-764">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ff4ce-765">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-765">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="ff4ce-766">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ff4ce-766">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="ff4ce-767">Agrega un controlador de eventos para un evento admitido.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-767">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="ff4ce-768">Actualmente son los tipos de eventos compatibles `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, y`Office.EventType.RecurrencePatternChanged`</span><span class="sxs-lookup"><span data-stu-id="ff4ce-768">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrencePatternChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-769">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-769">Parameters:</span></span>

| <span data-ttu-id="ff4ce-770">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-770">Name</span></span> | <span data-ttu-id="ff4ce-771">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-771">Type</span></span> | <span data-ttu-id="ff4ce-772">Atributos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-772">Attributes</span></span> | <span data-ttu-id="ff4ce-773">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-773">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="ff4ce-774">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="ff4ce-774">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="ff4ce-775">El evento que debe invocar el controlador.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-775">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="ff4ce-776">Función</span><span class="sxs-lookup"><span data-stu-id="ff4ce-776">Function</span></span> || <span data-ttu-id="ff4ce-p138">La función que va a controlar el evento. La función debe aceptar un único parámetro, que es un literal de objeto. La propiedad `type` del parámetro coincidirá con el parámetro `eventType` que se ha pasado a `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="ff4ce-780">Objeto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-780">Object</span></span> | <span data-ttu-id="ff4ce-781">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-781">&lt;optional&gt;</span></span> | <span data-ttu-id="ff4ce-782">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-782">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="ff4ce-783">Objeto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-783">Object</span></span> | <span data-ttu-id="ff4ce-784">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-784">&lt;optional&gt;</span></span> | <span data-ttu-id="ff4ce-785">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-785">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="ff4ce-786">función</span><span class="sxs-lookup"><span data-stu-id="ff4ce-786">function</span></span>| <span data-ttu-id="ff4ce-787">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-787">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-788">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-788">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-789">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-789">Requirements</span></span>

|<span data-ttu-id="ff4ce-790">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-790">Requirement</span></span>| <span data-ttu-id="ff4ce-791">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-791">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-792">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-792">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff4ce-793">Vista previa</span><span class="sxs-lookup"><span data-stu-id="ff4ce-793">Preview</span></span> |
|[<span data-ttu-id="ff4ce-794">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-794">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff4ce-795">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-795">ReadItem</span></span> |
|[<span data-ttu-id="ff4ce-796">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-796">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ff4ce-797">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-797">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="ff4ce-798">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-798">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrencePatternChanged, loadNewItem, function (result) {
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="ff4ce-799">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ff4ce-799">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ff4ce-800">Agrega un elemento de Exchange (por ejemplo, un mensaje) como datos adjuntos al mensaje o a la cita.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-800">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="ff4ce-p139">El método `addItemAttachmentAsync` asocia el elemento con el identificador especificado de Exchange al elemento en el formulario de redacción. Si se especifica un método de devolución de llamada, se llama al método con un parámetro, `asyncResult`, que contiene el identificador de datos adjuntos o un código que indique cualquier error que se ha producido al adjuntar el elemento. Si es necesario, puede usar el parámetro `options` para pasar información de estado al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="ff4ce-804">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-804">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="ff4ce-805">Si el complemento de Office se está ejecutando en Outlook Web App, el `addItemAttachmentAsync` método puede adjuntar elementos a los elementos que no sea el elemento que se está editando; Sin embargo, esto no se admite y no se recomienda.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-805">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-806">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-806">Parameters:</span></span>

|<span data-ttu-id="ff4ce-807">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-807">Name</span></span>|<span data-ttu-id="ff4ce-808">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-808">Type</span></span>|<span data-ttu-id="ff4ce-809">Atributos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-809">Attributes</span></span>|<span data-ttu-id="ff4ce-810">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-810">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="ff4ce-811">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-811">String</span></span>||<span data-ttu-id="ff4ce-p140">El identificador de Exchange del elemento que debe adjuntarse. La longitud máxima es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="ff4ce-814">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-814">String</span></span>||<span data-ttu-id="ff4ce-p141">El asunto del elemento que debe adjuntarse. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="ff4ce-817">Object</span><span class="sxs-lookup"><span data-stu-id="ff4ce-817">Object</span></span>|<span data-ttu-id="ff4ce-818">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-818">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-819">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-819">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ff4ce-820">Objeto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-820">Object</span></span>|<span data-ttu-id="ff4ce-821">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-821">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-822">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-822">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ff4ce-823">función</span><span class="sxs-lookup"><span data-stu-id="ff4ce-823">function</span></span>|<span data-ttu-id="ff4ce-824">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-824">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-825">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-825">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ff4ce-826">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-826">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ff4ce-827">Si se produce un error al agregar los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-827">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ff4ce-828">Errores</span><span class="sxs-lookup"><span data-stu-id="ff4ce-828">Errors</span></span>

|<span data-ttu-id="ff4ce-829">Código de error</span><span class="sxs-lookup"><span data-stu-id="ff4ce-829">Error code</span></span>|<span data-ttu-id="ff4ce-830">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-830">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="ff4ce-831">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-831">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-832">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-832">Requirements</span></span>

|<span data-ttu-id="ff4ce-833">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-833">Requirement</span></span>|<span data-ttu-id="ff4ce-834">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-835">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-835">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-836">1.1</span><span class="sxs-lookup"><span data-stu-id="ff4ce-836">1.1</span></span>|
|[<span data-ttu-id="ff4ce-837">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-837">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-838">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-838">ReadWriteItem</span></span>|
|[<span data-ttu-id="ff4ce-839">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-839">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-840">Redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-840">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-841">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-841">Example</span></span>

<span data-ttu-id="ff4ce-842">En el ejemplo siguiente, se agrega un elemento existente de Outlook como un archivo adjunto con el nombre `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-842">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="ff4ce-843">close()</span><span class="sxs-lookup"><span data-stu-id="ff4ce-843">close()</span></span>

<span data-ttu-id="ff4ce-844">Cierra el elemento actual que se está redactando.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-844">Closes the current item that is being composed.</span></span>

<span data-ttu-id="ff4ce-p142">El comportamiento del método `close` depende del estado actual del elemento que se está redactando. Si el elemento tiene cambios sin guardar, el cliente solicita al usuario que guarde, descarte o cancele la acción Cerrar.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-847">En Outlook, en el sitio web, si el elemento es una cita y anteriormente se ha guardado con `saveAsync`, se pide al usuario al guardar, descartar o cancelar incluso si no ha habido cambios desde la último el elemento guardado.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-847">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="ff4ce-848">En el cliente de escritorio de Outlook, si el mensaje es una respuesta directa, el método `close` no tiene ningún efecto.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-848">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-849">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-849">Requirements</span></span>

|<span data-ttu-id="ff4ce-850">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-850">Requirement</span></span>|<span data-ttu-id="ff4ce-851">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-851">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-852">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-852">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-853">1.3</span><span class="sxs-lookup"><span data-stu-id="ff4ce-853">1.3</span></span>|
|[<span data-ttu-id="ff4ce-854">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-854">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-855">Restringido</span><span class="sxs-lookup"><span data-stu-id="ff4ce-855">Restricted</span></span>|
|[<span data-ttu-id="ff4ce-856">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-856">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-857">Redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-857">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="ff4ce-858">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-858">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="ff4ce-859">Muestra un formulario de respuesta que incluye el remitente y todos los destinatarios del mensaje seleccionado o el organizador y todos los asistentes de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-859">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-860">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-860">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ff4ce-861">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-861">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ff4ce-862">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyAllForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-862">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="ff4ce-p143">Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-866">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-866">Parameters:</span></span>

|<span data-ttu-id="ff4ce-867">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-867">Name</span></span>|<span data-ttu-id="ff4ce-868">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-868">Type</span></span>|<span data-ttu-id="ff4ce-869">Atributos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-869">Attributes</span></span>|<span data-ttu-id="ff4ce-870">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-870">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="ff4ce-871">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ff4ce-871">String &#124; Object</span></span>||<span data-ttu-id="ff4ce-p144">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ff4ce-874">**O**</span><span class="sxs-lookup"><span data-stu-id="ff4ce-874">**OR**</span></span><br/><span data-ttu-id="ff4ce-p145">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="ff4ce-877">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-877">String</span></span>|<span data-ttu-id="ff4ce-878">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-878">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-p146">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="ff4ce-881">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-881">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="ff4ce-882">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-882">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-883">Una matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-883">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="ff4ce-884">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-884">String</span></span>||<span data-ttu-id="ff4ce-p147">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="ff4ce-887">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-887">String</span></span>||<span data-ttu-id="ff4ce-888">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-888">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="ff4ce-889">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-889">String</span></span>||<span data-ttu-id="ff4ce-p148">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="ff4ce-892">Booleano</span><span class="sxs-lookup"><span data-stu-id="ff4ce-892">Boolean</span></span>||<span data-ttu-id="ff4ce-p149">Solo se usa si `type` se establece en `file`. Si `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="ff4ce-895">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-895">String</span></span>||<span data-ttu-id="ff4ce-p150">Solo se usa si `type` se establece en `item`. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="ff4ce-899">función</span><span class="sxs-lookup"><span data-stu-id="ff4ce-899">function</span></span>|<span data-ttu-id="ff4ce-900">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-900">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-901">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-901">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-902">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-902">Requirements</span></span>

|<span data-ttu-id="ff4ce-903">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-903">Requirement</span></span>|<span data-ttu-id="ff4ce-904">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-905">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-906">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-906">1.0</span></span>|
|[<span data-ttu-id="ff4ce-907">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-908">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-909">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-910">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-910">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ff4ce-911">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-911">Examples</span></span>

<span data-ttu-id="ff4ce-912">El código siguiente pasa una cadena a la función `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-912">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="ff4ce-913">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-913">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="ff4ce-914">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-914">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ff4ce-915">Responder con un cuerpo y datos adjuntos de archivo.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-915">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="ff4ce-916">Responder con un cuerpo y datos adjuntos de elemento.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-916">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="ff4ce-917">Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-917">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="ff4ce-918">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-918">displayReplyForm(formData)</span></span>

<span data-ttu-id="ff4ce-919">Muestra un formulario de respuesta que incluye solo el remitente del mensaje seleccionado o el organizador de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-919">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-920">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-920">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ff4ce-921">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-921">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ff4ce-922">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-922">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="ff4ce-p151">Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-926">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-926">Parameters:</span></span>

|<span data-ttu-id="ff4ce-927">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-927">Name</span></span>|<span data-ttu-id="ff4ce-928">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-928">Type</span></span>|<span data-ttu-id="ff4ce-929">Atributos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-929">Attributes</span></span>|<span data-ttu-id="ff4ce-930">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-930">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="ff4ce-931">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ff4ce-931">String &#124; Object</span></span>||<span data-ttu-id="ff4ce-p152">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ff4ce-934">**O**</span><span class="sxs-lookup"><span data-stu-id="ff4ce-934">**OR**</span></span><br/><span data-ttu-id="ff4ce-p153">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="ff4ce-937">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-937">String</span></span>|<span data-ttu-id="ff4ce-938">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-938">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-p154">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="ff4ce-941">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-941">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="ff4ce-942">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-942">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-943">Una matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-943">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="ff4ce-944">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-944">String</span></span>||<span data-ttu-id="ff4ce-p155">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="ff4ce-947">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-947">String</span></span>||<span data-ttu-id="ff4ce-948">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-948">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="ff4ce-949">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-949">String</span></span>||<span data-ttu-id="ff4ce-p156">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="ff4ce-952">Booleano</span><span class="sxs-lookup"><span data-stu-id="ff4ce-952">Boolean</span></span>||<span data-ttu-id="ff4ce-p157">Solo se usa si `type` se establece en `file`. Si `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="ff4ce-955">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-955">String</span></span>||<span data-ttu-id="ff4ce-p158">Solo se usa si `type` se establece en `item`. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="ff4ce-959">función</span><span class="sxs-lookup"><span data-stu-id="ff4ce-959">function</span></span>|<span data-ttu-id="ff4ce-960">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-960">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-961">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-961">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-962">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-962">Requirements</span></span>

|<span data-ttu-id="ff4ce-963">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-963">Requirement</span></span>|<span data-ttu-id="ff4ce-964">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-964">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-965">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-965">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-966">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-966">1.0</span></span>|
|[<span data-ttu-id="ff4ce-967">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-967">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-968">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-968">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-969">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-969">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-970">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-970">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ff4ce-971">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-971">Examples</span></span>

<span data-ttu-id="ff4ce-972">El código siguiente pasa una cadena a la función `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-972">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="ff4ce-973">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-973">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="ff4ce-974">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-974">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ff4ce-975">Responder con un cuerpo y datos adjuntos de archivo.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-975">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="ff4ce-976">Responder con un cuerpo y datos adjuntos de elemento.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-976">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="ff4ce-977">Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-977">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="ff4ce-978">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="ff4ce-978">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="ff4ce-979">Obtiene las entidades que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-979">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-980">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-980">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-981">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-981">Requirements</span></span>

|<span data-ttu-id="ff4ce-982">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-982">Requirement</span></span>|<span data-ttu-id="ff4ce-983">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-983">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-984">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-984">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-985">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-985">1.0</span></span>|
|[<span data-ttu-id="ff4ce-986">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-986">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-987">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-987">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-988">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-988">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-989">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-989">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ff4ce-990">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-990">Returns:</span></span>

<span data-ttu-id="ff4ce-991">Tipo: [Entidades](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-991">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="ff4ce-992">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-992">Example</span></span>

<span data-ttu-id="ff4ce-993">En el ejemplo siguiente se obtiene acceso a las entidades de los contactos en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-993">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="ff4ce-994">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="ff4ce-994">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="ff4ce-995">Obtiene una matriz de todas las entidades del tipo de entidad especificado que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-995">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-996">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-996">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-997">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-997">Parameters:</span></span>

|<span data-ttu-id="ff4ce-998">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-998">Name</span></span>|<span data-ttu-id="ff4ce-999">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-999">Type</span></span>|<span data-ttu-id="ff4ce-1000">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1000">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="ff4ce-1001">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1001">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="ff4ce-1002">Uno de los valores de enumeración de EntityType.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1002">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-1003">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1003">Requirements</span></span>

|<span data-ttu-id="ff4ce-1004">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1004">Requirement</span></span>|<span data-ttu-id="ff4ce-1005">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-1006">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1006">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1007">1.0</span></span>|
|[<span data-ttu-id="ff4ce-1008">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1008">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-1009">Restringido</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1009">Restricted</span></span>|
|[<span data-ttu-id="ff4ce-1010">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1010">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-1011">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ff4ce-1012">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1012">Returns:</span></span>

<span data-ttu-id="ff4ce-1013">Si el valor que se pasa a `entityType` no es un miembro válido de la enumeración de `EntityType`, el método devuelve un valor NULL.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1013">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="ff4ce-1014">Si no existen entidades del tipo especificado están presentes en el cuerpo del elemento, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1014">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="ff4ce-1015">De otro modo, el tipo de los objetos en la matriz devuelta depende del tipo de entidad solicitada en el parámetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1015">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="ff4ce-1016">Aunque el nivel de permiso mínimo para usar este método es **Restricted**, algunos tipos de entidad requieren **ReadItem** para tener acceso, como se especifica en la tabla siguiente.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1016">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="ff4ce-1017">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1017">Value of `entityType`</span></span>|<span data-ttu-id="ff4ce-1018">Tipo de objetos de la matriz devuelta</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1018">Type of objects in returned array</span></span>|<span data-ttu-id="ff4ce-1019">Nivel de permiso requerido</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1019">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="ff4ce-1020">Cadena</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1020">String</span></span>|<span data-ttu-id="ff4ce-1021">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1021">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="ff4ce-1022">Contacto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1022">Contact</span></span>|<span data-ttu-id="ff4ce-1023">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1023">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="ff4ce-1024">Cadena</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1024">String</span></span>|<span data-ttu-id="ff4ce-1025">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1025">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="ff4ce-1026">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1026">MeetingSuggestion</span></span>|<span data-ttu-id="ff4ce-1027">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1027">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="ff4ce-1028">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1028">PhoneNumber</span></span>|<span data-ttu-id="ff4ce-1029">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1029">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="ff4ce-1030">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1030">TaskSuggestion</span></span>|<span data-ttu-id="ff4ce-1031">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1031">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="ff4ce-1032">Cadena</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1032">String</span></span>|<span data-ttu-id="ff4ce-1033">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1033">**Restricted**</span></span>|

<span data-ttu-id="ff4ce-1034">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="ff4ce-1034">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="ff4ce-1035">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1035">Example</span></span>

<span data-ttu-id="ff4ce-1036">En el ejemplo siguiente se muestra cómo obtener acceso a una matriz de cadenas que representan direcciones postales en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1036">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="ff4ce-1037">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1037">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="ff4ce-1038">Devuelve entidades conocidas en el elemento seleccionado que pasan el filtro con nombre definido en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1038">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-1039">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1039">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ff4ce-1040">El método `getFilteredEntitiesByName` devuelve las entidades que coinciden con la expresión regular definida en el elemento de regla [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) del archivo XML de manifiesto con el valor de elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1040">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-1041">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1041">Parameters:</span></span>

|<span data-ttu-id="ff4ce-1042">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1042">Name</span></span>|<span data-ttu-id="ff4ce-1043">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1043">Type</span></span>|<span data-ttu-id="ff4ce-1044">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1044">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="ff4ce-1045">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1045">String</span></span>|<span data-ttu-id="ff4ce-1046">El nombre del elemento de regla `ItemHasKnownEntity` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1046">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-1047">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1047">Requirements</span></span>

|<span data-ttu-id="ff4ce-1048">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1048">Requirement</span></span>|<span data-ttu-id="ff4ce-1049">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1049">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-1050">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1050">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-1051">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1051">1.0</span></span>|
|[<span data-ttu-id="ff4ce-1052">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1052">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-1053">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1053">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-1054">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1054">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-1055">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1055">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ff4ce-1056">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1056">Returns:</span></span>

<span data-ttu-id="ff4ce-p160">Si no existe ningún elemento `ItemHasKnownEntity` en el manifiesto con un valor de elemento `FilterName` que coincida con el parámetro `name`, el método devuelve `null`. Si el parámetro `name` no coincide con un elemento `ItemHasKnownEntity` del manifiesto, pero no existen entidades en el elemento actual que coincidan, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="ff4ce-1059">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="ff4ce-1059">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="ff4ce-1060">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1060">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="ff4ce-1061">Obtiene datos de inicialización que se pasan cuando [un mensaje accionable activa](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message) el complemento.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1061">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-1062">Este método sólo es compatible con Outlook 2016 para Windows (versiones de Click-to-Run mayores 16.0.8413.1000) y Outlook en el sitio web de Office 365.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1062">This method is only supported by Outlook 2016 for Windows (Click-to-Run versions greater than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-1063">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1063">Parameters:</span></span>
|<span data-ttu-id="ff4ce-1064">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1064">Name</span></span>|<span data-ttu-id="ff4ce-1065">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1065">Type</span></span>|<span data-ttu-id="ff4ce-1066">Atributos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1066">Attributes</span></span>|<span data-ttu-id="ff4ce-1067">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1067">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="ff4ce-1068">Objeto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1068">Object</span></span>|<span data-ttu-id="ff4ce-1069">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-1070">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1070">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ff4ce-1071">Objeto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1071">Object</span></span>|<span data-ttu-id="ff4ce-1072">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1072">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-1073">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1073">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ff4ce-1074">función</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1074">function</span></span>|<span data-ttu-id="ff4ce-1075">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1075">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-1076">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ff4ce-1077">En caso de éxito, los datos de inicialización se proporcionan en el `asyncResult.value` (propiedad) como una cadena.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1077">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="ff4ce-1078">Si no hay ningún contexto de inicialización, el objeto `asyncResult` contendrá un objeto `Error` con la propiedad `code` establecida en `9020` y la propiedad `name` establecida en `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1078">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-1079">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1079">Requirements</span></span>

|<span data-ttu-id="ff4ce-1080">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1080">Requirement</span></span>|<span data-ttu-id="ff4ce-1081">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1081">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-1082">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1082">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-1083">Vista previa</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1083">Preview</span></span>|
|[<span data-ttu-id="ff4ce-1084">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1084">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-1085">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1085">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-1086">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1086">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-1087">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1087">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-1088">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1088">Example</span></span>

```
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="ff4ce-1089">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1089">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="ff4ce-1090">Devuelve valores de cadena en el elemento seleccionado que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1090">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-1091">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1091">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ff4ce-p161">El método `getRegExMatches` devuelve las cadenas que coinciden con la expresión regular definida en cada elemento de regla `ItemHasRegularExpressionMatch` o `ItemHasKnownEntity` del archivo XML de manifiesto. Para una regla `ItemHasRegularExpressionMatch`, tiene que darse una cadena coincidente en la propiedad del elemento especificado por la regla. El tipo simple `PropertyName` define las propiedades admitidas.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ff4ce-1095">Por ejemplo, considere un manifiesto de complemento que tenga el siguiente elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1095">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ff4ce-1096">El objeto devuelto desde `getRegExMatches` tendría dos propiedades: `fruits` y `veggies`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1096">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="ff4ce-p162">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados. En su lugar, use el método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo el cuerpo.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-1100">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1100">Requirements</span></span>

|<span data-ttu-id="ff4ce-1101">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1101">Requirement</span></span>|<span data-ttu-id="ff4ce-1102">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1102">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-1103">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1103">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-1104">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1104">1.0</span></span>|
|[<span data-ttu-id="ff4ce-1105">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1105">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-1106">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1106">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-1107">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1107">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-1108">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1108">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ff4ce-1109">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1109">Returns:</span></span>

<span data-ttu-id="ff4ce-p163">Un objeto que contiene matrices de cadenas que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. El nombre de cada matriz es igual al valor correspondiente del atributo `RegExName` de la regla `ItemHasRegularExpressionMatch` coincidente o al atributo `FilterName` de la regla `ItemHasKnownEntity` coincidente.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="ff4ce-1112">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1112">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ff4ce-1113">Object</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1113">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ff4ce-1114">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1114">Example</span></span>

<span data-ttu-id="ff4ce-1115">En el siguiente ejemplo, se muestra cómo tener acceso a la matriz de coincidencias de los elementos de regla de expresión regular `fruits` y `veggies`, que se especifican en el manifiesto.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1115">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="ff4ce-1116">→ de getregexmatchesbyname (Name) (nullable) {matriz. < cadena >}</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1116">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="ff4ce-1117">Devuelve valores de cadena en el elemento seleccionado que coinciden con la expresión regular de nombre definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1117">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-1118">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1118">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ff4ce-1119">El método `getRegExMatchesByName` devuelve las cadenas que coinciden con la expresión regular definida en el elemento de regla `ItemHasRegularExpressionMatch` del archivo XML de manifiesto con el valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1119">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="ff4ce-p164">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-1122">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1122">Parameters:</span></span>

|<span data-ttu-id="ff4ce-1123">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1123">Name</span></span>|<span data-ttu-id="ff4ce-1124">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1124">Type</span></span>|<span data-ttu-id="ff4ce-1125">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1125">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="ff4ce-1126">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1126">String</span></span>|<span data-ttu-id="ff4ce-1127">El nombre del elemento de regla `ItemHasRegularExpressionMatch` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1127">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-1128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1128">Requirements</span></span>

|<span data-ttu-id="ff4ce-1129">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1129">Requirement</span></span>|<span data-ttu-id="ff4ce-1130">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1130">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-1131">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1131">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-1132">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1132">1.0</span></span>|
|[<span data-ttu-id="ff4ce-1133">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1133">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-1134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1134">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-1135">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1135">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-1136">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1136">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ff4ce-1137">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1137">Returns:</span></span>

<span data-ttu-id="ff4ce-1138">Una matriz que contiene las cadenas que coinciden con la expresión regular definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1138">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="ff4ce-1139">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1139">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ff4ce-1140">< Cadena > de matriz.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1140">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ff4ce-1141">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1141">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="ff4ce-1142">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1142">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="ff4ce-1143">Devuelve asincrónicamente datos seleccionados desde el asunto o el cuerpo de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1143">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="ff4ce-p165">Si no hay ninguna selección, pero el cursor está en el cuerpo o el asunto, el método devuelve null para los datos seleccionados. Si se selecciona un campo que no sea el cuerpo o el asunto, el método devuelve el error `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-1146">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1146">Parameters:</span></span>

|<span data-ttu-id="ff4ce-1147">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1147">Name</span></span>|<span data-ttu-id="ff4ce-1148">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1148">Type</span></span>|<span data-ttu-id="ff4ce-1149">Atributos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1149">Attributes</span></span>|<span data-ttu-id="ff4ce-1150">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1150">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="ff4ce-1151">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1151">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="ff4ce-p166">Solicita un formato para los datos. Si es Texto, el método devuelve el texto sin formato como cadena, quitando toda etiqueta HTML que hubiera. Si es HTML, el método devuelve el texto seleccionado, ya sea texto sin formato o HTML.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="ff4ce-1155">Object</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1155">Object</span></span>|<span data-ttu-id="ff4ce-1156">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1156">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-1157">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1157">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ff4ce-1158">Objeto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1158">Object</span></span>|<span data-ttu-id="ff4ce-1159">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-1160">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1160">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ff4ce-1161">función</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1161">function</span></span>||<span data-ttu-id="ff4ce-1162">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1162">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ff4ce-1163">Para tener acceso a los datos seleccionados desde el método de devolución de llamada, llame a `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1163">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="ff4ce-1164">Para obtener acceso a la propiedad de origen del que proviene la selección, llame a `asyncResult.value.sourceProperty`, que será cualquiera `body` o `subject`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1164">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-1165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1165">Requirements</span></span>

|<span data-ttu-id="ff4ce-1166">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1166">Requirement</span></span>|<span data-ttu-id="ff4ce-1167">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1167">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-1168">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-1169">1.2</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1169">1.2</span></span>|
|[<span data-ttu-id="ff4ce-1170">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-1171">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1171">ReadWriteItem</span></span>|
|[<span data-ttu-id="ff4ce-1172">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-1173">Redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1173">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="ff4ce-1174">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1174">Returns:</span></span>

<span data-ttu-id="ff4ce-1175">Los datos seleccionados como cadena con formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1175">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="ff4ce-1176">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1176">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ff4ce-1177">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1177">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ff4ce-1178">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1178">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="ff4ce-1179">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1179">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="ff4ce-p168">Obtiene las entidades que se han detectado en una coincidencia resaltada que un usuario ha seleccionado. Las coincidencias resaltadas se aplican a [complementos contextuales](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-1182">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1182">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-1183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1183">Requirements</span></span>

|<span data-ttu-id="ff4ce-1184">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1184">Requirement</span></span>|<span data-ttu-id="ff4ce-1185">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1185">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-1186">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1186">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-1187">1.6</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1187">1.6</span></span>|
|[<span data-ttu-id="ff4ce-1188">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1188">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-1189">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1189">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-1190">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-1191">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1191">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ff4ce-1192">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1192">Returns:</span></span>

<span data-ttu-id="ff4ce-1193">Tipo: [Entidades](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1193">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="ff4ce-1194">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1194">Example</span></span>

<span data-ttu-id="ff4ce-1195">En el siguiente ejemplo se accede a las entidades de direcciones en la coincidencia resaltada seleccionada por el usuario.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1195">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="ff4ce-1196">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1196">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="ff4ce-p169">Devuelve valores de cadena en una coincidencia resaltada que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. Las coincidencias resaltadas se aplican a [complementos contextuales](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-1199">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1199">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ff4ce-p170">El método `getSelectedRegExMatches` devuelve las cadenas que coinciden con la expresión regular definida en cada elemento de regla `ItemHasRegularExpressionMatch` o `ItemHasKnownEntity` del archivo XML de manifiesto. Para una regla `ItemHasRegularExpressionMatch`, tiene que darse una cadena coincidente en la propiedad del elemento especificado por la regla. El tipo simple `PropertyName` define las propiedades admitidas.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ff4ce-1203">Por ejemplo, considere un manifiesto de complemento que tenga el siguiente elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1203">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ff4ce-1204">El objeto devuelto desde `getRegExMatches` tendría dos propiedades: `fruits` y `veggies`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1204">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="ff4ce-p171">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados. En su lugar, use el método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo el cuerpo.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff4ce-1208">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1208">Requirements</span></span>

|<span data-ttu-id="ff4ce-1209">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1209">Requirement</span></span>|<span data-ttu-id="ff4ce-1210">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1210">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-1211">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1211">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-1212">1.6</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1212">1.6</span></span>|
|[<span data-ttu-id="ff4ce-1213">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1213">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-1214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1214">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-1215">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-1216">Lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1216">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ff4ce-1217">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1217">Returns:</span></span>

<span data-ttu-id="ff4ce-p172">Un objeto que contiene matrices de cadenas que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. El nombre de cada matriz es igual al valor correspondiente del atributo `RegExName` de la regla `ItemHasRegularExpressionMatch` coincidente o al atributo `FilterName` de la regla `ItemHasKnownEntity` coincidente.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="ff4ce-1220">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1220">Example</span></span>

<span data-ttu-id="ff4ce-1221">En el siguiente ejemplo, se muestra cómo tener acceso a la matriz de coincidencias de los elementos de regla de expresión regular `fruits` y `veggies`, que se especifican en el manifiesto.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1221">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="ff4ce-1222">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1222">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="ff4ce-1223">Carga de forma asincrónica las propiedades personalizadas de este complemento en el elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1223">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="ff4ce-p173">Las propiedades personalizadas se almacenan como pares de clave/valor según la aplicación y el elemento. Este método devuelve un objeto `CustomProperties` en la devolución de llamada, que proporciona métodos para tener acceso a las propiedades personalizadas específicas del elemento y el complemento actual. Las propiedades personalizadas no están cifradas en el elemento, por lo que no debería usarse como almacenamiento seguro.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-1227">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1227">Parameters:</span></span>

|<span data-ttu-id="ff4ce-1228">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1228">Name</span></span>|<span data-ttu-id="ff4ce-1229">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1229">Type</span></span>|<span data-ttu-id="ff4ce-1230">Atributos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1230">Attributes</span></span>|<span data-ttu-id="ff4ce-1231">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1231">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="ff4ce-1232">función</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1232">function</span></span>||<span data-ttu-id="ff4ce-1233">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1233">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ff4ce-1234">Las propiedades personalizadas se proporcionan como un objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1234">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="ff4ce-1235">Este objeto se puede utilizar para obtener, establecer y elimina las propiedades personalizadas del elemento y guarde los cambios en la propiedad personalizada que se vuelve a establecer en el servidor.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1235">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="ff4ce-1236">Objeto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1236">Object</span></span>|<span data-ttu-id="ff4ce-1237">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1237">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-1238">Los desarrolladores pueden proporcionar cualquier objeto que desean tener acceso en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1238">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="ff4ce-1239">Se puede tener acceso a este objeto por el `asyncResult.asyncContext` (propiedad) en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1239">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-1240">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1240">Requirements</span></span>

|<span data-ttu-id="ff4ce-1241">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1241">Requirement</span></span>|<span data-ttu-id="ff4ce-1242">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-1243">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1243">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-1244">1.0</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1244">1.0</span></span>|
|[<span data-ttu-id="ff4ce-1245">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1246">ReadItem</span></span>|
|[<span data-ttu-id="ff4ce-1247">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-1248">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-1249">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1249">Example</span></span>

<span data-ttu-id="ff4ce-p176">En el siguiente ejemplo de código se muestra cómo usar el método `loadCustomPropertiesAsync` para cargar de forma asincrónica propiedades personalizadas que son específicas del elemento actual. En el ejemplo también se muestra cómo usar el método `CustomProperties.saveAsync` para guardar estas propiedades en el servidor. Después de cargar las propiedades personalizadas, el ejemplo de código usa el método `CustomProperties.get` para leer la propiedad personalizada `myProp`, el método `CustomProperties.set` para escribir en la propiedad personalizada `otherProp` y finalmente se llama al método `saveAsync` para guardar las propiedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="ff4ce-1253">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1253">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="ff4ce-1254">Quita los datos adjuntos de un mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1254">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="ff4ce-p177">El método `removeAttachmentAsync` quita del elemento los datos adjuntos con el identificador especificado. Como práctica recomendada, debe usar el identificador de datos adjuntos para quitar datos adjuntos solo si la misma aplicación de correo ha agregado los datos adjuntos en la misma sesión. En Outlook Web App y OWA para los dispositivos, el identificador de datos adjuntos es válido solo en la misma sesión. Una sesión finaliza cuando el usuario cierra la aplicación, o si el usuario comienza a redactar en un formulario en línea y posteriormente se extrae el formulario en línea para continuar en una ventana independiente.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p177">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-1259">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1259">Parameters:</span></span>

|<span data-ttu-id="ff4ce-1260">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1260">Name</span></span>|<span data-ttu-id="ff4ce-1261">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1261">Type</span></span>|<span data-ttu-id="ff4ce-1262">Atributos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1262">Attributes</span></span>|<span data-ttu-id="ff4ce-1263">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1263">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="ff4ce-1264">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1264">String</span></span>||<span data-ttu-id="ff4ce-p178">El identificador de los datos adjuntos que se deben quitar. La longitud máxima de la cadena es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p178">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="ff4ce-1267">Object</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1267">Object</span></span>|<span data-ttu-id="ff4ce-1268">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-1269">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ff4ce-1270">Objeto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1270">Object</span></span>|<span data-ttu-id="ff4ce-1271">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-1272">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ff4ce-1273">función</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1273">function</span></span>|<span data-ttu-id="ff4ce-1274">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1274">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-1275">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1275">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ff4ce-1276">Si se produce un error en la eliminación de los datos adjuntos, la propiedad `asyncResult.error` contendrá un código de error con el motivo del error.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1276">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ff4ce-1277">Errores</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1277">Errors</span></span>

|<span data-ttu-id="ff4ce-1278">Código de error</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1278">Error code</span></span>|<span data-ttu-id="ff4ce-1279">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1279">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="ff4ce-1280">El identificador de datos adjuntos no existe.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1280">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-1281">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1281">Requirements</span></span>

|<span data-ttu-id="ff4ce-1282">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1282">Requirement</span></span>|<span data-ttu-id="ff4ce-1283">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1283">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-1284">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1284">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-1285">1.1</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1285">1.1</span></span>|
|[<span data-ttu-id="ff4ce-1286">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1286">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-1287">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1287">ReadWriteItem</span></span>|
|[<span data-ttu-id="ff4ce-1288">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1288">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-1289">Redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1289">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-1290">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1290">Example</span></span>

<span data-ttu-id="ff4ce-1291">El código siguiente quita un archivo adjunto con un identificador de "0".</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1291">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="ff4ce-1292">removeHandlerAsync (eventType, controlador, [opciones], [devolución de llamada])</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1292">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="ff4ce-1293">Quita un controlador de eventos para un evento compatible.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1293">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="ff4ce-1294">Actualmente son los tipos de eventos compatibles `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, y`Office.EventType.RecurrencePatternChanged`</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1294">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrencePatternChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-1295">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1295">Parameters:</span></span>

| <span data-ttu-id="ff4ce-1296">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1296">Name</span></span> | <span data-ttu-id="ff4ce-1297">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1297">Type</span></span> | <span data-ttu-id="ff4ce-1298">Atributos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1298">Attributes</span></span> | <span data-ttu-id="ff4ce-1299">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1299">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="ff4ce-1300">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1300">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="ff4ce-1301">El evento que debe invocar el controlador.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1301">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="ff4ce-1302">Función</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1302">Function</span></span> || <span data-ttu-id="ff4ce-p179">La función que va a controlar el evento. La función debe aceptar un único parámetro, que es un literal de objeto. La propiedad `type` del parámetro coincidirá con el parámetro `eventType` que se ha pasado a `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p179">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="ff4ce-1306">Objeto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1306">Object</span></span> | <span data-ttu-id="ff4ce-1307">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1307">&lt;optional&gt;</span></span> | <span data-ttu-id="ff4ce-1308">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1308">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="ff4ce-1309">Objeto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1309">Object</span></span> | <span data-ttu-id="ff4ce-1310">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1310">&lt;optional&gt;</span></span> | <span data-ttu-id="ff4ce-1311">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1311">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="ff4ce-1312">función</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1312">function</span></span>| <span data-ttu-id="ff4ce-1313">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1313">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-1314">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1314">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-1315">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1315">Requirements</span></span>

|<span data-ttu-id="ff4ce-1316">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1316">Requirement</span></span>| <span data-ttu-id="ff4ce-1317">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1317">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-1318">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1318">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff4ce-1319">Vista previa</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1319">Preview</span></span> |
|[<span data-ttu-id="ff4ce-1320">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1320">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff4ce-1321">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1321">ReadItem</span></span> |
|[<span data-ttu-id="ff4ce-1322">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1322">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ff4ce-1323">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1323">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="ff4ce-1324">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1324">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrencePatternChanged, loadNewItem, function (result) {
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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="ff4ce-1325">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1325">saveAsync([options], callback)</span></span>

<span data-ttu-id="ff4ce-1326">Guarda un elemento de forma asincrónica.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1326">Asynchronously saves an item.</span></span>

<span data-ttu-id="ff4ce-p180">Cuando se invoca, este método guarda el mensaje actual como un borrador y devuelve el identificador de elemento a través del método de devolución de llamada. En el modo en línea de Outlook Web App o Outlook, el elemento se guarda en el servidor. En modo en caché de Outlook, se guarda el elemento en la caché local.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p180">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-1330">Si el complemento llama a `saveAsync` en un elemento en modo de redacción con el fin de obtener una `itemId` para usar con EWS o la API de REST, tenga en cuenta que cuando Outlook está en modo en caché, puede tardar algún tiempo antes de que el elemento realmente se sincroniza con el servidor.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1330">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="ff4ce-1331">Hasta que se sincroniza el elemento, el uso de la `itemId` devolverá un error.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1331">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="ff4ce-p182">Dado que las citas no tienen ningún estado de borrador, si se llama a `saveAsync` en una cita en modo de redacción, el elemento se guardará como una cita normal en el calendario del usuario. Para las citas nuevas que todavía no se han guardado, no se enviará ninguna invitación. Si se guarda una cita existente, se enviará una actualización a los asistentes agregados o eliminados.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p182">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="ff4ce-1335">Los siguientes clientes tienen un comportamiento diferente para `saveAsync` en citas en modo de redacción:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1335">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="ff4ce-1336">Mac Outlook no admite `saveAsync` en una reunión en el modo de redacción.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1336">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="ff4ce-1337">Al llamar a `saveAsync` en una reunión en Mac Outlook devolverá un error.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1337">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="ff4ce-1338">Outlook en el web siempre envía una invitación o actualizar cuándo `saveAsync` se llama en una cita en modo de redacción.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1338">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-1339">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1339">Parameters:</span></span>

|<span data-ttu-id="ff4ce-1340">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1340">Name</span></span>|<span data-ttu-id="ff4ce-1341">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1341">Type</span></span>|<span data-ttu-id="ff4ce-1342">Atributos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1342">Attributes</span></span>|<span data-ttu-id="ff4ce-1343">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1343">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="ff4ce-1344">Objeto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1344">Object</span></span>|<span data-ttu-id="ff4ce-1345">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1345">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-1346">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1346">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ff4ce-1347">Objeto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1347">Object</span></span>|<span data-ttu-id="ff4ce-1348">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1348">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-1349">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1349">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="ff4ce-1350">función</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1350">function</span></span>||<span data-ttu-id="ff4ce-1351">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1351">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ff4ce-1352">En caso de éxito, se proporciona el identificador de elemento en el `asyncResult.value` (propiedad).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1352">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-1353">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1353">Requirements</span></span>

|<span data-ttu-id="ff4ce-1354">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1354">Requirement</span></span>|<span data-ttu-id="ff4ce-1355">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1355">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-1356">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1356">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-1357">1.3</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1357">1.3</span></span>|
|[<span data-ttu-id="ff4ce-1358">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1358">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-1359">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1359">ReadWriteItem</span></span>|
|[<span data-ttu-id="ff4ce-1360">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1360">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-1361">Redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1361">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ff4ce-1362">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1362">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="ff4ce-p184">A continuación, se muestra un ejemplo del parámetro `result` que se ha pasado a la función de devolución de llamada. La propiedad `value` contiene el identificador de elemento del elemento.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p184">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="ff4ce-1365">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1365">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="ff4ce-1366">Inserta asincrónicamente datos en el cuerpo o el asunto de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1366">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="ff4ce-p185">El método `setSelectedDataAsync` inserta la cadena especificada en la posición del cursor en el asunto o el cuerpo del elemento, o si se selecciona texto en el editor, reemplaza el texto seleccionado. Si el cursor no está en el cuerpo o el campo del asunto, se devuelve un error. Después de la inserción, el cursor se coloca al final del contenido insertado.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p185">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff4ce-1370">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1370">Parameters:</span></span>

|<span data-ttu-id="ff4ce-1371">Nombre</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1371">Name</span></span>|<span data-ttu-id="ff4ce-1372">Tipo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1372">Type</span></span>|<span data-ttu-id="ff4ce-1373">Atributos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1373">Attributes</span></span>|<span data-ttu-id="ff4ce-1374">Descripción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1374">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="ff4ce-1375">String</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1375">String</span></span>||<span data-ttu-id="ff4ce-p186">Datos que se van a insertar. Los datos no deben superar 1.000.000 de caracteres. Si se pasan más de 1.000.000 de caracteres, se produce una excepción `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p186">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="ff4ce-1379">Object</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1379">Object</span></span>|<span data-ttu-id="ff4ce-1380">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1380">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-1381">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1381">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="ff4ce-1382">Objeto</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1382">Object</span></span>|<span data-ttu-id="ff4ce-1383">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1383">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-1384">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1384">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="ff4ce-1385">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1385">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="ff4ce-1386">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1386">&lt;optional&gt;</span></span>|<span data-ttu-id="ff4ce-p187">Si `text`, se aplica el estilo actual en Outlook Web App y Outlook. Si el campo es un editor de HTML, se insertan solo los datos de texto, aunque los datos sean HTML.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p187">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="ff4ce-p188">Si `html` y el campo admiten HTML (el asunto no), el estilo actual se aplica en Outlook Web App y el estilo predeterminado en Outlook. Si el campo es un campo de texto, se devuelve un error `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-p188">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="ff4ce-1391">Si `coercionType` no está establecido, el resultado depende del campo: si el campo es HTML, se usa HTML; si el campo es texto, se usa texto sin formato.</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1391">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="ff4ce-1392">function</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1392">function</span></span>||<span data-ttu-id="ff4ce-1393">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1393">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff4ce-1394">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1394">Requirements</span></span>

|<span data-ttu-id="ff4ce-1395">Requirement</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1395">Requirement</span></span>|<span data-ttu-id="ff4ce-1396">Valor</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1396">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff4ce-1397">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1397">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="ff4ce-1398">1.2</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1398">1.2</span></span>|
|[<span data-ttu-id="ff4ce-1399">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1399">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="ff4ce-1400">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1400">ReadWriteItem</span></span>|
|[<span data-ttu-id="ff4ce-1401">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1401">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="ff4ce-1402">Redacción</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1402">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ff4ce-1403">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ff4ce-1403">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```