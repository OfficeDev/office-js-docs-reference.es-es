
# <a name="item"></a><span data-ttu-id="a3dd2-101">elemento</span><span class="sxs-lookup"><span data-stu-id="a3dd2-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="a3dd2-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="a3dd2-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="a3dd2-p101">El espacio de nombres `item` se usa para acceder al mensaje seleccionado actualmente, a una convocatoria de reunión o a una cita. Puede determinar el tipo del `item` mediante la propiedad [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-105">Requirements</span></span>

|<span data-ttu-id="a3dd2-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-106">Requirement</span></span>|<span data-ttu-id="a3dd2-107">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-108">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-109">1.0</span></span>|
|[<span data-ttu-id="a3dd2-110">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-111">Restringido</span><span class="sxs-lookup"><span data-stu-id="a3dd2-111">Restricted</span></span>|
|[<span data-ttu-id="a3dd2-112">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-113">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a3dd2-114">Miembros y métodos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-114">Members and methods</span></span>

| <span data-ttu-id="a3dd2-115">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-115">Member</span></span> | <span data-ttu-id="a3dd2-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a3dd2-117">attachments</span><span class="sxs-lookup"><span data-stu-id="a3dd2-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="a3dd2-118">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-118">Member</span></span> |
| [<span data-ttu-id="a3dd2-119">bcc</span><span class="sxs-lookup"><span data-stu-id="a3dd2-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a3dd2-120">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-120">Member</span></span> |
| [<span data-ttu-id="a3dd2-121">body</span><span class="sxs-lookup"><span data-stu-id="a3dd2-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="a3dd2-122">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-122">Member</span></span> |
| [<span data-ttu-id="a3dd2-123">cc</span><span class="sxs-lookup"><span data-stu-id="a3dd2-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a3dd2-124">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-124">Member</span></span> |
| [<span data-ttu-id="a3dd2-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="a3dd2-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="a3dd2-126">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-126">Member</span></span> |
| [<span data-ttu-id="a3dd2-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="a3dd2-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="a3dd2-128">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-128">Member</span></span> |
| [<span data-ttu-id="a3dd2-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="a3dd2-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="a3dd2-130">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-130">Member</span></span> |
| [<span data-ttu-id="a3dd2-131">end</span><span class="sxs-lookup"><span data-stu-id="a3dd2-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="a3dd2-132">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-132">Member</span></span> |
| [<span data-ttu-id="a3dd2-133">from</span><span class="sxs-lookup"><span data-stu-id="a3dd2-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="a3dd2-134">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-134">Member</span></span> |
| [<span data-ttu-id="a3dd2-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="a3dd2-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="a3dd2-136">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-136">Member</span></span> |
| [<span data-ttu-id="a3dd2-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="a3dd2-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="a3dd2-138">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-138">Member</span></span> |
| [<span data-ttu-id="a3dd2-139">itemId</span><span class="sxs-lookup"><span data-stu-id="a3dd2-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="a3dd2-140">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-140">Member</span></span> |
| [<span data-ttu-id="a3dd2-141">itemType</span><span class="sxs-lookup"><span data-stu-id="a3dd2-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="a3dd2-142">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-142">Member</span></span> |
| [<span data-ttu-id="a3dd2-143">location</span><span class="sxs-lookup"><span data-stu-id="a3dd2-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="a3dd2-144">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-144">Member</span></span> |
| [<span data-ttu-id="a3dd2-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="a3dd2-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="a3dd2-146">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-146">Member</span></span> |
| [<span data-ttu-id="a3dd2-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="a3dd2-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="a3dd2-148">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-148">Member</span></span> |
| [<span data-ttu-id="a3dd2-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="a3dd2-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a3dd2-150">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-150">Member</span></span> |
| [<span data-ttu-id="a3dd2-151">organizer</span><span class="sxs-lookup"><span data-stu-id="a3dd2-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="a3dd2-152">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-152">Member</span></span> |
| [<span data-ttu-id="a3dd2-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="a3dd2-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="a3dd2-154">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-154">Member</span></span> |
| [<span data-ttu-id="a3dd2-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="a3dd2-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a3dd2-156">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-156">Member</span></span> |
| [<span data-ttu-id="a3dd2-157">sender</span><span class="sxs-lookup"><span data-stu-id="a3dd2-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="a3dd2-158">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-158">Member</span></span> |
| [<span data-ttu-id="a3dd2-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="a3dd2-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="a3dd2-160">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-160">Member</span></span> |
| [<span data-ttu-id="a3dd2-161">start</span><span class="sxs-lookup"><span data-stu-id="a3dd2-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="a3dd2-162">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-162">Member</span></span> |
| [<span data-ttu-id="a3dd2-163">subject</span><span class="sxs-lookup"><span data-stu-id="a3dd2-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="a3dd2-164">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-164">Member</span></span> |
| [<span data-ttu-id="a3dd2-165">to</span><span class="sxs-lookup"><span data-stu-id="a3dd2-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a3dd2-166">Miembro	</span><span class="sxs-lookup"><span data-stu-id="a3dd2-166">Member</span></span> |
| [<span data-ttu-id="a3dd2-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a3dd2-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="a3dd2-168">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-168">Method</span></span> |
| [<span data-ttu-id="a3dd2-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="a3dd2-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="a3dd2-170">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-170">Method</span></span> |
| [<span data-ttu-id="a3dd2-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a3dd2-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="a3dd2-172">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-172">Method</span></span> |
| [<span data-ttu-id="a3dd2-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a3dd2-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="a3dd2-174">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-174">Method</span></span> |
| [<span data-ttu-id="a3dd2-175">close</span><span class="sxs-lookup"><span data-stu-id="a3dd2-175">close</span></span>](#close) | <span data-ttu-id="a3dd2-176">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-176">Method</span></span> |
| [<span data-ttu-id="a3dd2-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="a3dd2-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="a3dd2-178">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-178">Method</span></span> |
| [<span data-ttu-id="a3dd2-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="a3dd2-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="a3dd2-180">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-180">Method</span></span> |
| [<span data-ttu-id="a3dd2-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="a3dd2-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="a3dd2-182">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-182">Method</span></span> |
| [<span data-ttu-id="a3dd2-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="a3dd2-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="a3dd2-184">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-184">Method</span></span> |
| [<span data-ttu-id="a3dd2-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="a3dd2-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="a3dd2-186">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-186">Method</span></span> |
| [<span data-ttu-id="a3dd2-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="a3dd2-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="a3dd2-188">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-188">Method</span></span> |
| [<span data-ttu-id="a3dd2-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a3dd2-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="a3dd2-190">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-190">Method</span></span> |
| [<span data-ttu-id="a3dd2-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="a3dd2-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="a3dd2-192">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-192">Method</span></span> |
| [<span data-ttu-id="a3dd2-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a3dd2-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="a3dd2-194">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-194">Method</span></span> |
| [<span data-ttu-id="a3dd2-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="a3dd2-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="a3dd2-196">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-196">Method</span></span> |
| [<span data-ttu-id="a3dd2-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a3dd2-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="a3dd2-198">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-198">Method</span></span> |
| [<span data-ttu-id="a3dd2-199">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="a3dd2-199">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="a3dd2-200">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-200">Method</span></span> |
| [<span data-ttu-id="a3dd2-201">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="a3dd2-201">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="a3dd2-202">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-202">Method</span></span> |
| [<span data-ttu-id="a3dd2-203">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a3dd2-203">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="a3dd2-204">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-204">Method</span></span> |
| [<span data-ttu-id="a3dd2-205">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a3dd2-205">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="a3dd2-206">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-206">Method</span></span> |
| [<span data-ttu-id="a3dd2-207">saveAsync</span><span class="sxs-lookup"><span data-stu-id="a3dd2-207">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="a3dd2-208">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-208">Method</span></span> |
| [<span data-ttu-id="a3dd2-209">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a3dd2-209">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="a3dd2-210">Método</span><span class="sxs-lookup"><span data-stu-id="a3dd2-210">Method</span></span> |

### <a name="example"></a><span data-ttu-id="a3dd2-211">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-211">Example</span></span>

<span data-ttu-id="a3dd2-212">En el siguiente ejemplo de código de JavaScript, se muestra cómo tener acceso a la propiedad `subject` del elemento actual en Outlook.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-212">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="a3dd2-213">Miembros</span><span class="sxs-lookup"><span data-stu-id="a3dd2-213">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="a3dd2-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a3dd2-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="a3dd2-p102">Obtiene una matriz de datos adjuntos para el elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-217">Ciertos tipos de archivos están bloqueados por Outlook debido a problemas de seguridad y, por tanto, no se devuelven.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-217">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="a3dd2-218">Para obtener más información, vea [datos adjuntos bloqueados en Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-218">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-219">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-219">Type:</span></span>

*   <span data-ttu-id="a3dd2-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a3dd2-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-221">Requirements</span></span>

|<span data-ttu-id="a3dd2-222">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-222">Requirement</span></span>|<span data-ttu-id="a3dd2-223">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-224">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-225">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-225">1.0</span></span>|
|[<span data-ttu-id="a3dd2-226">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-227">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-228">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-229">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-230">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-230">Example</span></span>

<span data-ttu-id="a3dd2-231">El siguiente código crea una cadena HTML con los detalles de todos los datos adjuntos en el elemento actual.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-231">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a3dd2-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a3dd2-233">Obtiene un objeto que proporciona métodos para obtener o actualizar a los destinatarios de la línea de CCO (con copia oculta) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-233">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="a3dd2-234">Solo modo Redacción.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-234">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-235">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-235">Type:</span></span>

*   [<span data-ttu-id="a3dd2-236">Destinatarios</span><span class="sxs-lookup"><span data-stu-id="a3dd2-236">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="a3dd2-237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-237">Requirements</span></span>

|<span data-ttu-id="a3dd2-238">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-238">Requirement</span></span>|<span data-ttu-id="a3dd2-239">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-240">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-241">1.1</span><span class="sxs-lookup"><span data-stu-id="a3dd2-241">1.1</span></span>|
|[<span data-ttu-id="a3dd2-242">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-243">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-244">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-245">Redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-245">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-246">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-246">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="a3dd2-247">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-247">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="a3dd2-248">Obtiene un objeto que proporciona métodos para manipular el cuerpo de un elemento.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-248">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-249">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-249">Type:</span></span>

*   [<span data-ttu-id="a3dd2-250">Cuerpo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-250">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="a3dd2-251">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-251">Requirements</span></span>

|<span data-ttu-id="a3dd2-252">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-252">Requirement</span></span>|<span data-ttu-id="a3dd2-253">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-254">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-254">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-255">1.1</span><span class="sxs-lookup"><span data-stu-id="a3dd2-255">1.1</span></span>|
|[<span data-ttu-id="a3dd2-256">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-256">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-257">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-258">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-258">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-259">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-259">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a3dd2-260">cc: matriz. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-260">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a3dd2-261">Proporciona acceso a los destinatarios de Cc (con copia) de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-261">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="a3dd2-262">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-262">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a3dd2-263">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-263">Read mode</span></span>

<span data-ttu-id="a3dd2-p106">La propiedad `cc` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **CC** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a3dd2-266">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-266">Compose mode</span></span>

<span data-ttu-id="a3dd2-267">El `cc` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **Cc** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-267">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-268">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-268">Type:</span></span>

*   <span data-ttu-id="a3dd2-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-270">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-270">Requirements</span></span>

|<span data-ttu-id="a3dd2-271">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-271">Requirement</span></span>|<span data-ttu-id="a3dd2-272">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-273">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-274">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-274">1.0</span></span>|
|[<span data-ttu-id="a3dd2-275">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-276">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-277">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-278">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-278">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-279">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-279">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="a3dd2-280">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-280">(nullable) conversationId :String</span></span>

<span data-ttu-id="a3dd2-281">Obtiene un identificador de la conversación de correo electrónico que contiene un mensaje determinado.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-281">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="a3dd2-p107">Puede obtener un entero para esta propiedad si la aplicación de correo se activa en respuestas o formularios de lectura o en formularios de redacción. Si, después, el usuario cambia el asunto del mensaje de respuesta, al enviar dicha respuesta, el Id. de conversación de ese mensaje cambiará y el valor que se obtuvo anteriormente ya no será de aplicación.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="a3dd2-p108">Obtiene un valor NULL para esta propiedad para un nuevo elemento de un formulario de redacción. Si el usuario establece un asunto y guarda el elemento, la propiedad `conversationId` devolverá un valor.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-286">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-286">Type:</span></span>

*   <span data-ttu-id="a3dd2-287">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-287">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-288">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-288">Requirements</span></span>

|<span data-ttu-id="a3dd2-289">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-289">Requirement</span></span>|<span data-ttu-id="a3dd2-290">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-291">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-291">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-292">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-292">1.0</span></span>|
|[<span data-ttu-id="a3dd2-293">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-294">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-295">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-296">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-296">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="a3dd2-297">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="a3dd2-297">dateTimeCreated :Date</span></span>

<span data-ttu-id="a3dd2-p109">Obtiene la fecha y hora de creación de un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-300">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-300">Type:</span></span>

*   <span data-ttu-id="a3dd2-301">Fecha</span><span class="sxs-lookup"><span data-stu-id="a3dd2-301">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-302">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-302">Requirements</span></span>

|<span data-ttu-id="a3dd2-303">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-303">Requirement</span></span>|<span data-ttu-id="a3dd2-304">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-305">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-305">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-306">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-306">1.0</span></span>|
|[<span data-ttu-id="a3dd2-307">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-308">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-309">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-310">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-311">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-311">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="a3dd2-312">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="a3dd2-312">dateTimeModified :Date</span></span>

<span data-ttu-id="a3dd2-p110">Obtiene la fecha y hora en que se modificó por última vez un elemento. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-315">Este miembro no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-315">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-316">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-316">Type:</span></span>

*   <span data-ttu-id="a3dd2-317">Fecha</span><span class="sxs-lookup"><span data-stu-id="a3dd2-317">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-318">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-318">Requirements</span></span>

|<span data-ttu-id="a3dd2-319">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-319">Requirement</span></span>|<span data-ttu-id="a3dd2-320">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-320">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-321">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-321">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-322">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-322">1.0</span></span>|
|[<span data-ttu-id="a3dd2-323">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-323">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-324">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-324">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-325">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-325">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-326">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-326">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-327">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-327">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="a3dd2-328">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-328">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="a3dd2-329">Obtiene o establece la fecha y la hora de finalización de la cita.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-329">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="a3dd2-p111">La propiedad `end` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para convertir el valor de la propiedad final a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a3dd2-332">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-332">Read mode</span></span>

<span data-ttu-id="a3dd2-333">La propiedad `end` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-333">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a3dd2-334">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-334">Compose mode</span></span>

<span data-ttu-id="a3dd2-335">La propiedad `end` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-335">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="a3dd2-336">Si usa el método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para establecer la hora de finalización, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-336">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-337">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-337">Type:</span></span>

*   <span data-ttu-id="a3dd2-338">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-338">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-339">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-339">Requirements</span></span>

|<span data-ttu-id="a3dd2-340">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-340">Requirement</span></span>|<span data-ttu-id="a3dd2-341">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-342">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-342">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-343">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-343">1.0</span></span>|
|[<span data-ttu-id="a3dd2-344">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-344">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-345">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-346">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-346">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-347">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-347">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-348">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-348">Example</span></span>

<span data-ttu-id="a3dd2-349">En el ejemplo siguiente, se establece la hora de finalización de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-349">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="a3dd2-350">de:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[desde](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-350">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="a3dd2-351">Obtiene la dirección de correo electrónico del remitente de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-351">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="a3dd2-p112">Las propiedades `from` y [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-354">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `from` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-354">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a3dd2-355">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-355">Read mode</span></span>

<span data-ttu-id="a3dd2-356">El `from` (propiedad) devuelve un `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-356">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="a3dd2-357">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-357">Compose mode</span></span>

<span data-ttu-id="a3dd2-358">El `from` (propiedad) devuelve un `From` objeto que proporciona un método para obtener el valor de.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-358">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a3dd2-359">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-359">Type:</span></span>

*   <span data-ttu-id="a3dd2-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [desde](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-361">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-361">Requirements</span></span>

|<span data-ttu-id="a3dd2-362">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-362">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="a3dd2-363">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-363">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-364">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-364">1.0</span></span>|<span data-ttu-id="a3dd2-365">1.7</span><span class="sxs-lookup"><span data-stu-id="a3dd2-365">1.7</span></span>|
|[<span data-ttu-id="a3dd2-366">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-366">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-367">ReadItem</span></span>|<span data-ttu-id="a3dd2-368">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-368">ReadWriteItem</span></span>|
|[<span data-ttu-id="a3dd2-369">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-369">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-370">Read</span><span class="sxs-lookup"><span data-stu-id="a3dd2-370">Read</span></span>|<span data-ttu-id="a3dd2-371">Redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-371">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="a3dd2-372">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-372">internetMessageId :String</span></span>

<span data-ttu-id="a3dd2-p113">Obtiene el identificador de mensaje de Internet de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-375">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-375">Type:</span></span>

*   <span data-ttu-id="a3dd2-376">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-376">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-377">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-377">Requirements</span></span>

|<span data-ttu-id="a3dd2-378">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-378">Requirement</span></span>|<span data-ttu-id="a3dd2-379">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-380">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-380">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-381">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-381">1.0</span></span>|
|[<span data-ttu-id="a3dd2-382">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-382">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-383">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-384">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-384">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-385">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-385">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-386">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-386">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="a3dd2-387">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-387">itemClass :String</span></span>

<span data-ttu-id="a3dd2-p114">Obtiene la clase de elemento Servicios Web Exchange del elemento seleccionado. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="a3dd2-p115">La propiedad `itemClass` especifica la clase de mensaje del elemento seleccionado. Las siguientes son las clases de mensaje predeterminadas para el elemento de mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="a3dd2-392">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-392">Type</span></span>|<span data-ttu-id="a3dd2-393">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-393">Description</span></span>|<span data-ttu-id="a3dd2-394">Clase de elemento</span><span class="sxs-lookup"><span data-stu-id="a3dd2-394">item class</span></span>|
|---|---|---|
|<span data-ttu-id="a3dd2-395">Elementos de cita</span><span class="sxs-lookup"><span data-stu-id="a3dd2-395">Appointment items</span></span>|<span data-ttu-id="a3dd2-396">Estos son los elementos de calendario de la clase de elemento `IPM.Appointment` o `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-396">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="a3dd2-397">Elementos de mensaje</span><span class="sxs-lookup"><span data-stu-id="a3dd2-397">Message items</span></span>|<span data-ttu-id="a3dd2-398">Entre estos se encuentran los mensajes de correo electrónico que tienen la clase de mensaje predeterminada `IPM.Note` y las convocatorias de reunión, respuestas y cancelaciones que usan `IPM.Schedule.Meeting` como clase de mensaje base.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-398">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="a3dd2-399">Puede crear clases de mensaje personalizadas para ampliar una clase de mensaje predeterminada. Por ejemplo, puede crear una clase de mensaje de cita personalizada `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-399">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-400">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-400">Type:</span></span>

*   <span data-ttu-id="a3dd2-401">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-401">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-402">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-402">Requirements</span></span>

|<span data-ttu-id="a3dd2-403">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-403">Requirement</span></span>|<span data-ttu-id="a3dd2-404">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-404">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-405">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-405">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-406">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-406">1.0</span></span>|
|[<span data-ttu-id="a3dd2-407">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-407">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-408">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-408">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-409">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-409">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-410">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-410">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-411">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-411">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="a3dd2-412">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-412">(nullable) itemId :String</span></span>

<span data-ttu-id="a3dd2-p116">Obtiene el identificador de elemento de Servicios Web Exchange para el elemento actual. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-415">El identificador que devuelve la propiedad `itemId` es el mismo que el identificador de elemento de Servicios Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-415">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a3dd2-416">El `itemId` (propiedad) no es idéntico al identificador de entrada de Outlook o el identificador de la API de REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-416">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="a3dd2-417">Antes de realizar llamadas a la API de REST con este valor, se debe convertir mediante [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-417">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a3dd2-418">Para obtener más información, vea [usar las API de REST de Outlook desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-418">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="a3dd2-p118">La propiedad `itemId` no está disponible en el modo redacción. Si se requiere un identificador de elemento, el método [`saveAsync`](#saveasyncoptions-callback) puede usarse para guardar el elemento en el servidor, que devolverá el identificador de elemento en el parámetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) de la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-421">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-421">Type:</span></span>

*   <span data-ttu-id="a3dd2-422">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-422">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-423">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-423">Requirements</span></span>

|<span data-ttu-id="a3dd2-424">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-424">Requirement</span></span>|<span data-ttu-id="a3dd2-425">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-426">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-426">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-427">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-427">1.0</span></span>|
|[<span data-ttu-id="a3dd2-428">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-428">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-429">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-430">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-430">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-431">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-432">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-432">Example</span></span>

<span data-ttu-id="a3dd2-p119">El siguiente código comprueba la presencia de un identificador de elemento. Si la propiedad `itemId` devuelve `null` o `undefined`, guarda el elemento en el servidor y obtiene el identificador de elemento desde el resultado asincrónico.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="a3dd2-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="a3dd2-436">Obtiene el tipo de elemento que representa una instancia.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-436">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="a3dd2-437">La propiedad `itemType` devuelve uno de los valores de enumeración de `ItemType`, lo que indica si la instancia del objeto `item` es un mensaje o una cita.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-437">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-438">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-438">Type:</span></span>

*   [<span data-ttu-id="a3dd2-439">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="a3dd2-439">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="a3dd2-440">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-440">Requirements</span></span>

|<span data-ttu-id="a3dd2-441">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-441">Requirement</span></span>|<span data-ttu-id="a3dd2-442">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-443">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-443">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-444">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-444">1.0</span></span>|
|[<span data-ttu-id="a3dd2-445">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-446">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-447">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-448">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-449">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-449">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="a3dd2-450">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-450">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="a3dd2-451">Obtiene o establece la ubicación de una cita.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-451">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a3dd2-452">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-452">Read mode</span></span>

<span data-ttu-id="a3dd2-453">La propiedad `location` devuelve una cadena que contiene la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-453">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a3dd2-454">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-454">Compose mode</span></span>

<span data-ttu-id="a3dd2-455">La propiedad `location` devuelve un objeto `Location` que proporciona métodos para obtener y establecer la ubicación de la cita.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-455">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-456">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-456">Type:</span></span>

*   <span data-ttu-id="a3dd2-457">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-457">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-458">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-458">Requirements</span></span>

|<span data-ttu-id="a3dd2-459">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-459">Requirement</span></span>|<span data-ttu-id="a3dd2-460">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-460">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-461">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-461">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-462">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-462">1.0</span></span>|
|[<span data-ttu-id="a3dd2-463">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-463">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-464">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-464">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-465">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-465">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-466">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-466">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-467">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-467">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="a3dd2-468">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-468">normalizedSubject :String</span></span>

<span data-ttu-id="a3dd2-p120">Obtiene el asunto de un elemento, con todos los prefijos quitados (incluidos `RE:` y `FWD:`). Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="a3dd2-p121">La propiedad normalizedSubject obtiene el asunto del elemento, con los prefijos estándar (como `RE:` y `FW:`) agregados por los programas de correo electrónico. Para obtener el asunto del elemento con los prefijos intactos, use la propiedad [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-473">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-473">Type:</span></span>

*   <span data-ttu-id="a3dd2-474">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-474">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-475">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-475">Requirements</span></span>

|<span data-ttu-id="a3dd2-476">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-476">Requirement</span></span>|<span data-ttu-id="a3dd2-477">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-477">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-478">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-478">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-479">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-479">1.0</span></span>|
|[<span data-ttu-id="a3dd2-480">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-480">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-481">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-481">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-482">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-482">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-483">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-483">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-484">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-484">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="a3dd2-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="a3dd2-486">Obtiene los mensajes de notificación de un elemento.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-486">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-487">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-487">Type:</span></span>

*   [<span data-ttu-id="a3dd2-488">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="a3dd2-488">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="a3dd2-489">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-489">Requirements</span></span>

|<span data-ttu-id="a3dd2-490">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-490">Requirement</span></span>|<span data-ttu-id="a3dd2-491">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-491">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-492">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-492">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-493">1.3</span><span class="sxs-lookup"><span data-stu-id="a3dd2-493">1.3</span></span>|
|[<span data-ttu-id="a3dd2-494">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-494">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-495">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-495">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-496">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-496">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-497">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-497">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a3dd2-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a3dd2-499">Proporciona acceso a los asistentes opcionales de un evento.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-499">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="a3dd2-500">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-500">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a3dd2-501">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-501">Read mode</span></span>

<span data-ttu-id="a3dd2-502">La propiedad `optionalAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente opcional a la reunión.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-502">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a3dd2-503">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-503">Compose mode</span></span>

<span data-ttu-id="a3dd2-504">El `optionalAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes opcionales de una reunión.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-504">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-505">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-505">Type:</span></span>

*   <span data-ttu-id="a3dd2-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-507">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-507">Requirements</span></span>

|<span data-ttu-id="a3dd2-508">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-508">Requirement</span></span>|<span data-ttu-id="a3dd2-509">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-510">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-510">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-511">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-511">1.0</span></span>|
|[<span data-ttu-id="a3dd2-512">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-512">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-513">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-514">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-514">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-515">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-515">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-516">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-516">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="a3dd2-517">Organizador:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizador](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-517">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="a3dd2-518">Obtiene la dirección de correo electrónico del organizador de una reunión especificada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-518">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a3dd2-519">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-519">Read mode</span></span>

<span data-ttu-id="a3dd2-520">El `organizer` (propiedad) devuelve un objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa el organizador de la reunión.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-520">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a3dd2-521">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-521">Compose mode</span></span>

<span data-ttu-id="a3dd2-522">El `organizer` (propiedad) devuelve un objeto de [Organizador](/javascript/api/outlook/office.organizer) que proporciona un método para obtener el valor del organizador.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-522">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-523">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-523">Type:</span></span>

*   <span data-ttu-id="a3dd2-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizador](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-525">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-525">Requirements</span></span>

|<span data-ttu-id="a3dd2-526">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-526">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="a3dd2-527">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-527">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-528">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-528">1.0</span></span>|<span data-ttu-id="a3dd2-529">1.7</span><span class="sxs-lookup"><span data-stu-id="a3dd2-529">1.7</span></span>|
|[<span data-ttu-id="a3dd2-530">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-530">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-531">ReadItem</span></span>|<span data-ttu-id="a3dd2-532">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-532">ReadWriteItem</span></span>|
|[<span data-ttu-id="a3dd2-533">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-533">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-534">Read</span><span class="sxs-lookup"><span data-stu-id="a3dd2-534">Read</span></span>|<span data-ttu-id="a3dd2-535">Redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-536">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-536">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="a3dd2-537">periodicidad (nullable):[Periodicidad](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-537">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="a3dd2-538">Obtiene o establece el patrón de periodicidad de una cita.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-538">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="a3dd2-539">Obtiene el patrón de periodicidad de una convocatoria de reunión.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-539">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="a3dd2-540">Lectura y redacción modos para elementos de cita.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-540">Read and compose modes for appointment items.</span></span> <span data-ttu-id="a3dd2-541">Modo de lectura para los elementos de solicitud de reunión.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-541">Read mode for meeting request items.</span></span>

<span data-ttu-id="a3dd2-542">El `recurrence` (propiedad) devuelve un objeto de [Periodicidad](/javascript/api/outlook/office.recurrence) para solicitudes de reuniones o citas periódicas si un elemento es una serie o una instancia de una serie.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-542">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="a3dd2-543">`null`se devuelve solo citas y convocatorias de reunión de citas únicas.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-543">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="a3dd2-544">`undefined`se devuelve para los mensajes que no son las convocatorias de reunión.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-544">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="a3dd2-545">Nota: Las convocatorias de reunión tienen un `itemClass` valor de IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-545">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="a3dd2-546">Nota: Si el objeto de periodicidad es `null`, esto indica que el objeto es una cita única o una convocatoria de reunión de una cita única y no forma parte de una serie.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-546">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-547">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-547">Type:</span></span>

* [<span data-ttu-id="a3dd2-548">Periodicidad</span><span class="sxs-lookup"><span data-stu-id="a3dd2-548">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="a3dd2-549">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-549">Requirement</span></span>|<span data-ttu-id="a3dd2-550">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-551">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-551">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-552">1.7</span><span class="sxs-lookup"><span data-stu-id="a3dd2-552">1.7</span></span>|
|[<span data-ttu-id="a3dd2-553">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-554">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-555">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-556">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-556">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a3dd2-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a3dd2-558">Proporciona acceso a los asistentes necesarios de un evento.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-558">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="a3dd2-559">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-559">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a3dd2-560">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-560">Read mode</span></span>

<span data-ttu-id="a3dd2-561">La propiedad `requiredAttendees` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada asistente requerido a la reunión.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-561">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a3dd2-562">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-562">Compose mode</span></span>

<span data-ttu-id="a3dd2-563">El `requiredAttendees` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los asistentes necesarios de una reunión.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-563">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-564">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-564">Type:</span></span>

*   <span data-ttu-id="a3dd2-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-566">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-566">Requirements</span></span>

|<span data-ttu-id="a3dd2-567">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-567">Requirement</span></span>|<span data-ttu-id="a3dd2-568">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-568">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-569">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-569">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-570">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-570">1.0</span></span>|
|[<span data-ttu-id="a3dd2-571">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-571">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-572">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-572">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-573">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-573">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-574">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-574">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-575">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-575">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="a3dd2-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="a3dd2-p126">Obtiene la dirección de correo electrónico del remitente de un mensaje de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="a3dd2-p127">Las propiedades [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) y `sender` representan a la misma persona, a menos que un delegado envíe el mensaje. En ese caso, la propiedad `from` representa a la persona que delega y la propiedad sender representa al delegado.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-581">El `recipientType` (propiedad) de la `EmailAddressDetails` de objetos en el `sender` propiedad es `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-582">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-582">Type:</span></span>

*   [<span data-ttu-id="a3dd2-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a3dd2-583">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a3dd2-584">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-584">Requirements</span></span>

|<span data-ttu-id="a3dd2-585">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-585">Requirement</span></span>|<span data-ttu-id="a3dd2-586">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-587">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-587">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-588">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-588">1.0</span></span>|
|[<span data-ttu-id="a3dd2-589">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-589">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-590">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-591">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-591">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-592">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-593">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-593">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="a3dd2-594">seriesId (nullable): cadena</span><span class="sxs-lookup"><span data-stu-id="a3dd2-594">(nullable) seriesId :String</span></span>

<span data-ttu-id="a3dd2-595">Obtiene el identificador de la serie a la que pertenece una instancia.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="a3dd2-596">En OWA y Outlook, el `seriesId` devuelve el identificador de Exchange Web Services (EWS) del elemento primario (serie) al que pertenece este producto.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="a3dd2-597">Sin embargo, en iOS y Android, el `seriesId` devuelve el identificador del resto del elemento primario.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-598">El identificador que devuelve la propiedad `seriesId` es el mismo que el identificador de elemento de Servicios Web Exchange.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a3dd2-599">El `seriesId` propiedad no es idéntica a los identificadores de Outlook utilizadas por la API de REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="a3dd2-600">Antes de realizar llamadas a la API de REST con este valor, se debe convertir mediante [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a3dd2-601">Para obtener más información, vea [usar las API de REST de Outlook desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="a3dd2-602">El `seriesId` propiedad devuelve `null` para los elementos que no tienen elementos primarios, como único citas, elementos de serie, o solicitudes de reunión y devuelve `undefined` para todos los elementos que no son las convocatorias de reunión.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-603">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-603">Type:</span></span>

* <span data-ttu-id="a3dd2-604">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-605">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-605">Requirements</span></span>

|<span data-ttu-id="a3dd2-606">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-606">Requirement</span></span>|<span data-ttu-id="a3dd2-607">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-608">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-608">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-609">1.7</span><span class="sxs-lookup"><span data-stu-id="a3dd2-609">1.7</span></span>|
|[<span data-ttu-id="a3dd2-610">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-611">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-612">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-613">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-613">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-614">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-614">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="a3dd2-615">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-615">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="a3dd2-616">Obtiene o establece la fecha y la hora de inicio de la cita.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="a3dd2-p130">La propiedad `start` se expresa como un valor de fecha y hora de la hora universal coordinada (UTC). Puede usar el método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para convertir el valor a la fecha y hora local del cliente.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a3dd2-619">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-619">Read mode</span></span>

<span data-ttu-id="a3dd2-620">La propiedad `start` devuelve un objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-620">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a3dd2-621">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-621">Compose mode</span></span>

<span data-ttu-id="a3dd2-622">La propiedad `start` devuelve un objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="a3dd2-623">Si usa el método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para establecer la hora de inicio, use el método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para convertir la hora local del cliente en un valor UTC para el servidor.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-623">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-624">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-624">Type:</span></span>

*   <span data-ttu-id="a3dd2-625">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-625">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-626">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-626">Requirements</span></span>

|<span data-ttu-id="a3dd2-627">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-627">Requirement</span></span>|<span data-ttu-id="a3dd2-628">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-629">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-629">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-630">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-630">1.0</span></span>|
|[<span data-ttu-id="a3dd2-631">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-632">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-633">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-634">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-634">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-635">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-635">Example</span></span>

<span data-ttu-id="a3dd2-636">En el ejemplo siguiente, se establece la hora de inicio de una cita en el modo Redacción mediante el método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) del objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-636">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="a3dd2-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="a3dd2-638">Obtiene o establece la descripción que se muestra en el campo de asunto de un elemento.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-638">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="a3dd2-639">La propiedad `subject` obtiene o establece el asunto completo del elemento, como lo ha enviado el servidor de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-639">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a3dd2-640">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-640">Read mode</span></span>

<span data-ttu-id="a3dd2-p131">La propiedad `subject` devuelve una cadena. Use la propiedad [`normalizedSubject`](#normalizedsubject-string) para obtener el asunto menos cualquier prefijo inicial como `RE:` y `FW:`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="a3dd2-643">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-643">Compose mode</span></span>

<span data-ttu-id="a3dd2-644">La propiedad `subject` devuelve un objeto `Subject` que proporciona métodos para obtener y establecer el asunto.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a3dd2-645">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-645">Type:</span></span>

*   <span data-ttu-id="a3dd2-646">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-646">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-647">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-647">Requirements</span></span>

|<span data-ttu-id="a3dd2-648">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-648">Requirement</span></span>|<span data-ttu-id="a3dd2-649">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-650">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-650">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-651">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-651">1.0</span></span>|
|[<span data-ttu-id="a3dd2-652">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-653">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-654">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-655">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-655">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a3dd2-656">para: matriz. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[destinatarios](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a3dd2-657">Proporciona acceso a los destinatarios en la línea **para** de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="a3dd2-658">El tipo de objeto y el nivel de acceso dependen del modo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a3dd2-659">Modo de lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-659">Read mode</span></span>

<span data-ttu-id="a3dd2-p133">La propiedad `to` devuelve una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los destinatarios que aparecen en la línea **Para** del mensaje. La colección está limitada a un máximo de 100 miembros.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a3dd2-662">Modo de redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-662">Compose mode</span></span>

<span data-ttu-id="a3dd2-663">El `to` (propiedad) devuelve un `Recipients` objeto que proporciona métodos para obtener o actualizar los destinatarios en la línea **para** del mensaje.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a3dd2-664">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-664">Type:</span></span>

*   <span data-ttu-id="a3dd2-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-666">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-666">Requirements</span></span>

|<span data-ttu-id="a3dd2-667">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-667">Requirement</span></span>|<span data-ttu-id="a3dd2-668">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-669">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-669">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-670">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-670">1.0</span></span>|
|[<span data-ttu-id="a3dd2-671">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-672">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-673">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-674">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-674">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-675">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-675">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="a3dd2-676">Métodos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-676">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="a3dd2-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a3dd2-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a3dd2-678">Agrega un archivo a un mensaje o cita como datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-678">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a3dd2-679">El método `addFileAttachmentAsync` carga el archivo en el URI especificado y lo asocia al elemento en el formulario de redacción.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-679">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="a3dd2-680">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-680">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-681">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-681">Parameters:</span></span>
|<span data-ttu-id="a3dd2-682">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-682">Name</span></span>|<span data-ttu-id="a3dd2-683">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-683">Type</span></span>|<span data-ttu-id="a3dd2-684">Atributos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-684">Attributes</span></span>|<span data-ttu-id="a3dd2-685">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-685">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="a3dd2-686">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-686">String</span></span>||<span data-ttu-id="a3dd2-p134">El URI que proporciona la ubicación del archivo que se va a adjuntar al mensaje o a la cita. La longitud máxima es de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="a3dd2-689">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-689">String</span></span>||<span data-ttu-id="a3dd2-p135">El nombre de los datos adjuntos que se muestra mientras estos se cargan. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a3dd2-692">Object</span><span class="sxs-lookup"><span data-stu-id="a3dd2-692">Object</span></span>|<span data-ttu-id="a3dd2-693">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-693">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-694">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-694">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a3dd2-695">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-695">Object</span></span>|<span data-ttu-id="a3dd2-696">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-696">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-697">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-697">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="a3dd2-698">Booleano</span><span class="sxs-lookup"><span data-stu-id="a3dd2-698">Boolean</span></span>|<span data-ttu-id="a3dd2-699">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-699">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-700">Si está establecido en `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-700">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="a3dd2-701">Función</span><span class="sxs-lookup"><span data-stu-id="a3dd2-701">function</span></span>|<span data-ttu-id="a3dd2-702">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-702">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-703">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-703">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a3dd2-704">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-704">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a3dd2-705">Si se produce un error en la carga de los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-705">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a3dd2-706">Errores</span><span class="sxs-lookup"><span data-stu-id="a3dd2-706">Errors</span></span>

|<span data-ttu-id="a3dd2-707">Código de error</span><span class="sxs-lookup"><span data-stu-id="a3dd2-707">Error code</span></span>|<span data-ttu-id="a3dd2-708">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-708">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="a3dd2-709">Los datos adjuntos son más grandes de lo permitido.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-709">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="a3dd2-710">Los datos adjuntos tienen una extensión que no está permitida.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-710">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a3dd2-711">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-711">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-712">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-712">Requirements</span></span>

|<span data-ttu-id="a3dd2-713">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-713">Requirement</span></span>|<span data-ttu-id="a3dd2-714">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-715">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-715">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-716">1.1</span><span class="sxs-lookup"><span data-stu-id="a3dd2-716">1.1</span></span>|
|[<span data-ttu-id="a3dd2-717">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-717">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-718">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-718">ReadWriteItem</span></span>|
|[<span data-ttu-id="a3dd2-719">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-719">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-720">Redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-720">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a3dd2-721">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-721">Examples</span></span>

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

<span data-ttu-id="a3dd2-722">En el siguiente ejemplo, se agrega un archivo de imagen como un dato adjunto en línea y hace referencia a los datos adjuntos del cuerpo del mensaje.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-722">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="a3dd2-723">addFileAttachmentFromBase64Async (base64File, attachmentName, [opciones], [devolución de llamada])</span><span class="sxs-lookup"><span data-stu-id="a3dd2-723">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a3dd2-724">Agrega un archivo desde el base64 de codificación para un mensaje o cita como datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-724">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a3dd2-725">El `addFileAttachmentFromBase64Async` método carga el archivo de la codificación base64 y se adjunta al elemento en el formulario de redacción.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-725">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="a3dd2-726">Este método devuelve el identificador de datos adjuntos en el objeto AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-726">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="a3dd2-727">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-727">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-728">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-728">Parameters:</span></span>
|<span data-ttu-id="a3dd2-729">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-729">Name</span></span>|<span data-ttu-id="a3dd2-730">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-730">Type</span></span>|<span data-ttu-id="a3dd2-731">Atributos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-731">Attributes</span></span>|<span data-ttu-id="a3dd2-732">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-732">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="a3dd2-733">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-733">String</span></span>||<span data-ttu-id="a3dd2-734">La base64 codificado contenido de una imagen o un archivo que se agregarán a un correo electrónico o un evento.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-734">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="a3dd2-735">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-735">String</span></span>||<span data-ttu-id="a3dd2-p137">El nombre de los datos adjuntos que se muestra mientras estos se cargan. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a3dd2-738">Object</span><span class="sxs-lookup"><span data-stu-id="a3dd2-738">Object</span></span>|<span data-ttu-id="a3dd2-739">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-739">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-740">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a3dd2-741">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-741">Object</span></span>|<span data-ttu-id="a3dd2-742">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-742">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-743">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="a3dd2-744">Booleano</span><span class="sxs-lookup"><span data-stu-id="a3dd2-744">Boolean</span></span>|<span data-ttu-id="a3dd2-745">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-745">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-746">Si está establecido en `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="a3dd2-747">Función</span><span class="sxs-lookup"><span data-stu-id="a3dd2-747">function</span></span>|<span data-ttu-id="a3dd2-748">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-748">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-749">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a3dd2-750">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a3dd2-751">Si se produce un error en la carga de los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a3dd2-752">Errores</span><span class="sxs-lookup"><span data-stu-id="a3dd2-752">Errors</span></span>

|<span data-ttu-id="a3dd2-753">Código de error</span><span class="sxs-lookup"><span data-stu-id="a3dd2-753">Error code</span></span>|<span data-ttu-id="a3dd2-754">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="a3dd2-755">Los datos adjuntos son más grandes de lo permitido.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="a3dd2-756">Los datos adjuntos tienen una extensión que no está permitida.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a3dd2-757">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-758">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-758">Requirements</span></span>

|<span data-ttu-id="a3dd2-759">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-759">Requirement</span></span>|<span data-ttu-id="a3dd2-760">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-761">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-761">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-762">Vista previa</span><span class="sxs-lookup"><span data-stu-id="a3dd2-762">Preview</span></span>|
|[<span data-ttu-id="a3dd2-763">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="a3dd2-765">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-766">Redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a3dd2-767">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-767">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="a3dd2-768">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a3dd2-768">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="a3dd2-769">Agrega un controlador de eventos para un evento admitido.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-769">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="a3dd2-770">Actualmente son los tipos de eventos compatibles `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, y`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="a3dd2-770">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-771">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-771">Parameters:</span></span>

| <span data-ttu-id="a3dd2-772">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-772">Name</span></span> | <span data-ttu-id="a3dd2-773">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-773">Type</span></span> | <span data-ttu-id="a3dd2-774">Atributos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-774">Attributes</span></span> | <span data-ttu-id="a3dd2-775">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-775">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a3dd2-776">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a3dd2-776">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a3dd2-777">El evento que debe invocar el controlador.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-777">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="a3dd2-778">Función</span><span class="sxs-lookup"><span data-stu-id="a3dd2-778">Function</span></span> || <span data-ttu-id="a3dd2-p138">La función que va a controlar el evento. La función debe aceptar un único parámetro, que es un literal de objeto. La propiedad `type` del parámetro coincidirá con el parámetro `eventType` que se ha pasado a `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="a3dd2-782">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-782">Object</span></span> | <span data-ttu-id="a3dd2-783">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-783">&lt;optional&gt;</span></span> | <span data-ttu-id="a3dd2-784">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-784">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a3dd2-785">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-785">Object</span></span> | <span data-ttu-id="a3dd2-786">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-786">&lt;optional&gt;</span></span> | <span data-ttu-id="a3dd2-787">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-787">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a3dd2-788">función</span><span class="sxs-lookup"><span data-stu-id="a3dd2-788">function</span></span>| <span data-ttu-id="a3dd2-789">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-789">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-790">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-790">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-791">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-791">Requirements</span></span>

|<span data-ttu-id="a3dd2-792">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-792">Requirement</span></span>| <span data-ttu-id="a3dd2-793">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-793">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-794">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-794">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a3dd2-795">1.7</span><span class="sxs-lookup"><span data-stu-id="a3dd2-795">1.7</span></span> |
|[<span data-ttu-id="a3dd2-796">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-796">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a3dd2-797">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-797">ReadItem</span></span> |
|[<span data-ttu-id="a3dd2-798">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-798">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a3dd2-799">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-799">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="a3dd2-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a3dd2-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a3dd2-801">Agrega un elemento de Exchange (por ejemplo, un mensaje) como datos adjuntos al mensaje o a la cita.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-801">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="a3dd2-p139">El método `addItemAttachmentAsync` asocia el elemento con el identificador especificado de Exchange al elemento en el formulario de redacción. Si se especifica un método de devolución de llamada, se llama al método con un parámetro, `asyncResult`, que contiene el identificador de datos adjuntos o un código que indique cualquier error que se ha producido al adjuntar el elemento. Si es necesario, puede usar el parámetro `options` para pasar información de estado al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="a3dd2-805">Después, puede usar el identificador con el método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para quitar los datos adjuntos en la misma sesión.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-805">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="a3dd2-806">Si el complemento de Office se está ejecutando en Outlook Web App, el `addItemAttachmentAsync` método puede adjuntar elementos a los elementos que no sea el elemento que se está editando; Sin embargo, esto no se admite y no se recomienda.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-806">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-807">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-807">Parameters:</span></span>

|<span data-ttu-id="a3dd2-808">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-808">Name</span></span>|<span data-ttu-id="a3dd2-809">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-809">Type</span></span>|<span data-ttu-id="a3dd2-810">Atributos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-810">Attributes</span></span>|<span data-ttu-id="a3dd2-811">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-811">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="a3dd2-812">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-812">String</span></span>||<span data-ttu-id="a3dd2-p140">El identificador de Exchange del elemento que debe adjuntarse. La longitud máxima es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="a3dd2-815">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-815">String</span></span>||<span data-ttu-id="a3dd2-p141">El asunto del elemento que debe adjuntarse. La longitud máxima es de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a3dd2-818">Object</span><span class="sxs-lookup"><span data-stu-id="a3dd2-818">Object</span></span>|<span data-ttu-id="a3dd2-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-819">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-820">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-820">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a3dd2-821">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-821">Object</span></span>|<span data-ttu-id="a3dd2-822">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-822">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-823">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-823">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a3dd2-824">función</span><span class="sxs-lookup"><span data-stu-id="a3dd2-824">function</span></span>|<span data-ttu-id="a3dd2-825">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-825">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-826">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-826">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a3dd2-827">Si se realiza correctamente, se proporcionará el identificador de los datos adjuntos en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-827">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a3dd2-828">Si se produce un error al agregar los datos adjuntos, el objeto `asyncResult` contendrá un objeto `Error` que proporciona una descripción del error.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-828">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a3dd2-829">Errores</span><span class="sxs-lookup"><span data-stu-id="a3dd2-829">Errors</span></span>

|<span data-ttu-id="a3dd2-830">Código de error</span><span class="sxs-lookup"><span data-stu-id="a3dd2-830">Error code</span></span>|<span data-ttu-id="a3dd2-831">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-831">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a3dd2-832">El mensaje o cita tiene demasiados datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-832">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-833">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-833">Requirements</span></span>

|<span data-ttu-id="a3dd2-834">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-834">Requirement</span></span>|<span data-ttu-id="a3dd2-835">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-836">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-836">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-837">1.1</span><span class="sxs-lookup"><span data-stu-id="a3dd2-837">1.1</span></span>|
|[<span data-ttu-id="a3dd2-838">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-839">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-839">ReadWriteItem</span></span>|
|[<span data-ttu-id="a3dd2-840">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-841">Redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-841">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-842">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-842">Example</span></span>

<span data-ttu-id="a3dd2-843">En el ejemplo siguiente, se agrega un elemento existente de Outlook como un archivo adjunto con el nombre `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-843">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="a3dd2-844">close()</span><span class="sxs-lookup"><span data-stu-id="a3dd2-844">close()</span></span>

<span data-ttu-id="a3dd2-845">Cierra el elemento actual que se está redactando.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-845">Closes the current item that is being composed.</span></span>

<span data-ttu-id="a3dd2-p142">El comportamiento del método `close` depende del estado actual del elemento que se está redactando. Si el elemento tiene cambios sin guardar, el cliente solicita al usuario que guarde, descarte o cancele la acción Cerrar.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-848">En Outlook, en el sitio web, si el elemento es una cita y anteriormente se ha guardado con `saveAsync`, se pide al usuario al guardar, descartar o cancelar incluso si no ha habido cambios desde la último el elemento guardado.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-848">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="a3dd2-849">En el cliente de escritorio de Outlook, si el mensaje es una respuesta directa, el método `close` no tiene ningún efecto.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-849">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-850">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-850">Requirements</span></span>

|<span data-ttu-id="a3dd2-851">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-851">Requirement</span></span>|<span data-ttu-id="a3dd2-852">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-852">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-853">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-853">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-854">1.3</span><span class="sxs-lookup"><span data-stu-id="a3dd2-854">1.3</span></span>|
|[<span data-ttu-id="a3dd2-855">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-855">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-856">Restringido</span><span class="sxs-lookup"><span data-stu-id="a3dd2-856">Restricted</span></span>|
|[<span data-ttu-id="a3dd2-857">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-857">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-858">Redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-858">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="a3dd2-859">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-859">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="a3dd2-860">Muestra un formulario de respuesta que incluye el remitente y todos los destinatarios del mensaje seleccionado o el organizador y todos los asistentes de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-860">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-861">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-861">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a3dd2-862">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-862">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a3dd2-863">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyAllForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-863">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="a3dd2-p143">Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-867">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-867">Parameters:</span></span>

|<span data-ttu-id="a3dd2-868">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-868">Name</span></span>|<span data-ttu-id="a3dd2-869">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-869">Type</span></span>|<span data-ttu-id="a3dd2-870">Atributos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-870">Attributes</span></span>|<span data-ttu-id="a3dd2-871">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-871">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="a3dd2-872">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a3dd2-872">String &#124; Object</span></span>||<span data-ttu-id="a3dd2-p144">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a3dd2-875">**O**</span><span class="sxs-lookup"><span data-stu-id="a3dd2-875">**OR**</span></span><br/><span data-ttu-id="a3dd2-p145">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="a3dd2-878">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-878">String</span></span>|<span data-ttu-id="a3dd2-879">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-879">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-p146">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="a3dd2-882">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-882">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="a3dd2-883">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-883">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-884">Una matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-884">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="a3dd2-885">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-885">String</span></span>||<span data-ttu-id="a3dd2-p147">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="a3dd2-888">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-888">String</span></span>||<span data-ttu-id="a3dd2-889">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-889">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="a3dd2-890">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-890">String</span></span>||<span data-ttu-id="a3dd2-p148">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="a3dd2-893">Booleano</span><span class="sxs-lookup"><span data-stu-id="a3dd2-893">Boolean</span></span>||<span data-ttu-id="a3dd2-p149">Solo se usa si `type` se establece en `file`. Si `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="a3dd2-896">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-896">String</span></span>||<span data-ttu-id="a3dd2-p150">Solo se usa si `type` se establece en `item`. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="a3dd2-900">función</span><span class="sxs-lookup"><span data-stu-id="a3dd2-900">function</span></span>|<span data-ttu-id="a3dd2-901">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-901">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-902">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-902">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-903">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-903">Requirements</span></span>

|<span data-ttu-id="a3dd2-904">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-904">Requirement</span></span>|<span data-ttu-id="a3dd2-905">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-906">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-906">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-907">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-907">1.0</span></span>|
|[<span data-ttu-id="a3dd2-908">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-909">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-910">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-911">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-911">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a3dd2-912">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-912">Examples</span></span>

<span data-ttu-id="a3dd2-913">El código siguiente pasa una cadena a la función `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-913">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="a3dd2-914">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-914">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="a3dd2-915">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-915">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a3dd2-916">Responder con un cuerpo y datos adjuntos de archivo.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-916">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a3dd2-917">Responder con un cuerpo y datos adjuntos de elemento.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-917">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a3dd2-918">Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-918">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="a3dd2-919">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-919">displayReplyForm(formData)</span></span>

<span data-ttu-id="a3dd2-920">Muestra un formulario de respuesta que incluye solo el remitente del mensaje seleccionado o el organizador de la cita seleccionada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-920">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-921">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a3dd2-922">En Outlook Web App, el formulario de respuestas aparece como formulario emergente en la vista de tres columnas y en la vista de una o de dos columnas.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-922">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a3dd2-923">Si cualquiera de los parámetros de cadena supera sus límites, `displayReplyForm` produce una excepción.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-923">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="a3dd2-p151">Cuando se especifican los datos adjuntos en el parámetro `formData.attachments`, Outlook y Outlook Web App intentan descargar todos los datos adjuntos y adjuntarlos al formulario de respuesta. Si se produce un error al agregar los datos adjuntos, se muestra un error en la interfaz de usuario del formulario. Si esto no es posible, no se produce ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-927">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-927">Parameters:</span></span>

|<span data-ttu-id="a3dd2-928">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-928">Name</span></span>|<span data-ttu-id="a3dd2-929">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-929">Type</span></span>|<span data-ttu-id="a3dd2-930">Atributos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-930">Attributes</span></span>|<span data-ttu-id="a3dd2-931">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-931">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="a3dd2-932">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a3dd2-932">String &#124; Object</span></span>||<span data-ttu-id="a3dd2-p152">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a3dd2-935">**O**</span><span class="sxs-lookup"><span data-stu-id="a3dd2-935">**OR**</span></span><br/><span data-ttu-id="a3dd2-p153">Un objeto que contiene datos del cuerpo o datos adjuntos y una función de devolución de llamada. El objeto se define de la manera siguiente.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="a3dd2-938">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-938">String</span></span>|<span data-ttu-id="a3dd2-939">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-939">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-p154">Una cadena que contiene texto y HTML y que representa el cuerpo del formulario de respuestas. La cadena puede tener un máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="a3dd2-942">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-942">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="a3dd2-943">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-943">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-944">Una matriz de objetos JSON que son elementos o archivos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-944">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="a3dd2-945">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-945">String</span></span>||<span data-ttu-id="a3dd2-p155">Indica el tipo de datos adjuntos. Debe ser `file` para un archivo adjunto o `item` para un elemento adjunto.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="a3dd2-948">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-948">String</span></span>||<span data-ttu-id="a3dd2-949">Una cadena que contiene el nombre de los datos adjuntos, hasta 255 caracteres de longitud.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-949">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="a3dd2-950">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-950">String</span></span>||<span data-ttu-id="a3dd2-p156">Solo se usa si `type` se establece en `file`. El URI de la ubicación del archivo.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="a3dd2-953">Booleano</span><span class="sxs-lookup"><span data-stu-id="a3dd2-953">Boolean</span></span>||<span data-ttu-id="a3dd2-p157">Solo se usa si `type` se establece en `file`. Si `true`, indica que los datos adjuntos se mostrarán en línea en el cuerpo del mensaje, y no deben mostrarse en la lista de datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="a3dd2-956">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-956">String</span></span>||<span data-ttu-id="a3dd2-p158">Solo se usa si `type` se establece en `item`. El identificador de elemento EWS de los datos adjuntos. Se trata de una cadena de hasta 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="a3dd2-960">función</span><span class="sxs-lookup"><span data-stu-id="a3dd2-960">function</span></span>|<span data-ttu-id="a3dd2-961">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-961">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-962">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-962">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-963">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-963">Requirements</span></span>

|<span data-ttu-id="a3dd2-964">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-964">Requirement</span></span>|<span data-ttu-id="a3dd2-965">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-966">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-966">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-967">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-967">1.0</span></span>|
|[<span data-ttu-id="a3dd2-968">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-968">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-969">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-970">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-970">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-971">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-971">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a3dd2-972">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-972">Examples</span></span>

<span data-ttu-id="a3dd2-973">El código siguiente pasa una cadena a la función `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-973">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="a3dd2-974">Responder con un cuerpo vacío.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-974">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="a3dd2-975">Responder solo con un cuerpo.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-975">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a3dd2-976">Responder con un cuerpo y datos adjuntos de archivo.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-976">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a3dd2-977">Responder con un cuerpo y datos adjuntos de elemento.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-977">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a3dd2-978">Responder con un cuerpo, datos adjuntos de archivo, datos adjuntos de elemento y una devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-978">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="a3dd2-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a3dd2-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="a3dd2-980">Obtiene las entidades que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-980">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-981">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-981">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-982">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-982">Requirements</span></span>

|<span data-ttu-id="a3dd2-983">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-983">Requirement</span></span>|<span data-ttu-id="a3dd2-984">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-984">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-985">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-985">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-986">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-986">1.0</span></span>|
|[<span data-ttu-id="a3dd2-987">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-987">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-988">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-988">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-989">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-989">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-990">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-990">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a3dd2-991">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-991">Returns:</span></span>

<span data-ttu-id="a3dd2-992">Tipo: [Entidades](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-992">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a3dd2-993">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-993">Example</span></span>

<span data-ttu-id="a3dd2-994">En el ejemplo siguiente se obtiene acceso a las entidades de los contactos en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-994">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="a3dd2-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a3dd2-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a3dd2-996">Obtiene una matriz de todas las entidades del tipo de entidad especificado que se encuentran en el cuerpo del elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-996">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-997">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-997">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-998">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-998">Parameters:</span></span>

|<span data-ttu-id="a3dd2-999">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-999">Name</span></span>|<span data-ttu-id="a3dd2-1000">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1000">Type</span></span>|<span data-ttu-id="a3dd2-1001">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1001">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="a3dd2-1002">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1002">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="a3dd2-1003">Uno de los valores de enumeración de EntityType.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1003">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-1004">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1004">Requirements</span></span>

|<span data-ttu-id="a3dd2-1005">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1005">Requirement</span></span>|<span data-ttu-id="a3dd2-1006">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1006">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-1007">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1007">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-1008">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1008">1.0</span></span>|
|[<span data-ttu-id="a3dd2-1009">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1009">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-1010">Restringido</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1010">Restricted</span></span>|
|[<span data-ttu-id="a3dd2-1011">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1011">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-1012">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1012">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a3dd2-1013">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1013">Returns:</span></span>

<span data-ttu-id="a3dd2-1014">Si el valor que se pasa a `entityType` no es un miembro válido de la enumeración de `EntityType`, el método devuelve un valor NULL.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1014">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="a3dd2-1015">Si no existen entidades del tipo especificado están presentes en el cuerpo del elemento, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1015">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="a3dd2-1016">De otro modo, el tipo de los objetos en la matriz devuelta depende del tipo de entidad solicitada en el parámetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1016">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="a3dd2-1017">Aunque el nivel de permiso mínimo para usar este método es **Restricted**, algunos tipos de entidad requieren **ReadItem** para tener acceso, como se especifica en la tabla siguiente.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1017">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="a3dd2-1018">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1018">Value of `entityType`</span></span>|<span data-ttu-id="a3dd2-1019">Tipo de objetos de la matriz devuelta</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1019">Type of objects in returned array</span></span>|<span data-ttu-id="a3dd2-1020">Nivel de permiso requerido</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1020">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="a3dd2-1021">Cadena</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1021">String</span></span>|<span data-ttu-id="a3dd2-1022">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1022">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="a3dd2-1023">Contacto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1023">Contact</span></span>|<span data-ttu-id="a3dd2-1024">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1024">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="a3dd2-1025">Cadena</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1025">String</span></span>|<span data-ttu-id="a3dd2-1026">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1026">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="a3dd2-1027">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1027">MeetingSuggestion</span></span>|<span data-ttu-id="a3dd2-1028">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1028">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="a3dd2-1029">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1029">PhoneNumber</span></span>|<span data-ttu-id="a3dd2-1030">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1030">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="a3dd2-1031">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1031">TaskSuggestion</span></span>|<span data-ttu-id="a3dd2-1032">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1032">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="a3dd2-1033">Cadena</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1033">String</span></span>|<span data-ttu-id="a3dd2-1034">**Restringido**</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1034">**Restricted**</span></span>|

<span data-ttu-id="a3dd2-1035">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a3dd2-1035">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="a3dd2-1036">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1036">Example</span></span>

<span data-ttu-id="a3dd2-1037">En el ejemplo siguiente se muestra cómo obtener acceso a una matriz de cadenas que representan direcciones postales en el cuerpo del elemento actual.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1037">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="a3dd2-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a3dd2-1039">Devuelve entidades conocidas en el elemento seleccionado que pasan el filtro con nombre definido en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1039">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-1040">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1040">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a3dd2-1041">El método `getFilteredEntitiesByName` devuelve las entidades que coinciden con la expresión regular definida en el elemento de regla [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) del archivo XML de manifiesto con el valor de elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1041">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-1042">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1042">Parameters:</span></span>

|<span data-ttu-id="a3dd2-1043">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1043">Name</span></span>|<span data-ttu-id="a3dd2-1044">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1044">Type</span></span>|<span data-ttu-id="a3dd2-1045">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1045">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="a3dd2-1046">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1046">String</span></span>|<span data-ttu-id="a3dd2-1047">El nombre del elemento de regla `ItemHasKnownEntity` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1047">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-1048">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1048">Requirements</span></span>

|<span data-ttu-id="a3dd2-1049">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1049">Requirement</span></span>|<span data-ttu-id="a3dd2-1050">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-1051">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1051">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-1052">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1052">1.0</span></span>|
|[<span data-ttu-id="a3dd2-1053">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1053">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1054">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-1055">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1055">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-1056">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1056">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a3dd2-1057">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1057">Returns:</span></span>

<span data-ttu-id="a3dd2-p160">Si no existe ningún elemento `ItemHasKnownEntity` en el manifiesto con un valor de elemento `FilterName` que coincida con el parámetro `name`, el método devuelve `null`. Si el parámetro `name` no coincide con un elemento `ItemHasKnownEntity` del manifiesto, pero no existen entidades en el elemento actual que coincidan, el método devuelve una matriz vacía.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="a3dd2-1060">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a3dd2-1060">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="a3dd2-1061">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1061">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="a3dd2-1062">Obtiene datos de inicialización que se pasan cuando [un mensaje accionable activa](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message) el complemento.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1062">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-1063">Este método sólo se admite por Outlook 2016 o posterior para Windows (posterior a 16.0.8413.1000 Click-to-Run versiones) y Outlook en la web para Office 365.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1063">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-1064">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1064">Parameters:</span></span>
|<span data-ttu-id="a3dd2-1065">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1065">Name</span></span>|<span data-ttu-id="a3dd2-1066">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1066">Type</span></span>|<span data-ttu-id="a3dd2-1067">Atributos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1067">Attributes</span></span>|<span data-ttu-id="a3dd2-1068">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1068">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a3dd2-1069">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1069">Object</span></span>|<span data-ttu-id="a3dd2-1070">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1070">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1071">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1071">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a3dd2-1072">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1072">Object</span></span>|<span data-ttu-id="a3dd2-1073">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1073">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1074">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1074">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a3dd2-1075">función</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1075">function</span></span>|<span data-ttu-id="a3dd2-1076">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1076">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1077">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1077">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a3dd2-1078">En caso de éxito, los datos de inicialización se proporcionan en el `asyncResult.value` (propiedad) como una cadena.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1078">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="a3dd2-1079">Si no hay ningún contexto de inicialización, el objeto `asyncResult` contendrá un objeto `Error` con la propiedad `code` establecida en `9020` y la propiedad `name` establecida en `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1079">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-1080">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1080">Requirements</span></span>

|<span data-ttu-id="a3dd2-1081">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1081">Requirement</span></span>|<span data-ttu-id="a3dd2-1082">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-1083">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1083">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-1084">Vista previa</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1084">Preview</span></span>|
|[<span data-ttu-id="a3dd2-1085">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1085">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1086">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-1087">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1087">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-1088">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1088">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-1089">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1089">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="a3dd2-1090">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1090">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="a3dd2-1091">Devuelve valores de cadena en el elemento seleccionado que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1091">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-1092">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1092">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a3dd2-p161">El método `getRegExMatches` devuelve las cadenas que coinciden con la expresión regular definida en cada elemento de regla `ItemHasRegularExpressionMatch` o `ItemHasKnownEntity` del archivo XML de manifiesto. Para una regla `ItemHasRegularExpressionMatch`, tiene que darse una cadena coincidente en la propiedad del elemento especificado por la regla. El tipo simple `PropertyName` define las propiedades admitidas.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a3dd2-1096">Por ejemplo, considere un manifiesto de complemento que tenga el siguiente elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1096">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a3dd2-1097">El objeto devuelto desde `getRegExMatches` tendría dos propiedades: `fruits` y `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1097">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a3dd2-p162">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados. En su lugar, use el método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo el cuerpo.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-1101">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1101">Requirements</span></span>

|<span data-ttu-id="a3dd2-1102">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1102">Requirement</span></span>|<span data-ttu-id="a3dd2-1103">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1103">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-1104">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1104">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-1105">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1105">1.0</span></span>|
|[<span data-ttu-id="a3dd2-1106">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1106">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-1107">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1107">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-1108">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1108">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-1109">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1109">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a3dd2-1110">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1110">Returns:</span></span>

<span data-ttu-id="a3dd2-p163">Un objeto que contiene matrices de cadenas que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. El nombre de cada matriz es igual al valor correspondiente del atributo `RegExName` de la regla `ItemHasRegularExpressionMatch` coincidente o al atributo `FilterName` de la regla `ItemHasKnownEntity` coincidente.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="a3dd2-1113">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1113">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a3dd2-1114">Object</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1114">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a3dd2-1115">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1115">Example</span></span>

<span data-ttu-id="a3dd2-1116">En el siguiente ejemplo, se muestra cómo tener acceso a la matriz de coincidencias de los elementos de regla de expresión regular `fruits` y `veggies`, que se especifican en el manifiesto.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1116">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="a3dd2-1117">→ de getregexmatchesbyname (Name) (nullable) {matriz. < cadena >}</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1117">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="a3dd2-1118">Devuelve valores de cadena en el elemento seleccionado que coinciden con la expresión regular de nombre definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1118">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-1119">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1119">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a3dd2-1120">El método `getRegExMatchesByName` devuelve las cadenas que coinciden con la expresión regular definida en el elemento de regla `ItemHasRegularExpressionMatch` del archivo XML de manifiesto con el valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1120">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="a3dd2-p164">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-1123">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1123">Parameters:</span></span>

|<span data-ttu-id="a3dd2-1124">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1124">Name</span></span>|<span data-ttu-id="a3dd2-1125">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1125">Type</span></span>|<span data-ttu-id="a3dd2-1126">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1126">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="a3dd2-1127">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1127">String</span></span>|<span data-ttu-id="a3dd2-1128">El nombre del elemento de regla `ItemHasRegularExpressionMatch` que define el filtro para que coincidan.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1128">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-1129">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1129">Requirements</span></span>

|<span data-ttu-id="a3dd2-1130">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1130">Requirement</span></span>|<span data-ttu-id="a3dd2-1131">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1131">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-1132">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-1133">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1133">1.0</span></span>|
|[<span data-ttu-id="a3dd2-1134">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-1135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1135">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-1136">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-1137">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1137">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a3dd2-1138">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1138">Returns:</span></span>

<span data-ttu-id="a3dd2-1139">Una matriz que contiene las cadenas que coinciden con la expresión regular definida en el archivo XML de manifiesto.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1139">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="a3dd2-1140">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1140">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a3dd2-1141">< Cadena > de matriz.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1141">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a3dd2-1142">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1142">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="a3dd2-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="a3dd2-1144">Devuelve asincrónicamente datos seleccionados desde el asunto o el cuerpo de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1144">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="a3dd2-p165">Si no hay ninguna selección, pero el cursor está en el cuerpo o el asunto, el método devuelve null para los datos seleccionados. Si se selecciona un campo que no sea el cuerpo o el asunto, el método devuelve el error `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-1147">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1147">Parameters:</span></span>

|<span data-ttu-id="a3dd2-1148">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1148">Name</span></span>|<span data-ttu-id="a3dd2-1149">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1149">Type</span></span>|<span data-ttu-id="a3dd2-1150">Atributos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1150">Attributes</span></span>|<span data-ttu-id="a3dd2-1151">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1151">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="a3dd2-1152">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1152">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="a3dd2-p166">Solicita un formato para los datos. Si es Texto, el método devuelve el texto sin formato como cadena, quitando toda etiqueta HTML que hubiera. Si es HTML, el método devuelve el texto seleccionado, ya sea texto sin formato o HTML.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="a3dd2-1156">Object</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1156">Object</span></span>|<span data-ttu-id="a3dd2-1157">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1158">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1158">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a3dd2-1159">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1159">Object</span></span>|<span data-ttu-id="a3dd2-1160">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1161">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1161">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a3dd2-1162">función</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1162">function</span></span>||<span data-ttu-id="a3dd2-1163">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1163">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a3dd2-1164">Para tener acceso a los datos seleccionados desde el método de devolución de llamada, llame a `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1164">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="a3dd2-1165">Para obtener acceso a la propiedad de origen del que proviene la selección, llame a `asyncResult.value.sourceProperty`, que será cualquiera `body` o `subject`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1165">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-1166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1166">Requirements</span></span>

|<span data-ttu-id="a3dd2-1167">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1167">Requirement</span></span>|<span data-ttu-id="a3dd2-1168">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1168">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-1169">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1169">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-1170">1.2</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1170">1.2</span></span>|
|[<span data-ttu-id="a3dd2-1171">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1171">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-1172">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1172">ReadWriteItem</span></span>|
|[<span data-ttu-id="a3dd2-1173">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1173">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-1174">Redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1174">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="a3dd2-1175">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1175">Returns:</span></span>

<span data-ttu-id="a3dd2-1176">Los datos seleccionados como cadena con formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1176">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="a3dd2-1177">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1177">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a3dd2-1178">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1178">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a3dd2-1179">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1179">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="a3dd2-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="a3dd2-p168">Obtiene las entidades que se han detectado en una coincidencia resaltada que un usuario ha seleccionado. Las coincidencias resaltadas se aplican a [complementos contextuales](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-1183">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1183">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-1184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1184">Requirements</span></span>

|<span data-ttu-id="a3dd2-1185">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1185">Requirement</span></span>|<span data-ttu-id="a3dd2-1186">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1186">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-1187">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1187">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-1188">1.6</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1188">1.6</span></span>|
|[<span data-ttu-id="a3dd2-1189">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1189">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-1190">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1190">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-1191">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-1192">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1192">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a3dd2-1193">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1193">Returns:</span></span>

<span data-ttu-id="a3dd2-1194">Tipo: [Entidades](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1194">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a3dd2-1195">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1195">Example</span></span>

<span data-ttu-id="a3dd2-1196">En el siguiente ejemplo se accede a las entidades de direcciones en la coincidencia resaltada seleccionada por el usuario.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1196">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="a3dd2-1197">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1197">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="a3dd2-p169">Devuelve valores de cadena en una coincidencia resaltada que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. Las coincidencias resaltadas se aplican a [complementos contextuales](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-1200">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1200">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a3dd2-p170">El método `getSelectedRegExMatches` devuelve las cadenas que coinciden con la expresión regular definida en cada elemento de regla `ItemHasRegularExpressionMatch` o `ItemHasKnownEntity` del archivo XML de manifiesto. Para una regla `ItemHasRegularExpressionMatch`, tiene que darse una cadena coincidente en la propiedad del elemento especificado por la regla. El tipo simple `PropertyName` define las propiedades admitidas.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a3dd2-1204">Por ejemplo, considere un manifiesto de complemento que tenga el siguiente elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1204">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a3dd2-1205">El objeto devuelto desde `getRegExMatches` tendría dos propiedades: `fruits` y `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1205">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a3dd2-p171">Si especifica una regla `ItemHasRegularExpressionMatch` en la propiedad body de un elemento, la expresión regular debe filtrar aún más el cuerpo y no debe intentar devolver todo el cuerpo del elemento. Usar una expresión regular como `.*` para obtener todo el cuerpo de un elemento no siempre devuelve los resultados esperados. En su lugar, use el método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo el cuerpo.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3dd2-1209">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1209">Requirements</span></span>

|<span data-ttu-id="a3dd2-1210">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1210">Requirement</span></span>|<span data-ttu-id="a3dd2-1211">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1211">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-1212">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1212">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-1213">1.6</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1213">1.6</span></span>|
|[<span data-ttu-id="a3dd2-1214">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1214">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-1215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1215">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-1216">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-1217">Lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1217">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a3dd2-1218">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1218">Returns:</span></span>

<span data-ttu-id="a3dd2-p172">Un objeto que contiene matrices de cadenas que coinciden con las expresiones regulares definidas en el archivo XML de manifiesto. El nombre de cada matriz es igual al valor correspondiente del atributo `RegExName` de la regla `ItemHasRegularExpressionMatch` coincidente o al atributo `FilterName` de la regla `ItemHasKnownEntity` coincidente.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="a3dd2-1221">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1221">Example</span></span>

<span data-ttu-id="a3dd2-1222">En el siguiente ejemplo, se muestra cómo tener acceso a la matriz de coincidencias de los elementos de regla de expresión regular `fruits` y `veggies`, que se especifican en el manifiesto.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1222">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="a3dd2-1223">getSharedPropertiesAsync ([opciones], devolución de llamada)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1223">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="a3dd2-1224">Obtiene las propiedades de la cita seleccionada o un mensaje en una carpeta compartida, calendario o buzón de correo.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1224">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-1225">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1225">Parameters:</span></span>

|<span data-ttu-id="a3dd2-1226">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1226">Name</span></span>|<span data-ttu-id="a3dd2-1227">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1227">Type</span></span>|<span data-ttu-id="a3dd2-1228">Atributos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1228">Attributes</span></span>|<span data-ttu-id="a3dd2-1229">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1229">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a3dd2-1230">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1230">Object</span></span>|<span data-ttu-id="a3dd2-1231">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1231">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1232">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1232">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a3dd2-1233">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1233">Object</span></span>|<span data-ttu-id="a3dd2-1234">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1234">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1235">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1235">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a3dd2-1236">función</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1236">function</span></span>||<span data-ttu-id="a3dd2-1237">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1237">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a3dd2-1238">Las propiedades compartidas se proporcionan como un [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) objeto en el `asyncResult.value` (propiedad).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1238">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a3dd2-1239">Este objeto se puede usar para obtener las propiedades del elemento compartido.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1239">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-1240">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1240">Requirements</span></span>

|<span data-ttu-id="a3dd2-1241">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1241">Requirement</span></span>|<span data-ttu-id="a3dd2-1242">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-1243">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1243">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-1244">Vista previa</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1244">Preview</span></span>|
|[<span data-ttu-id="a3dd2-1245">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1246">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-1247">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-1248">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-1249">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1249">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="a3dd2-1250">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1250">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="a3dd2-1251">Carga de forma asincrónica las propiedades personalizadas de este complemento en el elemento seleccionado.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1251">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="a3dd2-p174">Las propiedades personalizadas se almacenan como pares de clave/valor según la aplicación y el elemento. Este método devuelve un objeto `CustomProperties` en la devolución de llamada, que proporciona métodos para tener acceso a las propiedades personalizadas específicas del elemento y el complemento actual. Las propiedades personalizadas no están cifradas en el elemento, por lo que no debería usarse como almacenamiento seguro.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p174">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-1255">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1255">Parameters:</span></span>

|<span data-ttu-id="a3dd2-1256">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1256">Name</span></span>|<span data-ttu-id="a3dd2-1257">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1257">Type</span></span>|<span data-ttu-id="a3dd2-1258">Atributos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1258">Attributes</span></span>|<span data-ttu-id="a3dd2-1259">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1259">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="a3dd2-1260">función</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1260">function</span></span>||<span data-ttu-id="a3dd2-1261">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1261">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a3dd2-1262">Las propiedades personalizadas se proporcionan como un objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1262">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a3dd2-1263">Este objeto se puede utilizar para obtener, establecer y elimina las propiedades personalizadas del elemento y guarde los cambios en la propiedad personalizada que se vuelve a establecer en el servidor.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1263">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="a3dd2-1264">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1264">Object</span></span>|<span data-ttu-id="a3dd2-1265">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1266">Los desarrolladores pueden proporcionar cualquier objeto que desean tener acceso en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1266">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="a3dd2-1267">Se puede tener acceso a este objeto por el `asyncResult.asyncContext` (propiedad) en la función de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1267">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-1268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1268">Requirements</span></span>

|<span data-ttu-id="a3dd2-1269">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1269">Requirement</span></span>|<span data-ttu-id="a3dd2-1270">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-1271">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1271">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1272">1.0</span></span>|
|[<span data-ttu-id="a3dd2-1273">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1274">ReadItem</span></span>|
|[<span data-ttu-id="a3dd2-1275">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-1276">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-1277">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1277">Example</span></span>

<span data-ttu-id="a3dd2-p177">En el siguiente ejemplo de código se muestra cómo usar el método `loadCustomPropertiesAsync` para cargar de forma asincrónica propiedades personalizadas que son específicas del elemento actual. En el ejemplo también se muestra cómo usar el método `CustomProperties.saveAsync` para guardar estas propiedades en el servidor. Después de cargar las propiedades personalizadas, el ejemplo de código usa el método `CustomProperties.get` para leer la propiedad personalizada `myProp`, el método `CustomProperties.set` para escribir en la propiedad personalizada `otherProp` y finalmente se llama al método `saveAsync` para guardar las propiedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p177">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="a3dd2-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="a3dd2-1282">Quita los datos adjuntos de un mensaje o cita.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1282">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="a3dd2-p178">El método `removeAttachmentAsync` quita del elemento los datos adjuntos con el identificador especificado. Como práctica recomendada, debe usar el identificador de datos adjuntos para quitar datos adjuntos solo si la misma aplicación de correo ha agregado los datos adjuntos en la misma sesión. En Outlook Web App y OWA para los dispositivos, el identificador de datos adjuntos es válido solo en la misma sesión. Una sesión finaliza cuando el usuario cierra la aplicación, o si el usuario comienza a redactar en un formulario en línea y posteriormente se extrae el formulario en línea para continuar en una ventana independiente.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p178">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-1287">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1287">Parameters:</span></span>

|<span data-ttu-id="a3dd2-1288">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1288">Name</span></span>|<span data-ttu-id="a3dd2-1289">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1289">Type</span></span>|<span data-ttu-id="a3dd2-1290">Atributos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1290">Attributes</span></span>|<span data-ttu-id="a3dd2-1291">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1291">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="a3dd2-1292">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1292">String</span></span>||<span data-ttu-id="a3dd2-p179">El identificador de los datos adjuntos que se deben quitar. La longitud máxima de la cadena es de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p179">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="a3dd2-1295">Object</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1295">Object</span></span>|<span data-ttu-id="a3dd2-1296">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1296">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1297">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1297">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a3dd2-1298">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1298">Object</span></span>|<span data-ttu-id="a3dd2-1299">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1299">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1300">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1300">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a3dd2-1301">función</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1301">function</span></span>|<span data-ttu-id="a3dd2-1302">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1303">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1303">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a3dd2-1304">Si se produce un error en la eliminación de los datos adjuntos, la propiedad `asyncResult.error` contendrá un código de error con el motivo del error.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1304">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a3dd2-1305">Errores</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1305">Errors</span></span>

|<span data-ttu-id="a3dd2-1306">Código de error</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1306">Error code</span></span>|<span data-ttu-id="a3dd2-1307">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1307">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="a3dd2-1308">El identificador de datos adjuntos no existe.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1308">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-1309">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1309">Requirements</span></span>

|<span data-ttu-id="a3dd2-1310">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1310">Requirement</span></span>|<span data-ttu-id="a3dd2-1311">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1311">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-1312">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1312">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-1313">1.1</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1313">1.1</span></span>|
|[<span data-ttu-id="a3dd2-1314">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-1315">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1315">ReadWriteItem</span></span>|
|[<span data-ttu-id="a3dd2-1316">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-1317">Redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1317">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-1318">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1318">Example</span></span>

<span data-ttu-id="a3dd2-1319">El código siguiente quita un archivo adjunto con un identificador de "0".</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1319">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="a3dd2-1320">removeHandlerAsync (eventType, controlador, [opciones], [devolución de llamada])</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1320">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="a3dd2-1321">Quita un controlador de eventos para un evento compatible.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1321">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="a3dd2-1322">Actualmente son los tipos de eventos compatibles `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, y`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1322">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-1323">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1323">Parameters:</span></span>

| <span data-ttu-id="a3dd2-1324">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1324">Name</span></span> | <span data-ttu-id="a3dd2-1325">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1325">Type</span></span> | <span data-ttu-id="a3dd2-1326">Atributos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1326">Attributes</span></span> | <span data-ttu-id="a3dd2-1327">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1327">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a3dd2-1328">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1328">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a3dd2-1329">El evento que debe invocar el controlador.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1329">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="a3dd2-1330">Función</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1330">Function</span></span> || <span data-ttu-id="a3dd2-p180">La función que va a controlar el evento. La función debe aceptar un único parámetro, que es un literal de objeto. La propiedad `type` del parámetro coincidirá con el parámetro `eventType` que se ha pasado a `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p180">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="a3dd2-1334">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1334">Object</span></span> | <span data-ttu-id="a3dd2-1335">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1335">&lt;optional&gt;</span></span> | <span data-ttu-id="a3dd2-1336">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1336">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a3dd2-1337">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1337">Object</span></span> | <span data-ttu-id="a3dd2-1338">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1338">&lt;optional&gt;</span></span> | <span data-ttu-id="a3dd2-1339">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1339">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a3dd2-1340">función</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1340">function</span></span>| <span data-ttu-id="a3dd2-1341">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1341">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1342">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-1343">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1343">Requirements</span></span>

|<span data-ttu-id="a3dd2-1344">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1344">Requirement</span></span>| <span data-ttu-id="a3dd2-1345">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1345">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-1346">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1346">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a3dd2-1347">1.7</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1347">1.7</span></span> |
|[<span data-ttu-id="a3dd2-1348">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1348">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a3dd2-1349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1349">ReadItem</span></span> |
|[<span data-ttu-id="a3dd2-1350">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1350">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a3dd2-1351">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1351">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="a3dd2-1352">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1352">saveAsync([options], callback)</span></span>

<span data-ttu-id="a3dd2-1353">Guarda un elemento de forma asincrónica.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1353">Asynchronously saves an item.</span></span>

<span data-ttu-id="a3dd2-p181">Cuando se invoca, este método guarda el mensaje actual como un borrador y devuelve el identificador de elemento a través del método de devolución de llamada. En el modo en línea de Outlook Web App o Outlook, el elemento se guarda en el servidor. En modo en caché de Outlook, se guarda el elemento en la caché local.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-1357">Si el complemento llama a `saveAsync` en un elemento en modo de redacción con el fin de obtener una `itemId` para usar con EWS o la API de REST, tenga en cuenta que cuando Outlook está en modo en caché, puede tardar algún tiempo antes de que el elemento realmente se sincroniza con el servidor.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1357">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="a3dd2-1358">Hasta que se sincroniza el elemento, el uso de la `itemId` devolverá un error.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1358">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="a3dd2-p183">Dado que las citas no tienen ningún estado de borrador, si se llama a `saveAsync` en una cita en modo de redacción, el elemento se guardará como una cita normal en el calendario del usuario. Para las citas nuevas que todavía no se han guardado, no se enviará ninguna invitación. Si se guarda una cita existente, se enviará una actualización a los asistentes agregados o eliminados.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="a3dd2-1362">Los siguientes clientes tienen un comportamiento diferente para `saveAsync` en citas en modo de redacción:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1362">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="a3dd2-1363">Mac Outlook no admite `saveAsync` en una reunión en el modo de redacción.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1363">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="a3dd2-1364">Al llamar a `saveAsync` en una reunión en Mac Outlook devolverá un error.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1364">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="a3dd2-1365">Outlook en el web siempre envía una invitación o actualizar cuándo `saveAsync` se llama en una cita en modo de redacción.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1365">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-1366">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1366">Parameters:</span></span>

|<span data-ttu-id="a3dd2-1367">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1367">Name</span></span>|<span data-ttu-id="a3dd2-1368">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1368">Type</span></span>|<span data-ttu-id="a3dd2-1369">Atributos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1369">Attributes</span></span>|<span data-ttu-id="a3dd2-1370">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1370">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a3dd2-1371">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1371">Object</span></span>|<span data-ttu-id="a3dd2-1372">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1372">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1373">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1373">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a3dd2-1374">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1374">Object</span></span>|<span data-ttu-id="a3dd2-1375">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1375">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1376">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1376">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a3dd2-1377">función</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1377">function</span></span>||<span data-ttu-id="a3dd2-1378">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1378">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a3dd2-1379">En caso de éxito, se proporciona el identificador de elemento en el `asyncResult.value` (propiedad).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1379">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-1380">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1380">Requirements</span></span>

|<span data-ttu-id="a3dd2-1381">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1381">Requirement</span></span>|<span data-ttu-id="a3dd2-1382">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1382">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-1383">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1383">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-1384">1.3</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1384">1.3</span></span>|
|[<span data-ttu-id="a3dd2-1385">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-1386">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1386">ReadWriteItem</span></span>|
|[<span data-ttu-id="a3dd2-1387">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-1388">Redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1388">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a3dd2-1389">Ejemplos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1389">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="a3dd2-p185">A continuación, se muestra un ejemplo del parámetro `result` que se ha pasado a la función de devolución de llamada. La propiedad `value` contiene el identificador de elemento del elemento.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="a3dd2-1392">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1392">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="a3dd2-1393">Inserta asincrónicamente datos en el cuerpo o el asunto de un mensaje.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1393">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="a3dd2-p186">El método `setSelectedDataAsync` inserta la cadena especificada en la posición del cursor en el asunto o el cuerpo del elemento, o si se selecciona texto en el editor, reemplaza el texto seleccionado. Si el cursor no está en el cuerpo o el campo del asunto, se devuelve un error. Después de la inserción, el cursor se coloca al final del contenido insertado.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a3dd2-1397">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1397">Parameters:</span></span>

|<span data-ttu-id="a3dd2-1398">Nombre</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1398">Name</span></span>|<span data-ttu-id="a3dd2-1399">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1399">Type</span></span>|<span data-ttu-id="a3dd2-1400">Atributos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1400">Attributes</span></span>|<span data-ttu-id="a3dd2-1401">Descripción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1401">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="a3dd2-1402">String</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1402">String</span></span>||<span data-ttu-id="a3dd2-p187">Datos que se van a insertar. Los datos no deben superar 1.000.000 de caracteres. Si se pasan más de 1.000.000 de caracteres, se produce una excepción `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="a3dd2-1406">Object</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1406">Object</span></span>|<span data-ttu-id="a3dd2-1407">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1407">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1408">Un objeto literal que contiene una o más de las siguientes propiedades.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1408">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a3dd2-1409">Objeto</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1409">Object</span></span>|<span data-ttu-id="a3dd2-1410">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1410">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-1411">Los desarrolladores pueden proporcionar cualquier objeto que quieran para tener acceso al método de devolución de llamada.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1411">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="a3dd2-1412">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1412">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="a3dd2-1413">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1413">&lt;optional&gt;</span></span>|<span data-ttu-id="a3dd2-p188">Si `text`, se aplica el estilo actual en Outlook Web App y Outlook. Si el campo es un editor de HTML, se insertan solo los datos de texto, aunque los datos sean HTML.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="a3dd2-p189">Si `html` y el campo admiten HTML (el asunto no), el estilo actual se aplica en Outlook Web App y el estilo predeterminado en Outlook. Si el campo es un campo de texto, se devuelve un error `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="a3dd2-1418">Si `coercionType` no está establecido, el resultado depende del campo: si el campo es HTML, se usa HTML; si el campo es texto, se usa texto sin formato.</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1418">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="a3dd2-1419">function</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1419">function</span></span>||<span data-ttu-id="a3dd2-1420">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1420">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3dd2-1421">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1421">Requirements</span></span>

|<span data-ttu-id="a3dd2-1422">Requirement</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1422">Requirement</span></span>|<span data-ttu-id="a3dd2-1423">Valor</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1423">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3dd2-1424">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a3dd2-1425">1.2</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1425">1.2</span></span>|
|[<span data-ttu-id="a3dd2-1426">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a3dd2-1427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1427">ReadWriteItem</span></span>|
|[<span data-ttu-id="a3dd2-1428">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a3dd2-1429">Redacción</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1429">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a3dd2-1430">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="a3dd2-1430">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```