
# <a name="mailbox"></a><span data-ttu-id="29588-101">buzón de correo</span><span class="sxs-lookup"><span data-stu-id="29588-101">mailbox</span></span>

### <span data-ttu-id="29588-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="29588-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="29588-104">Proporciona acceso al modelo de objeto de complemento de Outlook para Microsoft Outlook y Microsoft Outlook en el web.</span><span class="sxs-lookup"><span data-stu-id="29588-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="29588-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="29588-105">Requirements</span></span>

|<span data-ttu-id="29588-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="29588-106">Requirement</span></span>| <span data-ttu-id="29588-107">Valor</span><span class="sxs-lookup"><span data-stu-id="29588-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="29588-108">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="29588-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29588-109">1.0</span><span class="sxs-lookup"><span data-stu-id="29588-109">1.0</span></span>|
|[<span data-ttu-id="29588-110">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="29588-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29588-111">Restringido</span><span class="sxs-lookup"><span data-stu-id="29588-111">Restricted</span></span>|
|[<span data-ttu-id="29588-112">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="29588-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29588-113">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="29588-113">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="29588-114">Espacios de nombres</span><span class="sxs-lookup"><span data-stu-id="29588-114">Namespaces</span></span>

<span data-ttu-id="29588-115">[diagnostics](Office.context.mailbox.diagnostics.md): Proporciona información de diagnóstico a un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="29588-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="29588-116">[item](Office.context.mailbox.item.md): Proporciona métodos y propiedades para tener acceso a un mensaje o cita en un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="29588-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="29588-117">[userProfile](Office.context.mailbox.userProfile.md): Proporciona información sobre el usuario en un complemento de Outlook.</span><span class="sxs-lookup"><span data-stu-id="29588-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="29588-118">Miembros</span><span class="sxs-lookup"><span data-stu-id="29588-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="29588-119">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="29588-119">ewsUrl :String</span></span>

<span data-ttu-id="29588-p102">Obtiene la dirección URL del punto de conexión de Servicios Web Exchange (EWS) para esta cuenta de correo electrónico. Solo modo Lectura.</span><span class="sxs-lookup"><span data-stu-id="29588-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="29588-122">Este miembro no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="29588-122">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="29588-p103">El valor `ewsUrl` puede usarse por un servicio remoto para realizar llamadas EWS al buzón del usuario. Por ejemplo, puede crear un servicio remoto para [obtener datos adjuntos desde el elemento seleccionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="29588-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="29588-125">Tipo:</span><span class="sxs-lookup"><span data-stu-id="29588-125">Type:</span></span>

*   <span data-ttu-id="29588-126">String</span><span class="sxs-lookup"><span data-stu-id="29588-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="29588-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="29588-127">Requirements</span></span>

|<span data-ttu-id="29588-128">Requirement</span><span class="sxs-lookup"><span data-stu-id="29588-128">Requirement</span></span>| <span data-ttu-id="29588-129">Valor</span><span class="sxs-lookup"><span data-stu-id="29588-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="29588-130">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="29588-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29588-131">1.0</span><span class="sxs-lookup"><span data-stu-id="29588-131">1.0</span></span>|
|[<span data-ttu-id="29588-132">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="29588-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29588-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29588-133">ReadItem</span></span>|
|[<span data-ttu-id="29588-134">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="29588-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29588-135">Lectura</span><span class="sxs-lookup"><span data-stu-id="29588-135">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="29588-136">Métodos</span><span class="sxs-lookup"><span data-stu-id="29588-136">Methods</span></span>

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime"></a><span data-ttu-id="29588-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="29588-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)}</span></span>

<span data-ttu-id="29588-138">Obtiene un diccionario con información de tiempo en el tiempo del cliente local.</span><span class="sxs-lookup"><span data-stu-id="29588-138">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="29588-p104">Las fechas y horas usadas por una aplicación de correo para Outlook o Outlook Web App pueden usar distintas zonas horarias. Outlook usa la zona horaria del equipo cliente; Outlook Web App usa la zona horaria definida en el Centro de administración de Exchange (EAC). Debería tratar los valores de fecha y hora de modo que los valores que aparezcan en la interfaz de usuario sean siempre coherentes con la zona horaria que el usuario espera.</span><span class="sxs-lookup"><span data-stu-id="29588-p104">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="29588-p105">Si se está ejecutando la aplicación de correo en Outlook, el método `convertToLocalClientTime` devolverá un objeto de diccionario con los valores establecidos para la zona horaria del equipo cliente. Si se está ejecutando la aplicación de correo en Outlook Web App, el método `convertToLocalClientTime` devolverá un objeto de diccionario con los valores establecidos para la zona horaria especificada en el CEF.</span><span class="sxs-lookup"><span data-stu-id="29588-p105">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29588-144">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="29588-144">Parameters:</span></span>

|<span data-ttu-id="29588-145">Nombre</span><span class="sxs-lookup"><span data-stu-id="29588-145">Name</span></span>| <span data-ttu-id="29588-146">Tipo</span><span class="sxs-lookup"><span data-stu-id="29588-146">Type</span></span>| <span data-ttu-id="29588-147">Descripción</span><span class="sxs-lookup"><span data-stu-id="29588-147">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="29588-148">Fecha</span><span class="sxs-lookup"><span data-stu-id="29588-148">Date</span></span>|<span data-ttu-id="29588-149">Un objeto Date</span><span class="sxs-lookup"><span data-stu-id="29588-149">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="29588-150">Requisitos</span><span class="sxs-lookup"><span data-stu-id="29588-150">Requirements</span></span>

|<span data-ttu-id="29588-151">Requirement</span><span class="sxs-lookup"><span data-stu-id="29588-151">Requirement</span></span>| <span data-ttu-id="29588-152">Valor</span><span class="sxs-lookup"><span data-stu-id="29588-152">Value</span></span>|
|---|---|
|[<span data-ttu-id="29588-153">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="29588-153">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29588-154">1.0</span><span class="sxs-lookup"><span data-stu-id="29588-154">1.0</span></span>|
|[<span data-ttu-id="29588-155">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="29588-155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29588-156">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29588-156">ReadItem</span></span>|
|[<span data-ttu-id="29588-157">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="29588-157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29588-158">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="29588-158">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="29588-159">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="29588-159">Returns:</span></span>

<span data-ttu-id="29588-160">Tipo: [LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="29588-160">Type: [LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)</span></span>

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="29588-161">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="29588-161">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="29588-162">Obtiene un objeto Date del diccionario que contiene información de tiempo.</span><span class="sxs-lookup"><span data-stu-id="29588-162">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="29588-163">El método `convertToUtcClientTime` convierte un diccionario que contiene la fecha y la hora locales en un objeto Date con los valores correctos para la fecha y la hora locales.</span><span class="sxs-lookup"><span data-stu-id="29588-163">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29588-164">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="29588-164">Parameters:</span></span>

|<span data-ttu-id="29588-165">Nombre</span><span class="sxs-lookup"><span data-stu-id="29588-165">Name</span></span>| <span data-ttu-id="29588-166">Tipo</span><span class="sxs-lookup"><span data-stu-id="29588-166">Type</span></span>| <span data-ttu-id="29588-167">Descripción</span><span class="sxs-lookup"><span data-stu-id="29588-167">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="29588-168">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="29588-168">LocalClientTime</span></span>](/javascript/api/outlook_1_2/office.LocalClientTime)|<span data-ttu-id="29588-169">El valor de la hora local para convertir.</span><span class="sxs-lookup"><span data-stu-id="29588-169">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="29588-170">Requisitos</span><span class="sxs-lookup"><span data-stu-id="29588-170">Requirements</span></span>

|<span data-ttu-id="29588-171">Requirement</span><span class="sxs-lookup"><span data-stu-id="29588-171">Requirement</span></span>| <span data-ttu-id="29588-172">Valor</span><span class="sxs-lookup"><span data-stu-id="29588-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="29588-173">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="29588-173">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29588-174">1.0</span><span class="sxs-lookup"><span data-stu-id="29588-174">1.0</span></span>|
|[<span data-ttu-id="29588-175">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="29588-175">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29588-176">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29588-176">ReadItem</span></span>|
|[<span data-ttu-id="29588-177">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="29588-177">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29588-178">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="29588-178">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="29588-179">Valores devueltos:</span><span class="sxs-lookup"><span data-stu-id="29588-179">Returns:</span></span>

<span data-ttu-id="29588-180">Objeto Date con el tiempo expresado en UTC.</span><span class="sxs-lookup"><span data-stu-id="29588-180">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="29588-181">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="29588-181">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="29588-182">Fecha</span><span class="sxs-lookup"><span data-stu-id="29588-182">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="29588-183">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="29588-183">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="29588-184">Muestra una cita de calendario existente.</span><span class="sxs-lookup"><span data-stu-id="29588-184">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="29588-185">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="29588-185">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="29588-186">El método `displayAppointmentForm` abre una cita de calendario existente en una nueva ventana del escritorio o en un cuadro de diálogo en los dispositivos móviles.</span><span class="sxs-lookup"><span data-stu-id="29588-186">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="29588-p106">En Outlook para Mac, puede usar este método para mostrar una cita que no forme parte de una serie periódica o la cita principal de una serie periódica, pero no puede mostrar una instancia de la serie. Esto es porque en Outlook para Mac no se puede tener acceso a las propiedades (incluido el identificador de elemento) de las instancias de una serie periódica.</span><span class="sxs-lookup"><span data-stu-id="29588-p106">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="29588-189">En Outlook Web App, este método abre el formato especificado solo si el cuerpo del formulario es inferior o igual a 32 KB en el número de caracteres.</span><span class="sxs-lookup"><span data-stu-id="29588-189">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="29588-190">Si el identificador de elemento especificado no identifica una cita existente, se abrirá una página en blanco en el dispositivo o equipo cliente y no se generará ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="29588-190">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29588-191">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="29588-191">Parameters:</span></span>

|<span data-ttu-id="29588-192">Nombre</span><span class="sxs-lookup"><span data-stu-id="29588-192">Name</span></span>| <span data-ttu-id="29588-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="29588-193">Type</span></span>| <span data-ttu-id="29588-194">Descripción</span><span class="sxs-lookup"><span data-stu-id="29588-194">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="29588-195">String</span><span class="sxs-lookup"><span data-stu-id="29588-195">String</span></span>|<span data-ttu-id="29588-196">Identificador de los servicios web de Exchange (EWS) para una cita de calendario existente.</span><span class="sxs-lookup"><span data-stu-id="29588-196">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="29588-197">Requisitos</span><span class="sxs-lookup"><span data-stu-id="29588-197">Requirements</span></span>

|<span data-ttu-id="29588-198">Requirement</span><span class="sxs-lookup"><span data-stu-id="29588-198">Requirement</span></span>| <span data-ttu-id="29588-199">Valor</span><span class="sxs-lookup"><span data-stu-id="29588-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="29588-200">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="29588-200">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29588-201">1.0</span><span class="sxs-lookup"><span data-stu-id="29588-201">1.0</span></span>|
|[<span data-ttu-id="29588-202">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="29588-202">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29588-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29588-203">ReadItem</span></span>|
|[<span data-ttu-id="29588-204">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="29588-204">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29588-205">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="29588-205">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="29588-206">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="29588-206">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="29588-207">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="29588-207">displayMessageForm(itemId)</span></span>

<span data-ttu-id="29588-208">Muestra un mensaje existente.</span><span class="sxs-lookup"><span data-stu-id="29588-208">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="29588-209">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="29588-209">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="29588-210">El método `displayMessageForm` abre un mensaje existente en una nueva ventana del escritorio o en un cuadro de diálogo en los dispositivos móviles.</span><span class="sxs-lookup"><span data-stu-id="29588-210">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="29588-211">En Outlook Web App, este método abre el formato especificado solo si el cuerpo del formulario es inferior o igual a 32 KB en el número de caracteres.</span><span class="sxs-lookup"><span data-stu-id="29588-211">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="29588-212">Si el identificador de elemento especificado no identifica un mensaje existente, no se mostrará ningún mensaje en el equipo cliente ni se generará ningún mensaje de error.</span><span class="sxs-lookup"><span data-stu-id="29588-212">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="29588-p107">No use el valor `displayMessageForm` con un `itemId` que represente una cita. Use el método `displayAppointmentForm` para mostrar una cita existente y `displayNewAppointmentForm` para mostrar un formulario para crear una cita nueva.</span><span class="sxs-lookup"><span data-stu-id="29588-p107">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29588-215">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="29588-215">Parameters:</span></span>

|<span data-ttu-id="29588-216">Nombre</span><span class="sxs-lookup"><span data-stu-id="29588-216">Name</span></span>| <span data-ttu-id="29588-217">Tipo</span><span class="sxs-lookup"><span data-stu-id="29588-217">Type</span></span>| <span data-ttu-id="29588-218">Descripción</span><span class="sxs-lookup"><span data-stu-id="29588-218">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="29588-219">String</span><span class="sxs-lookup"><span data-stu-id="29588-219">String</span></span>|<span data-ttu-id="29588-220">Identificador de los servicios web de Exchange (EWS) para un mensaje existente.</span><span class="sxs-lookup"><span data-stu-id="29588-220">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="29588-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="29588-221">Requirements</span></span>

|<span data-ttu-id="29588-222">Requirement</span><span class="sxs-lookup"><span data-stu-id="29588-222">Requirement</span></span>| <span data-ttu-id="29588-223">Valor</span><span class="sxs-lookup"><span data-stu-id="29588-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="29588-224">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="29588-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29588-225">1.0</span><span class="sxs-lookup"><span data-stu-id="29588-225">1.0</span></span>|
|[<span data-ttu-id="29588-226">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="29588-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29588-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29588-227">ReadItem</span></span>|
|[<span data-ttu-id="29588-228">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="29588-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29588-229">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="29588-229">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="29588-230">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="29588-230">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="29588-231">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="29588-231">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="29588-232">Muestra un formulario para crear una nueva cita de calendario.</span><span class="sxs-lookup"><span data-stu-id="29588-232">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="29588-233">Este método no se admite en Outlook para iOS o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="29588-233">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="29588-p108">El método `displayNewAppointmentForm` abre un formulario que permite al usuario crear una nueva cita o reunión. Si se especifican parámetros, los campos de formulario de cita se rellenan automáticamente con el contenido de los parámetros.</span><span class="sxs-lookup"><span data-stu-id="29588-p108">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="29588-p109">En Outlook Web App y OWA para dispositivos, este método muestra siempre un formulario con un campo de asistentes. Si no especifica ningún asistente como argumento de entrada, el método muestra un formulario con un botón **Guardar**. Si ha especificado asistentes, el formulario incluirá a los asistentes y un botón **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="29588-p109">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="29588-p110">En el cliente enriquecido de Outlook y Outlook RT, si se especifica cualquier asistente o recurso en los parámetros `requiredAttendees`, `optionalAttendees` o `resources`, este método muestra un formulario de reunión con un botón **Enviar**. Si no se especifica ningún destinatario, este método muestra un formulario de cita con un botón **Guardar y cerrar**.</span><span class="sxs-lookup"><span data-stu-id="29588-p110">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="29588-241">Si cualquiera de los parámetros supera los límites de tamaño especificados o si se especifica un nombre de parámetro desconocido, se genera una excepción.</span><span class="sxs-lookup"><span data-stu-id="29588-241">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29588-242">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="29588-242">Parameters:</span></span>

|<span data-ttu-id="29588-243">Nombre</span><span class="sxs-lookup"><span data-stu-id="29588-243">Name</span></span>| <span data-ttu-id="29588-244">Tipo</span><span class="sxs-lookup"><span data-stu-id="29588-244">Type</span></span>| <span data-ttu-id="29588-245">Descripción</span><span class="sxs-lookup"><span data-stu-id="29588-245">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="29588-246">Object</span><span class="sxs-lookup"><span data-stu-id="29588-246">Object</span></span> | <span data-ttu-id="29588-247">Un diccionario de parámetros que describen la nueva cita.</span><span class="sxs-lookup"><span data-stu-id="29588-247">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="29588-248">Matriz.&lt;Cadena&gt; &#124; Matriz.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="29588-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="29588-p111">Una matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los asistentes necesarios de la cita. La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="29588-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="29588-251">Matriz.&lt;Cadena&gt; &#124; Matriz.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="29588-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="29588-p112">Una matriz de cadenas que contiene las direcciones de correo electrónico o una matriz que contiene un objeto `EmailAddressDetails` para cada uno de los asistentes opcionales de la cita. La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="29588-p112">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="29588-254">Fecha</span><span class="sxs-lookup"><span data-stu-id="29588-254">Date</span></span> | <span data-ttu-id="29588-255">Un objeto `Date` que especifica la fecha y hora de inicio de la cita.</span><span class="sxs-lookup"><span data-stu-id="29588-255">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="29588-256">Fecha</span><span class="sxs-lookup"><span data-stu-id="29588-256">Date</span></span> | <span data-ttu-id="29588-257">Un objeto `Date` que especifica la fecha y hora de finalización de la cita.</span><span class="sxs-lookup"><span data-stu-id="29588-257">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="29588-258">String</span><span class="sxs-lookup"><span data-stu-id="29588-258">String</span></span> | <span data-ttu-id="29588-p113">Una cadena que contiene la ubicación de la cita. La cadena está limitada a un máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="29588-p113">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="29588-261">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="29588-261">Array.&lt;String&gt;</span></span> | <span data-ttu-id="29588-p114">Una matriz de cadenas que contiene los recursos necesarios para la cita. La matriz está limitada a un máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="29588-p114">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="29588-264">String</span><span class="sxs-lookup"><span data-stu-id="29588-264">String</span></span> | <span data-ttu-id="29588-p115">Una cadena que contiene al asunto de la cita. La cadena está limitada a un máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="29588-p115">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="29588-267">String</span><span class="sxs-lookup"><span data-stu-id="29588-267">String</span></span> | <span data-ttu-id="29588-p116">El cuerpo de la cita. El contenido del cuerpo está limitado a un tamaño máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="29588-p116">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="29588-270">Requisitos</span><span class="sxs-lookup"><span data-stu-id="29588-270">Requirements</span></span>

|<span data-ttu-id="29588-271">Requirement</span><span class="sxs-lookup"><span data-stu-id="29588-271">Requirement</span></span>| <span data-ttu-id="29588-272">Valor</span><span class="sxs-lookup"><span data-stu-id="29588-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="29588-273">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="29588-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29588-274">1.0</span><span class="sxs-lookup"><span data-stu-id="29588-274">1.0</span></span>|
|[<span data-ttu-id="29588-275">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="29588-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29588-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29588-276">ReadItem</span></span>|
|[<span data-ttu-id="29588-277">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="29588-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29588-278">Lectura</span><span class="sxs-lookup"><span data-stu-id="29588-278">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="29588-279">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="29588-279">Example</span></span>

```
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="29588-280">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="29588-280">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="29588-281">Obtiene una cadena que contiene un token usado para obtener datos adjuntos o un elemento de Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="29588-281">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="29588-p117">El método `getCallbackTokenAsync` realiza una llamada asincrónica para obtener un token opaco desde Exchange Server que hospeda el buzón del usuario. La duración del token de devolución de llamada es de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="29588-p117">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="29588-p118">Puede pasar el token y un identificador de archivo adjunto o el identificador del elemento a un sistema de terceros. El sistema de terceros usa el token como token de autorización del portador para llamar a los Servicios web de Exchange (EWS) o a las operaciones [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) o [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) para devolver datos adjuntos o un elemento. Por ejemplo, puede crear un servicio remoto para [obtener datos adjuntos desde el elemento seleccionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="29588-p118">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="29588-287">La aplicación debe tener el permiso **ReadItem** especificado en su manifiesto para poder llamar al método `getCallbackTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="29588-287">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29588-288">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="29588-288">Parameters:</span></span>

|<span data-ttu-id="29588-289">Nombre</span><span class="sxs-lookup"><span data-stu-id="29588-289">Name</span></span>| <span data-ttu-id="29588-290">Tipo</span><span class="sxs-lookup"><span data-stu-id="29588-290">Type</span></span>| <span data-ttu-id="29588-291">Atributos</span><span class="sxs-lookup"><span data-stu-id="29588-291">Attributes</span></span>| <span data-ttu-id="29588-292">Descripción</span><span class="sxs-lookup"><span data-stu-id="29588-292">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="29588-293">función</span><span class="sxs-lookup"><span data-stu-id="29588-293">function</span></span>||<span data-ttu-id="29588-294">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="29588-294">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="29588-295">El token se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="29588-295">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="29588-296">Object</span><span class="sxs-lookup"><span data-stu-id="29588-296">Object</span></span>| <span data-ttu-id="29588-297">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29588-297">&lt;optional&gt;</span></span>|<span data-ttu-id="29588-298">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="29588-298">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="29588-299">Requisitos</span><span class="sxs-lookup"><span data-stu-id="29588-299">Requirements</span></span>

|<span data-ttu-id="29588-300">Requirement</span><span class="sxs-lookup"><span data-stu-id="29588-300">Requirement</span></span>| <span data-ttu-id="29588-301">Valor</span><span class="sxs-lookup"><span data-stu-id="29588-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="29588-302">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="29588-302">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29588-303">1.0</span><span class="sxs-lookup"><span data-stu-id="29588-303">1.0</span></span>|
|[<span data-ttu-id="29588-304">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="29588-304">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29588-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29588-305">ReadItem</span></span>|
|[<span data-ttu-id="29588-306">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="29588-306">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29588-307">Lectura</span><span class="sxs-lookup"><span data-stu-id="29588-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="29588-308">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="29588-308">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="29588-309">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="29588-309">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="29588-310">Obtiene un token que identifica al usuario y al complemento de Office.</span><span class="sxs-lookup"><span data-stu-id="29588-310">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="29588-311">El método `getUserIdentityTokenAsync` devuelve un token que puede usar para identificar y [autenticar el complemento y el usuario mediante un sistema de terceros](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="29588-311">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="29588-312">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="29588-312">Parameters:</span></span>

|<span data-ttu-id="29588-313">Nombre</span><span class="sxs-lookup"><span data-stu-id="29588-313">Name</span></span>| <span data-ttu-id="29588-314">Tipo</span><span class="sxs-lookup"><span data-stu-id="29588-314">Type</span></span>| <span data-ttu-id="29588-315">Atributos</span><span class="sxs-lookup"><span data-stu-id="29588-315">Attributes</span></span>| <span data-ttu-id="29588-316">Descripción</span><span class="sxs-lookup"><span data-stu-id="29588-316">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="29588-317">función</span><span class="sxs-lookup"><span data-stu-id="29588-317">function</span></span>||<span data-ttu-id="29588-318">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="29588-318">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="29588-319">El token se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="29588-319">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="29588-320">Object</span><span class="sxs-lookup"><span data-stu-id="29588-320">Object</span></span>| <span data-ttu-id="29588-321">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29588-321">&lt;optional&gt;</span></span>|<span data-ttu-id="29588-322">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="29588-322">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="29588-323">Requisitos</span><span class="sxs-lookup"><span data-stu-id="29588-323">Requirements</span></span>

|<span data-ttu-id="29588-324">Requirement</span><span class="sxs-lookup"><span data-stu-id="29588-324">Requirement</span></span>| <span data-ttu-id="29588-325">Valor</span><span class="sxs-lookup"><span data-stu-id="29588-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="29588-326">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="29588-326">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29588-327">1.0</span><span class="sxs-lookup"><span data-stu-id="29588-327">1.0</span></span>|
|[<span data-ttu-id="29588-328">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="29588-328">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29588-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="29588-329">ReadItem</span></span>|
|[<span data-ttu-id="29588-330">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="29588-330">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29588-331">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="29588-331">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="29588-332">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="29588-332">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="29588-333">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="29588-333">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="29588-334">Realiza una solicitud asincrónica a un servicio de Servicios Web Exchange (EWS) en el servidor Exchange que hospeda el buzón del usuario.</span><span class="sxs-lookup"><span data-stu-id="29588-334">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="29588-335">Este método no se admite en los siguientes escenarios.</span><span class="sxs-lookup"><span data-stu-id="29588-335">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="29588-336">En Outlook para iOS o Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="29588-336">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="29588-337">Cuando el complemento se carga en un buzón de Gmail</span><span class="sxs-lookup"><span data-stu-id="29588-337">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="29588-338">En estos casos, complementos deben [usar API de REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) para tener acceso al buzón del usuario en su lugar.</span><span class="sxs-lookup"><span data-stu-id="29588-338">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="29588-339">El método `makeEwsRequestAsync` envía una solicitud de EWS en nombre del complemento a Exchange.</span><span class="sxs-lookup"><span data-stu-id="29588-339">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="29588-340">Para obtener una lista de las operaciones de EWS compatibles, vea [llamar a servicios web desde un complemento de Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="29588-340">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="29588-341">No puede solicitar elementos asociados de las carpetas con el método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="29588-341">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="29588-342">La solicitud XML debe especificar codificación UTF-8.</span><span class="sxs-lookup"><span data-stu-id="29588-342">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="29588-p120">Su complemento debe tener el permiso **ReadWriteMailbox** para usar el método `makeEwsRequestAsync`. Para obtener información sobre cómo usar el permiso **ReadWriteMailbox** y sobre las operaciones EWS que puede llamar con el método `makeEwsRequestAsync`, consulte [Especificar permisos para el acceso del complemento de correo al buzón del usuario](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="29588-p120">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="29588-345">El administrador del servidor debe establecer `OAuthAuthentication` en true en el directorio EWS del servidor de acceso de cliente para habilitar la `makeEwsRequestAsync` método para realizar EWS solicita.</span><span class="sxs-lookup"><span data-stu-id="29588-345">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="29588-346">Diferencias de versión</span><span class="sxs-lookup"><span data-stu-id="29588-346">Version differences</span></span>

<span data-ttu-id="29588-347">Si usa el método `makeEwsRequestAsync` en aplicaciones de correo que se ejecutan en versiones de Outlook anteriores a 15.0.4535.1004, debe establecer el valor de codificación a `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="29588-347">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="29588-p121">No es necesario establecer el valor de codificación si la aplicación de correo se ejecuta en Outlook en la web. Puede averiguar si su aplicación de correo se ejecuta en Outlook o en Outlook en la web usando la propiedad mailbox.diagnostics.hostName. Para averiguar qué versión de Outlook se está ejecutando, use la propiedad mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="29588-p121">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="29588-351">Parámetros:</span><span class="sxs-lookup"><span data-stu-id="29588-351">Parameters:</span></span>

|<span data-ttu-id="29588-352">Nombre</span><span class="sxs-lookup"><span data-stu-id="29588-352">Name</span></span>| <span data-ttu-id="29588-353">Tipo</span><span class="sxs-lookup"><span data-stu-id="29588-353">Type</span></span>| <span data-ttu-id="29588-354">Atributos</span><span class="sxs-lookup"><span data-stu-id="29588-354">Attributes</span></span>| <span data-ttu-id="29588-355">Descripción</span><span class="sxs-lookup"><span data-stu-id="29588-355">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="29588-356">String</span><span class="sxs-lookup"><span data-stu-id="29588-356">String</span></span>||<span data-ttu-id="29588-357">La solicitud de EWS.</span><span class="sxs-lookup"><span data-stu-id="29588-357">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="29588-358">función</span><span class="sxs-lookup"><span data-stu-id="29588-358">function</span></span>||<span data-ttu-id="29588-359">Cuando el método finaliza, la función que se pasa en el parámetro `callback` se llama con un único parámetro, `asyncResult`, que es un objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="29588-359">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="29588-360">El resultado XML de la llamada EWS se proporciona como una cadena en la propiedad `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="29588-360">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="29588-361">Si el resultado supera 1 MB de tamaño, se devuelve un mensaje de error en su lugar.</span><span class="sxs-lookup"><span data-stu-id="29588-361">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="29588-362">Objeto</span><span class="sxs-lookup"><span data-stu-id="29588-362">Object</span></span>| <span data-ttu-id="29588-363">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="29588-363">&lt;optional&gt;</span></span>|<span data-ttu-id="29588-364">Cualquier dato de estado que se pasa al método asincrónico.</span><span class="sxs-lookup"><span data-stu-id="29588-364">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="29588-365">Requisitos</span><span class="sxs-lookup"><span data-stu-id="29588-365">Requirements</span></span>

|<span data-ttu-id="29588-366">Requirement</span><span class="sxs-lookup"><span data-stu-id="29588-366">Requirement</span></span>| <span data-ttu-id="29588-367">Valor</span><span class="sxs-lookup"><span data-stu-id="29588-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="29588-368">Versión del conjunto de requisitos mínimos del buzón</span><span class="sxs-lookup"><span data-stu-id="29588-368">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="29588-369">1.0</span><span class="sxs-lookup"><span data-stu-id="29588-369">1.0</span></span>|
|[<span data-ttu-id="29588-370">Nivel de permisos mínimo</span><span class="sxs-lookup"><span data-stu-id="29588-370">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="29588-371">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="29588-371">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="29588-372">Modo de Outlook aplicable</span><span class="sxs-lookup"><span data-stu-id="29588-372">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="29588-373">Redacción o lectura</span><span class="sxs-lookup"><span data-stu-id="29588-373">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="29588-374">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="29588-374">Example</span></span>

<span data-ttu-id="29588-375">En el siguiente ejemplo, se llama a `makeEwsRequestAsync` para usar la operación `GetItem` para obtener el asunto de un elemento.</span><span class="sxs-lookup"><span data-stu-id="29588-375">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item.
   var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
}
```