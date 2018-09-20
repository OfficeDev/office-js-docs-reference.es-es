# <a name="extensionpoint-element"></a><span data-ttu-id="ce329-101">Elemento ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="ce329-101">ExtensionPoint element</span></span>

 <span data-ttu-id="ce329-102">Define dónde expone su función un complemento en la interfaz de usuario de Office.</span><span class="sxs-lookup"><span data-stu-id="ce329-102">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="ce329-103">El elemento **ExtensionPoint** es un elemento secundario de [AllFormFactors](allformfactors.md), de [DesktopFormFactor](desktopformfactor.md) o de [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="ce329-103">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> 

## <a name="attributes"></a><span data-ttu-id="ce329-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="ce329-104">Attributes</span></span>

|  <span data-ttu-id="ce329-105">Atributo</span><span class="sxs-lookup"><span data-stu-id="ce329-105">Attribute</span></span>  |  <span data-ttu-id="ce329-106">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="ce329-106">Required</span></span>  |  <span data-ttu-id="ce329-107">Descripción</span><span class="sxs-lookup"><span data-stu-id="ce329-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="ce329-108">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="ce329-108">**xsi:type**</span></span>  |  <span data-ttu-id="ce329-109">Sí</span><span class="sxs-lookup"><span data-stu-id="ce329-109">Yes</span></span>  | <span data-ttu-id="ce329-110">El tipo de punto de extensión que se está definiendo.</span><span class="sxs-lookup"><span data-stu-id="ce329-110">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="ce329-111">Puntos de extensión solo para Excel</span><span class="sxs-lookup"><span data-stu-id="ce329-111">Extension points for Excel only</span></span>

- <span data-ttu-id="ce329-112">**CustomFunctions** - una función personalizada que se escriben en JavaScript para Excel.</span><span class="sxs-lookup"><span data-stu-id="ce329-112">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="ce329-113">[XML de este ejemplo de código](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml) se muestra cómo usar el elemento **ExtensionPoint** con el valor del atributo **CustomFunctions** y los elementos secundarios que se usará.</span><span class="sxs-lookup"><span data-stu-id="ce329-113">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="ce329-114">Puntos de extensión para comandos de complemento de Word, Excel, PowerPoint y OneNote</span><span class="sxs-lookup"><span data-stu-id="ce329-114">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="ce329-115">**PrimaryCommandSurface**: la cinta de opciones en Office.</span><span class="sxs-lookup"><span data-stu-id="ce329-115">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="ce329-116">**ContextMenu**: el menú contextual que aparece cuando se hace clic con el botón derecho en la interfaz de usuario de Office.</span><span class="sxs-lookup"><span data-stu-id="ce329-116">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="ce329-117">Los ejemplos siguientes muestran cómo usar el elemento  **ExtensionPoint** con los valores de atributo **PrimaryCommandSurface** y **ContextMenu**, así como los elementos secundarios que hay que usar con cada uno de ellos.</span><span class="sxs-lookup"><span data-stu-id="ce329-117">The following examples show how to use the  **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="ce329-118">Para los elementos que contienen un atributo ID, asegúrese de que proporcionar un identificador único.</span><span class="sxs-lookup"><span data-stu-id="ce329-118">For elements that contain an ID attribute, make sure you provide a unique ID.</span></span> <span data-ttu-id="ce329-119">Le recomendamos que use el nombre de su compañía con su id.</span><span class="sxs-lookup"><span data-stu-id="ce329-119">We recommend that you use your company's name along with your ID.</span></span> <span data-ttu-id="ce329-120">Por ejemplo, use el formato siguiente.</span><span class="sxs-lookup"><span data-stu-id="ce329-120">For example, use the following format.</span></span> <CustomTab id="mycompanyname.mygroupname">

```XML
<ExtensionPoint xsi:type="PrimaryCommandSurface">
          <CustomTab id="Contoso Tab">
          <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
            <!-- <OfficeTab id="TabData"> -->
            <Label resid="residLabel4" />
            <Group id="Group1Id12">
              <Label resid="residLabel4" />
              <Icon>
                <bt:Image size="16" resid="icon1_32x32" />
                <bt:Image size="32" resid="icon1_32x32" />
                <bt:Image size="80" resid="icon1_32x32" />
              </Icon>
              <Tooltip resid="residToolTip" />
              <Control xsi:type="Button" id="Button1Id1">

                  <!-- information about the control -->
              </Control>
              <!-- other controls, as needed -->
            </Group>
          </CustomTab>
        </ExtensionPoint>

      <ExtensionPoint xsi:type="ContextMenu">
        <OfficeMenu id="ContextMenuCell">
          <Control xsi:type="Menu" id="ContextMenu2">
                  <!-- information about the control -->
          </Control>
          <!-- other controls, as needed -->
        </OfficeMenu>
        </ExtensionPoint>
```

#### <a name="child-elements"></a><span data-ttu-id="ce329-121">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="ce329-121">Child elements</span></span>
 
|<span data-ttu-id="ce329-122">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="ce329-122">**Element**</span></span>|<span data-ttu-id="ce329-123">**Descripción**</span><span class="sxs-lookup"><span data-stu-id="ce329-123">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="ce329-124">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="ce329-124">**CustomTab**</span></span>|<span data-ttu-id="ce329-p103">Es obligatorio si quiere agregar una pestaña personalizada a la cinta de opciones (con  **PrimaryCommandSurface**). Si usa el elemento  **CustomTab**, no puede usar el elemento  **OfficeTab**. El atributo  **id** es obligatorio.</span><span class="sxs-lookup"><span data-stu-id="ce329-p103">Required if you want to add a custom tab to the ribbon (using  **PrimaryCommandSurface**). If you use the  **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="ce329-128">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="ce329-128">**OfficeTab**</span></span>|<span data-ttu-id="ce329-p104">Es obligatorio si quiere extender una pestaña de cinta de Office predeterminada (mediante **PrimaryCommandSurface**). Si usa el elemento **OfficeTab**, no puede usar el elemento **CustomTab**. Para obtener información detallada, consulte [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="ce329-p104">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the  **OfficeTab** element, you can't use the **CustomTab** element. For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="ce329-132">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="ce329-132">**OfficeMenu**</span></span>|<span data-ttu-id="ce329-p105">Es obligatorio si agrega comandos de complemento a un menú contextual predeterminado (con **ContextMenu**). El atributo **id** debe establecerse en: </span><span class="sxs-lookup"><span data-stu-id="ce329-p105">Required if you're adding add-in commands to a default context menu (using  **ContextMenu**). The  **id** attribute must be set to: </span></span><br/> <span data-ttu-id="ce329-p106">- **ContextMenuText** para Excel o Word. Muestra el elemento en el menú contextual cuando se selecciona un texto y luego el usuario hace clic en él con el botón derecho. </span><span class="sxs-lookup"><span data-stu-id="ce329-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="ce329-p107">- **ContextMenuCell** para Excel. Muestra el elemento en el menú contextual cuando el usuario hace clic con el botón derecho en una celda de la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="ce329-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="ce329-139">**Group**</span><span class="sxs-lookup"><span data-stu-id="ce329-139">**Group**</span></span>|<span data-ttu-id="ce329-p108">Un grupo de puntos de extensión de interfaz de usuario en una pestaña. Un grupo puede tener hasta seis controles. El atributo  **id** es obligatorio. Es una cadena con un máximo de 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ce329-p108">A group of user interface extension points on a tab. A group can have up to six controls. The  **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="ce329-143">**Label**</span><span class="sxs-lookup"><span data-stu-id="ce329-143">**Label**</span></span>|<span data-ttu-id="ce329-p109">Obligatorio. La etiqueta del grupo. El atributo  **resid** debe establecerse en el valor del atributo **id** de un elemento **String**. El elemento  **String** es un elemento secundario del elemento **ShortStrings**, que a su vez lo es de  **Resources**.</span><span class="sxs-lookup"><span data-stu-id="ce329-p109">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="ce329-148">**Icon**</span><span class="sxs-lookup"><span data-stu-id="ce329-148">**Icon**</span></span>|<span data-ttu-id="ce329-p110">Obligatorio. Especifica el icono del grupo que se usará en dispositivos de factor de forma pequeños o cuando se muestren demasiados botones. El atributo  **resid** debe establecerse en el valor del atributo **id** de un elemento **Image**. El elemento  **Image** es un elemento secundario de **Images**, que a su vez lo es de  **Resources**. El atributo **size** determina el tamaño de la imagen en píxeles. Se necesitan tres tamaños de imagen: 16, 32 y 80. También se admiten cinco tamaños opcionales: 20, 24, 40, 48 y 64.</span><span class="sxs-lookup"><span data-stu-id="ce329-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="ce329-156">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="ce329-156">**Tooltip**</span></span>|<span data-ttu-id="ce329-p111">Opcional. La información sobre herramientas del grupo. El atributo  **resid** debe establecerse en el valor del atributo **id** de un elemento **String**. El elemento  **String** es un elemento secundario del elemento **LongStrings**, que a su vez lo es de  **Resources**.</span><span class="sxs-lookup"><span data-stu-id="ce329-p111">Optional. The tooltip of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="ce329-161">**Control**</span><span class="sxs-lookup"><span data-stu-id="ce329-161">**Control**</span></span>|<span data-ttu-id="ce329-162">Cada grupo necesita como mínimo un control.</span><span class="sxs-lookup"><span data-stu-id="ce329-162">Each group requires at least one control.</span></span> <span data-ttu-id="ce329-163">Un elemento **Control** puede ser un **botón** o un **menú**.</span><span class="sxs-lookup"><span data-stu-id="ce329-163">A  **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="ce329-164">Use **Menu** para especificar una lista desplegable de controles de botón.</span><span class="sxs-lookup"><span data-stu-id="ce329-164">Use  **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="ce329-165">Actualmente, solo se admiten botones y menús.</span><span class="sxs-lookup"><span data-stu-id="ce329-165">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="ce329-166">Vea las secciones [Controles de botón](control.md#button-control) y [Controles de menú (de botones desplegables)](control.md#menu-dropdown-button-controls) para obtener más información.</span><span class="sxs-lookup"><span data-stu-id="ce329-166">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="ce329-167">**Nota:**  Para hacer que la solución de problemas, se recomienda que un elemento de **Control** y los elementos secundarios de **recursos** relacionados se agregará uno a la vez.</span><span class="sxs-lookup"><span data-stu-id="ce329-167">**Note:**  To make troubleshooting easier, we recommend that a  **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="ce329-168">**Script**</span><span class="sxs-lookup"><span data-stu-id="ce329-168">**Script**</span></span>|<span data-ttu-id="ce329-169">Vincula al archivo JavaScript con el código de registro y la definición de función personalizada.</span><span class="sxs-lookup"><span data-stu-id="ce329-169">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="ce329-170">Este elemento no se usa en la vista previa para desarrolladores.</span><span class="sxs-lookup"><span data-stu-id="ce329-170">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="ce329-171">En su lugar, la página HTML es responsable de cargar todos los archivos de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ce329-171">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="ce329-172">**Page**</span><span class="sxs-lookup"><span data-stu-id="ce329-172">**Page**</span></span>|<span data-ttu-id="ce329-173">Vínculos a la página HTML para sus funciones personalizadas.</span><span class="sxs-lookup"><span data-stu-id="ce329-173">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="ce329-174">Puntos de extensión para Outlook</span><span class="sxs-lookup"><span data-stu-id="ce329-174">Extension points for Outlook</span></span>

- [<span data-ttu-id="ce329-175">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ce329-175">MessageReadCommandSurface</span></span>](#messagereadcommandsurface) 
- [<span data-ttu-id="ce329-176">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ce329-176">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface) 
- [<span data-ttu-id="ce329-177">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ce329-177">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface) 
- [<span data-ttu-id="ce329-178">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ce329-178">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="ce329-179">[Module](#module) (solo puede usarse en el elemento [DesktopFormFactor](desktopformfactor.md)).</span><span class="sxs-lookup"><span data-stu-id="ce329-179">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="ce329-180">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ce329-180">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="ce329-181">Events</span><span class="sxs-lookup"><span data-stu-id="ce329-181">Events</span></span>](#events)
- [<span data-ttu-id="ce329-182">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="ce329-182">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="ce329-183">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ce329-183">MessageReadCommandSurface</span></span>
<span data-ttu-id="ce329-p114">Este punto de extensión coloca botones en la superficie del comando de la vista de lectura de correo. En el escritorio de Outlook aparece en la cinta.</span><span class="sxs-lookup"><span data-stu-id="ce329-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="ce329-186">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="ce329-186">Child elements</span></span>

|  <span data-ttu-id="ce329-187">Elemento</span><span class="sxs-lookup"><span data-stu-id="ce329-187">Element</span></span> |  <span data-ttu-id="ce329-188">Descripción</span><span class="sxs-lookup"><span data-stu-id="ce329-188">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ce329-189">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ce329-189">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="ce329-190">Agrega los comandos a la pestaña de la cinta predeterminada.</span><span class="sxs-lookup"><span data-stu-id="ce329-190">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="ce329-191">CustomTab</span><span class="sxs-lookup"><span data-stu-id="ce329-191">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="ce329-192">Agrega los comandos a la pestaña de la cinta personalizada.</span><span class="sxs-lookup"><span data-stu-id="ce329-192">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="ce329-193">Ejemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ce329-193">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="ce329-194">Ejemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="ce329-194">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="ce329-195">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ce329-195">MessageComposeCommandSurface</span></span>
<span data-ttu-id="ce329-196">Este punto de extensión coloca botones en la cinta para complementos mediante el formulario de redacción de correo.</span><span class="sxs-lookup"><span data-stu-id="ce329-196">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="ce329-197">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="ce329-197">Child elements</span></span>

|  <span data-ttu-id="ce329-198">Elemento</span><span class="sxs-lookup"><span data-stu-id="ce329-198">Element</span></span> |  <span data-ttu-id="ce329-199">Descripción</span><span class="sxs-lookup"><span data-stu-id="ce329-199">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ce329-200">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ce329-200">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="ce329-201">Agrega los comandos a la pestaña de la cinta predeterminada.</span><span class="sxs-lookup"><span data-stu-id="ce329-201">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="ce329-202">CustomTab</span><span class="sxs-lookup"><span data-stu-id="ce329-202">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="ce329-203">Agrega los comandos a la pestaña de la cinta personalizada.</span><span class="sxs-lookup"><span data-stu-id="ce329-203">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="ce329-204">Ejemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ce329-204">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="ce329-205">Ejemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="ce329-205">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="ce329-206">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ce329-206">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="ce329-207">Este punto de extensión coloca botones en la cinta para el formulario que se muestra al organizador de la reunión.</span><span class="sxs-lookup"><span data-stu-id="ce329-207">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="ce329-208">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="ce329-208">Child elements</span></span>

|  <span data-ttu-id="ce329-209">Elemento</span><span class="sxs-lookup"><span data-stu-id="ce329-209">Element</span></span> |  <span data-ttu-id="ce329-210">Descripción</span><span class="sxs-lookup"><span data-stu-id="ce329-210">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ce329-211">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ce329-211">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="ce329-212">Agrega los comandos a la pestaña de la cinta predeterminada.</span><span class="sxs-lookup"><span data-stu-id="ce329-212">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="ce329-213">CustomTab</span><span class="sxs-lookup"><span data-stu-id="ce329-213">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="ce329-214">Agrega los comandos a la pestaña de la cinta personalizada.</span><span class="sxs-lookup"><span data-stu-id="ce329-214">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="ce329-215">Ejemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ce329-215">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="ce329-216">Ejemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="ce329-216">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="ce329-217">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ce329-217">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="ce329-218">Este punto de extensión coloca botones en la cinta para el formulario que se muestra al asistente de la reunión.</span><span class="sxs-lookup"><span data-stu-id="ce329-218">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="ce329-219">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="ce329-219">Child elements</span></span>

|  <span data-ttu-id="ce329-220">Elemento</span><span class="sxs-lookup"><span data-stu-id="ce329-220">Element</span></span> |  <span data-ttu-id="ce329-221">Descripción</span><span class="sxs-lookup"><span data-stu-id="ce329-221">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ce329-222">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ce329-222">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="ce329-223">Agrega los comandos a la pestaña de la cinta predeterminada.</span><span class="sxs-lookup"><span data-stu-id="ce329-223">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="ce329-224">CustomTab</span><span class="sxs-lookup"><span data-stu-id="ce329-224">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="ce329-225">Agrega los comandos a la pestaña de la cinta personalizada.</span><span class="sxs-lookup"><span data-stu-id="ce329-225">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="ce329-226">Ejemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ce329-226">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="ce329-227">Ejemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="ce329-227">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="ce329-228">Módulo</span><span class="sxs-lookup"><span data-stu-id="ce329-228">Module</span></span>

<span data-ttu-id="ce329-229">Este punto de extensión coloca botones en la cinta para la extensión de módulo.</span><span class="sxs-lookup"><span data-stu-id="ce329-229">This extension point puts buttons on the ribbon for the module extension.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="ce329-230">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="ce329-230">Child elements</span></span>

|  <span data-ttu-id="ce329-231">Elemento</span><span class="sxs-lookup"><span data-stu-id="ce329-231">Element</span></span> |  <span data-ttu-id="ce329-232">Descripción</span><span class="sxs-lookup"><span data-stu-id="ce329-232">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ce329-233">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ce329-233">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="ce329-234">Agrega los comandos a la pestaña de la cinta predeterminada.</span><span class="sxs-lookup"><span data-stu-id="ce329-234">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="ce329-235">CustomTab</span><span class="sxs-lookup"><span data-stu-id="ce329-235">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="ce329-236">Agrega los comandos a la pestaña de la cinta personalizada.</span><span class="sxs-lookup"><span data-stu-id="ce329-236">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="ce329-237">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ce329-237">MobileMessageReadCommandSurface</span></span>
<span data-ttu-id="ce329-238">Este punto de extensión coloca botones en la superficie del comando de la vista de lectura de correo en el factor de forma móvil.</span><span class="sxs-lookup"><span data-stu-id="ce329-238">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="ce329-239">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="ce329-239">Child elements</span></span>

|  <span data-ttu-id="ce329-240">Elemento</span><span class="sxs-lookup"><span data-stu-id="ce329-240">Element</span></span> |  <span data-ttu-id="ce329-241">Descripción</span><span class="sxs-lookup"><span data-stu-id="ce329-241">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ce329-242">Group</span><span class="sxs-lookup"><span data-stu-id="ce329-242">Group</span></span>](group.md) |  <span data-ttu-id="ce329-243">Agrega un grupo de botones a la superficie del comando.</span><span class="sxs-lookup"><span data-stu-id="ce329-243">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="ce329-244">Los elementos **ExtensionPoint** de este tipo solo pueden tener un elemento secundario: un elemento **Group**.</span><span class="sxs-lookup"><span data-stu-id="ce329-244">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="ce329-245">Los elementos **Control** que se incluyen en este punto de extensión deben tener el atributo **xsi:type** establecido en `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="ce329-245">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="ce329-246">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ce329-246">Example</span></span>
```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="mobileGroupID">
    <Label resid="residAppName"/>
      <Control id="mobileButton1" xsi:type="MobileButton">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="events"></a><span data-ttu-id="ce329-247">Eventos</span><span class="sxs-lookup"><span data-stu-id="ce329-247">Events</span></span>

<span data-ttu-id="ce329-248">Este punto de extensión agrega un controlador de eventos para un evento especificado.</span><span class="sxs-lookup"><span data-stu-id="ce329-248">This extension point adds an event handler for a specified event.</span></span>

> [!NOTE]
> <span data-ttu-id="ce329-249">Este tipo de elemento sólo es compatible con Outlook en la web en Office 365.</span><span class="sxs-lookup"><span data-stu-id="ce329-249">This element type is only supported by Outlook on the web in Office 365.</span></span>

| <span data-ttu-id="ce329-250">Elemento</span><span class="sxs-lookup"><span data-stu-id="ce329-250">Element</span></span> | <span data-ttu-id="ce329-251">Descripción</span><span class="sxs-lookup"><span data-stu-id="ce329-251">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ce329-252">Event</span><span class="sxs-lookup"><span data-stu-id="ce329-252">Event</span></span>](event.md) |  <span data-ttu-id="ce329-253">Especifica el evento y la función del controlador de eventos.</span><span class="sxs-lookup"><span data-stu-id="ce329-253">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="ce329-254">Ejemplo del evento ItemSend</span><span class="sxs-lookup"><span data-stu-id="ce329-254">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events"> 
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
</ExtensionPoint> 
```

### <a name="detectedentity"></a><span data-ttu-id="ce329-255">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="ce329-255">DetectedEntity</span></span>

<span data-ttu-id="ce329-256">Este punto de extensión agrega una activación de complemento contextual en un tipo de entidad especificado.</span><span class="sxs-lookup"><span data-stu-id="ce329-256">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="ce329-257">El elemento que contenga [VersionOverrides](versionoverrides.md) debe tener un valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="ce329-257">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="ce329-258">Este tipo de elemento sólo es compatible con Outlook en la web en Office 365.</span><span class="sxs-lookup"><span data-stu-id="ce329-258">This element type is only supported by Outlook on the web in Office 365.</span></span>

|  <span data-ttu-id="ce329-259">Elemento</span><span class="sxs-lookup"><span data-stu-id="ce329-259">Element</span></span> |  <span data-ttu-id="ce329-260">Descripción</span><span class="sxs-lookup"><span data-stu-id="ce329-260">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ce329-261">Label</span><span class="sxs-lookup"><span data-stu-id="ce329-261">Label</span></span>](#label) |  <span data-ttu-id="ce329-262">Especifica la etiqueta del complemento en la ventana contextual.</span><span class="sxs-lookup"><span data-stu-id="ce329-262">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="ce329-263">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="ce329-263">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="ce329-264">Especifica la URL para la ventana contextual.</span><span class="sxs-lookup"><span data-stu-id="ce329-264">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="ce329-265">Rule</span><span class="sxs-lookup"><span data-stu-id="ce329-265">Rule</span></span>](rule.md) |  <span data-ttu-id="ce329-266">Especifica la regla o reglas que determinan cuándo se activa un complemento.</span><span class="sxs-lookup"><span data-stu-id="ce329-266">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="ce329-267">Label</span><span class="sxs-lookup"><span data-stu-id="ce329-267">Label</span></span>

<span data-ttu-id="ce329-p115">Obligatorio. La etiqueta del grupo. El atributo  **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento **ShortStrings** del elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="ce329-p115">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="ce329-271">Requisitos de resaltado</span><span class="sxs-lookup"><span data-stu-id="ce329-271">Highlight requirements</span></span>

<span data-ttu-id="ce329-p116">La única manera que tiene un usuario de activar un complemento contextual es para interactuar con una entidad resaltada. Los desarrolladores pueden controlar qué entidades están resaltadas con el atributo `Highlight` del elemento `Rule` para los tipos de regla `ItemHasKnownEntity` y `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="ce329-p116">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="ce329-p117">En cambio, hay algunas limitaciones que tener en cuenta. Estas limitaciones se usan para garantizar que siempre habrá una entidad resaltada en citas o mensajes aplicables para proporcionar al usuario una manera de activar el complemento.</span><span class="sxs-lookup"><span data-stu-id="ce329-p117">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="ce329-276">Los tipos de entidad `EmailAddress` y `Url` no pueden resaltarse, y por lo tanto, no pueden usarse para activar un complemento.</span><span class="sxs-lookup"><span data-stu-id="ce329-276">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="ce329-277">Si se usa una única regla, `Highlight` DEBE establecerse en `all`.</span><span class="sxs-lookup"><span data-stu-id="ce329-277">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="ce329-278">Si se usa un tipo de regla `RuleCollection` con `Mode="AND"` para combinar varias reglas, al menos una de las reglas DEBE tener `Highlight` establecido en `all`.</span><span class="sxs-lookup"><span data-stu-id="ce329-278">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="ce329-279">Si se usa un tipo de regla `RuleCollection` con `Mode="OR"` para combinar varias reglas, todas las reglas DEBEN tener `Highlight` establecido en `all`.</span><span class="sxs-lookup"><span data-stu-id="ce329-279">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="ce329-280">Ejemplo de evento DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="ce329-280">DetectedEntity event example</span></span>

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint> 
```