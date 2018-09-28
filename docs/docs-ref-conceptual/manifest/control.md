# <a name="control-element"></a><span data-ttu-id="2e73d-101">Elemento Control</span><span class="sxs-lookup"><span data-stu-id="2e73d-101">Control element</span></span>

<span data-ttu-id="2e73d-p101">Define una función de JavaScript que ejecuta una acción o inicia un panel de tareas. Un elemento **Control** puede ser un botón o una opción de menú. Al menos un **Control** debe estar incluido en un elemento [Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="2e73d-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="2e73d-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="2e73d-105">Attributes</span></span>

|  <span data-ttu-id="2e73d-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="2e73d-106">Attribute</span></span>  |  <span data-ttu-id="2e73d-107">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="2e73d-107">Required</span></span>  |  <span data-ttu-id="2e73d-108">Descripción</span><span class="sxs-lookup"><span data-stu-id="2e73d-108">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="2e73d-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="2e73d-109">**xsi:type**</span></span>|<span data-ttu-id="2e73d-110">Sí</span><span class="sxs-lookup"><span data-stu-id="2e73d-110">Yes</span></span>|<span data-ttu-id="2e73d-p102">El tipo de control que se está definiendo. Puede ser `Button`, `Menu` o `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="2e73d-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="2e73d-113">**id**</span><span class="sxs-lookup"><span data-stu-id="2e73d-113">**id**</span></span>|<span data-ttu-id="2e73d-114">No</span><span class="sxs-lookup"><span data-stu-id="2e73d-114">No</span></span>|<span data-ttu-id="2e73d-p103">El identificador del elemento de control. Puede tener 125 caracteres como máximo.</span><span class="sxs-lookup"><span data-stu-id="2e73d-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="2e73d-117">El `MobileButton` el valor de **xsi: Type** se define en el esquema de VersionOverrides 1.1.</span><span class="sxs-lookup"><span data-stu-id="2e73d-117">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="2e73d-118">Sólo se aplica a los elementos de **Control de** contenido dentro de un elemento [MobileFormFactor](mobileformfactor.md) .</span><span class="sxs-lookup"><span data-stu-id="2e73d-118">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="2e73d-119">Control de botón</span><span class="sxs-lookup"><span data-stu-id="2e73d-119">Button control</span></span>

<span data-ttu-id="2e73d-p105">Un botón realiza una sola acción cuando el usuario lo selecciona. Puede ejecutar una función o mostrar un panel de tareas. Cada control de botón ha que tener un `id` único para el manifiesto.</span><span class="sxs-lookup"><span data-stu-id="2e73d-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="2e73d-123">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="2e73d-123">Child elements</span></span>
|  <span data-ttu-id="2e73d-124">Elemento</span><span class="sxs-lookup"><span data-stu-id="2e73d-124">Element</span></span> |  <span data-ttu-id="2e73d-125">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="2e73d-125">Required</span></span>  |  <span data-ttu-id="2e73d-126">Descripción</span><span class="sxs-lookup"><span data-stu-id="2e73d-126">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2e73d-127">**Label**</span><span class="sxs-lookup"><span data-stu-id="2e73d-127">**Label**</span></span>     | <span data-ttu-id="2e73d-128">Sí</span><span class="sxs-lookup"><span data-stu-id="2e73d-128">Yes</span></span> |  <span data-ttu-id="2e73d-p106">El texto del botón. El atributo **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento **ShortStrings** del elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="2e73d-p106">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="2e73d-131">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="2e73d-131">**ToolTip**</span></span>  |<span data-ttu-id="2e73d-132">No</span><span class="sxs-lookup"><span data-stu-id="2e73d-132">No</span></span>|<span data-ttu-id="2e73d-p107">La información sobre herramientas del botón. El atributo **resid** tiene que establecerse en el valor del atributo **id** de un elemento **String**. El elemento **String** es un elemento secundario del elemento **LongStrings**, que es un elemento secundario del elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="2e73d-p107">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="2e73d-136">Supertip</span><span class="sxs-lookup"><span data-stu-id="2e73d-136">Supertip</span></span>](supertip.md)  | <span data-ttu-id="2e73d-137">Sí</span><span class="sxs-lookup"><span data-stu-id="2e73d-137">Yes</span></span> |  <span data-ttu-id="2e73d-138">La sugerencia del botón.</span><span class="sxs-lookup"><span data-stu-id="2e73d-138">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="2e73d-139">Icon</span><span class="sxs-lookup"><span data-stu-id="2e73d-139">Icon</span></span>](icon.md)      | <span data-ttu-id="2e73d-140">Sí</span><span class="sxs-lookup"><span data-stu-id="2e73d-140">Yes</span></span> |  <span data-ttu-id="2e73d-141">Una imagen del botón.</span><span class="sxs-lookup"><span data-stu-id="2e73d-141">An image for the button.</span></span>         |
|  [<span data-ttu-id="2e73d-142">Action</span><span class="sxs-lookup"><span data-stu-id="2e73d-142">Action</span></span>](action.md)    | <span data-ttu-id="2e73d-143">Sí</span><span class="sxs-lookup"><span data-stu-id="2e73d-143">Yes</span></span> |  <span data-ttu-id="2e73d-144">Especifica la acción que se realizará.</span><span class="sxs-lookup"><span data-stu-id="2e73d-144">Specifies the action to perform.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="2e73d-145">Ejemplo de botón ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="2e73d-145">ExecuteFunction button example</span></span>

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-button-example"></a><span data-ttu-id="2e73d-146">Ejemplo de botón ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="2e73d-146">ShowTaskpane button example</span></span>

```xml
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="2e73d-147">Controles de menú (botón desplegable)</span><span class="sxs-lookup"><span data-stu-id="2e73d-147">Menu (dropdown button) controls</span></span>

<span data-ttu-id="2e73d-p108">Un menú define una lista estática de opciones. Cada elemento de menú ejecuta una función o muestra un panel de tareas. No se admiten submenús.</span><span class="sxs-lookup"><span data-stu-id="2e73d-p108">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="2e73d-151">Cuando se usa con un [punto de extensión](extensionpoint.md) **PrimaryCommandSurface** o **ContextMenu**, el control de menú define:</span><span class="sxs-lookup"><span data-stu-id="2e73d-151">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="2e73d-152">Un elemento de menú de nivel de raíz.</span><span class="sxs-lookup"><span data-stu-id="2e73d-152">A root-level menu item.</span></span>

- <span data-ttu-id="2e73d-153">Una lista de elementos de submenú.</span><span class="sxs-lookup"><span data-stu-id="2e73d-153">A list of submenu items.</span></span>

<span data-ttu-id="2e73d-p109">Cuando se usa con  **PrimaryCommandSurface**, el elemento de menú raíz se muestra como un botón en la cinta de opciones. Cuando el botón está seleccionado, el submenú se muestra como una lista desplegable. Cuando se usa con  **ContextMenu**, se inserta un elemento de menú con un submenú en el menú contextual. En ambos casos, los elementos individuales del submenú pueden ejecutar una función de JavaScript o mostrar un panel de tareas. Por ahora solo se admite un nivel de submenús.</span><span class="sxs-lookup"><span data-stu-id="2e73d-p109">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with  **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="2e73d-p110">En el ejemplo siguiente se muestra cómo definir un elemento de menú con dos elementos de submenú. El primer elemento de submenú muestra un panel de tareas y el segundo elemento de submenú ejecuta una función JavaScript.</span><span class="sxs-lookup"><span data-stu-id="2e73d-p110">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

### <a name="child-elements"></a><span data-ttu-id="2e73d-161">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="2e73d-161">Child elements</span></span>

|  <span data-ttu-id="2e73d-162">Elemento</span><span class="sxs-lookup"><span data-stu-id="2e73d-162">Element</span></span> |  <span data-ttu-id="2e73d-163">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="2e73d-163">Required</span></span>  |  <span data-ttu-id="2e73d-164">Descripción</span><span class="sxs-lookup"><span data-stu-id="2e73d-164">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2e73d-165">**Label**</span><span class="sxs-lookup"><span data-stu-id="2e73d-165">**Label**</span></span>     | <span data-ttu-id="2e73d-166">Sí</span><span class="sxs-lookup"><span data-stu-id="2e73d-166">Yes</span></span> |  <span data-ttu-id="2e73d-p111">El texto del botón. El atributo **resid** debe establecerse en el valor del atributo **id** de un elemento **String** en el elemento **ShortStrings** del elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="2e73d-p111">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="2e73d-169">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="2e73d-169">**ToolTip**</span></span>  |<span data-ttu-id="2e73d-170">No</span><span class="sxs-lookup"><span data-stu-id="2e73d-170">No</span></span>|<span data-ttu-id="2e73d-p112">La información sobre herramientas del botón. El atributo **resid** tiene que establecerse en el valor del atributo **id** de un elemento **String**. El elemento **String** es un elemento secundario del elemento **LongStrings**, que es un elemento secundario del elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="2e73d-p112">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="2e73d-174">Supertip</span><span class="sxs-lookup"><span data-stu-id="2e73d-174">Supertip</span></span>](supertip.md)  | <span data-ttu-id="2e73d-175">Sí</span><span class="sxs-lookup"><span data-stu-id="2e73d-175">Yes</span></span> |  <span data-ttu-id="2e73d-176">La sugerencia de este botón.</span><span class="sxs-lookup"><span data-stu-id="2e73d-176">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="2e73d-177">Icon</span><span class="sxs-lookup"><span data-stu-id="2e73d-177">Icon</span></span>](icon.md)      | <span data-ttu-id="2e73d-178">Sí</span><span class="sxs-lookup"><span data-stu-id="2e73d-178">Yes</span></span> |  <span data-ttu-id="2e73d-179">Una imagen del botón.</span><span class="sxs-lookup"><span data-stu-id="2e73d-179">An image for the button.</span></span>         |
|  <span data-ttu-id="2e73d-180">**Items**</span><span class="sxs-lookup"><span data-stu-id="2e73d-180">**Items**</span></span>     | <span data-ttu-id="2e73d-181">Sí</span><span class="sxs-lookup"><span data-stu-id="2e73d-181">Yes</span></span> |  <span data-ttu-id="2e73d-p113">Una colección de botones que se mostrarán en el menú Contiene los elementos **Item** para cada elemento de submenú. Cada elemento **Item** contiene los elementos secundarios del [control Button](#button-control).</span><span class="sxs-lookup"><span data-stu-id="2e73d-p113">A collection of Buttons to display within the menu. Contains the  **Item** elements for each submenu item. Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="2e73d-185">Ejemplos de control de menú</span><span class="sxs-lookup"><span data-stu-id="2e73d-185">Menu control examples</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

```xml
<Control xsi:type="Menu" id="msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="msgReadMenuItem1">
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

## <a name="mobilebutton-control"></a><span data-ttu-id="2e73d-186">Control MobileButton</span><span class="sxs-lookup"><span data-stu-id="2e73d-186">MobileButton control</span></span>

<span data-ttu-id="2e73d-p114">Un botón móvil realiza una sola acción cuando el usuario lo selecciona. Puede ejecutar una función o mostrar un panel de tareas. Cada control de botón móvil debe tener un `id` único para el manifiesto.</span><span class="sxs-lookup"><span data-stu-id="2e73d-p114">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="2e73d-p115">El valor `MobileButton` de **xsi:type** se define en el esquema VersionOverrides 1.1. El elemento que contenga [VersionOverrides](versionoverrides.md) debe tener un valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="2e73d-p115">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="2e73d-192">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="2e73d-192">Child elements</span></span>
|  <span data-ttu-id="2e73d-193">Elemento</span><span class="sxs-lookup"><span data-stu-id="2e73d-193">Element</span></span> |  <span data-ttu-id="2e73d-194">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="2e73d-194">Required</span></span>  |  <span data-ttu-id="2e73d-195">Descripción</span><span class="sxs-lookup"><span data-stu-id="2e73d-195">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2e73d-196">**Label**</span><span class="sxs-lookup"><span data-stu-id="2e73d-196">**Label**</span></span>     | <span data-ttu-id="2e73d-197">Sí</span><span class="sxs-lookup"><span data-stu-id="2e73d-197">Yes</span></span> |  <span data-ttu-id="2e73d-p116">El texto del botón. El atributo **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento **ShortStrings** del elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="2e73d-p116">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="2e73d-200">Icon</span><span class="sxs-lookup"><span data-stu-id="2e73d-200">Icon</span></span>](icon.md)      | <span data-ttu-id="2e73d-201">Sí</span><span class="sxs-lookup"><span data-stu-id="2e73d-201">Yes</span></span> |  <span data-ttu-id="2e73d-202">Una imagen del botón.</span><span class="sxs-lookup"><span data-stu-id="2e73d-202">An image for the button.</span></span>         |
|  [<span data-ttu-id="2e73d-203">Action</span><span class="sxs-lookup"><span data-stu-id="2e73d-203">Action</span></span>](action.md)    | <span data-ttu-id="2e73d-204">Sí</span><span class="sxs-lookup"><span data-stu-id="2e73d-204">Yes</span></span> |  <span data-ttu-id="2e73d-205">Especifica la acción que se realizará.</span><span class="sxs-lookup"><span data-stu-id="2e73d-205">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="2e73d-206">Ejemplo de botón móvil ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="2e73d-206">ExecuteFunction mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="2e73d-207">Ejemplo de botón móvil ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="2e73d-207">ShowTaskpane mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```