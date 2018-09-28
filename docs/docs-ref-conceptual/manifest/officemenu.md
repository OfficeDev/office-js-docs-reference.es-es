# <a name="officemenu-element"></a><span data-ttu-id="cf42e-101">Elemento OfficeMenu</span><span class="sxs-lookup"><span data-stu-id="cf42e-101">OfficeMenu element</span></span>

<span data-ttu-id="cf42e-p101">Define una colección de controles que se agregarán al menú contextual de Office. Se aplica a los complementos de Word, Excel, PowerPoint y OneNote.</span><span class="sxs-lookup"><span data-stu-id="cf42e-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="cf42e-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="cf42e-104">Attributes</span></span>

| <span data-ttu-id="cf42e-105">Atributo</span><span class="sxs-lookup"><span data-stu-id="cf42e-105">Attribute</span></span>            | <span data-ttu-id="cf42e-106">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="cf42e-106">Required</span></span> | <span data-ttu-id="cf42e-107">Descripción</span><span class="sxs-lookup"><span data-stu-id="cf42e-107">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="cf42e-108">xsi:type</span><span class="sxs-lookup"><span data-stu-id="cf42e-108">xsi:type</span></span>](#xsitype) | <span data-ttu-id="cf42e-109">Sí</span><span class="sxs-lookup"><span data-stu-id="cf42e-109">Yes</span></span>      | <span data-ttu-id="cf42e-110">Tipo de OfficeMenu que va a definir.</span><span class="sxs-lookup"><span data-stu-id="cf42e-110">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="cf42e-111">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="cf42e-111">Child elements</span></span>

|  <span data-ttu-id="cf42e-112">Elemento</span><span class="sxs-lookup"><span data-stu-id="cf42e-112">Element</span></span> |  <span data-ttu-id="cf42e-113">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="cf42e-113">Required</span></span>  |  <span data-ttu-id="cf42e-114">Descripción</span><span class="sxs-lookup"><span data-stu-id="cf42e-114">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cf42e-115">Control</span><span class="sxs-lookup"><span data-stu-id="cf42e-115">Control</span></span>](#control)    | <span data-ttu-id="cf42e-116">Sí</span><span class="sxs-lookup"><span data-stu-id="cf42e-116">Yes</span></span> |  <span data-ttu-id="cf42e-117">Una colección de uno o varios objetos Control.</span><span class="sxs-lookup"><span data-stu-id="cf42e-117">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="cf42e-118">xsi:type</span><span class="sxs-lookup"><span data-stu-id="cf42e-118">xsi:type</span></span>

<span data-ttu-id="cf42e-119">Especifica un menú integrado de la aplicación de cliente de Office en el que se va a agregar este complemento de Office.</span><span class="sxs-lookup"><span data-stu-id="cf42e-119">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="cf42e-p102">`ContextMenuText` - Muestra el elemento en el menú contextual cuando se selecciona un texto y el usuario abre el menú contextual (con el botón derecho) sobre el texto seleccionado. Se aplica a Word, Excel, PowerPoint y OneNote.</span><span class="sxs-lookup"><span data-stu-id="cf42e-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="cf42e-p103">`ContextMenuCell` - Muestra el elemento en el menú contextual cuando el usuario abre el menú contextual (hace clic con el botón derecho) en una celda de la hoja de cálculo. Se aplica a Excel.</span><span class="sxs-lookup"><span data-stu-id="cf42e-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="cf42e-124">Control</span><span class="sxs-lookup"><span data-stu-id="cf42e-124">Control</span></span>

<span data-ttu-id="cf42e-125">Cada elemento **OfficeMenu** requiere por lo menos uno o varios controles de [menú](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="cf42e-125">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="cf42e-126">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="cf42e-126">Example</span></span>

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>   
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>    
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>    
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />    
          </Action>    
        </Item>
      </Items>
    </Control>   
</OfficeMenu>
```
