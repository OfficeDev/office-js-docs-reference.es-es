# <a name="group-element"></a><span data-ttu-id="5f75b-101">Elemento Group</span><span class="sxs-lookup"><span data-stu-id="5f75b-101">Group element</span></span>

<span data-ttu-id="5f75b-p101">Define un grupo de controles de interfaz de usuario en un pestaña.  En las pestañas personalizadas, el complemento puede crear hasta 10 grupos. Cada grupo está limitado a 6 controles, independientemente de la pestaña donde aparezca. Los complementos están limitados a una pestaña personalizada.</span><span class="sxs-lookup"><span data-stu-id="5f75b-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="5f75b-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="5f75b-105">Attributes</span></span>

|  <span data-ttu-id="5f75b-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="5f75b-106">Attribute</span></span>  |  <span data-ttu-id="5f75b-107">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="5f75b-107">Required</span></span>  |  <span data-ttu-id="5f75b-108">Descripción</span><span class="sxs-lookup"><span data-stu-id="5f75b-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5f75b-109">id</span><span class="sxs-lookup"><span data-stu-id="5f75b-109">id</span></span>](#id-attribute)  |  <span data-ttu-id="5f75b-110">Sí</span><span class="sxs-lookup"><span data-stu-id="5f75b-110">Yes</span></span>  | <span data-ttu-id="5f75b-111">Un identificador único para el grupo.</span><span class="sxs-lookup"><span data-stu-id="5f75b-111">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="5f75b-112">Atributo id</span><span class="sxs-lookup"><span data-stu-id="5f75b-112">id attribute</span></span>

<span data-ttu-id="5f75b-p102">Necesario. Identificador único para el grupo. Es una cadena con un máximo de 125 caracteres. Debe ser único dentro del manifiesto o el grupo no podrá procesarse.</span><span class="sxs-lookup"><span data-stu-id="5f75b-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="5f75b-117">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="5f75b-117">Child elements</span></span>
|  <span data-ttu-id="5f75b-118">Elemento</span><span class="sxs-lookup"><span data-stu-id="5f75b-118">Element</span></span> |  <span data-ttu-id="5f75b-119">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="5f75b-119">Required</span></span>  |  <span data-ttu-id="5f75b-120">Descripción</span><span class="sxs-lookup"><span data-stu-id="5f75b-120">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5f75b-121">Label</span><span class="sxs-lookup"><span data-stu-id="5f75b-121">Label</span></span>](#label)      | <span data-ttu-id="5f75b-122">Sí</span><span class="sxs-lookup"><span data-stu-id="5f75b-122">Yes</span></span> |  <span data-ttu-id="5f75b-123">La etiqueta de CustomTab o de un grupo.</span><span class="sxs-lookup"><span data-stu-id="5f75b-123">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="5f75b-124">Control</span><span class="sxs-lookup"><span data-stu-id="5f75b-124">Control</span></span>](#control)    | <span data-ttu-id="5f75b-125">Sí</span><span class="sxs-lookup"><span data-stu-id="5f75b-125">Yes</span></span> |  <span data-ttu-id="5f75b-126">Colección de uno o más objetos Control.</span><span class="sxs-lookup"><span data-stu-id="5f75b-126">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="5f75b-127">Label</span><span class="sxs-lookup"><span data-stu-id="5f75b-127">Label</span></span> 

<span data-ttu-id="5f75b-p103">Obligatorio. La etiqueta del grupo. El atributo  **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento **ShortStrings** del elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="5f75b-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="5f75b-131">Control</span><span class="sxs-lookup"><span data-stu-id="5f75b-131">Control</span></span>
<span data-ttu-id="5f75b-132">Un grupo necesita como mínimo un control.</span><span class="sxs-lookup"><span data-stu-id="5f75b-132">A group requires at least one control.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```