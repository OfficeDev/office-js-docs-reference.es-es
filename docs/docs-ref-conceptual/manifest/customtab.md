# <a name="customtab-element"></a><span data-ttu-id="26678-101">Elemento CustomTab</span><span class="sxs-lookup"><span data-stu-id="26678-101">CustomTab element</span></span>

<span data-ttu-id="26678-p101">En la cinta, especifique la pestaña y el grupo para sus comandos de complemento. Puede ser en la pestaña predeterminada (**Inicio**, **Mensaje** o **Reunión**), o en una pestaña personalizada definida por el complemento.</span><span class="sxs-lookup"><span data-stu-id="26678-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="26678-p102">En las pestañas personalizadas, el complemento puede crear hasta 10 grupos. Cada grupo está limitado a 6 controles, independientemente de la pestaña donde aparezca. Los complementos están limitados a una pestaña personalizada.</span><span class="sxs-lookup"><span data-stu-id="26678-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="26678-107">El atributo  **id** debe ser único en el manifiesto.</span><span class="sxs-lookup"><span data-stu-id="26678-107">The  **id** attribute must be unique within the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="26678-108">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="26678-108">Child elements</span></span>

|  <span data-ttu-id="26678-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="26678-109">Element</span></span> |  <span data-ttu-id="26678-110">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="26678-110">Required</span></span>  |  <span data-ttu-id="26678-111">Descripción</span><span class="sxs-lookup"><span data-stu-id="26678-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="26678-112">Group</span><span class="sxs-lookup"><span data-stu-id="26678-112">Group</span></span>](group.md)      | <span data-ttu-id="26678-113">Sí</span><span class="sxs-lookup"><span data-stu-id="26678-113">Yes</span></span> |  <span data-ttu-id="26678-114">Define un grupo de comandos.</span><span class="sxs-lookup"><span data-stu-id="26678-114">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="26678-115">Label</span><span class="sxs-lookup"><span data-stu-id="26678-115">Label</span></span>](#label-tab)      | <span data-ttu-id="26678-116">Sí</span><span class="sxs-lookup"><span data-stu-id="26678-116">Yes</span></span> |  <span data-ttu-id="26678-117">La etiqueta de CustomTab o de un grupo.</span><span class="sxs-lookup"><span data-stu-id="26678-117">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="26678-118">Control</span><span class="sxs-lookup"><span data-stu-id="26678-118">Control</span></span>](control.md)    | <span data-ttu-id="26678-119">Sí</span><span class="sxs-lookup"><span data-stu-id="26678-119">Yes</span></span> |  <span data-ttu-id="26678-120">Una colección de uno o más objetos Control.</span><span class="sxs-lookup"><span data-stu-id="26678-120">A collection of one or more Control objects.</span></span>  |

### <a name="group"></a><span data-ttu-id="26678-121">Group</span><span class="sxs-lookup"><span data-stu-id="26678-121">Group</span></span>

<span data-ttu-id="26678-p103">Necesario. Consulte [elemento Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="26678-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="26678-124">Etiqueta (pestaña)</span><span class="sxs-lookup"><span data-stu-id="26678-124">Label (Tab)</span></span>

<span data-ttu-id="26678-p104">Obligatorio. La etiqueta de la pestaña personalizada. El atributo  **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento **ShortStrings** del elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="26678-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="26678-127">Ejemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="26678-127">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```