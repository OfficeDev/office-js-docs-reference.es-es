# <a name="icon-element"></a><span data-ttu-id="b29af-101">Elemento Icon</span><span class="sxs-lookup"><span data-stu-id="b29af-101">Icon element</span></span>

<span data-ttu-id="b29af-102">Defina los elementos **Image** para los controles [Button](control.md#button-control) o [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="b29af-102">Defines **Image** elements for [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="b29af-103">Atributos</span><span class="sxs-lookup"><span data-stu-id="b29af-103">Attributes</span></span>

|  <span data-ttu-id="b29af-104">Atributo</span><span class="sxs-lookup"><span data-stu-id="b29af-104">Attribute</span></span>  |  <span data-ttu-id="b29af-105">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="b29af-105">Required</span></span>  |  <span data-ttu-id="b29af-106">Descripción</span><span class="sxs-lookup"><span data-stu-id="b29af-106">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="b29af-107">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="b29af-107">**xsi:type**</span></span>  |  <span data-ttu-id="b29af-108">No</span><span class="sxs-lookup"><span data-stu-id="b29af-108">No</span></span>  | <span data-ttu-id="b29af-p101">El tipo de icono que se está definiendo. Esto se aplica a los iconos de factores de forma móvil. Los elementos **Icon** que se incluyen en un elemento [MobileFormFactor](mobileformfactor.md) deben tener este atributo establecido en `bt:MobileIconList`.</span><span class="sxs-lookup"><span data-stu-id="b29af-p101">The type of icon being defined. This is only applicable to icons in mobile form factors. **Icon** elements contained within a [MobileFormFactor](mobileformfactor.md) element must have this attribute set to `bt:MobileIconList`.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="b29af-112">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="b29af-112">Child elements</span></span>

|  <span data-ttu-id="b29af-113">Elemento</span><span class="sxs-lookup"><span data-stu-id="b29af-113">Element</span></span> |  <span data-ttu-id="b29af-114">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="b29af-114">Required</span></span>  |  <span data-ttu-id="b29af-115">Descripción</span><span class="sxs-lookup"><span data-stu-id="b29af-115">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b29af-116">Image</span><span class="sxs-lookup"><span data-stu-id="b29af-116">Image</span></span>](#image)        | <span data-ttu-id="b29af-117">Sí</span><span class="sxs-lookup"><span data-stu-id="b29af-117">Yes</span></span> |   <span data-ttu-id="b29af-118">resid de una imagen que se usará</span><span class="sxs-lookup"><span data-stu-id="b29af-118">resid of an image to use</span></span>         |

### <a name="image"></a><span data-ttu-id="b29af-119">Image</span><span class="sxs-lookup"><span data-stu-id="b29af-119">Image</span></span>

<span data-ttu-id="b29af-p102">Imagen del botón. El atributo **resid** tiene que establecerse en el valor del atributo **id** de un elemento **Image** en el elemento **Images** del elemento [Resources](resources.md). El atributo **size** indica el tamaño en píxeles de la imagen. Se necesitan tres tamaños de imágenes (16, 32 y 80 píxeles), mientras que se admiten otros cinco tamaños (20, 24, 40, 48 y 64 píxeles).|</span><span class="sxs-lookup"><span data-stu-id="b29af-p102">An image for the button. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the [Resources](resources.md) element. The **size** attribute indicates the size in pixels of the image. Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).|</span></span>

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

## <a name="additional-requirements-for-mobile-form-factors"></a><span data-ttu-id="b29af-124">Requisitos adicionales para los factores de forma móvil</span><span class="sxs-lookup"><span data-stu-id="b29af-124">Additional requirements for mobile form factors</span></span>

<span data-ttu-id="b29af-p103">Cuando el elemento **Icon** primario es un descendiente de un elemento [MobileFormFactor](mobileformfactor.md), los tamaños mínimos necesarios son ligeramente diferentes. El manifiesto debe proporcionar como mínimo tamaños de 25, 32 y 48 píxeles. Cada tamaño proporcionado debe aparecer tres veces, con un atributo `scale` establecido en `1`, `2` o `3`.</span><span class="sxs-lookup"><span data-stu-id="b29af-p103">When the parent **Icon** element is a descendant of a [MobileFormFactor](mobileformfactor.md) element, the minimum required sizes are slightly different. The manifest must minimally provide 25, 32, and 48 pixel sizes. Each size provided must appear three times, with a `scale` attribute set to `1`, `2`, or `3`.</span></span>

```xml
<Icon xsi:type="bt:MobileIconList">
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
```