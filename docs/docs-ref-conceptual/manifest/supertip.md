# <a name="supertip"></a><span data-ttu-id="b3c84-101">Supertip</span><span class="sxs-lookup"><span data-stu-id="b3c84-101">Supertip</span></span>

<span data-ttu-id="b3c84-p101">Defina una información sobre herramientas enriquecida (título y descripción). Lo usan los controles [Button](control.md#button-control) y [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="b3c84-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b3c84-104">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="b3c84-104">Child elements</span></span>

|  <span data-ttu-id="b3c84-105">Elemento</span><span class="sxs-lookup"><span data-stu-id="b3c84-105">Element</span></span> |  <span data-ttu-id="b3c84-106">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="b3c84-106">Required</span></span>  |  <span data-ttu-id="b3c84-107">Descripción</span><span class="sxs-lookup"><span data-stu-id="b3c84-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b3c84-108">Título</span><span class="sxs-lookup"><span data-stu-id="b3c84-108">Title</span></span>](#title)        | <span data-ttu-id="b3c84-109">Sí</span><span class="sxs-lookup"><span data-stu-id="b3c84-109">Yes</span></span> |   <span data-ttu-id="b3c84-110">El texto de la sugerencia.</span><span class="sxs-lookup"><span data-stu-id="b3c84-110">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="b3c84-111">Descripción</span><span class="sxs-lookup"><span data-stu-id="b3c84-111">Description</span></span>](#description)  | <span data-ttu-id="b3c84-112">Sí</span><span class="sxs-lookup"><span data-stu-id="b3c84-112">Yes</span></span> |  <span data-ttu-id="b3c84-113">La descripción de la sugerencia.</span><span class="sxs-lookup"><span data-stu-id="b3c84-113">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="b3c84-114">Título</span><span class="sxs-lookup"><span data-stu-id="b3c84-114">Title</span></span>

<span data-ttu-id="b3c84-p102">Obligatorio. Texto para la sugerencia falsa. El atributo  **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento **ShortStrings** del elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="b3c84-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="b3c84-118">Descripción</span><span class="sxs-lookup"><span data-stu-id="b3c84-118">Description</span></span>

<span data-ttu-id="b3c84-p103">Obligatorio. Descripción para la sugerencia falsa. El atributo  **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento **LongStrings** del elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="b3c84-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="b3c84-122">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="b3c84-122">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
