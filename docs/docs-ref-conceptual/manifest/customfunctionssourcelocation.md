# <a name="sourcelocation-element"></a><span data-ttu-id="f6593-101">Elemento SourceLocation</span><span class="sxs-lookup"><span data-stu-id="f6593-101">SourceLocation element</span></span>

<span data-ttu-id="f6593-102">Define la ubicación de un recurso que requieren los elementos Script o Page usados por funciones personalizadas de Excel.</span><span class="sxs-lookup"><span data-stu-id="f6593-102">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="f6593-103">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6593-103">Attributes</span></span>

| <span data-ttu-id="f6593-104">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="f6593-104">**Attribute**</span></span> | <span data-ttu-id="f6593-105">**Necesario**</span><span class="sxs-lookup"><span data-stu-id="f6593-105">**Required**</span></span> | <span data-ttu-id="f6593-106">**Descripción**</span><span class="sxs-lookup"><span data-stu-id="f6593-106">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="f6593-107">resid</span><span class="sxs-lookup"><span data-stu-id="f6593-107">resid</span></span>         | <span data-ttu-id="f6593-108">Sí</span><span class="sxs-lookup"><span data-stu-id="f6593-108">Yes</span></span>          | <span data-ttu-id="f6593-109">Nombre de un recurso de URL definido en la sección &lt;Recursos&gt; del manifiesto.</span><span class="sxs-lookup"><span data-stu-id="f6593-109">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="f6593-110">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="f6593-110">Child elements</span></span>

<span data-ttu-id="f6593-111">Ninguno</span><span class="sxs-lookup"><span data-stu-id="f6593-111">None</span></span>

## <a name="example"></a><span data-ttu-id="f6593-112">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="f6593-112">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```