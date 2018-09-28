# <a name="metadata-element"></a><span data-ttu-id="aca7b-101">Elemento de metadatos</span><span class="sxs-lookup"><span data-stu-id="aca7b-101">Metadata element</span></span>

<span data-ttu-id="aca7b-102">Define la configuración de metadatos usada por una función personalizada en Excel.</span><span class="sxs-lookup"><span data-stu-id="aca7b-102">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="aca7b-103">Atributos</span><span class="sxs-lookup"><span data-stu-id="aca7b-103">Attributes</span></span>

<span data-ttu-id="aca7b-104">Ninguno</span><span class="sxs-lookup"><span data-stu-id="aca7b-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="aca7b-105">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="aca7b-105">Child elements</span></span>

|  <span data-ttu-id="aca7b-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="aca7b-106">Element</span></span>  |  <span data-ttu-id="aca7b-107">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="aca7b-107">Required</span></span>  |  <span data-ttu-id="aca7b-108">Descripción</span><span class="sxs-lookup"><span data-stu-id="aca7b-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="aca7b-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="aca7b-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="aca7b-110">Sí</span><span class="sxs-lookup"><span data-stu-id="aca7b-110">Yes</span></span>  | <span data-ttu-id="aca7b-111">Cadena con el identificador de recurso del archivo JSON usado por funciones personalizadas.</span><span class="sxs-lookup"><span data-stu-id="aca7b-111">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="aca7b-112">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="aca7b-112">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
