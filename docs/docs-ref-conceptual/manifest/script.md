# <a name="script-element"></a><span data-ttu-id="d7609-101">Elemento Script</span><span class="sxs-lookup"><span data-stu-id="d7609-101">Script element</span></span>

<span data-ttu-id="d7609-102">Define la configuración de scripts que usa una función personalizada en Excel.</span><span class="sxs-lookup"><span data-stu-id="d7609-102">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="d7609-103">Atributos</span><span class="sxs-lookup"><span data-stu-id="d7609-103">Attributes</span></span>

<span data-ttu-id="d7609-104">Ninguno</span><span class="sxs-lookup"><span data-stu-id="d7609-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="d7609-105">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="d7609-105">Child elements</span></span>

|<span data-ttu-id="d7609-106">Elementos</span><span class="sxs-lookup"><span data-stu-id="d7609-106">Elements</span></span>  |  <span data-ttu-id="d7609-107">Necesario</span><span class="sxs-lookup"><span data-stu-id="d7609-107">Required</span></span>  |  <span data-ttu-id="d7609-108">Descripción</span><span class="sxs-lookup"><span data-stu-id="d7609-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d7609-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="d7609-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="d7609-110">Sí</span><span class="sxs-lookup"><span data-stu-id="d7609-110">Yes</span></span>  | <span data-ttu-id="d7609-111">Cadena con el identificador de recurso del archivo JavaScript usado por funciones personalizadas.</span><span class="sxs-lookup"><span data-stu-id="d7609-111">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="d7609-112">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="d7609-112">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
