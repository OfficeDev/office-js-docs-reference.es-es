# <a name="script-element"></a><span data-ttu-id="cf621-101">Elemento Script</span><span class="sxs-lookup"><span data-stu-id="cf621-101">Script element</span></span>

<span data-ttu-id="cf621-102">Define la configuración de scripts que usa una función personalizada en Excel.</span><span class="sxs-lookup"><span data-stu-id="cf621-102">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="cf621-103">Atributos</span><span class="sxs-lookup"><span data-stu-id="cf621-103">Attributes</span></span>

<span data-ttu-id="cf621-104">Ninguno</span><span class="sxs-lookup"><span data-stu-id="cf621-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="cf621-105">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="cf621-105">Child elements</span></span>

|<span data-ttu-id="cf621-106">Elementos</span><span class="sxs-lookup"><span data-stu-id="cf621-106">Elements</span></span>  |  <span data-ttu-id="cf621-107">Necesario</span><span class="sxs-lookup"><span data-stu-id="cf621-107">Required</span></span>  |  <span data-ttu-id="cf621-108">Descripción</span><span class="sxs-lookup"><span data-stu-id="cf621-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cf621-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="cf621-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="cf621-110">Sí</span><span class="sxs-lookup"><span data-stu-id="cf621-110">Yes</span></span>  | <span data-ttu-id="cf621-111">Cadena con el identificador de recurso del archivo JavaScript usado por funciones personalizadas.</span><span class="sxs-lookup"><span data-stu-id="cf621-111">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="cf621-112">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="cf621-112">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
