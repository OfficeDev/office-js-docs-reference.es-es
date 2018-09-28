# <a name="allformfactors-element"></a><span data-ttu-id="6cb08-101">Elemento AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="6cb08-101">AllFormFactors element</span></span>

<span data-ttu-id="6cb08-102">Especifica la configuración de un complemento para todos los factores de forma.</span><span class="sxs-lookup"><span data-stu-id="6cb08-102">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="6cb08-103">Actualmente, la única característica con **AllFormFactors** es funciones personalizadas.</span><span class="sxs-lookup"><span data-stu-id="6cb08-103">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="6cb08-104">**AllFormFactors** es un elemento obligatorio al uso de funciones personalizadas.</span><span class="sxs-lookup"><span data-stu-id="6cb08-104">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="6cb08-105">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="6cb08-105">Child elements</span></span>

|  <span data-ttu-id="6cb08-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="6cb08-106">Element</span></span> |  <span data-ttu-id="6cb08-107">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="6cb08-107">Required</span></span>  |  <span data-ttu-id="6cb08-108">Descripción</span><span class="sxs-lookup"><span data-stu-id="6cb08-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="6cb08-109">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="6cb08-109">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="6cb08-110">Sí</span><span class="sxs-lookup"><span data-stu-id="6cb08-110">Yes</span></span> |  <span data-ttu-id="6cb08-111">Define dónde expone su funcionalidad un complemento.</span><span class="sxs-lookup"><span data-stu-id="6cb08-111">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="6cb08-112">Ejemplo de AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="6cb08-112">AllFormFactors example</span></span>

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
