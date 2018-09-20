# <a name="desktopformfactor-element"></a><span data-ttu-id="eba30-101">Elemento DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="eba30-101">DesktopFormFactor element</span></span>

<span data-ttu-id="eba30-p101">Especifica la configuración de un complemento para un factor de forma de escritorio. El factor de forma de escritorio incluye Office para Windows, Office para Mac y Office Online. Contiene toda la información sobre complementos para dicho factor de forma, excepto para el nodo **Resources**.</span><span class="sxs-lookup"><span data-stu-id="eba30-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="eba30-p102">Cada definición DesktopFormFactor contiene el elemento **FunctionFile** y uno o más elementos **ExtensionPoint**. Para obtener más información, consulte [Elemento FunctionFile](functionfile.md) y [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="eba30-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span> 

## <a name="child-elements"></a><span data-ttu-id="eba30-107">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="eba30-107">Child elements</span></span>

| <span data-ttu-id="eba30-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="eba30-108">Element</span></span>                               | <span data-ttu-id="eba30-109">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="eba30-109">Required</span></span> | <span data-ttu-id="eba30-110">Descripción</span><span class="sxs-lookup"><span data-stu-id="eba30-110">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="eba30-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="eba30-111">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="eba30-112">Sí</span><span class="sxs-lookup"><span data-stu-id="eba30-112">Yes</span></span>      | <span data-ttu-id="eba30-113">Define dónde expone su funcionalidad un complemento.</span><span class="sxs-lookup"><span data-stu-id="eba30-113">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="eba30-114">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="eba30-114">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="eba30-115">Sí</span><span class="sxs-lookup"><span data-stu-id="eba30-115">Yes</span></span>      | <span data-ttu-id="eba30-116">Una dirección URL de un archivo que contiene funciones de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="eba30-116">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="eba30-117">GetStarted</span><span class="sxs-lookup"><span data-stu-id="eba30-117">GetStarted</span></span>](getstarted.md)         | <span data-ttu-id="eba30-118">No</span><span class="sxs-lookup"><span data-stu-id="eba30-118">No</span></span>       | <span data-ttu-id="eba30-119">Define la llamada que aparece cuando se instala el complemento en hosts de Word, Excel o PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="eba30-119">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="eba30-120">Ejemplo de DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="eba30-120">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
