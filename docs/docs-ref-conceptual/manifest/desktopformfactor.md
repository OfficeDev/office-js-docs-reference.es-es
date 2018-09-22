# <a name="desktopformfactor-element"></a><span data-ttu-id="31ab8-101">Elemento DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="31ab8-101">DesktopFormFactor element</span></span>

<span data-ttu-id="31ab8-p101">Especifica la configuración de un complemento para un factor de forma de escritorio. El factor de forma de escritorio incluye Office para Windows, Office para Mac y Office Online. Contiene toda la información sobre complementos para dicho factor de forma, excepto para el nodo **Resources**.</span><span class="sxs-lookup"><span data-stu-id="31ab8-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="31ab8-p102">Cada definición DesktopFormFactor contiene el elemento **FunctionFile** y uno o más elementos **ExtensionPoint**. Para obtener más información, consulte [Elemento FunctionFile](functionfile.md) y [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="31ab8-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="31ab8-107">El elemento SupportsSharedFolders sólo está disponible en el conjunto del requisito Preview de complementos de Outlook con Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="31ab8-107">The SupportsSharedFolders element is only available in the Outlook add-ins Preview Requirement Set against Exchange Online.</span></span>
> <span data-ttu-id="31ab8-108">Complementos que usan este elemento no están permitidos en la tienda de Office o la implementación centralizada.</span><span class="sxs-lookup"><span data-stu-id="31ab8-108">Add-ins that use this element aren't allowed in the Office Store or Centralized Deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="31ab8-109">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="31ab8-109">Child elements</span></span>

| <span data-ttu-id="31ab8-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="31ab8-110">Element</span></span>                               | <span data-ttu-id="31ab8-111">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="31ab8-111">Required</span></span> | <span data-ttu-id="31ab8-112">Descripción</span><span class="sxs-lookup"><span data-stu-id="31ab8-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="31ab8-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="31ab8-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="31ab8-114">Sí</span><span class="sxs-lookup"><span data-stu-id="31ab8-114">Yes</span></span>      | <span data-ttu-id="31ab8-115">Define dónde expone su funcionalidad un complemento.</span><span class="sxs-lookup"><span data-stu-id="31ab8-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="31ab8-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="31ab8-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="31ab8-117">Sí</span><span class="sxs-lookup"><span data-stu-id="31ab8-117">Yes</span></span>      | <span data-ttu-id="31ab8-118">Una dirección URL de un archivo que contiene funciones de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="31ab8-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="31ab8-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="31ab8-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="31ab8-120">No</span><span class="sxs-lookup"><span data-stu-id="31ab8-120">No</span></span>       | <span data-ttu-id="31ab8-121">Define la llamada que aparece cuando se instala el complemento en hosts de Word, Excel o PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="31ab8-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| <span data-ttu-id="31ab8-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="31ab8-122">SupportsSharedFolders</span></span>                 | <span data-ttu-id="31ab8-123">No</span><span class="sxs-lookup"><span data-stu-id="31ab8-123">No</span></span>       | <span data-ttu-id="31ab8-124">Define si el complemento de Outlook está disponible en escenarios de delegado y se establece en *false* de forma predeterminada.</span><span class="sxs-lookup"><span data-stu-id="31ab8-124">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> <span data-ttu-id="31ab8-125">Obtener una vista previa de conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="31ab8-125">Preview requirement set.</span></span>|

## <a name="desktopformfactor-example"></a><span data-ttu-id="31ab8-126">Ejemplo de DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="31ab8-126">DesktopFormFactor example</span></span>

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
