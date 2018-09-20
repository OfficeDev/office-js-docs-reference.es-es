# <a name="mobileformfactor-element"></a><span data-ttu-id="4fa52-101">Elemento MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="4fa52-101">MobileFormFactor element</span></span>

<span data-ttu-id="4fa52-p101">Especifica la configuración de un complemento para el factor de forma móvil. Contiene toda la información sobre complementos para dicho factor de forma móvil, excepto para el nodo **Resources**.</span><span class="sxs-lookup"><span data-stu-id="4fa52-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="4fa52-p102">Cada definición **MobileFormFactor** contiene el elemento **FunctionFile** y uno o más elementos **ExtensionPoint**. Para obtener más información, vea [Elemento FunctionFile](functionfile.md) y [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="4fa52-p102">Each **MobileFormFactor** definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="4fa52-p103">El elemento **MobileFormFactor** se define en el esquema VersionOverrides 1.1. El elemento que contenga [VersionOverrides](versionoverrides.md) debe tener un valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="4fa52-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="4fa52-108">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="4fa52-108">Child elements</span></span>

| <span data-ttu-id="4fa52-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="4fa52-109">Element</span></span>                               | <span data-ttu-id="4fa52-110">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="4fa52-110">Required</span></span> | <span data-ttu-id="4fa52-111">Descripción</span><span class="sxs-lookup"><span data-stu-id="4fa52-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="4fa52-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="4fa52-112">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="4fa52-113">Sí</span><span class="sxs-lookup"><span data-stu-id="4fa52-113">Yes</span></span>      | <span data-ttu-id="4fa52-114">Define dónde expone su funcionalidad un complemento.</span><span class="sxs-lookup"><span data-stu-id="4fa52-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="4fa52-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="4fa52-115">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="4fa52-116">Sí</span><span class="sxs-lookup"><span data-stu-id="4fa52-116">Yes</span></span>      | <span data-ttu-id="4fa52-117">Una dirección URL de un archivo que contiene funciones de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="4fa52-117">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="4fa52-118">Ejemplo MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="4fa52-118">MobileFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
