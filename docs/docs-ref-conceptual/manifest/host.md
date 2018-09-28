# <a name="host-element"></a><span data-ttu-id="cada4-101">Elemento Host</span><span class="sxs-lookup"><span data-stu-id="cada4-101">Host element</span></span>

<span data-ttu-id="cada4-102">Especifica un tipo individual de aplicación de Office en el que se debe activar el complemento.</span><span class="sxs-lookup"><span data-stu-id="cada4-102">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="cada4-103">La sintaxis de elemento de **Host** varía en función de si el elemento se ha definido dentro del [manifiesto básica](#basic-manifest) o dentro del nodo [VersionOverrides](#versionoverrides-node) .</span><span class="sxs-lookup"><span data-stu-id="cada4-103">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="cada4-104">Sin embargo, las funciones son las mismas.</span><span class="sxs-lookup"><span data-stu-id="cada4-104">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="cada4-105">Manifiesto básico</span><span class="sxs-lookup"><span data-stu-id="cada4-105">Basic manifest</span></span>

<span data-ttu-id="cada4-106">Cuando se define en el manifiesto básico (bajo [OfficeApp](officeapp.md)), el tipo de host está determinado en el atributo `Name`.</span><span class="sxs-lookup"><span data-stu-id="cada4-106">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="cada4-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="cada4-107">Attributes</span></span>

| <span data-ttu-id="cada4-108">Atributo</span><span class="sxs-lookup"><span data-stu-id="cada4-108">Attribute</span></span>     | <span data-ttu-id="cada4-109">Tipo</span><span class="sxs-lookup"><span data-stu-id="cada4-109">Type</span></span>   | <span data-ttu-id="cada4-110">Necesario</span><span class="sxs-lookup"><span data-stu-id="cada4-110">Required</span></span> | <span data-ttu-id="cada4-111">Descripción</span><span class="sxs-lookup"><span data-stu-id="cada4-111">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="cada4-112">Name</span><span class="sxs-lookup"><span data-stu-id="cada4-112">Name</span></span>](#name) | <span data-ttu-id="cada4-113">string</span><span class="sxs-lookup"><span data-stu-id="cada4-113">string</span></span> | <span data-ttu-id="cada4-114">necesario</span><span class="sxs-lookup"><span data-stu-id="cada4-114">required</span></span> | <span data-ttu-id="cada4-115">El nombre del tipo de aplicación host de Office.</span><span class="sxs-lookup"><span data-stu-id="cada4-115">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="cada4-116">Nombre</span><span class="sxs-lookup"><span data-stu-id="cada4-116">Name</span></span>
<span data-ttu-id="cada4-p102">Especifica el tipo de host al que se dirige este complemento. El valor debe ser uno de los siguientes:</span><span class="sxs-lookup"><span data-stu-id="cada4-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="cada4-119">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="cada4-119">`Document` (Word)</span></span>
- <span data-ttu-id="cada4-120">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="cada4-120">`Database` (Access)</span></span>
- <span data-ttu-id="cada4-121">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="cada4-121">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="cada4-122">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="cada4-122">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="cada4-123">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="cada4-123">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="cada4-124">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="cada4-124">`Project` (Project)</span></span>
- <span data-ttu-id="cada4-125">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="cada4-125">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="cada4-126">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="cada4-126">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="cada4-127">Nodo VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="cada4-127">VersionOverrides node</span></span>
<span data-ttu-id="cada4-128">Cuando se define en [VersionOverrides](versionoverrides.md), el tipo de host está determinado en el atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="cada4-128">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="cada4-129">Atributos</span><span class="sxs-lookup"><span data-stu-id="cada4-129">Attributes</span></span>

|  <span data-ttu-id="cada4-130">Atributo</span><span class="sxs-lookup"><span data-stu-id="cada4-130">Attribute</span></span>  |  <span data-ttu-id="cada4-131">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="cada4-131">Required</span></span>  |  <span data-ttu-id="cada4-132">Descripción</span><span class="sxs-lookup"><span data-stu-id="cada4-132">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cada4-133">xsi:type</span><span class="sxs-lookup"><span data-stu-id="cada4-133">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="cada4-134">Sí</span><span class="sxs-lookup"><span data-stu-id="cada4-134">Yes</span></span>  | <span data-ttu-id="cada4-135">Describe el host de Office donde se aplica esta configuración.</span><span class="sxs-lookup"><span data-stu-id="cada4-135">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="cada4-136">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="cada4-136">Child elements</span></span>

|  <span data-ttu-id="cada4-137">Elemento</span><span class="sxs-lookup"><span data-stu-id="cada4-137">Element</span></span> |  <span data-ttu-id="cada4-138">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="cada4-138">Required</span></span>  |  <span data-ttu-id="cada4-139">Descripción</span><span class="sxs-lookup"><span data-stu-id="cada4-139">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cada4-140">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="cada4-140">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="cada4-141">Sí</span><span class="sxs-lookup"><span data-stu-id="cada4-141">Yes</span></span>   |  <span data-ttu-id="cada4-142">Define la configuración del factor de forma de escritorio.</span><span class="sxs-lookup"><span data-stu-id="cada4-142">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="cada4-143">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="cada4-143">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="cada4-144">No</span><span class="sxs-lookup"><span data-stu-id="cada4-144">No</span></span>   |  <span data-ttu-id="cada4-p103">Define la configuración del factor de forma móvil. **Nota:** Este elemento solo se admite en Outlook para iOS.</span><span class="sxs-lookup"><span data-stu-id="cada4-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="cada4-147">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="cada4-147">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="cada4-148">No</span><span class="sxs-lookup"><span data-stu-id="cada4-148">No</span></span>   |  <span data-ttu-id="cada4-149">Define la configuración de todos los factores de forma.</span><span class="sxs-lookup"><span data-stu-id="cada4-149">Defines the settings for all form factors.</span></span> <span data-ttu-id="cada4-150">Solo lo usan las funciones personalizadas en Excel.</span><span class="sxs-lookup"><span data-stu-id="cada4-150">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="cada4-151">xsi:type</span><span class="sxs-lookup"><span data-stu-id="cada4-151">xsi:type</span></span>

<span data-ttu-id="cada4-152">Controla a qué host de Office (Word, Excel, PowerPoint, Outlook, OneNote) se aplica la configuración contenida.</span><span class="sxs-lookup"><span data-stu-id="cada4-152">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="cada4-153">El valor debe ser uno de los siguientes:</span><span class="sxs-lookup"><span data-stu-id="cada4-153">The value must be one of the following:</span></span>

- <span data-ttu-id="cada4-154">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="cada4-154">`Document` (Word)</span></span>
- <span data-ttu-id="cada4-155">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="cada4-155">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="cada4-156">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="cada4-156">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="cada4-157">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="cada4-157">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="cada4-158">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="cada4-158">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="cada4-159">Ejemplo de host</span><span class="sxs-lookup"><span data-stu-id="cada4-159">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
