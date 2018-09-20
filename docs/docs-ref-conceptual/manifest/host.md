# <a name="host-element"></a><span data-ttu-id="e97d8-101">Elemento Host</span><span class="sxs-lookup"><span data-stu-id="e97d8-101">Host element</span></span>

<span data-ttu-id="e97d8-102">Especifica un tipo individual de aplicación de Office en el que se debe activar el complemento.</span><span class="sxs-lookup"><span data-stu-id="e97d8-102">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="e97d8-103">La sintaxis de elemento de **Host** varía en función de si el elemento se ha definido dentro del [manifiesto básica](#basic-manifest) o dentro del nodo [VersionOverrides](#versionoverrides-node) .</span><span class="sxs-lookup"><span data-stu-id="e97d8-103">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="e97d8-104">Sin embargo, las funciones son las mismas.</span><span class="sxs-lookup"><span data-stu-id="e97d8-104">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="e97d8-105">Manifiesto básico</span><span class="sxs-lookup"><span data-stu-id="e97d8-105">Basic manifest</span></span>

<span data-ttu-id="e97d8-106">Cuando se define en el manifiesto básico (bajo [OfficeApp](officeapp.md)), el tipo de host está determinado en el atributo `Name`.</span><span class="sxs-lookup"><span data-stu-id="e97d8-106">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="e97d8-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="e97d8-107">Attributes</span></span>

| <span data-ttu-id="e97d8-108">Atributo</span><span class="sxs-lookup"><span data-stu-id="e97d8-108">Attribute</span></span>     | <span data-ttu-id="e97d8-109">Tipo</span><span class="sxs-lookup"><span data-stu-id="e97d8-109">Type</span></span>   | <span data-ttu-id="e97d8-110">Necesario</span><span class="sxs-lookup"><span data-stu-id="e97d8-110">Required</span></span> | <span data-ttu-id="e97d8-111">Descripción</span><span class="sxs-lookup"><span data-stu-id="e97d8-111">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="e97d8-112">Name</span><span class="sxs-lookup"><span data-stu-id="e97d8-112">Name</span></span>](#name) | <span data-ttu-id="e97d8-113">string</span><span class="sxs-lookup"><span data-stu-id="e97d8-113">string</span></span> | <span data-ttu-id="e97d8-114">necesario</span><span class="sxs-lookup"><span data-stu-id="e97d8-114">required</span></span> | <span data-ttu-id="e97d8-115">El nombre del tipo de aplicación host de Office.</span><span class="sxs-lookup"><span data-stu-id="e97d8-115">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="e97d8-116">Nombre</span><span class="sxs-lookup"><span data-stu-id="e97d8-116">Name</span></span>
<span data-ttu-id="e97d8-p102">Especifica el tipo de host al que se dirige este complemento. El valor debe ser uno de los siguientes:</span><span class="sxs-lookup"><span data-stu-id="e97d8-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="e97d8-119">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="e97d8-119">`Document` (Word)</span></span>
- <span data-ttu-id="e97d8-120">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="e97d8-120">`Database` (Access)</span></span>
- <span data-ttu-id="e97d8-121">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="e97d8-121">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="e97d8-122">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="e97d8-122">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="e97d8-123">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="e97d8-123">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="e97d8-124">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="e97d8-124">`Project` (Project)</span></span>
- <span data-ttu-id="e97d8-125">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="e97d8-125">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="e97d8-126">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="e97d8-126">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="e97d8-127">Nodo VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="e97d8-127">VersionOverrides node</span></span>
<span data-ttu-id="e97d8-128">Cuando se define en [VersionOverrides](versionoverrides.md), el tipo de host está determinado en el atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="e97d8-128">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="e97d8-129">Atributos</span><span class="sxs-lookup"><span data-stu-id="e97d8-129">Attributes</span></span>

|  <span data-ttu-id="e97d8-130">Atributo</span><span class="sxs-lookup"><span data-stu-id="e97d8-130">Attribute</span></span>  |  <span data-ttu-id="e97d8-131">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="e97d8-131">Required</span></span>  |  <span data-ttu-id="e97d8-132">Descripción</span><span class="sxs-lookup"><span data-stu-id="e97d8-132">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e97d8-133">xsi:type</span><span class="sxs-lookup"><span data-stu-id="e97d8-133">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="e97d8-134">Sí</span><span class="sxs-lookup"><span data-stu-id="e97d8-134">Yes</span></span>  | <span data-ttu-id="e97d8-135">Describe el host de Office donde se aplica esta configuración.</span><span class="sxs-lookup"><span data-stu-id="e97d8-135">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="e97d8-136">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="e97d8-136">Child elements</span></span>

|  <span data-ttu-id="e97d8-137">Elemento</span><span class="sxs-lookup"><span data-stu-id="e97d8-137">Element</span></span> |  <span data-ttu-id="e97d8-138">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="e97d8-138">Required</span></span>  |  <span data-ttu-id="e97d8-139">Descripción</span><span class="sxs-lookup"><span data-stu-id="e97d8-139">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e97d8-140">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="e97d8-140">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="e97d8-141">Sí</span><span class="sxs-lookup"><span data-stu-id="e97d8-141">Yes</span></span>   |  <span data-ttu-id="e97d8-142">Define la configuración del factor de forma de escritorio.</span><span class="sxs-lookup"><span data-stu-id="e97d8-142">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="e97d8-143">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="e97d8-143">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="e97d8-144">No</span><span class="sxs-lookup"><span data-stu-id="e97d8-144">No</span></span>   |  <span data-ttu-id="e97d8-p103">Define la configuración del factor de forma móvil. **Nota:** Este elemento solo se admite en Outlook para iOS.</span><span class="sxs-lookup"><span data-stu-id="e97d8-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="e97d8-147">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="e97d8-147">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="e97d8-148">No</span><span class="sxs-lookup"><span data-stu-id="e97d8-148">No</span></span>   |  <span data-ttu-id="e97d8-149">Define la configuración de todos los factores de forma.</span><span class="sxs-lookup"><span data-stu-id="e97d8-149">Defines the settings for all form factors.</span></span> <span data-ttu-id="e97d8-150">Solo lo usan las funciones personalizadas en Excel.</span><span class="sxs-lookup"><span data-stu-id="e97d8-150">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="e97d8-151">xsi:type</span><span class="sxs-lookup"><span data-stu-id="e97d8-151">xsi:type</span></span>

<span data-ttu-id="e97d8-152">Controla a qué host de Office (Word, Excel, PowerPoint, Outlook, OneNote) se aplica la configuración contenida.</span><span class="sxs-lookup"><span data-stu-id="e97d8-152">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="e97d8-153">El valor debe ser uno de los siguientes:</span><span class="sxs-lookup"><span data-stu-id="e97d8-153">The value must be one of the following:</span></span>

- <span data-ttu-id="e97d8-154">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="e97d8-154">`Document` (Word)</span></span>
- <span data-ttu-id="e97d8-155">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="e97d8-155">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="e97d8-156">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="e97d8-156">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="e97d8-157">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="e97d8-157">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="e97d8-158">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="e97d8-158">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="e97d8-159">Ejemplo de host</span><span class="sxs-lookup"><span data-stu-id="e97d8-159">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
