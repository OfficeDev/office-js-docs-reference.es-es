# <a name="resources-element"></a><span data-ttu-id="c24ca-101">Elemento Resources</span><span class="sxs-lookup"><span data-stu-id="c24ca-101">Resources element</span></span>

<span data-ttu-id="c24ca-p101">Contiene iconos, cadenas y las direcciones URL para el nodo [VersionOverrides](versionoverrides.md). Un elemento del manifiesto especifica un recurso por medio del **identificador** del recurso. Así se contribuye a mantener el tamaño del manifiesto manejable, en especial cuando los recursos tienen diferentes versiones para diferentes configuraciones regionales. El **identificador** debe ser único dentro del manifiesto y puede tener un máximo de 32 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c24ca-p101">Contains icons, strings, and URLs for the [VersionOverrides](versionoverrides.md) node. A manifest element specifies a resource by using the **id** of the resource. This helps to keep the size of the manifest manageable, especially when resources have versions for different locales. An **id** must be unique within the manifest and can have a maximum of 32 characters.</span></span>

<span data-ttu-id="c24ca-106">Cada recurso puede tener uno o varios elementos secundarios **Override** para definir un recurso diferente en una configuración regional determinada.</span><span class="sxs-lookup"><span data-stu-id="c24ca-106">Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c24ca-107">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="c24ca-107">Child elements</span></span>

|  <span data-ttu-id="c24ca-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="c24ca-108">Element</span></span> |  <span data-ttu-id="c24ca-109">Tipo</span><span class="sxs-lookup"><span data-stu-id="c24ca-109">Type</span></span>  |  <span data-ttu-id="c24ca-110">Descripción</span><span class="sxs-lookup"><span data-stu-id="c24ca-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c24ca-111">Imágenes</span><span class="sxs-lookup"><span data-stu-id="c24ca-111">Images</span></span>](#images)            |  <span data-ttu-id="c24ca-112">image</span><span class="sxs-lookup"><span data-stu-id="c24ca-112">image</span></span>   |  <span data-ttu-id="c24ca-113">Proporciona la dirección URL HTTPS a una imagen para un icono.</span><span class="sxs-lookup"><span data-stu-id="c24ca-113">Provides the HTTPS URL to an image for an icon.</span></span> |
|  <span data-ttu-id="c24ca-114">**Urls**</span><span class="sxs-lookup"><span data-stu-id="c24ca-114">**Urls**</span></span>                |  <span data-ttu-id="c24ca-115">url</span><span class="sxs-lookup"><span data-stu-id="c24ca-115">url</span></span>     |  <span data-ttu-id="c24ca-p102">Proporciona una ubicación URL HTTPS. La dirección URL puede tener 2048 caracteres como máximo.</span><span class="sxs-lookup"><span data-stu-id="c24ca-p102">Provides an HTTPS URL location. A URL can have a maximum of 2048 characters.</span></span> |
|  <span data-ttu-id="c24ca-118">**ShortStrings**</span><span class="sxs-lookup"><span data-stu-id="c24ca-118">**ShortStrings**</span></span> |  <span data-ttu-id="c24ca-119">string</span><span class="sxs-lookup"><span data-stu-id="c24ca-119">string</span></span>  |  <span data-ttu-id="c24ca-p103">Texto de los elementos **Label** y **Title**. Cada **cadena** contiene un máximo de 125 caracteres. </span><span class="sxs-lookup"><span data-stu-id="c24ca-p103">The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters.</span></span>|
|  <span data-ttu-id="c24ca-122">**LongStrings**</span><span class="sxs-lookup"><span data-stu-id="c24ca-122">**LongStrings**</span></span>  |  <span data-ttu-id="c24ca-123">string</span><span class="sxs-lookup"><span data-stu-id="c24ca-123">string</span></span>  | <span data-ttu-id="c24ca-p104">Texto de los atributos **Description**. Cada **cadena** contiene un máximo de 250 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c24ca-p104">The text for **Description** attributes. Each **String** contains a maximum of 250 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="c24ca-126">Debe usar capa de Sockets seguros (SSL) para todas las direcciones URL en los elementos de **imagen** y **dirección Url** .</span><span class="sxs-lookup"><span data-stu-id="c24ca-126">You must use Secure Sockets Layer (SSL) for all URLs in the  **Image** and **Url** elements.</span></span>

### <a name="images"></a><span data-ttu-id="c24ca-127">Imágenes</span><span class="sxs-lookup"><span data-stu-id="c24ca-127">Images</span></span>
<span data-ttu-id="c24ca-128">Cada icono debe tener tres elementos **Image**, uno por cada uno de los tres tamaños obligatorios:</span><span class="sxs-lookup"><span data-stu-id="c24ca-128">Each icon must have three  **Images** elements, one for each of the three mandatory sizes:</span></span>

- <span data-ttu-id="c24ca-129">16x16</span><span class="sxs-lookup"><span data-stu-id="c24ca-129">16x16</span></span>
- <span data-ttu-id="c24ca-130">32x32</span><span class="sxs-lookup"><span data-stu-id="c24ca-130">32x32</span></span>
- <span data-ttu-id="c24ca-131">80x80</span><span class="sxs-lookup"><span data-stu-id="c24ca-131">80x80</span></span>

<span data-ttu-id="c24ca-132">También se admiten los siguientes tamaños adicionales, pero no son obligatorios:</span><span class="sxs-lookup"><span data-stu-id="c24ca-132">The following additional sizes are also supported, but not required:</span></span>

- <span data-ttu-id="c24ca-133">20x20</span><span class="sxs-lookup"><span data-stu-id="c24ca-133">20x20</span></span>
- <span data-ttu-id="c24ca-134">24x24</span><span class="sxs-lookup"><span data-stu-id="c24ca-134">24x24</span></span>
- <span data-ttu-id="c24ca-135">40x40</span><span class="sxs-lookup"><span data-stu-id="c24ca-135">40x40</span></span>
- <span data-ttu-id="c24ca-136">48x48</span><span class="sxs-lookup"><span data-stu-id="c24ca-136">48x48</span></span>
- <span data-ttu-id="c24ca-137">64x64</span><span class="sxs-lookup"><span data-stu-id="c24ca-137">64x64</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="c24ca-138">Outlook requiere la capacidad de recursos de imagen de la memoria caché para aumentar el rendimiento.</span><span class="sxs-lookup"><span data-stu-id="c24ca-138">Outlook requires the ability to cache image resources for performance purposes.</span></span> <span data-ttu-id="c24ca-139">Por este motivo, el servidor donde se hospede un recurso de imagen no tiene que agregar directivas Cache-Control al encabezado de respuesta.</span><span class="sxs-lookup"><span data-stu-id="c24ca-139">For this reason, the server hosting an image resource must not add any CACHE-CONTROL directives to the response header.</span></span> <span data-ttu-id="c24ca-140">Si lo hace, Outlook lo sustituirá automáticamente por una imagen genérica o una imagen predeterminada.</span><span class="sxs-lookup"><span data-stu-id="c24ca-140">This will result in Outlook automatically substituting a generic or default image.</span></span>    

## <a name="resources-examples"></a><span data-ttu-id="c24ca-141">Ejemplos de recursos</span><span class="sxs-lookup"><span data-stu-id="c24ca-141">Resources examples</span></span> 

```XML
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp32-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp80-icon_default.png" />
        </bt:Image>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
        </bt:Url>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residLabel" DefaultValue="GetData">
          <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residToolTip" DefaultValue="Get data for your document.">
          <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
        </bt:String>
      </bt:LongStrings>
    </Resources>
```

```xml
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER//blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/blue-80.png"/>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
    <!-- other URLs -->
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="groupLabel" DefaultValue="Add-in Demo">
      <bt:Override Locale="ar-sa" Value="<Localized text>" />
    </bt:String>
    <!-- Other short strings -->
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment.">
      <bt:Override Locale="ar-sa" Value="<Localized text>." />
    </bt:String>
    <!-- Other long strings -->
  </bt:LongStrings>
</Resources>
```
