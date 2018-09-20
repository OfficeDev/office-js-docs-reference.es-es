# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="840ce-101">Conjuntos de requisitos de la API de JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="840ce-101">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="840ce-102">Los conjuntos de requisitos son grupos de miembros de la API con nombre.</span><span class="sxs-lookup"><span data-stu-id="840ce-102">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="840ce-103">Complementos de Office utilizan conjuntos de requisitos especificados en el manifiesto o una comprobación en tiempo de ejecución para determinar si un host de Office admite las API que un complemento necesita.</span><span class="sxs-lookup"><span data-stu-id="840ce-103">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="840ce-104">Para obtener más información, vea [conjuntos de versiones de Office y los requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="840ce-104">For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="840ce-105">En la siguiente tabla se enumeran los conjuntos de requisitos de OneNote, las aplicaciones de host de Office que admiten esos conjuntos de requisitos y las versiones de la compilación o fecha de disponibilidad.</span><span class="sxs-lookup"><span data-stu-id="840ce-105">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="840ce-106">Conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="840ce-106">Requirement set</span></span>  |  <span data-ttu-id="840ce-107">Office Online</span><span class="sxs-lookup"><span data-stu-id="840ce-107">Office Online</span></span> | 
|:-----|:-----|
| <span data-ttu-id="840ce-108">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="840ce-108">OneNoteApi 1.1</span></span>  | <span data-ttu-id="840ce-109">Septiembre de 2016</span><span class="sxs-lookup"><span data-stu-id="840ce-109">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="840ce-110">Conjuntos de requisitos comunes de la API de Office</span><span class="sxs-lookup"><span data-stu-id="840ce-110">Office common API requirement sets</span></span>

<span data-ttu-id="840ce-111">Para obtener información sobre los conjuntos de requisitos comunes de la API, consulte [Office common API requirement sets (Conjuntos de requisitos comunes de la API de Office)](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="840ce-111">For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="840ce-112">API de JavaScript de OneNote 1.1</span><span class="sxs-lookup"><span data-stu-id="840ce-112">OneNote JavaScript API 1.1</span></span> 

<span data-ttu-id="840ce-113">API de JavaScript de OneNote 1.1 es la primera versión de la API.</span><span class="sxs-lookup"><span data-stu-id="840ce-113">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="840ce-114">Para obtener información detallada acerca de la API, vea la [información general sobre programación de API de JavaScript de OneNote](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span><span class="sxs-lookup"><span data-stu-id="840ce-114">For details about the API, see the [OneNote JavaScript API programming overview](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="840ce-115">Comprobación de compatibilidad de los requisitos en tiempo de ejecución</span><span class="sxs-lookup"><span data-stu-id="840ce-115">Runtime requirement support check</span></span>

<span data-ttu-id="840ce-116">Durante el tiempo de ejecución, los complementos pueden comprobar si un host determinado admite un conjunto de requisitos de la API establecido realizando la siguiente comprobación:</span><span class="sxs-lookup"><span data-stu-id="840ce-116">During the runtime, add-ins can check if a particular host supports an API requirement set by doing the following-check:</span></span> 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="840ce-117">Comprobación de compatibilidad de los requisitos basada en un manifiesto</span><span class="sxs-lookup"><span data-stu-id="840ce-117">Manifest-based requirement support check</span></span>

<span data-ttu-id="840ce-p103">Use el elemento Requirements en el manifiesto del complemento para especificar conjuntos de requisitos críticos o miembros de la API que debe usar el complemento. Si la plataforma o el host de Office no son compatibles con los conjuntos de requisitos o los miembros de la API especificados en el elemento Requirements, el complemento no se ejecutará en la plataforma o el host ni se mostrará en Mis complementos.</span><span class="sxs-lookup"><span data-stu-id="840ce-p103">Use the Requirements element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the Requirements element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="840ce-120">En el siguiente ejemplo de código se muestra un complemento que se carga en todas las aplicaciones host de Office que admiten el conjunto de requisitos OneNoteApi, versión 1.1.</span><span class="sxs-lookup"><span data-stu-id="840ce-120">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a><span data-ttu-id="840ce-121">Vea también</span><span class="sxs-lookup"><span data-stu-id="840ce-121">See also</span></span>

- [<span data-ttu-id="840ce-122">Versiones de Office y conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="840ce-122">Office versions and requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="840ce-123">Especificar los hosts de Office y los requisitos de la API</span><span class="sxs-lookup"><span data-stu-id="840ce-123">Specify Office hosts and API requirements</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="840ce-124">Manifiesto XML de complementos para Office</span><span class="sxs-lookup"><span data-stu-id="840ce-124">Office Add-ins XML manifest</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
