# <a name="sets-element"></a><span data-ttu-id="6dcbb-101">Elemento Sets</span><span class="sxs-lookup"><span data-stu-id="6dcbb-101">Sets element</span></span>

<span data-ttu-id="6dcbb-102">Especifica el subconjunto mínimo de la API de JavaScript para Office que su complemento de Office necesita para activarse.</span><span class="sxs-lookup"><span data-stu-id="6dcbb-102">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="6dcbb-103">**Tipo de complemento:** Contenido, panel de tareas, correo</span><span class="sxs-lookup"><span data-stu-id="6dcbb-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6dcbb-104">Sintaxis</span><span class="sxs-lookup"><span data-stu-id="6dcbb-104">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="6dcbb-105">Contenidos en</span><span class="sxs-lookup"><span data-stu-id="6dcbb-105">Contained in</span></span>

[<span data-ttu-id="6dcbb-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6dcbb-106">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="6dcbb-107">Puede contener</span><span class="sxs-lookup"><span data-stu-id="6dcbb-107">Can contain</span></span>

[<span data-ttu-id="6dcbb-108">Set</span><span class="sxs-lookup"><span data-stu-id="6dcbb-108">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="6dcbb-109">Atributos</span><span class="sxs-lookup"><span data-stu-id="6dcbb-109">Attributes</span></span>

|<span data-ttu-id="6dcbb-110">**Attribute**</span><span class="sxs-lookup"><span data-stu-id="6dcbb-110">**Attribute**</span></span>|<span data-ttu-id="6dcbb-111">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="6dcbb-111">**Type**</span></span>|<span data-ttu-id="6dcbb-112">**Obligatorio**</span><span class="sxs-lookup"><span data-stu-id="6dcbb-112">**Required**</span></span>|<span data-ttu-id="6dcbb-113">**Descripción**</span><span class="sxs-lookup"><span data-stu-id="6dcbb-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="6dcbb-114">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="6dcbb-114">DefaultMinVersion</span></span>|<span data-ttu-id="6dcbb-115">string</span><span class="sxs-lookup"><span data-stu-id="6dcbb-115">string</span></span>|<span data-ttu-id="6dcbb-116">opcional</span><span class="sxs-lookup"><span data-stu-id="6dcbb-116">optional</span></span>|<span data-ttu-id="6dcbb-p101">Especifica el valor de atributo predeterminado de **MinVersion** elementos [Set](set.md) secundarios. El valor predeterminado es "1.1".</span><span class="sxs-lookup"><span data-stu-id="6dcbb-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="6dcbb-119">Comentarios</span><span class="sxs-lookup"><span data-stu-id="6dcbb-119">Remarks</span></span>

<span data-ttu-id="6dcbb-120">Para obtener más información acerca de los conjuntos de requisitos, vea [conjuntos de versiones de Office y los requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="6dcbb-120">For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="6dcbb-121">Para obtener más información sobre el atributo **MinVersion** del elemento **Set**  y del atributo **DefaultMinVersion** del elemento **Sets**, consulte [Definir el elemento Requirements en el manifiesto](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="6dcbb-121">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

