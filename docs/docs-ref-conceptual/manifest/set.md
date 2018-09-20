# <a name="set-element"></a><span data-ttu-id="8d1ff-101">Elemento Set</span><span class="sxs-lookup"><span data-stu-id="8d1ff-101">Set element</span></span>

<span data-ttu-id="8d1ff-102">Especifica un conjunto de requisitos de la API de JavaScript para Office que su complemento de Office necesita para activarse.</span><span class="sxs-lookup"><span data-stu-id="8d1ff-102">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="8d1ff-103">**Tipo de complemento:** Contenido, panel de tareas, correo</span><span class="sxs-lookup"><span data-stu-id="8d1ff-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="8d1ff-104">Sintaxis</span><span class="sxs-lookup"><span data-stu-id="8d1ff-104">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="8d1ff-105">Contenidos en</span><span class="sxs-lookup"><span data-stu-id="8d1ff-105">Contained in</span></span>

[<span data-ttu-id="8d1ff-106">Conjuntos</span><span class="sxs-lookup"><span data-stu-id="8d1ff-106">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="8d1ff-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="8d1ff-107">Attributes</span></span>

|<span data-ttu-id="8d1ff-108">**Attribute**</span><span class="sxs-lookup"><span data-stu-id="8d1ff-108">**Attribute**</span></span>|<span data-ttu-id="8d1ff-109">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="8d1ff-109">**Type**</span></span>|<span data-ttu-id="8d1ff-110">**Obligatorio**</span><span class="sxs-lookup"><span data-stu-id="8d1ff-110">**Required**</span></span>|<span data-ttu-id="8d1ff-111">**Descripción**</span><span class="sxs-lookup"><span data-stu-id="8d1ff-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="8d1ff-112">Nombre</span><span class="sxs-lookup"><span data-stu-id="8d1ff-112">Name</span></span>|<span data-ttu-id="8d1ff-113">string</span><span class="sxs-lookup"><span data-stu-id="8d1ff-113">string</span></span>|<span data-ttu-id="8d1ff-114">necesario</span><span class="sxs-lookup"><span data-stu-id="8d1ff-114">required</span></span>|<span data-ttu-id="8d1ff-115">El nombre de un [conjunto de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="8d1ff-115">The name of a [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="8d1ff-116">MinVersion</span><span class="sxs-lookup"><span data-stu-id="8d1ff-116">MinVersion</span></span>|<span data-ttu-id="8d1ff-117">string</span><span class="sxs-lookup"><span data-stu-id="8d1ff-117">string</span></span>|<span data-ttu-id="8d1ff-118">opcional</span><span class="sxs-lookup"><span data-stu-id="8d1ff-118">optional</span></span>|<span data-ttu-id="8d1ff-p101">Especifica la versión mínima del conjunto de API que necesita el complemento. Reemplaza el valor **DefaultMinVersion** si se especifica en el elemento primario [Sets](sets.md).</span><span class="sxs-lookup"><span data-stu-id="8d1ff-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="8d1ff-121">Comentarios</span><span class="sxs-lookup"><span data-stu-id="8d1ff-121">Remarks</span></span>

<span data-ttu-id="8d1ff-122">Para obtener más información acerca de los conjuntos de requisitos, vea [conjuntos de versiones de Office y los requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="8d1ff-122">For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="8d1ff-123">Para obtener más información sobre el atributo **MinVersion** del elemento **Set**  y del atributo **DefaultMinVersion** del elemento **Sets**, consulte [Definir el elemento Requirements en el manifiesto](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="8d1ff-123">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="8d1ff-124">Para complementos de correo, sólo hay una `"Mailbox"` conjuntos de requisito.</span><span class="sxs-lookup"><span data-stu-id="8d1ff-124">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="8d1ff-125">En este conjunto de requisitos contiene el subconjunto completo de API admitidas en los complementos de correo para Outlook, y debe especificar el `"Mailbox"` requisito establecido en manifiesto del su correo complemento (no es opcional como es el caso de tareas y de contenido complementos del panel).</span><span class="sxs-lookup"><span data-stu-id="8d1ff-125">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="8d1ff-126">Además, no se puede declarar compatibilidad para métodos específicos en los complementos de correo.</span><span class="sxs-lookup"><span data-stu-id="8d1ff-126">Also, you can't declare support for specific methods in mail add-ins.</span></span>
