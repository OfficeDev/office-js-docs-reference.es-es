# <a name="officetab-element"></a><span data-ttu-id="d5421-101">Elemento OfficeTab</span><span class="sxs-lookup"><span data-stu-id="d5421-101">OfficeTab element</span></span>

<span data-ttu-id="d5421-p101">Define la ficha de la cinta en la que aparece el comando del complemento. Puede ser la pestaña predeterminada (**Inicio**, **Mensaje** o **Reunión**), o una pestaña personalizada definida por el complemento. Se requiere este elemento.</span><span class="sxs-lookup"><span data-stu-id="d5421-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="d5421-105">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="d5421-105">Child elements</span></span>

|  <span data-ttu-id="d5421-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="d5421-106">Element</span></span> |  <span data-ttu-id="d5421-107">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="d5421-107">Required</span></span>  |  <span data-ttu-id="d5421-108">Descripción</span><span class="sxs-lookup"><span data-stu-id="d5421-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="d5421-109">Grupo</span><span class="sxs-lookup"><span data-stu-id="d5421-109">Group</span></span>      | <span data-ttu-id="d5421-110">Sí</span><span class="sxs-lookup"><span data-stu-id="d5421-110">Yes</span></span> |  <span data-ttu-id="d5421-p102">Define un grupo de comandos. Solo se puede agregar un grupo por cada complemento a la ficha predeterminada.</span><span class="sxs-lookup"><span data-stu-id="d5421-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="d5421-113">Los siguientes valores son `id` valores de ficha válidos por host:</span><span class="sxs-lookup"><span data-stu-id="d5421-113">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="d5421-114">Los valores en **negrita** son compatibles con escritorio y en línea (por ejemplo, Word 2016 o posterior para Windows y Word en línea).</span><span class="sxs-lookup"><span data-stu-id="d5421-114">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="d5421-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="d5421-115">Outlook</span></span>

- <span data-ttu-id="d5421-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="d5421-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="d5421-117">Word</span><span class="sxs-lookup"><span data-stu-id="d5421-117">Word</span></span>

- <span data-ttu-id="d5421-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="d5421-118">**TabHome**</span></span>
- <span data-ttu-id="d5421-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="d5421-119">**TabInsert**</span></span>
- <span data-ttu-id="d5421-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="d5421-120">TabWordDesign</span></span>
- <span data-ttu-id="d5421-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="d5421-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="d5421-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="d5421-122">TabReferences</span></span>
- <span data-ttu-id="d5421-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="d5421-123">TabMailings</span></span>
- <span data-ttu-id="d5421-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="d5421-124">TabReviewWord</span></span>
- <span data-ttu-id="d5421-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="d5421-125">**TabView**</span></span>
- <span data-ttu-id="d5421-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="d5421-126">TabDeveloper</span></span>
- <span data-ttu-id="d5421-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="d5421-127">TabAddIns</span></span>
- <span data-ttu-id="d5421-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="d5421-128">TabBlogPost</span></span>
- <span data-ttu-id="d5421-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="d5421-129">TabBlogInsert</span></span>
- <span data-ttu-id="d5421-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="d5421-130">TabPrintPreview</span></span>
- <span data-ttu-id="d5421-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="d5421-131">TabOutlining</span></span>
- <span data-ttu-id="d5421-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="d5421-132">TabConflicts</span></span>
- <span data-ttu-id="d5421-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="d5421-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="d5421-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="d5421-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="d5421-135">Excel</span><span class="sxs-lookup"><span data-stu-id="d5421-135">Excel</span></span>

- <span data-ttu-id="d5421-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="d5421-136">**TabHome**</span></span>
- <span data-ttu-id="d5421-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="d5421-137">**TabInsert**</span></span>
- <span data-ttu-id="d5421-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="d5421-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="d5421-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="d5421-139">TabFormulas</span></span>
- <span data-ttu-id="d5421-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="d5421-140">**TabData**</span></span>
- <span data-ttu-id="d5421-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="d5421-141">**TabReview**</span></span>
- <span data-ttu-id="d5421-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="d5421-142">**TabView**</span></span>
- <span data-ttu-id="d5421-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="d5421-143">TabDeveloper</span></span>
- <span data-ttu-id="d5421-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="d5421-144">TabAddIns</span></span>
- <span data-ttu-id="d5421-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="d5421-145">TabPrintPreview</span></span>
- <span data-ttu-id="d5421-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="d5421-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="d5421-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d5421-147">PowerPoint</span></span>

- <span data-ttu-id="d5421-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="d5421-148">**TabHome**</span></span>
- <span data-ttu-id="d5421-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="d5421-149">**TabInsert**</span></span>
- <span data-ttu-id="d5421-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="d5421-150">**TabDesign**</span></span>
- <span data-ttu-id="d5421-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="d5421-151">**TabTransitions**</span></span>
- <span data-ttu-id="d5421-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="d5421-152">**TabAnimations**</span></span>
- <span data-ttu-id="d5421-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="d5421-153">TabSlideShow</span></span>
- <span data-ttu-id="d5421-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="d5421-154">TabReview</span></span>
- <span data-ttu-id="d5421-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="d5421-155">**TabView**</span></span>
- <span data-ttu-id="d5421-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="d5421-156">TabDeveloper</span></span>
- <span data-ttu-id="d5421-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="d5421-157">TabAddIns</span></span>
- <span data-ttu-id="d5421-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="d5421-158">TabPrintPreview</span></span>
- <span data-ttu-id="d5421-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="d5421-159">TabMerge</span></span>
- <span data-ttu-id="d5421-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="d5421-160">TabGrayscale</span></span>
- <span data-ttu-id="d5421-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="d5421-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="d5421-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="d5421-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="d5421-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="d5421-163">TabSlideMaster</span></span>
- <span data-ttu-id="d5421-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="d5421-164">TabHandoutMaster</span></span>
- <span data-ttu-id="d5421-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="d5421-165">TabNotesMaster</span></span>
- <span data-ttu-id="d5421-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="d5421-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="d5421-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="d5421-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="d5421-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="d5421-168">OneNote</span></span>

- <span data-ttu-id="d5421-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="d5421-169">**TabHome**</span></span>
- <span data-ttu-id="d5421-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="d5421-170">**TabInsert**</span></span>
- <span data-ttu-id="d5421-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="d5421-171">**TabView**</span></span>
- <span data-ttu-id="d5421-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="d5421-172">TabDeveloper</span></span>
- <span data-ttu-id="d5421-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="d5421-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="d5421-174">Group</span><span class="sxs-lookup"><span data-stu-id="d5421-174">Group</span></span>

<span data-ttu-id="d5421-p104">Un grupo de puntos de extensión de UI en una ficha. Un grupo puede tener hasta seis controles. El atributo **id** es obligatorio y cada **id** debe ser único en el manifiesto. El **id** es una cadena de 125 caracteres como máximo. Consulte el [elemento Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="d5421-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="d5421-179">Ejemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="d5421-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
