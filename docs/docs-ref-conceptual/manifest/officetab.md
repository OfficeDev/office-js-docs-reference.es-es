# <a name="officetab-element"></a><span data-ttu-id="7a41e-101">Elemento OfficeTab</span><span class="sxs-lookup"><span data-stu-id="7a41e-101">OfficeTab element</span></span>

<span data-ttu-id="7a41e-p101">Define la ficha de la cinta en la que aparece el comando del complemento. Puede ser la pestaña predeterminada (**Inicio**, **Mensaje** o **Reunión**), o una pestaña personalizada definida por el complemento. Se requiere este elemento.</span><span class="sxs-lookup"><span data-stu-id="7a41e-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="7a41e-105">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="7a41e-105">Child elements</span></span>

|  <span data-ttu-id="7a41e-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="7a41e-106">Element</span></span> |  <span data-ttu-id="7a41e-107">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="7a41e-107">Required</span></span>  |  <span data-ttu-id="7a41e-108">Descripción</span><span class="sxs-lookup"><span data-stu-id="7a41e-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="7a41e-109">Grupo</span><span class="sxs-lookup"><span data-stu-id="7a41e-109">Group</span></span>      | <span data-ttu-id="7a41e-110">Sí</span><span class="sxs-lookup"><span data-stu-id="7a41e-110">Yes</span></span> |  <span data-ttu-id="7a41e-p102">Define un grupo de comandos. Solo se puede agregar un grupo por cada complemento a la ficha predeterminada.</span><span class="sxs-lookup"><span data-stu-id="7a41e-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="7a41e-113">Los siguientes valores son `id` valores de ficha válidos por host:</span><span class="sxs-lookup"><span data-stu-id="7a41e-113">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="7a41e-114">Los valores en **negrita** son compatibles con escritorio y en línea (por ejemplo, Word 2016 o posterior para Windows y Word en línea).</span><span class="sxs-lookup"><span data-stu-id="7a41e-114">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="7a41e-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="7a41e-115">Outlook</span></span>

- <span data-ttu-id="7a41e-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="7a41e-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="7a41e-117">Word</span><span class="sxs-lookup"><span data-stu-id="7a41e-117">Word</span></span>

- <span data-ttu-id="7a41e-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="7a41e-118">**TabHome**</span></span>
- <span data-ttu-id="7a41e-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="7a41e-119">**TabInsert**</span></span>
- <span data-ttu-id="7a41e-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="7a41e-120">TabWordDesign</span></span>
- <span data-ttu-id="7a41e-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="7a41e-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="7a41e-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="7a41e-122">TabReferences</span></span>
- <span data-ttu-id="7a41e-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="7a41e-123">TabMailings</span></span>
- <span data-ttu-id="7a41e-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="7a41e-124">TabReviewWord</span></span>
- <span data-ttu-id="7a41e-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="7a41e-125">**TabView**</span></span>
- <span data-ttu-id="7a41e-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="7a41e-126">TabDeveloper</span></span>
- <span data-ttu-id="7a41e-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="7a41e-127">TabAddIns</span></span>
- <span data-ttu-id="7a41e-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="7a41e-128">TabBlogPost</span></span>
- <span data-ttu-id="7a41e-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="7a41e-129">TabBlogInsert</span></span>
- <span data-ttu-id="7a41e-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="7a41e-130">TabPrintPreview</span></span>
- <span data-ttu-id="7a41e-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="7a41e-131">TabOutlining</span></span>
- <span data-ttu-id="7a41e-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="7a41e-132">TabConflicts</span></span>
- <span data-ttu-id="7a41e-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="7a41e-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="7a41e-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="7a41e-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="7a41e-135">Excel</span><span class="sxs-lookup"><span data-stu-id="7a41e-135">Excel</span></span>

- <span data-ttu-id="7a41e-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="7a41e-136">**TabHome**</span></span>
- <span data-ttu-id="7a41e-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="7a41e-137">**TabInsert**</span></span>
- <span data-ttu-id="7a41e-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="7a41e-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="7a41e-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="7a41e-139">TabFormulas</span></span>
- <span data-ttu-id="7a41e-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="7a41e-140">**TabData**</span></span>
- <span data-ttu-id="7a41e-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="7a41e-141">**TabReview**</span></span>
- <span data-ttu-id="7a41e-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="7a41e-142">**TabView**</span></span>
- <span data-ttu-id="7a41e-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="7a41e-143">TabDeveloper</span></span>
- <span data-ttu-id="7a41e-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="7a41e-144">TabAddIns</span></span>
- <span data-ttu-id="7a41e-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="7a41e-145">TabPrintPreview</span></span>
- <span data-ttu-id="7a41e-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="7a41e-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="7a41e-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="7a41e-147">PowerPoint</span></span>

- <span data-ttu-id="7a41e-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="7a41e-148">**TabHome**</span></span>
- <span data-ttu-id="7a41e-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="7a41e-149">**TabInsert**</span></span>
- <span data-ttu-id="7a41e-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="7a41e-150">**TabDesign**</span></span>
- <span data-ttu-id="7a41e-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="7a41e-151">**TabTransitions**</span></span>
- <span data-ttu-id="7a41e-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="7a41e-152">**TabAnimations**</span></span>
- <span data-ttu-id="7a41e-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="7a41e-153">TabSlideShow</span></span>
- <span data-ttu-id="7a41e-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="7a41e-154">TabReview</span></span>
- <span data-ttu-id="7a41e-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="7a41e-155">**TabView**</span></span>
- <span data-ttu-id="7a41e-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="7a41e-156">TabDeveloper</span></span>
- <span data-ttu-id="7a41e-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="7a41e-157">TabAddIns</span></span>
- <span data-ttu-id="7a41e-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="7a41e-158">TabPrintPreview</span></span>
- <span data-ttu-id="7a41e-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="7a41e-159">TabMerge</span></span>
- <span data-ttu-id="7a41e-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="7a41e-160">TabGrayscale</span></span>
- <span data-ttu-id="7a41e-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="7a41e-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="7a41e-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="7a41e-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="7a41e-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="7a41e-163">TabSlideMaster</span></span>
- <span data-ttu-id="7a41e-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="7a41e-164">TabHandoutMaster</span></span>
- <span data-ttu-id="7a41e-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="7a41e-165">TabNotesMaster</span></span>
- <span data-ttu-id="7a41e-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="7a41e-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="7a41e-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="7a41e-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="7a41e-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="7a41e-168">OneNote</span></span>

- <span data-ttu-id="7a41e-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="7a41e-169">**TabHome**</span></span>
- <span data-ttu-id="7a41e-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="7a41e-170">**TabInsert**</span></span>
- <span data-ttu-id="7a41e-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="7a41e-171">**TabView**</span></span>
- <span data-ttu-id="7a41e-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="7a41e-172">TabDeveloper</span></span>
- <span data-ttu-id="7a41e-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="7a41e-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="7a41e-174">Group</span><span class="sxs-lookup"><span data-stu-id="7a41e-174">Group</span></span>

<span data-ttu-id="7a41e-p104">Un grupo de puntos de extensión de UI en una ficha. Un grupo puede tener hasta seis controles. El atributo **id** es obligatorio y cada **id** debe ser único en el manifiesto. El **id** es una cadena de 125 caracteres como máximo. Consulte el [elemento Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="7a41e-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="7a41e-179">Ejemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="7a41e-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
