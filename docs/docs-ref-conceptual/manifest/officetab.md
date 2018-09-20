# <a name="officetab-element"></a><span data-ttu-id="bb071-101">Elemento OfficeTab</span><span class="sxs-lookup"><span data-stu-id="bb071-101">OfficeTab element</span></span>

<span data-ttu-id="bb071-p101">Define la ficha de la cinta en la que aparece el comando del complemento. Puede ser la pestaña predeterminada (**Inicio**, **Mensaje** o **Reunión**), o una pestaña personalizada definida por el complemento. Se requiere este elemento.</span><span class="sxs-lookup"><span data-stu-id="bb071-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="bb071-105">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="bb071-105">Child elements</span></span>

|  <span data-ttu-id="bb071-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="bb071-106">Element</span></span> |  <span data-ttu-id="bb071-107">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="bb071-107">Required</span></span>  |  <span data-ttu-id="bb071-108">Descripción</span><span class="sxs-lookup"><span data-stu-id="bb071-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bb071-109">Grupo</span><span class="sxs-lookup"><span data-stu-id="bb071-109">Group</span></span>      | <span data-ttu-id="bb071-110">Sí</span><span class="sxs-lookup"><span data-stu-id="bb071-110">Yes</span></span> |  <span data-ttu-id="bb071-p102">Define un grupo de comandos. Solo se puede agregar un grupo por cada complemento a la ficha predeterminada.</span><span class="sxs-lookup"><span data-stu-id="bb071-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="bb071-p103">Los siguientes valores son valores `id` de pestaña válidos por host: Los valores en **negrita** son compatibles en el escritorio y en línea (por ejemplo, Word 2016 para Windows y Word Online).</span><span class="sxs-lookup"><span data-stu-id="bb071-p103">The following are valid tab `id` values by host. Values in **bold** are supported in both desktop and online (for example, Word 2016 for Windows and Word Online).</span></span> 

### <a name="outlook"></a><span data-ttu-id="bb071-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="bb071-115">Outlook</span></span> 

- <span data-ttu-id="bb071-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="bb071-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="bb071-117">Word</span><span class="sxs-lookup"><span data-stu-id="bb071-117">Word</span></span>

- <span data-ttu-id="bb071-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bb071-118">**TabHome**</span></span>
- <span data-ttu-id="bb071-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bb071-119">**TabInsert**</span></span>
- <span data-ttu-id="bb071-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="bb071-120">TabWordDesign</span></span>
- <span data-ttu-id="bb071-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="bb071-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="bb071-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="bb071-122">TabReferences</span></span>
- <span data-ttu-id="bb071-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="bb071-123">TabMailings</span></span>
- <span data-ttu-id="bb071-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="bb071-124">TabReviewWord</span></span>
- <span data-ttu-id="bb071-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bb071-125">**TabView**</span></span>
- <span data-ttu-id="bb071-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bb071-126">TabDeveloper</span></span>
- <span data-ttu-id="bb071-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bb071-127">TabAddIns</span></span>
- <span data-ttu-id="bb071-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="bb071-128">TabBlogPost</span></span>
- <span data-ttu-id="bb071-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="bb071-129">TabBlogInsert</span></span>
- <span data-ttu-id="bb071-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="bb071-130">TabPrintPreview</span></span>
- <span data-ttu-id="bb071-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="bb071-131">TabOutlining</span></span>
- <span data-ttu-id="bb071-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="bb071-132">TabConflicts</span></span>
- <span data-ttu-id="bb071-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="bb071-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="bb071-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="bb071-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="bb071-135">Excel</span><span class="sxs-lookup"><span data-stu-id="bb071-135">Excel</span></span>

- <span data-ttu-id="bb071-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bb071-136">**TabHome**</span></span>
- <span data-ttu-id="bb071-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bb071-137">**TabInsert**</span></span>
- <span data-ttu-id="bb071-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="bb071-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="bb071-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="bb071-139">TabFormulas</span></span>
- <span data-ttu-id="bb071-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="bb071-140">**TabData**</span></span>
- <span data-ttu-id="bb071-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="bb071-141">**TabReview**</span></span>
- <span data-ttu-id="bb071-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bb071-142">**TabView**</span></span>
- <span data-ttu-id="bb071-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bb071-143">TabDeveloper</span></span>
- <span data-ttu-id="bb071-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bb071-144">TabAddIns</span></span>
- <span data-ttu-id="bb071-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="bb071-145">TabPrintPreview</span></span>
- <span data-ttu-id="bb071-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="bb071-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="bb071-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="bb071-147">PowerPoint</span></span>

- <span data-ttu-id="bb071-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bb071-148">**TabHome**</span></span>
- <span data-ttu-id="bb071-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bb071-149">**TabInsert**</span></span>
- <span data-ttu-id="bb071-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="bb071-150">**TabDesign**</span></span>
- <span data-ttu-id="bb071-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="bb071-151">**TabTransitions**</span></span>
- <span data-ttu-id="bb071-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="bb071-152">**TabAnimations**</span></span>
- <span data-ttu-id="bb071-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="bb071-153">TabSlideShow</span></span>
- <span data-ttu-id="bb071-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="bb071-154">TabReview</span></span>
- <span data-ttu-id="bb071-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bb071-155">**TabView**</span></span>
- <span data-ttu-id="bb071-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bb071-156">TabDeveloper</span></span>
- <span data-ttu-id="bb071-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bb071-157">TabAddIns</span></span>
- <span data-ttu-id="bb071-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="bb071-158">TabPrintPreview</span></span>
- <span data-ttu-id="bb071-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="bb071-159">TabMerge</span></span>
- <span data-ttu-id="bb071-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="bb071-160">TabGrayscale</span></span>
- <span data-ttu-id="bb071-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="bb071-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="bb071-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="bb071-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="bb071-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="bb071-163">TabSlideMaster</span></span>
- <span data-ttu-id="bb071-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="bb071-164">TabHandoutMaster</span></span>
- <span data-ttu-id="bb071-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="bb071-165">TabNotesMaster</span></span>
- <span data-ttu-id="bb071-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="bb071-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="bb071-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="bb071-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="bb071-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="bb071-168">OneNote</span></span>

- <span data-ttu-id="bb071-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bb071-169">**TabHome**</span></span>
- <span data-ttu-id="bb071-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bb071-170">**TabInsert**</span></span>
- <span data-ttu-id="bb071-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bb071-171">**TabView**</span></span>
- <span data-ttu-id="bb071-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bb071-172">TabDeveloper</span></span>
- <span data-ttu-id="bb071-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bb071-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="bb071-174">Group</span><span class="sxs-lookup"><span data-stu-id="bb071-174">Group</span></span>

<span data-ttu-id="bb071-p104">Un grupo de puntos de extensión de UI en una ficha. Un grupo puede tener hasta seis controles. El atributo **id** es obligatorio y cada **id** debe ser único en el manifiesto. El **id** es una cadena de 125 caracteres como máximo. Consulte el [elemento Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="bb071-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="bb071-179">Ejemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="bb071-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
