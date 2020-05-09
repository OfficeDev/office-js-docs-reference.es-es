---
title: Referencia de la API de JavaScript de Office
description: El host establece el conjunto de requisitos de las API de JavaScript de Office.
ms.date: 05/05/2020
ms.openlocfilehash: 3a32c47b23fd6635c4c2b44b58ee9b351fffd8d5
ms.sourcegitcommit: 23d9a58660cb1dedf0bc414849a5aec519b419b3
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/07/2020
ms.locfileid: "44128550"
---
# <a name="office-javascript-api-reference"></a><span data-ttu-id="9bfa8-103">Referencia de la API de JavaScript de Office</span><span class="sxs-lookup"><span data-stu-id="9bfa8-103">Office JavaScript API reference</span></span>

<span data-ttu-id="9bfa8-104">La API de JavaScript para Office le permite crear aplicaciones web que interactúen con los modelos de objetos de las aplicaciones host de Office.</span><span class="sxs-lookup"><span data-stu-id="9bfa8-104">The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications.</span></span> <span data-ttu-id="9bfa8-105">Use esta sección para obtener más información sobre las clases, los métodos y otros tipos disponibles para crear complementos de Office.</span><span class="sxs-lookup"><span data-stu-id="9bfa8-105">Use this section to learn more about the classes, methods, and other types available for building Office Add-ins.</span></span>

<span data-ttu-id="9bfa8-106">A continuación se muestra una lista de conjuntos de requisitos específicos del host (y las API comunes entre hosts).</span><span class="sxs-lookup"><span data-stu-id="9bfa8-106">The following is a list of host-specific requirement sets (and the cross-host Common APIs).</span></span> <span data-ttu-id="9bfa8-107">Cada elemento se vincula a una versión de la documentación de referencia de la API que es compatible con el conjunto de requisitos (por ejemplo, ExcelApi 1,3 muestra las API de ExcelApi 1,1, 1,2, 1,3 y la API común).</span><span class="sxs-lookup"><span data-stu-id="9bfa8-107">Each item links to a version of the API reference documentation that is supported by that requirement set (e.g. ExcelApi 1.3 shows APIs in ExcelApi 1.1, 1.2, 1.3 as well as the Common API).</span></span>

<span data-ttu-id="9bfa8-108">`ExcelApiOnline 1.1`es un conjunto de requisitos especial.</span><span class="sxs-lookup"><span data-stu-id="9bfa8-108">`ExcelApiOnline 1.1` is a special requirement set.</span></span> <span data-ttu-id="9bfa8-109">Contiene las API más recientes para Excel en la web, pero es posible que todavía no se admitan por completo en todas las plataformas.</span><span class="sxs-lookup"><span data-stu-id="9bfa8-109">It contains the latest APIs for Excel on the web, but those APIs may not yet be fully supported across all platforms.</span></span> <span data-ttu-id="9bfa8-110">Para obtener más información [, vea el conjunto de requisitos de la API de JavaScript de Excel en línea](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set) .</span><span class="sxs-lookup"><span data-stu-id="9bfa8-110">See [Excel JavaScript API online-only requirement set](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set) for more information.</span></span>

> [!TIP]
> <span data-ttu-id="9bfa8-111">Elija un vínculo en esta página para ver la documentación de referencia de las API compatibles con el conjunto de requisitos especificado o use el menú desplegable de selección de filtro que se encuentra encima de la tabla de contenido para cambiar el conjunto de requisitos en cualquier momento.</span><span class="sxs-lookup"><span data-stu-id="9bfa8-111">Choose a link on this page to view reference documentation for APIs supported by the specified requirement set, or use the filter selection drop-down menu above the table of contents to change the requirement set at any time.</span></span>

## <a name="excel"></a><span data-ttu-id="9bfa8-112">Excel</span><span class="sxs-lookup"><span data-stu-id="9bfa8-112">Excel</span></span>

- [<span data-ttu-id="9bfa8-113">Vista previa de ExcelApi</span><span class="sxs-lookup"><span data-stu-id="9bfa8-113">ExcelApi Preview</span></span>](/javascript/api/excel?view=excel-js-preview)
- [<span data-ttu-id="9bfa8-114">ExcelApiOnline 1,1</span><span class="sxs-lookup"><span data-stu-id="9bfa8-114">ExcelApiOnline 1.1</span></span>](/javascript/api/excel?view=excel-js-online)
- [<span data-ttu-id="9bfa8-115">ExcelApi 1,11</span><span class="sxs-lookup"><span data-stu-id="9bfa8-115">ExcelApi 1.11</span></span>](/javascript/api/excel?view=excel-js-1.11)
- [<span data-ttu-id="9bfa8-116">ExcelApi 1.10</span><span class="sxs-lookup"><span data-stu-id="9bfa8-116">ExcelApi 1.10</span></span>](/javascript/api/excel?view=excel-js-1.10)
- [<span data-ttu-id="9bfa8-117">ExcelApi 1.9</span><span class="sxs-lookup"><span data-stu-id="9bfa8-117">ExcelApi 1.9</span></span>](/javascript/api/excel?view=excel-js-1.9)
- [<span data-ttu-id="9bfa8-118">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="9bfa8-118">ExcelApi 1.8</span></span>](/javascript/api/excel?view=excel-js-1.8)
- [<span data-ttu-id="9bfa8-119">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="9bfa8-119">ExcelApi 1.7</span></span>](/javascript/api/excel?view=excel-js-1.7)
- [<span data-ttu-id="9bfa8-120">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="9bfa8-120">ExcelApi 1.6</span></span>](/javascript/api/excel?view=excel-js-1.6)
- [<span data-ttu-id="9bfa8-121">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="9bfa8-121">ExcelApi 1.5</span></span>](/javascript/api/excel?view=excel-js-1.5)
- [<span data-ttu-id="9bfa8-122">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="9bfa8-122">ExcelApi 1.4</span></span>](/javascript/api/excel?view=excel-js-1.4)
- [<span data-ttu-id="9bfa8-123">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="9bfa8-123">ExcelApi 1.3</span></span>](/javascript/api/excel?view=excel-js-1.3)
- [<span data-ttu-id="9bfa8-124">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="9bfa8-124">ExcelApi 1.2</span></span>](/javascript/api/excel?view=excel-js-1.2)
- [<span data-ttu-id="9bfa8-125">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="9bfa8-125">ExcelApi 1.1</span></span>](/javascript/api/excel?view=excel-js-1.1)

## <a name="onenote"></a><span data-ttu-id="9bfa8-126">OneNote</span><span class="sxs-lookup"><span data-stu-id="9bfa8-126">OneNote</span></span>

- [<span data-ttu-id="9bfa8-127">OneNote 1,1</span><span class="sxs-lookup"><span data-stu-id="9bfa8-127">OneNote 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)

## <a name="outlook"></a><span data-ttu-id="9bfa8-128">Outlook</span><span class="sxs-lookup"><span data-stu-id="9bfa8-128">Outlook</span></span>

- [<span data-ttu-id="9bfa8-129">Vista previa de buzón</span><span class="sxs-lookup"><span data-stu-id="9bfa8-129">Mailbox Preview</span></span>](/javascript/api/outlook?view=outlook-js-preview)
- [<span data-ttu-id="9bfa8-130">Mailbox 1.8</span><span class="sxs-lookup"><span data-stu-id="9bfa8-130">Mailbox 1.8</span></span>](/javascript/api/outlook?view=outlook-js-1.8)
- [<span data-ttu-id="9bfa8-131">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="9bfa8-131">Mailbox 1.7</span></span>](/javascript/api/outlook?view=outlook-js-1.7)
- [<span data-ttu-id="9bfa8-132">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="9bfa8-132">Mailbox 1.6</span></span>](/javascript/api/outlook?view=outlook-js-1.6)
- [<span data-ttu-id="9bfa8-133">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="9bfa8-133">Mailbox 1.5</span></span>](/javascript/api/outlook?view=outlook-js-1.5)
- [<span data-ttu-id="9bfa8-134">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="9bfa8-134">Mailbox 1.4</span></span>](/javascript/api/outlook?view=outlook-js-1.4)
- [<span data-ttu-id="9bfa8-135">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="9bfa8-135">Mailbox 1.3</span></span>](/javascript/api/outlook?view=outlook-js-1.3)
- [<span data-ttu-id="9bfa8-136">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="9bfa8-136">Mailbox 1.2</span></span>](/javascript/api/outlook?view=outlook-js-1.2)
- [<span data-ttu-id="9bfa8-137">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="9bfa8-137">Mailbox 1.1</span></span>](/javascript/api/outlook?view=outlook-js-1.1)

## <a name="powerpoint"></a><span data-ttu-id="9bfa8-138">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="9bfa8-138">PowerPoint</span></span>

- [<span data-ttu-id="9bfa8-139">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="9bfa8-139">PowerPointApi 1.1</span></span>](/javascript/api/powerpoint?view=powerpoint-js-1.1)

## <a name="visio"></a><span data-ttu-id="9bfa8-140">Visio</span><span class="sxs-lookup"><span data-stu-id="9bfa8-140">Visio</span></span>

- [<span data-ttu-id="9bfa8-141">VisioApi 1,1</span><span class="sxs-lookup"><span data-stu-id="9bfa8-141">VisioApi 1.1</span></span>](/javascript/api/visio?view=visio-js-1.1)

## <a name="word"></a><span data-ttu-id="9bfa8-142">Word</span><span class="sxs-lookup"><span data-stu-id="9bfa8-142">Word</span></span>

- [<span data-ttu-id="9bfa8-143">Vista previa de Word</span><span class="sxs-lookup"><span data-stu-id="9bfa8-143">Word Preview</span></span>](/javascript/api/word?view=word-js-preview)
- [<span data-ttu-id="9bfa8-144">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="9bfa8-144">WordApi 1.3</span></span>](/javascript/api/word?view=word-js-1.3)
- [<span data-ttu-id="9bfa8-145">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="9bfa8-145">WordApi 1.2</span></span>](/javascript/api/word?view=word-js-1.2)
- [<span data-ttu-id="9bfa8-146">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="9bfa8-146">WordApi 1.1</span></span>](/javascript/api/word?view=word-js-1.1)

## <a name="common-api"></a><span data-ttu-id="9bfa8-147">API común</span><span class="sxs-lookup"><span data-stu-id="9bfa8-147">Common API</span></span>

- [<span data-ttu-id="9bfa8-148">API común</span><span class="sxs-lookup"><span data-stu-id="9bfa8-148">Common API</span></span>](/javascript/api/office?view=common-js)

## <a name="see-also"></a><span data-ttu-id="9bfa8-149">Consulte también</span><span class="sxs-lookup"><span data-stu-id="9bfa8-149">See also</span></span>

- [<span data-ttu-id="9bfa8-150">Acerca de los complementos de Office</span><span class="sxs-lookup"><span data-stu-id="9bfa8-150">About Office Add-ins</span></span>](/office/dev/add-ins/overview)
- [<span data-ttu-id="9bfa8-151">Disponibilidad de plataformas y hosts de los complementos de Office</span><span class="sxs-lookup"><span data-stu-id="9bfa8-151">Office Add-in host and platform availability</span></span>](/office/dev/add-ins/overview/office-add-in-availability)
- [<span data-ttu-id="9bfa8-152">Versiones de Office y conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="9bfa8-152">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- <span data-ttu-id="9bfa8-153">[Explorar la API de JavaScript de Office con Script Lab](/office/dev/add-ins/overview/explore-with-script-lab).</span><span class="sxs-lookup"><span data-stu-id="9bfa8-153">[Explore Office JavaScript API using Script Lab](/office/dev/add-ins/overview/explore-with-script-lab)</span></span>
