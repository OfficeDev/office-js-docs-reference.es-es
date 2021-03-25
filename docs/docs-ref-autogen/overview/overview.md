---
layout: LandingPage
ms.topic: landing-page
title: Referencia de la API de JavaScript de Office
description: Las API de JavaScript de Office por host y versión.
author: o365devx
ms.author: o365devx
ms.prod: non-product-specific
localization_priority: Priority
ms.date: 03/17/2021
ms.openlocfilehash: 63d2efa4d4b9406208f3227ec96df467ea0d8877
ms.sourcegitcommit: a5c179dc927ce89d1ada071d388d582191d3fa1e
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 03/24/2021
ms.locfileid: "51183543"
---
# <a name="office-add-ins-javascript-api-reference"></a>Referencia de la API de JavaScript de complementos de Office

La API de JavaScript de Office le permite crear aplicaciones web que interactúen con los modelos de objetos en las aplicaciones host de Office. Use esta sección para obtener más información sobre las clases, los métodos y otros tipos disponibles para crear complementos de Office.

La siguiente es una lista de las API para las [aplicaciones host de Office compatibles](/office/dev/add-ins/overview/office-add-in-availability). Los vínculos de API comunes incluyen todas las API no atribuidas a un host en particular (como se detalla en [Conjunto de requisitos de la API común de Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)). Otros elementos vinculan a una versión de la documentación de referencia de la API para ese host en función de un conjunto de requisitos. La documentación de referencia se incluirá en una versión para incluir todas las API existentes hasta el momento e incluirá el conjunto de requisitos (por ejemplo, ExcelApi 1.3 muestra las API de ExcelApi 1.1, 1.2, 1.3 y las API comunes).

`ExcelApiOnline 1.1` es un conjunto de requisitos especial. Contiene las API más recientes de Excel en la web, pero es posible que esas API aún no sean totalmente compatibles con todas las plataformas. Para obtener más información, vea [Conjunto de requisitos solo en línea de la API de JavaScript de Excel](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set).

> [!TIP]
> Puede cambiar la versión de una página de referencia en cualquier momento mediante el menú desplegable de selección de filtro que está encima de la tabla de contenido. Si la página no existe en esa versión específica, se le devolverá a la versión actual.

<h2>Hosts de Office</h2>

<ul class="cardsK panelContent cols cols3">
    <li>
        <div class="cardImageOuter">
            <div class="cardImage">
                <img src="https://docs.microsoft.com/javascript/api/overview/images/logo-excel.svg" alt="Excel add-ins" height="140" />
            </div>
        </div>
        <div class="cardText">
            <h3>API de Excel</h3>
            <ul>
                <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-preview">ExcelApi versión preliminar</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-online">ExcelApiOnline 1.1</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.12">ExcelApi 1.12</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.11">ExcelApi 1.11</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.10">ExcelApi 1.10</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.9">ExcelApi 1.9</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.8">ExcelApi 1.8</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.7">ExcelApi 1.7</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.6">ExcelApi 1.6</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.5">ExcelApi 1.5</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.4">ExcelApi 1.4</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.3">ExcelApi 1.3</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.2">ExcelApi 1.2</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.1">ExcelApi 1.1</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/office?view=excel-js-preview">API comunes</a></li>
            </ul>
        </div>
    </li>
    <li>
        <div class="cardImageOuter">
            <div class="cardImage">
                <img src="https://docs.microsoft.com/javascript/api/overview/images/logo-outlook.svg" alt="Outlook add-ins" height="140" />
            </div>
        </div>
        <div class="cardText">
            <h3>API de Outlook</h3>
            <ul>
                <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-preview">Mailbox versión preliminar</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.9">Mailbox 1.9</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.8">Mailbox 1.8</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.7">Mailbox 1.7</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.6">Mailbox 1.6</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.5">Mailbox 1.5</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.4">Mailbox 1.4</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.3">Mailbox 1.3</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.2">Mailbox 1.2</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.1">Mailbox 1.1</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/office?view=outlook-js-preview">API comunes</a></li>
            </ul>
        </div>
    </li>
    <li>
        <div class="cardImageOuter">
            <div class="cardImage">
                <img src="https://docs.microsoft.com/javascript/api/overview/images/logo-word.svg" alt="Word add-ins" height="140" />
            </div>
        </div>
        <div class="cardText">
            <h3>API de Word</h3>
            <ul>
                <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-preview">WordApi versión preliminar</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.3">WordApi 1.3</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.2">WordApi 1.2</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.1">WordApi 1.1</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/office?view=word-js-preview">API comunes</a></li>
            </ul>
        </div>
    </li>
    <li>
        <div class="cardImageOuter">
            <div class="cardImage">
                <img src="https://docs.microsoft.com/javascript/api/overview/images/logo-powerpoint.svg" alt="PowerPoint add-ins" height="140" />
            </div>
        </div>
        <div class="cardText">
            <h3>API de PowerPoint</h3>
            <ul>
                <li><a style="font-size: 1rem;" href="/javascript/api/powerpoint?view=powerpoint-js-preview">PowerPointApi Preview</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/powerpoint?view=powerpoint-js-1.2">PowerPointApi 1.2</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/powerpoint?view=powerpoint-js-1.1">PowerPointApi 1.1</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/office?view=powerpoint-js-preview">API comunes</a></li>
            </ul>
        </div>
    </li>
    <li>
        <div class="cardImageOuter">
            <div class="cardImage">
                <img src="https://docs.microsoft.com/javascript/api/overview/images/logo-onenote.svg" alt="OneNote add-ins" height="140" />
            </div>
        </div>
        <div class="cardText">
            <h3>API de OneNote</h3>
            <ul>
                <li><a style="font-size: 1rem;" href="/javascript/api/onenote?view=onenote-js-1.1">OneNoteApi 1.1</a></li>
                <li><a style="font-size: 1rem;" href="/javascript/api/office?view=onenote-js-1.1">API comunes</a></li>
            </ul>
        </div>
    </li>
    <li>
        <div class="cardImageOuter">
            <div class="cardImage">
                <img src="https://docs.microsoft.com/javascript/api/overview/images/logo-visio.svg" alt="Visio add-ins" height="140" />
            </div>
        </div>
        <div class="cardText">
            <h3>API de Visio</h3>
            <ul>
                <li><a style="font-size: 1rem;" href="/javascript/api/visio?view=visio-js-1.1">VisioApi 1.1</a></li>
            </ul>
        </div>
    </li>
    <li>
        <div class="cardImageOuter">
            <div class="cardImage">
                <img src="https://docs.microsoft.com/javascript/api/overview/images/logo-project.svg" alt="Project add-ins" height="140" />
            </div>
        </div>
        <div class="cardText">
            <h3>API de Project</h3>
            <ul>
                <li><a style="font-size: 1rem;" href="/javascript/api/office?view=common-js">Solo API comunes</a></li>
            </ul>
        </div>
    </li>
</ul>

> [!NOTE]
> Si está buscando las API de JavaScript para desarrollar scripts de Office, visite la [Referencia de la API de scripts de Office](/javascript/api/office-scripts/overview).

## <a name="see-also"></a>Vea también

- [Acerca de los complementos de Office](/office/dev/add-ins/overview)
- [Disponibilidad de plataformas y host de los complementos de Office](/office/dev/add-ins/overview/office-add-in-availability)
- [Versiones y conjuntos de requisitos de Office](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Explorar la API de JavaScript de Office con Script Lab](/office/dev/add-ins/overview/explore-with-script-lab).
