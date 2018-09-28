# <a name="requestedheight-element"></a><span data-ttu-id="2edcb-101">Elemento RequestedHeight.</span><span class="sxs-lookup"><span data-stu-id="2edcb-101">RequestedHeight element</span></span>

<span data-ttu-id="2edcb-102">Especifica el alto inicial (en píxeles) de un contenido complemento o un complemento en el correo.</span><span class="sxs-lookup"><span data-stu-id="2edcb-102">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="2edcb-103">**Tipo de complemento:** Contenido, correo</span><span class="sxs-lookup"><span data-stu-id="2edcb-103">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="2edcb-104">Sintaxis</span><span class="sxs-lookup"><span data-stu-id="2edcb-104">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="2edcb-105">Contenidos en</span><span class="sxs-lookup"><span data-stu-id="2edcb-105">Contained in</span></span>

- <span data-ttu-id="2edcb-106">[DefaultSettings](defaultsettings.md) (Contenido de complementos) con un valor que puede estar entre 32 y 1.000</span><span class="sxs-lookup"><span data-stu-id="2edcb-106">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="2edcb-107">[DesktopSettings](desktopsettings.md) y [TabletSettings](tabletsettings.md) (complementos de correo) con un valor que puede estar comprendido entre 32 y 450</span><span class="sxs-lookup"><span data-stu-id="2edcb-107">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="2edcb-108">[ExtensionPoint](extensionpoint.md) (Complementos de correo contextual) con un valor que puede ser entre 140 y 450 para el punto de extensión **DetectedEntity** y entre 32 y 450 para el punto de extensión **CustomPane**</span><span class="sxs-lookup"><span data-stu-id="2edcb-108">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>