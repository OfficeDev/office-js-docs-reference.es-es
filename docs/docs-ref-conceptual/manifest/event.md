# <a name="event-element"></a><span data-ttu-id="ab34a-101">Elemento Event</span><span class="sxs-lookup"><span data-stu-id="ab34a-101">Event element</span></span>

<span data-ttu-id="ab34a-102">Define un controlador de eventos en un complemento.</span><span class="sxs-lookup"><span data-stu-id="ab34a-102">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="ab34a-103">El `Event` elemento actualmente es solo compatible con Outlook en la web en Office 365.</span><span class="sxs-lookup"><span data-stu-id="ab34a-103">The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="ab34a-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="ab34a-104">Attributes</span></span>

|  <span data-ttu-id="ab34a-105">Atributo</span><span class="sxs-lookup"><span data-stu-id="ab34a-105">Attribute</span></span>  |  <span data-ttu-id="ab34a-106">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="ab34a-106">Required</span></span>  |  <span data-ttu-id="ab34a-107">Descripción</span><span class="sxs-lookup"><span data-stu-id="ab34a-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ab34a-108">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab34a-108">Type</span></span>](#type-attribute)  |  <span data-ttu-id="ab34a-109">Sí</span><span class="sxs-lookup"><span data-stu-id="ab34a-109">Yes</span></span>  | <span data-ttu-id="ab34a-110">Especifica el evento que se va a controlar.</span><span class="sxs-lookup"><span data-stu-id="ab34a-110">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="ab34a-111">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="ab34a-111">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="ab34a-112">Sí</span><span class="sxs-lookup"><span data-stu-id="ab34a-112">Yes</span></span>  | <span data-ttu-id="ab34a-p101">Especifica el estilo de ejecución para el controlador de eventos, asincrónico o sincrónico. Actualmente solo se admiten los controladores de eventos sincrónicos.</span><span class="sxs-lookup"><span data-stu-id="ab34a-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="ab34a-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="ab34a-115">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="ab34a-116">Sí</span><span class="sxs-lookup"><span data-stu-id="ab34a-116">Yes</span></span>  | <span data-ttu-id="ab34a-117">Especifica el nombre de la función para el controlador de eventos.</span><span class="sxs-lookup"><span data-stu-id="ab34a-117">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="ab34a-118">Atributo Type</span><span class="sxs-lookup"><span data-stu-id="ab34a-118">Type attribute</span></span>

<span data-ttu-id="ab34a-p102">Necesario. Especifica qué evento invocará el controlador de eventos. Los posibles valores para este atributo se especifican en la tabla siguiente.</span><span class="sxs-lookup"><span data-stu-id="ab34a-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="ab34a-122">Tipo de evento</span><span class="sxs-lookup"><span data-stu-id="ab34a-122">Event type</span></span>  |  <span data-ttu-id="ab34a-123">Descripción</span><span class="sxs-lookup"><span data-stu-id="ab34a-123">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="ab34a-124">El controlador de eventos se invocará cuando el usuario envíe un mensaje o una invitación de reunión.</span><span class="sxs-lookup"><span data-stu-id="ab34a-124">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="ab34a-125">Atributo FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="ab34a-125">FunctionExecution attribute</span></span>

<span data-ttu-id="ab34a-p103">Necesario. DEBE establecerse en `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="ab34a-p103">Required. MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="ab34a-128">Atributo FunctionName</span><span class="sxs-lookup"><span data-stu-id="ab34a-128">FunctionName attribute</span></span>

<span data-ttu-id="ab34a-p104">Necesario. Especifica el nombre de la función del controlador de eventos. Este valor debe coincidir con un nombre de función en el [archivo de función](functionfile.md) del complemento.</span><span class="sxs-lookup"><span data-stu-id="ab34a-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```