# <a name="visio-javascript-api-overview"></a><span data-ttu-id="226b5-101">Introducción a la API de JavaScript de Visio</span><span class="sxs-lookup"><span data-stu-id="226b5-101">Visio JavaScript API overview</span></span>

<span data-ttu-id="226b5-102">Puede utilizar las API de JavaScript Visio para incrustar los diagramas de Visio en SharePoint Online.</span><span class="sxs-lookup"><span data-stu-id="226b5-102">You can use the Visio JavaScript APIs to embed Visio diagrams in SharePoint Online.</span></span> <span data-ttu-id="226b5-103">Un diagrama de Visio incrustado es un diagrama que se almacena en una biblioteca de documentos de SharePoint y que se muestra en una página de SharePoint.</span><span class="sxs-lookup"><span data-stu-id="226b5-103">An embedded Visio diagram is a diagram that is stored in a SharePoint document library and displayed on a SharePoint page.</span></span> <span data-ttu-id="226b5-104">Para incrustar un diagrama de Visio, mostrarlos en un elemento HTML `<iframe>` elemento.</span><span class="sxs-lookup"><span data-stu-id="226b5-104">To embed a Visio diagram, display it in an HTML `<iframe>` element.</span></span> <span data-ttu-id="226b5-105">A continuación, puede utilizar las API de JavaScript de Visio para trabajar mediante programación con el diagrama incrustado.</span><span class="sxs-lookup"><span data-stu-id="226b5-105">Then you can use Visio JavaScript APIs to programmatically work with the embedded diagram.</span></span>

![El diagrama de Visio en iframe en la página de SharePoint junto con el elemento web de editor de scripts](../images/visio-api-block-diagram.png)


<span data-ttu-id="226b5-107">Puede utilizar las API de JavaScript para Visio para:</span><span class="sxs-lookup"><span data-stu-id="226b5-107">You can use the Visio JavaScript APIs to:</span></span>

* <span data-ttu-id="226b5-108">Interactuar con los elementos de diagramas de Visio como páginas y formas.</span><span class="sxs-lookup"><span data-stu-id="226b5-108">Interact with Visio diagram elements like pages and shapes.</span></span>
* <span data-ttu-id="226b5-109">Crear marcado visual en el lienzo de diagrama de Visio.</span><span class="sxs-lookup"><span data-stu-id="226b5-109">Create visual markup on the Visio diagram canvas.</span></span>
* <span data-ttu-id="226b5-110">Escribir controladores personalizados para los eventos del mouse en el dibujo.</span><span class="sxs-lookup"><span data-stu-id="226b5-110">Write custom handlers for mouse events within the drawing.</span></span>
* <span data-ttu-id="226b5-111">Exponer los datos del diagrama, como texto de forma, datos de forma e hipervínculos, a su solución.</span><span class="sxs-lookup"><span data-stu-id="226b5-111">Expose diagram data, such as shape text, shape data, and hyperlinks, to your solution.</span></span>

<span data-ttu-id="226b5-p102">En este artículo se describe cómo usar las API de JavaScript para Visio con Visio Online para crear sus soluciones para SharePoint Online. Es una introducción a conceptos clave que son fundamentales para usar las API, como **EmbeddedSession**, **RequestContext** y objetos proxy de JavaScript, y los métodos **sync()**, **Visio.run()** y **load()**. Los ejemplos de código muestran cómo aplicar estos conceptos.</span><span class="sxs-lookup"><span data-stu-id="226b5-p102">This article describes how to use the Visio JavaScript APIs with Visio Online to build your solutions for SharePoint Online. It introduces key concepts that are fundamental to using the APIs, such as **EmbeddedSession**, **RequestContext**, and JavaScript proxy objects, and the **sync()**, **Visio.run()**, and **load()** methods. The code examples show you how to apply these concepts.</span></span>

## <a name="embeddedsession"></a><span data-ttu-id="226b5-115">EmbeddedSession</span><span class="sxs-lookup"><span data-stu-id="226b5-115">EmbeddedSession</span></span>

<span data-ttu-id="226b5-116">El objeto EmbeddedSession inicializa la comunicación entre el marco del desarrollador y el marco de Visio Online.</span><span class="sxs-lookup"><span data-stu-id="226b5-116">The EmbeddedSession object initializes communication between the developer frame and the Visio Online frame.</span></span>

```js
var session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
session.init().then(function () {    
    window.console.log("Session successfully initialized");
});
```

## <a name="visiorunsession-functioncontext--batch-"></a><span data-ttu-id="226b5-117">Visio.Run (sesión, function(context) {lote})</span><span class="sxs-lookup"><span data-stu-id="226b5-117">Visio.run(session, function(context) { batch })</span></span>

<span data-ttu-id="226b5-118">**Visio.run()** ejecuta un script por lotes que realiza acciones en el modelo de objetos de Visio.</span><span class="sxs-lookup"><span data-stu-id="226b5-118">**Visio.run()** executes a batch script that performs actions on the Visio object model.</span></span> <span data-ttu-id="226b5-119">Los comandos por lotes incluyen definiciones de objetos proxy locales de JavaScript y métodos **sync()** que sincronizan el estado entre los objetos locales y de Visio y la resolución de la promesa.</span><span class="sxs-lookup"><span data-stu-id="226b5-119">The batch commands include definitions of local JavaScript proxy objects and **sync()** methods that synchronize the state between local and Visio objects and promise resolution.</span></span> <span data-ttu-id="226b5-120">La ventaja de procesamiento por lotes de las solicitudes en **Visio.run()** es que, cuando se resuelve la promesa, los objetos de la página de los que se realiza el seguimiento y que se asignaron durante la ejecución se liberarán automáticamente.</span><span class="sxs-lookup"><span data-stu-id="226b5-120">The advantage of batching requests in **Visio.run()** is that when the promise is resolved, any tracked page objects that were allocated during the execution will be automatically released.</span></span> 

<span data-ttu-id="226b5-121">El método de ejecución toma en la sesión y objeto RequestContext y devuelve una promesa (normalmente, sólo el resultado de **context.sync()**).</span><span class="sxs-lookup"><span data-stu-id="226b5-121">The run method takes in session and RequestContext object and returns a promise (typically, just the result of **context.sync()**).</span></span> <span data-ttu-id="226b5-122">Es posible ejecutar la operación por lotes fuera de **Visio.run()**.</span><span class="sxs-lookup"><span data-stu-id="226b5-122">It is possible to run the batch operation outside of the **Visio.run()**.</span></span> <span data-ttu-id="226b5-123">Sin embargo, en este caso, todas las referencias a objetos de página deben seguirse y administrarse manualmente.</span><span class="sxs-lookup"><span data-stu-id="226b5-123">However, in such a scenario, any page object references needs to be manually tracked and managed.</span></span> 

## <a name="requestcontext"></a><span data-ttu-id="226b5-124">RequestContext</span><span class="sxs-lookup"><span data-stu-id="226b5-124">RequestContext</span></span>

<span data-ttu-id="226b5-125">El objeto RequestContext facilita las solicitudes a la aplicación de Visio.</span><span class="sxs-lookup"><span data-stu-id="226b5-125">The RequestContext object facilitates requests to the Visio application.</span></span> <span data-ttu-id="226b5-126">Debido a que el marco para desarrolladores y la aplicación de Visio en línea se ejecutan en dos IFRAME diferente, el objeto RequestContext (contexto en el ejemplo siguiente) es necesaria para obtener acceso a Visio y objetos relacionados como páginas y formas, desde el marco para desarrolladores.</span><span class="sxs-lookup"><span data-stu-id="226b5-126">Because the developer frame and the Visio Online application run in two different iframes, the RequestContext object (context in below example) is required to get access to Visio and related objects such as pages and shapes, from the developer frame.</span></span> 

```js
function hideToolbars() {
    Visio.run(session, function(context){
        var app = context.document.application;
        app.showToolbars = false;            
        return context.sync().then(function ()
        {
            window.console.log("Toolbars Hidden");
        });      
        }).catch(function(error)
    {
        window.console.log("Error: " + error);            
    });
};
```

## <a name="proxy-objects"></a><span data-ttu-id="226b5-127">Objetos proxy</span><span class="sxs-lookup"><span data-stu-id="226b5-127">Proxy objects</span></span>

<span data-ttu-id="226b5-p106">Los objetos de JavaScript de Visio declarados y usados en un complemento son objetos proxy para los objetos reales de un documento de Visio. Las acciones llevadas a cabo en los objetos proxy no se realizan en Visio y el estado del documento de Visio no se realiza en los objetos proxy mientras no se sincronice el estado del documento. El estado del documento se sincroniza cuando se ejecuta `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="226b5-p106">The Visio JavaScript objects declared and used in an add-in are proxy objects for the real objects in a Visio document. All actions taken on proxy objects are not realized in Visio, and the state of the Visio document is not realized in the proxy objects until the document state has been synchronized. The document state is synchronized when `context.sync()` is run.</span></span>

<span data-ttu-id="226b5-131">Por ejemplo, se declara el getActivePage local de objetos de JavaScript para hacer referencia a la página seleccionada.</span><span class="sxs-lookup"><span data-stu-id="226b5-131">For example, the local JavaScript object getActivePage is declared to reference the selected page.</span></span> <span data-ttu-id="226b5-132">Esto puede usarse para poner en cola la configuración de sus propiedades y métodos de invocación.</span><span class="sxs-lookup"><span data-stu-id="226b5-132">This can be used to queue the setting of its properties and invoking methods.</span></span> <span data-ttu-id="226b5-133">Las acciones en dichos objetos no se obtienen hasta que se ejecute el método **sync()** .</span><span class="sxs-lookup"><span data-stu-id="226b5-133">The actions on such objects are not realized until the **sync()** method is run.</span></span>

```js
var activePage = context.document.getActivePage();
```

## <a name="sync"></a><span data-ttu-id="226b5-134">sync()</span><span class="sxs-lookup"><span data-stu-id="226b5-134">sync()</span></span>

<span data-ttu-id="226b5-135">El método **sync()** sincroniza el estado entre objetos de servidor proxy de JavaScript y objetos reales de Visio mediante la ejecución de instrucciones en cola en el contexto y cargan de recuperar propiedades de objetos de Office para su uso en el código.</span><span class="sxs-lookup"><span data-stu-id="226b5-135">The **sync()** method synchronizes the state between JavaScript proxy objects and real objects in Visio by executing instructions queued on the context and retrieving properties of loaded Office objects for use in your code.</span></span> <span data-ttu-id="226b5-136">Este método devuelve una promesa, que se resuelve cuando se completa la sincronización.</span><span class="sxs-lookup"><span data-stu-id="226b5-136">This method returns a promise, which is resolved when synchronization is complete.</span></span> 

## <a name="load"></a><span data-ttu-id="226b5-137">load()</span><span class="sxs-lookup"><span data-stu-id="226b5-137">load()</span></span>

<span data-ttu-id="226b5-p109">El método **load()** se usa para rellenar los objetos proxy creados en la capa de JavaScript del complemento. Al intentar recuperar un objeto, como un documento, se crea en primer lugar un objeto proxy local en la capa de JavaScript. Dicho objeto puede usarse para poner en cola la configuración de sus propiedades y métodos de invocación. Sin embargo, para leer las propiedades o las relaciones de los objetos, deben invocarse primero los métodos **load()** y **sync()**. El método load() toma las propiedades y las relaciones que necesitan cargarse cuando se llama al método**sync()**.</span><span class="sxs-lookup"><span data-stu-id="226b5-p109">The **load()** method is used to fill in the proxy objects created in the add-in JavaScript layer. When trying to retrieve an object such as a document, a local proxy object is created first in the JavaScript layer. Such an object can be used to queue the setting of its properties and invoking methods. However, for reading object properties or relations, the **load()** and **sync()** methods need to be invoked first. The load() method takes in the properties and relations that need to be loaded when the **sync()** method is called.</span></span>

<span data-ttu-id="226b5-143">A continuación, se muestra la sintaxis para el método **load()**.</span><span class="sxs-lookup"><span data-stu-id="226b5-143">The following shows the syntax for the **load()** method.</span></span>

```js
object.load(string: properties); //or object.load(array: properties); //or object.load({loadOption});
```

1. <span data-ttu-id="226b5-144">**Propiedades** es la lista de nombres de propiedad para ser cargadas, especificadas como cadenas delimitado por comas o una matriz de nombres.</span><span class="sxs-lookup"><span data-stu-id="226b5-144">**properties** is the list of property names to be loaded, specified as comma-delimited strings or array of names.</span></span> <span data-ttu-id="226b5-145">Consulte los métodos **.load()** de cada objeto para obtener más detalles.</span><span class="sxs-lookup"><span data-stu-id="226b5-145">See **.load()** methods under each object for details.</span></span>

2. <span data-ttu-id="226b5-p111">**loadOption** especifica un objeto que describe las opciones Selection, Expansion, Top y Skip. Consulte las [opciones](/javascript/api/office/officeextension.loadoption) de carga de objetos para obtener más detalles.</span><span class="sxs-lookup"><span data-stu-id="226b5-p111">**loadOption** specifies an object that describes the selection, expansion, top, and skip options. See object load [options](/javascript/api/office/officeextension.loadoption) for details.</span></span>

## <a name="example-printing-all-shapes-text-in-active-page"></a><span data-ttu-id="226b5-148">Ejemplo: Impresión de todas las formas de texto en la página activa</span><span class="sxs-lookup"><span data-stu-id="226b5-148">Example: Printing all shapes text in active page</span></span>

<span data-ttu-id="226b5-149">En el ejemplo siguiente se muestra cómo imprimir el valor de texto de forma de un objeto de formas de matriz.</span><span class="sxs-lookup"><span data-stu-id="226b5-149">The following example shows you how to print shape text value from an array shapes object.</span></span> <span data-ttu-id="226b5-150">El método **Visio.run()** contiene un lote de instrucciones.</span><span class="sxs-lookup"><span data-stu-id="226b5-150">The **Visio.run()** method contains a batch of instructions.</span></span> <span data-ttu-id="226b5-151">Como parte de este lote, se crea un objeto proxy que hace referencia a formas del documento activo.</span><span class="sxs-lookup"><span data-stu-id="226b5-151">As part of this batch, a proxy object is created that references shapes on the active document.</span></span>

<span data-ttu-id="226b5-152">Todos estos comandos se ponen en cola y se ejecutan cuando se llama a **context.sync()** .</span><span class="sxs-lookup"><span data-stu-id="226b5-152">All these commands are queued and run when **context.sync()** is called.</span></span> <span data-ttu-id="226b5-153">El método **sync()** devuelve una promesa que puede usarse para encadenarla con otras operaciones.</span><span class="sxs-lookup"><span data-stu-id="226b5-153">The **sync()** method returns a promise that can be used to chain it with other operations.</span></span>

```js
Visio.run(session, function (context) {
   var page = context.document.getActivePage();
   var shapes = page.shapes;
   shapes.load();
   return context.sync().then(function () {
        for(var i=0; i<shapes.items.length;i++)
 {
            var shape = shapes.items[i];
     window.console.log("Shape Text: " + shape.text );
 }
});
}).catch(function(error) {
  window.console.log("Error: " + error);
  if (error instanceof OfficeExtension.Error) {
       window.console.log ("Debug info: " + JSON.stringify(error.debugInfo));
  }
});
```

## <a name="error-messages"></a><span data-ttu-id="226b5-154">Mensajes de error</span><span class="sxs-lookup"><span data-stu-id="226b5-154">Error messages</span></span>

<span data-ttu-id="226b5-p114">Los errores se devuelven usando un objeto de error que consta de un código y un mensaje. En la siguiente tabla se proporciona una lista de las posibles condiciones de error que pueden producirse.</span><span class="sxs-lookup"><span data-stu-id="226b5-p114">Errors are returned using an error object that consists of a code and a message. The following table provides a list of possible error conditions that can occur.</span></span>

| <span data-ttu-id="226b5-157">error.code</span><span class="sxs-lookup"><span data-stu-id="226b5-157">error.code</span></span>            | <span data-ttu-id="226b5-158">error.message</span><span class="sxs-lookup"><span data-stu-id="226b5-158">error.message</span></span> |
|-----------------------|----------------------------------------------------------------|
| <span data-ttu-id="226b5-159">ArgumentoNoV?lido</span><span class="sxs-lookup"><span data-stu-id="226b5-159">InvalidArgument</span></span>       | <span data-ttu-id="226b5-160">El argumento no es válido, o falta o tiene un formato incorrecto.</span><span class="sxs-lookup"><span data-stu-id="226b5-160">The argument is invalid or missing or has an incorrect format.</span></span> |
| <span data-ttu-id="226b5-161">GeneralException</span><span class="sxs-lookup"><span data-stu-id="226b5-161">GeneralException</span></span>      | <span data-ttu-id="226b5-162">Se produjo un error interno al procesar la solicitud.</span><span class="sxs-lookup"><span data-stu-id="226b5-162">There was an internal error while processing the request.</span></span> |
| <span data-ttu-id="226b5-163">NoImplementado</span><span class="sxs-lookup"><span data-stu-id="226b5-163">NotImplemented</span></span>        | <span data-ttu-id="226b5-164">La característica solicitada no se implementó.</span><span class="sxs-lookup"><span data-stu-id="226b5-164">The requested feature isn't implemented.</span></span>  |
| <span data-ttu-id="226b5-165">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="226b5-165">UnsupportedOperation</span></span>  | <span data-ttu-id="226b5-166">No se admite la operación que se está intentando.</span><span class="sxs-lookup"><span data-stu-id="226b5-166">The operation being attempted is not supported.</span></span> |
| <span data-ttu-id="226b5-167">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="226b5-167">AccessDenied</span></span>          | <span data-ttu-id="226b5-168">No se puede realizar la operaci?n solicitada.</span><span class="sxs-lookup"><span data-stu-id="226b5-168">You cannot perform the requested operation.</span></span> |
| <span data-ttu-id="226b5-169">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="226b5-169">ItemNotFound</span></span>          | <span data-ttu-id="226b5-170">El recurso solicitado no existe.</span><span class="sxs-lookup"><span data-stu-id="226b5-170">The requested resource doesn't exist.</span></span> |

## <a name="get-started"></a><span data-ttu-id="226b5-171">Introducción</span><span class="sxs-lookup"><span data-stu-id="226b5-171">Get started</span></span>

<span data-ttu-id="226b5-172">Puede utilizar el ejemplo de esta sección para empezar.</span><span class="sxs-lookup"><span data-stu-id="226b5-172">You can use the example in this section to get started.</span></span> <span data-ttu-id="226b5-173">En este ejemplo se muestra cómo mostrar mediante programación el texto de la forma de la forma seleccionada en un diagrama de Visio.</span><span class="sxs-lookup"><span data-stu-id="226b5-173">This example shows you how to programmatically display the shape text of the selected shape in a Visio diagram.</span></span> <span data-ttu-id="226b5-174">Para empezar, cree una página clásica en SharePoint Online o editar una página existente.</span><span class="sxs-lookup"><span data-stu-id="226b5-174">To begin, create a classic page in SharePoint Online or edit an existing page.</span></span> <span data-ttu-id="226b5-175">Agregar un elemento Web editor de secuencias de comandos en la página y copiar y pegar el código siguiente.</span><span class="sxs-lookup"><span data-stu-id="226b5-175">Add a script editor webpart on the page and copy-paste the following code.</span></span>

```js
<script src='https://appsforoffice.microsoft.com/embedded/1.0/visio-web-embedded.js' type='text/javascript'></script>

Enter Visio File Url:<br/>
<script language="javascript">
document.write("<input type='text' id='fileUrl' size='120'/>");
document.write("<input type='button' value='InitEmbeddedFrame' onclick='initEmbeddedFrame()' />");
document.write("<br />");
document.write("<input type='button' value='SelectedShapeText' onclick='getSelectedShapeText()' />");
document.write("<textarea id='ResultOutput' style='width:350px;height:60px'> </textarea>");
document.write("<div id='iframeHost' />");

let session; // Global variable to store the session and pass it afterwards in Visio.run()
var textArea;
// Loads the Visio application and Initializes communication between devloper frame and Visio online frame
function initEmbeddedFrame() {
        textArea = document.getElementById('ResultOutput');
    var url = document.getElementById('fileUrl').value;
    if (!url) {
        window.alert("File URL should not be empty");
    }
    // APIs are enabled for EmbedView action only.   
    url = url.replace("action=view","action=embedview");
    url = url.replace("action=interactivepreview","action=embedview");
    url = url.replace("action=default","action=embedview");
    url = url.replace("action=edit","action=embedview");
  
       session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
       return session.init().then(function () {
        // Initilization is successful 
        textArea.value  = "Initilization is successful";
    });
     }

// Code for getting selected Shape Text using the shapes collection object
function getSelectedShapeText() {
    Visio.run(session, function (context) {     
       var page = context.document.getActivePage();
       var shapes = page.shapes;
       shapes.load();
           return context.sync().then(function () {
               textArea.value = "Please select a Shape in the Diagram";
               for(var i=0; i<shapes.items.length;i++)
            {
              var shape = shapes.items[i];
                  if ( shape.select == true)
               {
                textArea.value = shape.text;
                    return;
                   }
            }
      });
     }).catch(function(error) {
        textArea.value = "Error: ";
        if (error instanceof OfficeExtension.Error) {
            textArea.value += "Debug info: " + JSON.stringify(error.debugInfo);
        }
    });
}
</script>
```

<span data-ttu-id="226b5-176">Después de eso, todo lo que necesita es la dirección URL de un diagrama de Visio que se desea trabajar con.</span><span class="sxs-lookup"><span data-stu-id="226b5-176">After that, all you need is the URL of a Visio diagram that you want to work with.</span></span> <span data-ttu-id="226b5-177">Acaba de cargar el diagrama de Visio en SharePoint Online y abrirlo en Visio en línea.</span><span class="sxs-lookup"><span data-stu-id="226b5-177">Just upload the Visio diagram to SharePoint Online and open it in Visio Online.</span></span> <span data-ttu-id="226b5-178">Desde allí, abra el cuadro de diálogo Insertar y usar la dirección URL incrustar en el ejemplo anterior.</span><span class="sxs-lookup"><span data-stu-id="226b5-178">From there, open the Embed dialog and use the Embed URL in the above example.</span></span>

![Copie la dirección URL del archivo de Visio desde el cuadro de diálogo de insertar](../images/Visio-embed-url.png)

<span data-ttu-id="226b5-180">Si usa Visio en línea en modo de edición, abra el cuadro de diálogo Insertar eligiendo **archivo** > **Compartir** > **Embed**.</span><span class="sxs-lookup"><span data-stu-id="226b5-180">If you are using Visio Online in Edit mode, open the Embed dialog by choosing **File** > **Share** > **Embed**.</span></span> <span data-ttu-id="226b5-181">Si usa Visio en línea en modo de vista, abra el cuadro de diálogo Insertar eligiendo '...' y, a continuación, **Insertar**.</span><span class="sxs-lookup"><span data-stu-id="226b5-181">If you are using Visio Online in View mode, open the Embed dialog by choosing '...' and then **Embed**.</span></span> 

## <a name="open-api-specifications"></a><span data-ttu-id="226b5-182">Especificaciones de la API pública</span><span class="sxs-lookup"><span data-stu-id="226b5-182">Open API specifications</span></span>

<span data-ttu-id="226b5-p118">Cuando diseñemos y desarrollemos nuevas API, estarán disponibles y nos podrá enviar sus comentarios en la página [Especificaciones de la API abierta](../openspec.md). Descubra las nuevas características que están en proceso y envíe sus comentarios sobre nuestras especificaciones de diseño.</span><span class="sxs-lookup"><span data-stu-id="226b5-p118">As we design and develop new APIs, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page. Find out what new features are in the pipeline, and provide your input on our design specifications.</span></span> 

## <a name="visio-javascript-api-reference"></a><span data-ttu-id="226b5-185">Referencia de la API de JavaScript de Visio</span><span class="sxs-lookup"><span data-stu-id="226b5-185">Visio JavaScript API reference</span></span>

<span data-ttu-id="226b5-186">Para obtener información detallada acerca de la API de JavaScript de Visio, vea la [documentación de referencia de la API de JavaScript de Visio](/javascript/api/visio).</span><span class="sxs-lookup"><span data-stu-id="226b5-186">For detailed information about Visio JavaScript API, see the [Visio JavaScript API reference documentation](/javascript/api/visio).</span></span>
