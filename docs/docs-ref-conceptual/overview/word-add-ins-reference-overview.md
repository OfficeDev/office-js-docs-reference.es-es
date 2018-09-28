# <a name="word-javascript-api-overview"></a><span data-ttu-id="c4e0d-101">Introducción a la API de JavaScript de Word</span><span class="sxs-lookup"><span data-stu-id="c4e0d-101">Word JavaScript API overview</span></span>

<span data-ttu-id="c4e0d-p101">Word proporciona un conjunto amplio de API que puede usar para crear complementos que interactúan con los metadatos y el contenido del documento. Use estas API para crear experiencias atractivas que se integran con Word y lo amplían. Puede importar y exportar contenido, ensamblar documentos nuevos desde distintos orígenes de datos e integrarlos con flujos de trabajo del documento para crear soluciones del documento personalizadas.</span><span class="sxs-lookup"><span data-stu-id="c4e0d-p101">Word provides a rich set of APIs that you can use to create add-ins that interact with document content and metadata. Use these APIs to create compelling experiences that integrate with and extend Word. You can import and export content, assemble new documents from different data sources, and integrate with document workflows to create custom document solutions.</span></span>

<span data-ttu-id="c4e0d-105">Puede usar dos API de JavaScript para interactuar con los objetos y los metadatos en un documento de Word:</span><span class="sxs-lookup"><span data-stu-id="c4e0d-105">You can use two JavaScript APIs to interact with the objects and metadata in a Word document:</span></span>

- <span data-ttu-id="c4e0d-106">API de JavaScript de Word: se introdujo en Office 2016.</span><span class="sxs-lookup"><span data-stu-id="c4e0d-106">Word JavaScript API - Introduced in Office 2016.</span></span>
- <span data-ttu-id="c4e0d-107">[API de JavaScript para Office](../javascript-api-for-office.md) (Office.js): se introdujo en Office 2013.</span><span class="sxs-lookup"><span data-stu-id="c4e0d-107">[JavaScript API for Office](../javascript-api-for-office.md) (Office.js) - Introduced in Office 2013.</span></span>

## <a name="word-javascript-api"></a><span data-ttu-id="c4e0d-108">API de JavaScript de Word</span><span class="sxs-lookup"><span data-stu-id="c4e0d-108">Word JavaScript API</span></span>

<span data-ttu-id="c4e0d-p102">La API de JavaScript de Word se carga mediante Office.js. La API de JavaScript de Word cambia la forma en que puede interactuar con objetos como documentos y párrafos. En lugar de proporcionar API asincrónicas individuales para recuperar y actualizar cada uno de estos objetos, la API de JavaScript de Word proporciona objetos "proxy" de JavaScript que corresponden a los objetos reales que se ejecutan en Word. Puede interactuar con estos objetos proxy leyendo y escribiendo sus propiedades de forma sincrónica y llamando a métodos sincrónicos para realizar operaciones en ellos. Estas interacciones con objetos proxy no se realizan inmediatamente en el script que se está ejecutando. El método **context.sync** sincroniza el estado entre el JavaScript que se está ejecutando y los objetos reales de Office. Para ello, ejecuta instrucciones situadas en la cola y recupera las propiedades de los objetos de Word cargados para usarlos en el script.</span><span class="sxs-lookup"><span data-stu-id="c4e0d-p102">The Word JavaScript API is loaded by Office.js. The Word JavaScript API changes the way that you can interact with objects like documents and paragraphs. Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the Word JavaScript API provides “proxy” JavaScript objects that correspond to the real objects running in Word. You can interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them. These interactions with proxy objects aren’t immediately realized in the running script. The **context.sync** method synchronizes the state between your running JavaScript and the real objects in Office by executing queued instructions and retrieving properties of loaded Word objects for use in your script.</span></span>

## <a name="javascript-api-for-office"></a><span data-ttu-id="c4e0d-115">API de JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="c4e0d-115">JavaScript API for Office</span></span>

<span data-ttu-id="c4e0d-116">Puede hacer referencia a Office.js desde las siguientes ubicaciones:</span><span class="sxs-lookup"><span data-stu-id="c4e0d-116">You can reference Office.js from the following locations:</span></span>

* <span data-ttu-id="c4e0d-117">https://appsforoffice.microsoft.com/lib/1/hosted/office.js-Utilice este recurso para complementos de producción.</span><span class="sxs-lookup"><span data-stu-id="c4e0d-117">https://appsforoffice.microsoft.com/lib/1/hosted/office.js - use this resource for production add-ins.</span></span>
* <span data-ttu-id="c4e0d-118">https://appsforoffice.microsoft.com/lib/beta/hosted/office.js-usar este recurso al que está tratando de las características de vista previa.</span><span class="sxs-lookup"><span data-stu-id="c4e0d-118">https://appsforoffice.microsoft.com/lib/beta/hosted/office.js - use this resource when you're trying out preview features.</span></span>

<span data-ttu-id="c4e0d-p103">Si está usando [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs), puede descargar [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) para obtener las plantillas de proyecto que incluye Office.js.  También puede usar [nuget para obtener Office.js](https://www.nuget.org/packages/Microsoft.Office.js/).</span><span class="sxs-lookup"><span data-stu-id="c4e0d-p103">If you're using [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs), you can download the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) to get project templates that include Office.js.  You can also use [nuget to get Office.js](https://www.nuget.org/packages/Microsoft.Office.js/).</span></span>

<span data-ttu-id="c4e0d-121">Si usa TypeScript y tiene npm, puede obtener las definiciones de TypeScript al escribir lo siguiente en la interfaz de la línea de comandos: `typings install office-js --ambient`.</span><span class="sxs-lookup"><span data-stu-id="c4e0d-121">If you use TypeScript and have npm, you can get the the TypeScript definitions by typing this in your command line interface: `typings install office-js --ambient`.</span></span>

## <a name="running-word-add-ins"></a><span data-ttu-id="c4e0d-122">Ejecutar complementos de Word</span><span class="sxs-lookup"><span data-stu-id="c4e0d-122">Running Word add-ins</span></span>

<span data-ttu-id="c4e0d-p104">Para ejecutar el complemento, use un controlador de eventos de Office.initialize. Para más información sobre la inicialización del complemento, vea [Información sobre la API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="c4e0d-p104">To run your add-in, use an Office.initialize event handler. For more information about add-in initialization, see [Understanding the API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office) .</span></span>

<span data-ttu-id="c4e0d-125">Complementos que Word 2016 de destino o ejecución más adelante, se pasa una función en el método **Word.run()** .</span><span class="sxs-lookup"><span data-stu-id="c4e0d-125">Add-ins that target Word 2016 or later execute by passing a function into the **Word.run()** method.</span></span> <span data-ttu-id="c4e0d-126">La función que se pasan al método **Ejecutar** debe tener un argumento de contexto.</span><span class="sxs-lookup"><span data-stu-id="c4e0d-126">The function passed into the **run** method must have a context argument.</span></span> <span data-ttu-id="c4e0d-127">Este [objeto de contexto](/javascript/api/word/word.requestcontext) es diferente del objeto de contexto que obtener desde el objeto de Office, pero también se usa para interactuar con el entorno de tiempo de ejecución de Word.</span><span class="sxs-lookup"><span data-stu-id="c4e0d-127">This [context object](/javascript/api/word/word.requestcontext) is different than the context object you get from the Office object, but it is also used to interact with the Word runtime environment.</span></span> <span data-ttu-id="c4e0d-128">El objeto context proporciona acceso al modelo de objetos de Word API de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c4e0d-128">The context object provides access to the Word JavaScript API object model.</span></span> <span data-ttu-id="c4e0d-129">En el ejemplo siguiente se muestra cómo inicializar y ejecutar una palabra complemento mediante el método **Word.run()** .</span><span class="sxs-lookup"><span data-stu-id="c4e0d-129">The following example shows how to initialize and execute a Word add-in by using the **Word.run()** method.</span></span>

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a><span data-ttu-id="c4e0d-130">Sincronizar documentos de Word con objetos proxy de la API de JavaScript de Word</span><span class="sxs-lookup"><span data-stu-id="c4e0d-130">Synchronizing Word documents with Word JavaScript API proxy objects</span></span>

<span data-ttu-id="c4e0d-p106">El modelo de objetos de la API de JavaScript de Word se acopla libremente a los objetos de Word. Los objetos de la API de JavaScript de Word son servidores proxy para objetos en un documento de Word. Las acciones que se realizan en los objetos proxy no se realizan en Word hasta que se haya sincronizado el estado del documento. En cambio, el estado del documento de Word no se realiza en los objetos proxy hasta que se haya sincronizado el estado del documento. Para sincronizar el estado del documento, ejecute el método **context.sync()**. En el siguiente ejemplo se crea un objeto proxy de cuerpo y un comando en la cola para cargar la propiedad del texto en el objeto proxy de cuerpo y se usa **context.sync()** para sincronizar el cuerpo del documento de Word con el objeto proxy de cuerpo.</span><span class="sxs-lookup"><span data-stu-id="c4e0d-p106">The Word JavaScript API object model is loosely coupled with the objects in Word. Word JavaScript API objects are proxies for objects in a Word document. Actions taken on proxy objects are not realized in Word until the document state has been synchronized. Conversely, the state of the Word document is not realized in the proxy objects until the document state has been synchronized. To synchronize the document state, you run the **context.sync()** method. The following example creates a proxy body object and a queued command to load the text property on the proxy body object, and uses the **context.sync()** method to synchronize the body of the Word document with the body proxy object.</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a><span data-ttu-id="c4e0d-137">Ejecutar un lote de comandos</span><span class="sxs-lookup"><span data-stu-id="c4e0d-137">Executing a batch of commands</span></span>

<span data-ttu-id="c4e0d-p107">Los objetos proxy de Word disponen de métodos para obtener acceso y actualizar el modelo de objetos. Estos métodos se ejecutan secuencialmente en el orden en el que se pusieron en la cola en el lote. Todos los comandos que están en la cola en el lote se ejecutan cuando se llama a context.sync().</span><span class="sxs-lookup"><span data-stu-id="c4e0d-p107">The Word proxy objects have methods for accessing and updating the object model. These methods are executed sequentially in the order in which they were queued in the batch. All of the commands that are queued in the batch are executed when context.sync() is called.</span></span>

<span data-ttu-id="c4e0d-p108">En el siguiente ejemplo se muestra cómo funciona la cola de comandos. Cuando se llama a **context.sync()**, el comando que va a cargar el texto del cuerpo se ejecuta en Word. Después se ejecuta el comando que va a insertar texto en el cuerpo en Word. Luego se devuelven los resultados al objeto proxy de cuerpo. El valor de la propiedad **body.text** en la API de JavaScript de Word será el valor del cuerpo del documento de Word <u>antes</u> de que el texto se insertase en el documento de Word.</span><span class="sxs-lookup"><span data-stu-id="c4e0d-p108">The following example shows how the command queue works. When **context.sync()** is called, the command to load the body text is executed in Word. Then, the command to insert text into the body in Word occurs. The results are then returned to the body proxy object. The value of the **body.text** property in the Word JavaScript API is the value of the Word document body <u>before</u> the text was inserted into Word document.</span></span>


```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    context.load(body, 'text');

    // Queue a command to insert text into the end of the Word document body.
    body.insertText('This is text inserted after loading the body.text property',
                    Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

## <a name="word-javascript-api-open-specifications"></a><span data-ttu-id="c4e0d-146">Especificaciones abiertas de API de JavaScript de Word</span><span class="sxs-lookup"><span data-stu-id="c4e0d-146">Word JavaScript API open specifications</span></span>

<span data-ttu-id="c4e0d-p109">Cuando diseñemos y desarrollemos nuevas API para complementos de Word, estarán disponibles y nos podrá enviar sus comentarios en la página [Especificaciones de la API abierta](../openspec.md). Descubra las nuevas características que están en proceso para las API de JavaScript para Word y envíe sus comentarios sobre nuestras especificaciones de diseño.</span><span class="sxs-lookup"><span data-stu-id="c4e0d-p109">As we design and develop new APIs for Word add-ins, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page. Find out what new features are in the pipeline for the Word JavaScript APIs, and provide your input on our design specifications.</span></span>

## <a name="word-javascript-api-reference"></a><span data-ttu-id="c4e0d-149">Referencia de la API de JavaScript de Word</span><span class="sxs-lookup"><span data-stu-id="c4e0d-149">Word JavaScript API reference</span></span>

<span data-ttu-id="c4e0d-150">Para obtener información detallada acerca de la API de JavaScript de Word, consulte la [documentación de referencia de la API de JavaScript de Word](/javascript/api/word).</span><span class="sxs-lookup"><span data-stu-id="c4e0d-150">For detailed information about the Word JavaScript API, see the [Word JavaScript API reference documentation](/javascript/api/word).</span></span>

## <a name="see-also"></a><span data-ttu-id="c4e0d-151">Ver también</span><span class="sxs-lookup"><span data-stu-id="c4e0d-151">See also</span></span>

* [<span data-ttu-id="c4e0d-152">Informaci?n general sobre los complementos de Word</span><span class="sxs-lookup"><span data-stu-id="c4e0d-152">Word add-ins overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/word/word-add-ins-programming-overview)
* [<span data-ttu-id="c4e0d-153">Informaci?n general sobre la plataforma de complementos de Office</span><span class="sxs-lookup"><span data-stu-id="c4e0d-153">Office Add-ins platform overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* [<span data-ttu-id="c4e0d-154">Ejemplos de complementos de Word en GitHub</span><span class="sxs-lookup"><span data-stu-id="c4e0d-154">Word add-in samples on GitHub</span></span>](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Word)
