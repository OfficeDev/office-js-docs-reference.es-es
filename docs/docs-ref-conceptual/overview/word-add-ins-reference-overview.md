# <a name="word-javascript-api-overview"></a>Introducción a la API de JavaScript de Word

Word proporciona un conjunto amplio de API que puede usar para crear complementos que interactúan con los metadatos y el contenido del documento. Use estas API para crear experiencias atractivas que se integran con Word y lo amplían. Puede importar y exportar contenido, ensamblar documentos nuevos desde distintos orígenes de datos e integrarlos con flujos de trabajo del documento para crear soluciones del documento personalizadas.

Puede usar dos API de JavaScript para interactuar con los objetos y los metadatos en un documento de Word:

- API de JavaScript de Word: se introdujo en Office 2016.
- [API de JavaScript para Office](../javascript-api-for-office.md) (Office.js): se introdujo en Office 2013.

## <a name="word-javascript-api"></a>API de JavaScript de Word

La API de JavaScript de Word se carga mediante Office.js. La API de JavaScript de Word cambia la forma en que puede interactuar con objetos como documentos y párrafos. En lugar de proporcionar API asincrónicas individuales para recuperar y actualizar cada uno de estos objetos, la API de JavaScript de Word proporciona objetos "proxy" de JavaScript que corresponden a los objetos reales que se ejecutan en Word. Puede interactuar con estos objetos proxy leyendo y escribiendo sus propiedades de forma sincrónica y llamando a métodos sincrónicos para realizar operaciones en ellos. Estas interacciones con objetos proxy no se realizan inmediatamente en el script que se está ejecutando. El método **context.sync** sincroniza el estado entre el JavaScript que se está ejecutando y los objetos reales de Office. Para ello, ejecuta instrucciones situadas en la cola y recupera las propiedades de los objetos de Word cargados para usarlos en el script.

## <a name="javascript-api-for-office"></a>API de JavaScript para Office

Puede hacer referencia a Office.js desde las siguientes ubicaciones:

* https://appsforoffice.microsoft.com/lib/1/hosted/office.js-Utilice este recurso para complementos de producción.
* https://appsforoffice.microsoft.com/lib/beta/hosted/office.js-usar este recurso al que está tratando de las características de vista previa.

Si está usando [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs), puede descargar [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) para obtener las plantillas de proyecto que incluye Office.js.  También puede usar [nuget para obtener Office.js](https://www.nuget.org/packages/Microsoft.Office.js/).

Si usa TypeScript y tiene npm, puede obtener las definiciones de TypeScript al escribir lo siguiente en la interfaz de la línea de comandos: `typings install office-js --ambient`.

## <a name="running-word-add-ins"></a>Ejecutar complementos de Word

Para ejecutar el complemento, use un controlador de eventos de Office.initialize. Para más información sobre la inicialización del complemento, vea [Información sobre la API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

Los complementos destinados a Word 2016 se ejecutan pasando una función al método **Word.run()**. La función que se pasa al método de **ejecución** debe tener un argumento de contexto. Este [objeto de contexto](/javascript/api/word/word.requestcontext) es diferente del objeto de contexto que se obtiene del objeto de Office, pero también se usa para interactuar con el entorno de runtime de Word. El objeto de contexto proporciona acceso al modelo de objetos de la API de JavaScript de Word. El siguiente ejemplo muestra cómo inicializar y ejecutar un complemento de Word mediante el método **Word.run()**.

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

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a>Sincronizar documentos de Word con objetos proxy de la API de JavaScript de Word

El modelo de objetos de la API de JavaScript de Word se acopla libremente a los objetos de Word. Los objetos de la API de JavaScript de Word son servidores proxy para objetos en un documento de Word. Las acciones que se realizan en los objetos proxy no se realizan en Word hasta que se haya sincronizado el estado del documento. En cambio, el estado del documento de Word no se realiza en los objetos proxy hasta que se haya sincronizado el estado del documento. Para sincronizar el estado del documento, ejecute el método **context.sync()**. En el siguiente ejemplo se crea un objeto proxy de cuerpo y un comando en la cola para cargar la propiedad del texto en el objeto proxy de cuerpo y se usa **context.sync()** para sincronizar el cuerpo del documento de Word con el objeto proxy de cuerpo.

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

### <a name="executing-a-batch-of-commands"></a>Ejecutar un lote de comandos

Los objetos proxy de Word disponen de métodos para obtener acceso y actualizar el modelo de objetos. Estos métodos se ejecutan secuencialmente en el orden en el que se pusieron en la cola en el lote. Todos los comandos que están en la cola en el lote se ejecutan cuando se llama a context.sync().

En el siguiente ejemplo se muestra cómo funciona la cola de comandos. Cuando se llama a **context.sync()**, el comando que va a cargar el texto del cuerpo se ejecuta en Word. Después se ejecuta el comando que va a insertar texto en el cuerpo en Word. Luego se devuelven los resultados al objeto proxy de cuerpo. El valor de la propiedad **body.text** en la API de JavaScript de Word será el valor del cuerpo del documento de Word <u>antes</u> de que el texto se insertase en el documento de Word.


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

## <a name="word-javascript-api-open-specifications"></a>Especificaciones abiertas de API de JavaScript de Word

Cuando diseñemos y desarrollemos nuevas API para complementos de Word, estarán disponibles y nos podrá enviar sus comentarios en la página [Especificaciones de la API abierta](../openspec.md). Descubra las nuevas características que están en proceso para las API de JavaScript para Word y envíe sus comentarios sobre nuestras especificaciones de diseño.

## <a name="word-javascript-api-reference"></a>Referencia de la API de JavaScript de Word

Para obtener información detallada acerca de la API de JavaScript de Word, consulte la [documentación de referencia de la API de JavaScript de Word](/javascript/api/word).

## <a name="see-also"></a>Ver también

* [Informaci?n general sobre los complementos de Word](https://docs.microsoft.com/office/dev/add-ins/word/word-add-ins-programming-overview)
* [Informaci?n general sobre la plataforma de complementos de Office](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* [Ejemplos de complementos de Word en GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Word)
