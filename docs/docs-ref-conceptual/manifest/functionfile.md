# <a name="functionfile-element"></a><span data-ttu-id="623ff-101">Elemento FunctionFile</span><span class="sxs-lookup"><span data-stu-id="623ff-101">FunctionFile element</span></span>

<span data-ttu-id="623ff-p101">Especifica el archivo de código fuente para las operaciones que expone un complemento a través de comandos que ejecutan una función de JavaScript en lugar de mostrar la interfaz de usuario. El elemento **FunctionFile** es un elemento secundario de [DesktopFormFactor](desktopformfactor.md) o de [MobileFormFactor](mobileformfactor.md). El atributo **resid** del elemento **FunctionFile** está establecido en el valor del atributo **id** de un elemento **Url** en el elemento **Resources** que contiene la dirección URL a un archivo HTML que contiene o carga todas las funciones de JavaScript que usan los botones de comandos del complemento sin interfaz de usuario, como define el [elemento Control](control.md).</span><span class="sxs-lookup"><span data-stu-id="623ff-p101">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI. The  **FunctionFile** element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md). The **resid** attribute of the **FunctionFile** element is set to the value of the **id** attribute of a **Url** element in the **Resources** element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="623ff-105">El siguiente es un ejemplo del elemento **FunctionFile** .</span><span class="sxs-lookup"><span data-stu-id="623ff-105">The following is an example of the  **FunctionFile** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="623ff-106">El código JavaScript en el archivo HTML que se indican por el elemento **FunctionFile** debe llamar `Office.initialize` y definir las funciones con nombre que toman un solo parámetro: `event`.</span><span class="sxs-lookup"><span data-stu-id="623ff-106">The JavaScript in the HTML file indicated by the  **FunctionFile** element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="623ff-107">Deben usar las funciones de la `item.notificationMessages` API para indicar el progreso, el éxito o el fracaso al usuario.</span><span class="sxs-lookup"><span data-stu-id="623ff-107">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="623ff-108">También debe llamar al `event.completed` cuando haya finalizado la ejecución.</span><span class="sxs-lookup"><span data-stu-id="623ff-108">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="623ff-109">El nombre de las funciones se usan en el elemento **nombrefunción** de los botones de la interfaz de usuario menor.</span><span class="sxs-lookup"><span data-stu-id="623ff-109">The name of the functions are used in the **FunctionName** element for UI-less buttons.</span></span>

<span data-ttu-id="623ff-110">A continuación se muestra un ejemplo de un archivo HTML que define una función **trackMessage**.</span><span class="sxs-lookup"><span data-stu-id="623ff-110">The following is an example of an HTML file defining a **trackMessage** function.</span></span>

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

<span data-ttu-id="623ff-111">El siguiente código muestra cómo implementar la función utilizada por **nombrefunción**.</span><span class="sxs-lookup"><span data-stu-id="623ff-111">The following code shows how to implement the function used by  **FunctionName**.</span></span>

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

> [!IMPORTANT]
> <span data-ttu-id="623ff-112">La llamada a **event.completed** indica que ha controlado el evento correctamente.</span><span class="sxs-lookup"><span data-stu-id="623ff-112">The call to  **event.completed** signals that you have successfully handled the event.</span></span> <span data-ttu-id="623ff-113">Al llamar a una función varias veces (por ejemplo, con varios clics en el mismo comando de complemento), todos los eventos se ponen automáticamente en cola.</span><span class="sxs-lookup"><span data-stu-id="623ff-113">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="623ff-114">El primer evento se ejecuta automáticamente, mientras que los demás eventos permanecen en la cola.</span><span class="sxs-lookup"><span data-stu-id="623ff-114">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="623ff-115">Cuando la función llama a **event.completed**, se ejecuta la siguiente llamada en cola a esa función.</span><span class="sxs-lookup"><span data-stu-id="623ff-115">When your function calls **event.completed**, the next queued call to that function runs.</span></span> <span data-ttu-id="623ff-116">Se debe llamar a **event.completed**; en caso contrario, no se ejecutará la función.</span><span class="sxs-lookup"><span data-stu-id="623ff-116">You must call **event.completed**; otherwise your function will not run.</span></span>