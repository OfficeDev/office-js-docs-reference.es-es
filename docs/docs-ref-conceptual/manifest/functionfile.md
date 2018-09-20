# <a name="functionfile-element"></a>Elemento FunctionFile

Especifica el archivo de código fuente para las operaciones que expone un complemento a través de comandos que ejecutan una función de JavaScript en lugar de mostrar la interfaz de usuario. El elemento **FunctionFile** es un elemento secundario de [DesktopFormFactor](desktopformfactor.md) o de [MobileFormFactor](mobileformfactor.md). El atributo **resid** del elemento **FunctionFile** está establecido en el valor del atributo **id** de un elemento **Url** en el elemento **Resources** que contiene la dirección URL a un archivo HTML que contiene o carga todas las funciones de JavaScript que usan los botones de comandos del complemento sin interfaz de usuario, como define el [elemento Control](control.md).

El siguiente es un ejemplo del elemento **FunctionFile** .

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

El código JavaScript en el archivo HTML que se indican por el elemento **FunctionFile** debe llamar `Office.initialize` y definir las funciones con nombre que toman un solo parámetro: `event`. Deben usar las funciones de la `item.notificationMessages` API para indicar el progreso, el éxito o el fracaso al usuario. También debe llamar al `event.completed` cuando haya finalizado la ejecución. El nombre de las funciones se usan en el elemento **nombrefunción** de los botones de la interfaz de usuario menor.

A continuación se muestra un ejemplo de un archivo HTML que define una función **trackMessage**.

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

El siguiente código muestra cómo implementar la función utilizada por **nombrefunción**.

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
> La llamada a **event.completed** indica que ha controlado el evento correctamente. Al llamar a una función varias veces (por ejemplo, con varios clics en el mismo comando de complemento), todos los eventos se ponen automáticamente en cola. El primer evento se ejecuta automáticamente, mientras que los demás eventos permanecen en la cola. Cuando la función llama a **event.completed**, se ejecuta la siguiente llamada en cola a esa función. Se debe llamar a **event.completed**; en caso contrario, no se ejecutará la función.