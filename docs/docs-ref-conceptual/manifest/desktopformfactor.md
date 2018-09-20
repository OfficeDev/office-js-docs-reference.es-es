# <a name="desktopformfactor-element"></a>Elemento DesktopFormFactor

Especifica la configuración de un complemento para un factor de forma de escritorio. El factor de forma de escritorio incluye Office para Windows, Office para Mac y Office Online. Contiene toda la información sobre complementos para dicho factor de forma, excepto para el nodo **Resources**.

Cada definición DesktopFormFactor contiene el elemento **FunctionFile** y uno o más elementos **ExtensionPoint**. Para obtener más información, consulte [Elemento FunctionFile](functionfile.md) y [Elemento ExtensionPoint](extensionpoint.md). 

## <a name="child-elements"></a>Elementos secundarios

| Elemento                               | Obligatorio | Descripción  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md) | Sí      | Define dónde expone su funcionalidad un complemento. |
| [FunctionFile](functionfile.md)     | Sí      | Una dirección URL de un archivo que contiene funciones de JavaScript.|
| [GetStarted](getstarted.md)         | No       | Define la llamada que aparece cuando se instala el complemento en hosts de Word, Excel o PowerPoint. |

## <a name="desktopformfactor-example"></a>Ejemplo de DesktopFormFactor

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
