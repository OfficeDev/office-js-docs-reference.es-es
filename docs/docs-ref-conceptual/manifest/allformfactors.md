# <a name="allformfactors-element"></a>Elemento AllFormFactors

Especifica la configuración de un complemento para todos los factores de forma. Actualmente, la única característica con **AllFormFactors** es funciones personalizadas. **AllFormFactors** es un elemento obligatorio al uso de funciones personalizadas.

## <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  Sí |  Define dónde expone su funcionalidad un complemento. |

## <a name="allformfactors-example"></a>Ejemplo de AllFormFactors

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
