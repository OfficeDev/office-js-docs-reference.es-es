# <a name="mobileformfactor-element"></a>Elemento MobileFormFactor

Especifica la configuración de un complemento para el factor de forma móvil. Contiene toda la información sobre complementos para dicho factor de forma móvil, excepto para el nodo **Resources**.

Cada definición **MobileFormFactor** contiene el elemento **FunctionFile** y uno o más elementos **ExtensionPoint**. Para obtener más información, vea [Elemento FunctionFile](functionfile.md) y [Elemento ExtensionPoint](extensionpoint.md).

El elemento **MobileFormFactor** se define en el esquema VersionOverrides 1.1. El elemento que contenga [VersionOverrides](versionoverrides.md) debe tener un valor de atributo `xsi:type` de `VersionOverridesV1_1`.

## <a name="child-elements"></a>Elementos secundarios

| Elemento                               | Obligatorio | Descripción  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md) | Sí      | Define dónde expone su funcionalidad un complemento. |
| [FunctionFile](functionfile.md)     | Sí      | Una dirección URL de un archivo que contiene funciones de JavaScript.|

## <a name="mobileformfactor-example"></a>Ejemplo MobileFormFactor

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
