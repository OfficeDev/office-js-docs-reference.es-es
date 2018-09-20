# <a name="versionoverrides-element"></a>Elemento VersionOverrides

Elemento raíz que contiene la información de los comandos del complemento implementados por el complemento. **VersionOverrides** es un elemento secundario del elemento [OfficeApp](./officeapp.md) del manifiesto. Este elemento se admite en la versión 1.1 del esquema del manifiesto y en las posteriores, pero se define en el esquema de la versión 1.0 o 1.1 de VersionOverrides.

## <a name="attributes"></a>Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  **xmlns**       |  Sí  |  Ubicación del esquema, que tiene que ser `http://schemas.microsoft.com/office/mailappversionoverrides` cuando `xsi:type` sea `VersionOverridesV1_0` y `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` cuando `xsi:type` sea `VersionOverridesV1_1`.|
|  **xsi:type**  |  Sí  | Versión del esquema. En este momento, los únicos valores válidos son `VersionOverridesV1_0` y `VersionOverridesV1_1`. |

> [!NOTE]
> Actualmente sólo 2016 de Outlook es compatible con el esquema v1.1 de VersionOverrides y la `VersionOverridesV1_1` tipo.

## <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  **Descripción**    |  No   |  Describe el complemento. Esto reemplaza el elemento `Description` en cualquier parte principal del manifiesto. El texto de la descripción está contenido en un elemento secundario del elemento **LongString**, contenido en el elemento [Resources](./resources.md). El atributo `resid` del elemento **Description** está establecido en el valor del atributo `id` del elemento `String` que contiene el texto.|
|  **Requisitos**  |  No   |  Especifica el conjunto de requisitos mínimos y la versión de Office.js que necesita el complemento. Esto reemplaza el elemento `Requirements` en la parte principal del manifiesto.|
|  [Hosts](./hosts.md)                |  Sí  |  Especifica una colección de hosts de Office. El elemento Hosts secundario reemplaza el elemento Hosts en cualquier parte principal del manifiesto.  |
|  [Recursos](./resources.md)    |  Sí  | Define una colección de recursos (cadenas, direcciones URL e imágenes) a las que hacen referencia otros elementos del manifiesto.|
|  **VersionOverrides**    |  No  | Define comandos de complemento en una versión más reciente del esquema. Consulte [Implementar varias versiones](#implementing-multiple-versions) para obtener detalles. |
|  **WebApplicationInfo**    |  No  | Especifica detalles sobre la aplicación web asociada al complemento. |



### <a name="versionoverrides-example"></a>Ejemplo de VersionOverrides
```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a>Implementar varias versiones

Un manifiesto puede implementar varias versiones del elemento `VersionOverrides` que admiten diferentes versiones del esquema de VersionOverrides. Esta acción puede realizarse para tener la opción de admitir nuevas características en un esquema más reciente, pero, a la vez, admitir clientes anteriores que no sean compatibles con las nuevas características.

Para poder implementar varias versiones, el elemento `VersionOverrides` de la versión más reciente deberá ser un elemento secundario del elemento `VersionOverrides` de la versión anterior. El elemento `VersionOverrides` secundario no hereda ningún valor del elemento primario.

Para implementar tanto el esquema de la versión 1.0 como el de la versión 1.1 de VersionOverrides, el manifiesto debería ser similar al del siguiente ejemplo:

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
...
</OfficeApp>
```
