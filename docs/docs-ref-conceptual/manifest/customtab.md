# <a name="customtab-element"></a>Elemento CustomTab

En la cinta, especifique la pestaña y el grupo para sus comandos de complemento. Puede ser en la pestaña predeterminada (**Inicio**, **Mensaje** o **Reunión**), o en una pestaña personalizada definida por el complemento.

En las pestañas personalizadas, el complemento puede crear hasta 10 grupos. Cada grupo está limitado a 6 controles, independientemente de la pestaña donde aparezca. Los complementos están limitados a una pestaña personalizada.

El atributo  **id** debe ser único en el manifiesto.

## <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Sí |  Define un grupo de comandos.  |
|  [Label](#label-tab)      | Sí |  La etiqueta de CustomTab o de un grupo.  |
|  [Control](control.md)    | Sí |  Una colección de uno o más objetos Control.  |

### <a name="group"></a>Group

Necesario. Consulte [elemento Group](group.md).

### <a name="label-tab"></a>Etiqueta (pestaña)

Obligatorio. La etiqueta de la pestaña personalizada. El atributo  **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento **ShortStrings** del elemento [Resources](resources.md).


## <a name="customtab-example"></a>Ejemplo de CustomTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```