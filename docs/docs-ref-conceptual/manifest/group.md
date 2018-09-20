# <a name="group-element"></a>Elemento Group

Define un grupo de controles de interfaz de usuario en un pestaña.  En las pestañas personalizadas, el complemento puede crear hasta 10 grupos. Cada grupo está limitado a 6 controles, independientemente de la pestaña donde aparezca. Los complementos están limitados a una pestaña personalizada.

## <a name="attributes"></a>Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Sí  | Un identificador único para el grupo.|

### <a name="id-attribute"></a>Atributo id

Necesario. Identificador único para el grupo. Es una cadena con un máximo de 125 caracteres. Debe ser único dentro del manifiesto o el grupo no podrá procesarse.

## <a name="child-elements"></a>Elementos secundarios
|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [Label](#label)      | Sí |  La etiqueta de CustomTab o de un grupo.  |
|  [Control](#control)    | Sí |  Colección de uno o más objetos Control.  |

### <a name="label"></a>Label 

Obligatorio. La etiqueta del grupo. El atributo  **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento **ShortStrings** del elemento [Resources](resources.md).

### <a name="control"></a>Control
Un grupo necesita como mínimo un control.

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```