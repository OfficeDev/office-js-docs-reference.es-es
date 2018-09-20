# <a name="icon-element"></a>Elemento Icon

Defina los elementos **Image** para los controles [Button](control.md#button-control) o [Menu](control.md#menu-dropdown-button-controls).

## <a name="attributes"></a>Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  **xsi:type**  |  No  | El tipo de icono que se está definiendo. Esto se aplica a los iconos de factores de forma móvil. Los elementos **Icon** que se incluyen en un elemento [MobileFormFactor](mobileformfactor.md) deben tener este atributo establecido en `bt:MobileIconList`. |

## <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [Image](#image)        | Sí |   resid de una imagen que se usará         |

### <a name="image"></a>Image

Imagen del botón. El atributo **resid** tiene que establecerse en el valor del atributo **id** de un elemento **Image** en el elemento **Images** del elemento [Resources](resources.md). El atributo **size** indica el tamaño en píxeles de la imagen. Se necesitan tres tamaños de imágenes (16, 32 y 80 píxeles), mientras que se admiten otros cinco tamaños (20, 24, 40, 48 y 64 píxeles).|

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

## <a name="additional-requirements-for-mobile-form-factors"></a>Requisitos adicionales para los factores de forma móvil

Cuando el elemento **Icon** primario es un descendiente de un elemento [MobileFormFactor](mobileformfactor.md), los tamaños mínimos necesarios son ligeramente diferentes. El manifiesto debe proporcionar como mínimo tamaños de 25, 32 y 48 píxeles. Cada tamaño proporcionado debe aparecer tres veces, con un atributo `scale` establecido en `1`, `2` o `3`.

```xml
<Icon xsi:type="bt:MobileIconList">
  <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
  <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
  <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
  <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
  <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
  <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
  <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
  <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
  <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
</Icon>
```