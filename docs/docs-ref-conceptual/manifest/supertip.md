# <a name="supertip"></a>Supertip

Defina una información sobre herramientas enriquecida (título y descripción). Lo usan los controles [Button](control.md#button-control) y [Menu](control.md#menu-dropdown-button-controls).

## <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [Título](#title)        | Sí |   El texto de la sugerencia.         |
|  [Descripción](#description)  | Sí |  La descripción de la sugerencia.    |

### <a name="title"></a>Título

Obligatorio. Texto para la sugerencia falsa. El atributo  **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento **ShortStrings** del elemento [Resources](resources.md).

### <a name="description"></a>Descripción

Obligatorio. Descripción para la sugerencia falsa. El atributo  **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento **LongStrings** del elemento [Resources](resources.md).

## <a name="example"></a>Ejemplo

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
