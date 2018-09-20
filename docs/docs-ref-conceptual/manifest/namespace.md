# <a name="namespace-element"></a>Elemento Namespace

Define el espacio de nombres usado por una función personalizada en Excel.

## <a name="attributes"></a>Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  **RESID = "espacio de nombres"**  |  Sí  | Debe coincidir con el título de ShortStrings para la función personalizada, especificado dentro del elemento de [recursos](resources.md) . |

## <a name="child-elements"></a>Elementos secundarios

Ninguno

## <a name="example"></a>Ejemplo

```xml
<Namespace resid="namespace" />
```
