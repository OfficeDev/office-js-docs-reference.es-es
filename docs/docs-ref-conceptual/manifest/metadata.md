# <a name="metadata-element"></a>Elemento de metadatos

Define la configuración de metadatos usada por una función personalizada en Excel.

## <a name="attributes"></a>Atributos

Ninguno

## <a name="child-elements"></a>Elementos secundarios

|  Elemento  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Sí  | Cadena con el identificador de recurso del archivo JSON usado por funciones personalizadas. |

## <a name="example"></a>Ejemplo

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
