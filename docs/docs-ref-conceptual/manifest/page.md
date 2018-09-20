# <a name="page-element"></a>Elemento Page

Define la configuración de páginas HTLM que usa una función personalizada en Excel.

## <a name="attributes"></a>Atributos

Ninguno

## <a name="child-elements"></a>Elementos secundarios

|  Elemento  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Sí  | Cadena con el identificador de recurso del archivo HTML usado por funciones personalizadas. |

## <a name="example"></a>Ejemplo

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
