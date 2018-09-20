# <a name="sourcelocation-element"></a>Elemento SourceLocation

Define la ubicación de un recurso que requieren los elementos Script o Page usados por funciones personalizadas de Excel.

## <a name="attributes"></a>Atributos

| **Atributo** | **Necesario** | **Descripción**                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| resid         | Sí          | Nombre de un recurso de URL definido en la sección &lt;Recursos&gt; del manifiesto. |

## <a name="child-elements"></a>Elementos secundarios

Ninguno

## <a name="example"></a>Ejemplo

```xml
<SourceLocation resid="pageURL"/>
```