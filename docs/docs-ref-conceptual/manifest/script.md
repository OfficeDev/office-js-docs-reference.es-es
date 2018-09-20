# <a name="script-element"></a>Elemento Script

Define la configuración de scripts que usa una función personalizada en Excel.

## <a name="attributes"></a>Atributos

Ninguno

## <a name="child-elements"></a>Elementos secundarios

|Elementos  |  Necesario  |  Descripción  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Sí  | Cadena con el identificador de recurso del archivo JavaScript usado por funciones personalizadas.|

## <a name="example"></a>Ejemplo

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
