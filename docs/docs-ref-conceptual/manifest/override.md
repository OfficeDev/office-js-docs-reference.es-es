# <a name="override-element"></a>Elemento Override

Permite especificar el valor de una opción para una configuración regional adicional.

**Tipo de complemento:** Contenido, panel de tareas, correo

## <a name="syntax"></a>Sintaxis

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a>Contenidos en

|**Element**|
|:-----|
|[CitationText](citationtext.md)|
|[Descripción](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[DisplayName](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|

## <a name="attributes"></a>Atributos

|**Attribute**|**Tipo**|**Obligatorio**|**Descripción**|
|:-----|:-----|:-----|:-----|
|Locale|string|necesario|Especifica el nombre de referencia cultural de la configuración regional para esta invalidación en el formato de etiqueta de lenguaje BCP 47, como en `"en-US"`.|
|Valor|string|necesario|Especifica el valor de la opción de configuración expresado para la configuración regional especificada.|

## <a name="see-also"></a>Recursos adicionales

- [Localización de complementos de Office](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
