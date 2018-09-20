# <a name="rule-element"></a>Elemento Rule

Especifica las reglas de activación que deberían evaluarse para este complemento de correo contextual.

**Tipo de complemento:** Complemento contextual de correo

## <a name="contained-in"></a>Contenidos en

- [OfficeApp](officeapp.md)
- [ExtensionPoint](extensionpoint.md)

## <a name="attributes"></a>Atributos

| Atributo | Obligatorio | Descripción |
|:-----|:-----|:-----|
| **xsi:type** | Sí | El tipo de regla que se está definiendo. |

El tipo de regla puede ser uno de los siguientes.

- [ItemIs](#itemis-rule)
- [ItemHasAttachment](#itemhasattachment-rule)
- [ItemHasKnownEntity](#itemhasknownentity-rule)
- [ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)
- [RuleCollection](#rulecollection)

## <a name="itemis-rule"></a>Regla ItemIs

Define una regla que evalúa en verdadero si el elemento seleccionado es del tipo especificado.

### <a name="attributes"></a>Atributos

| Atributo | Obligatorio | Descripción |
|:-----|:-----|:-----|
| **ItemType** | Sí | Especifica el tipo de elemento con el que debe coincidir. Puede ser `Message` o `Appointment`. El tipo de elemento `Message` incluye correo electrónico, convocatorias de reunión, respuestas a la reunión y cancelaciones de reunión. |
| **FormType** | No (dentro de [ExtensionPoint](extensionpoint.md)), Sí (dentro de [OfficeApp](officeapp.md)) | Especifica si la aplicación debe aparecer en el formulario de edición o lectura para el elemento. Puede ser uno de los siguientes: `Read`, `Edit` o `ReadOrEdit`. Si se especifica en una `Rule` dentro de un `ExtensionPoint`, este valor DEBE ser `Read`. |
| **ItemClass** | No | Especifica la clase de mensaje personalizada con la que debe coincidir. Para obtener más información, vea [Activar un complemento de correo de Outlook para una clase de mensaje específica](https://docs.microsoft.com/outlook/add-ins/activation-rules). |
| **IncludeSubClasses** | No | Especifica si la regla debería evaluar en verdadero si el elemento es de una subclase de la clase del mensaje especificada; el valor predeterminado es `false`. |

### <a name="example"></a>Ejemplo

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a>Regla ItemHasAttachment

Define una regla que evalúa en verdadero si el elemento contiene datos adjuntos.

### <a name="example"></a>Ejemplo

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a>Regla ItemHasKnownEntity

Define una regla que evalúa en verdadero si el elemento contiene texto del tipo de entidad especificado en el asunto o en el cuerpo.

### <a name="attributes"></a>Atributos

| Atributo | Obligatorio | Descripción |
|:-----|:-----|:-----|
| **EntityType** | Sí | Especifica el tipo de entidad que se tiene que encontrar para que la regla evalúe en verdadero. Puede ser uno de los siguientes: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` o `Contact`. |
| **RegExFilter** | No | Especifica una expresión regular que se debe ejecutar con esta entidad para su activación. |
| **NombreDeFiltro** | No | Especifica el nombre del filtro de expresión regular, de modo que después sea posible hacerle referencia en el código de su complemento. |
| **IgnoreCase** | No | Especifica que se ignoren las mayúsculas y minúsculas cuando se ejecute la expresión regular especificada por el atributo **RegExFilter**. |
| **Resaltar** | No | **Nota:** Esto solo se aplica en elementos **Rule** dentro de elementos **ExtensionPoint**. Especifica cómo debe resaltar el cliente las entidades coincidentes. Puede ser uno de los siguientes: `all` o `none`. Si no se especifica, el valor predeterminado es `all`. |

### <a name="example"></a>Ejemplo

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a>Regla ItemHasRegularExpressionMatch

Define una regla que evalúa en verdadero si se encuentra una coincidencia para la expresión regular especificada en la propiedad indicada del elemento.

### <a name="attributes"></a>Atributos

| Atributo | Obligatorio | Descripción |
|:-----|:-----|:-----|
| **RegExName** | Sí | Especifica el nombre de una expresión regular para que pueda hacer referencia a dicha expresión en el código de su complemento. |
| **RegExValue** | Sí | Especifica la expresión regular que se evaluará para determinar si se debe mostrar el complemento de correo. |
| **PropertyName** | Sí | Especifica el nombre de la propiedad contra la que se evaluará la expresión regular. Puede ser uno de los siguientes: `Subject`, `BodyAsPlaintext`, `BodyAsHtml` o `SenderSTMPAddress`. |
| **IgnoreCase** | No | Especifica que se ignoren las mayúsculas y minúsculas cuando se ejecute la expresión regular. |
| **Resaltar** | No | **Nota:** Esto solo se aplica en elementos **Rule** dentro de elementos **ExtensionPoint**. Especifica cómo debe resaltar el cliente el texto coincidente. Puede ser uno de los siguientes: `all` o `none`. Si no se especifica, el valor predeterminado es `all`. |

### <a name="example"></a>Ejemplo

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHtml" IgnoreCase="true" />
```

## <a name="rulecollection"></a>RuleCollection

Define una colección de reglas y el operador lógico que se debe usar cuando se evalúen.

### <a name="attributes"></a>Atributos

| Atributo | Obligatorio | Descripción |
|:-----|:-----|:-----|
| **Mode** | Sí | Especifica el operador lógico que se usará al evaluar esta colección de reglas. Puede ser: `And` o `Or`. |

### <a name="example"></a>Ejemplo

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a>Vea también

- [Definir reglas para mostrar una aplicación de correo en Outlook 2013 Preview](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [Coincidencia de cadenas en un elemento de Outlook como entidades conocidas](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [Usar expresiones regulares para mostrar una aplicación de correo en Outlook 2013 Preview](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)