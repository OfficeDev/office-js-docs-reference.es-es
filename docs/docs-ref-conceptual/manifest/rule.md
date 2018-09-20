# <a name="rule-element"></a><span data-ttu-id="ae38b-101">Elemento Rule</span><span class="sxs-lookup"><span data-stu-id="ae38b-101">Rule element</span></span>

<span data-ttu-id="ae38b-102">Especifica las reglas de activación que deberían evaluarse para este complemento de correo contextual.</span><span class="sxs-lookup"><span data-stu-id="ae38b-102">Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="ae38b-103">**Tipo de complemento:** Complemento contextual de correo</span><span class="sxs-lookup"><span data-stu-id="ae38b-103">**Add-in type:** Mail contextual add-in</span></span>

## <a name="contained-in"></a><span data-ttu-id="ae38b-104">Contenidos en</span><span class="sxs-lookup"><span data-stu-id="ae38b-104">Contained in</span></span>

- [<span data-ttu-id="ae38b-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="ae38b-105">OfficeApp</span></span>](officeapp.md)
- [<span data-ttu-id="ae38b-106">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="ae38b-106">ExtensionPoint</span></span>](extensionpoint.md)

## <a name="attributes"></a><span data-ttu-id="ae38b-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="ae38b-107">Attributes</span></span>

| <span data-ttu-id="ae38b-108">Atributo</span><span class="sxs-lookup"><span data-stu-id="ae38b-108">Attribute</span></span> | <span data-ttu-id="ae38b-109">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="ae38b-109">Required</span></span> | <span data-ttu-id="ae38b-110">Descripción</span><span class="sxs-lookup"><span data-stu-id="ae38b-110">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="ae38b-111">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="ae38b-111">**xsi:type**</span></span> | <span data-ttu-id="ae38b-112">Sí</span><span class="sxs-lookup"><span data-stu-id="ae38b-112">Yes</span></span> | <span data-ttu-id="ae38b-113">El tipo de regla que se está definiendo.</span><span class="sxs-lookup"><span data-stu-id="ae38b-113">The type of rule being defined.</span></span> |

<span data-ttu-id="ae38b-114">El tipo de regla puede ser uno de los siguientes.</span><span class="sxs-lookup"><span data-stu-id="ae38b-114">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="ae38b-115">ItemIs</span><span class="sxs-lookup"><span data-stu-id="ae38b-115">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="ae38b-116">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="ae38b-116">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="ae38b-117">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="ae38b-117">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="ae38b-118">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="ae38b-118">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="ae38b-119">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="ae38b-119">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="ae38b-120">Regla ItemIs</span><span class="sxs-lookup"><span data-stu-id="ae38b-120">ItemIs rule</span></span>

<span data-ttu-id="ae38b-121">Define una regla que evalúa en verdadero si el elemento seleccionado es del tipo especificado.</span><span class="sxs-lookup"><span data-stu-id="ae38b-121">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="ae38b-122">Atributos</span><span class="sxs-lookup"><span data-stu-id="ae38b-122">Attributes</span></span>

| <span data-ttu-id="ae38b-123">Atributo</span><span class="sxs-lookup"><span data-stu-id="ae38b-123">Attribute</span></span> | <span data-ttu-id="ae38b-124">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="ae38b-124">Required</span></span> | <span data-ttu-id="ae38b-125">Descripción</span><span class="sxs-lookup"><span data-stu-id="ae38b-125">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="ae38b-126">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="ae38b-126">**ItemType**</span></span> | <span data-ttu-id="ae38b-127">Sí</span><span class="sxs-lookup"><span data-stu-id="ae38b-127">Yes</span></span> | <span data-ttu-id="ae38b-p101">Especifica el tipo de elemento con el que debe coincidir. Puede ser `Message` o `Appointment`. El tipo de elemento `Message` incluye correo electrónico, convocatorias de reunión, respuestas a la reunión y cancelaciones de reunión.</span><span class="sxs-lookup"><span data-stu-id="ae38b-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="ae38b-131">**FormType**</span><span class="sxs-lookup"><span data-stu-id="ae38b-131">**FormType**</span></span> | <span data-ttu-id="ae38b-132">No (dentro de [ExtensionPoint](extensionpoint.md)), Sí (dentro de [OfficeApp](officeapp.md))</span><span class="sxs-lookup"><span data-stu-id="ae38b-132">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="ae38b-p102">Especifica si la aplicación debe aparecer en el formulario de edición o lectura para el elemento. Puede ser uno de los siguientes: `Read`, `Edit` o `ReadOrEdit`. Si se especifica en una `Rule` dentro de un `ExtensionPoint`, este valor DEBE ser `Read`.</span><span class="sxs-lookup"><span data-stu-id="ae38b-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="ae38b-136">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="ae38b-136">**ItemClass**</span></span> | <span data-ttu-id="ae38b-137">No</span><span class="sxs-lookup"><span data-stu-id="ae38b-137">No</span></span> | <span data-ttu-id="ae38b-p103">Especifica la clase de mensaje personalizada con la que debe coincidir. Para obtener más información, vea [Activar un complemento de correo de Outlook para una clase de mensaje específica](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span><span class="sxs-lookup"><span data-stu-id="ae38b-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span></span> |
| <span data-ttu-id="ae38b-140">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="ae38b-140">**IncludeSubClasses**</span></span> | <span data-ttu-id="ae38b-141">No</span><span class="sxs-lookup"><span data-stu-id="ae38b-141">No</span></span> | <span data-ttu-id="ae38b-142">Especifica si la regla debería evaluar en verdadero si el elemento es de una subclase de la clase del mensaje especificada; el valor predeterminado es `false`.</span><span class="sxs-lookup"><span data-stu-id="ae38b-142">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="ae38b-143">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ae38b-143">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="ae38b-144">Regla ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="ae38b-144">ItemHasAttachment rule</span></span>

<span data-ttu-id="ae38b-145">Define una regla que evalúa en verdadero si el elemento contiene datos adjuntos.</span><span class="sxs-lookup"><span data-stu-id="ae38b-145">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="ae38b-146">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ae38b-146">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="ae38b-147">Regla ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="ae38b-147">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="ae38b-148">Define una regla que evalúa en verdadero si el elemento contiene texto del tipo de entidad especificado en el asunto o en el cuerpo.</span><span class="sxs-lookup"><span data-stu-id="ae38b-148">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="ae38b-149">Atributos</span><span class="sxs-lookup"><span data-stu-id="ae38b-149">Attributes</span></span>

| <span data-ttu-id="ae38b-150">Atributo</span><span class="sxs-lookup"><span data-stu-id="ae38b-150">Attribute</span></span> | <span data-ttu-id="ae38b-151">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="ae38b-151">Required</span></span> | <span data-ttu-id="ae38b-152">Descripción</span><span class="sxs-lookup"><span data-stu-id="ae38b-152">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="ae38b-153">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="ae38b-153">**EntityType**</span></span> | <span data-ttu-id="ae38b-154">Sí</span><span class="sxs-lookup"><span data-stu-id="ae38b-154">Yes</span></span> | <span data-ttu-id="ae38b-p104">Especifica el tipo de entidad que se tiene que encontrar para que la regla evalúe en verdadero. Puede ser uno de los siguientes: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` o `Contact`.</span><span class="sxs-lookup"><span data-stu-id="ae38b-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="ae38b-157">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="ae38b-157">**RegExFilter**</span></span> | <span data-ttu-id="ae38b-158">No</span><span class="sxs-lookup"><span data-stu-id="ae38b-158">No</span></span> | <span data-ttu-id="ae38b-159">Especifica una expresión regular que se debe ejecutar con esta entidad para su activación.</span><span class="sxs-lookup"><span data-stu-id="ae38b-159">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="ae38b-160">**NombreDeFiltro**</span><span class="sxs-lookup"><span data-stu-id="ae38b-160">**FilterName**</span></span> | <span data-ttu-id="ae38b-161">No</span><span class="sxs-lookup"><span data-stu-id="ae38b-161">No</span></span> | <span data-ttu-id="ae38b-162">Especifica el nombre del filtro de expresión regular, de modo que después sea posible hacerle referencia en el código de su complemento.</span><span class="sxs-lookup"><span data-stu-id="ae38b-162">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="ae38b-163">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="ae38b-163">**IgnoreCase**</span></span> | <span data-ttu-id="ae38b-164">No</span><span class="sxs-lookup"><span data-stu-id="ae38b-164">No</span></span> | <span data-ttu-id="ae38b-165">Especifica que se ignoren las mayúsculas y minúsculas cuando se ejecute la expresión regular especificada por el atributo **RegExFilter**.</span><span class="sxs-lookup"><span data-stu-id="ae38b-165">Specifies to ignore case when running the regular expression specified by the  **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="ae38b-166">**Resaltar**</span><span class="sxs-lookup"><span data-stu-id="ae38b-166">**Highlight**</span></span> | <span data-ttu-id="ae38b-167">No</span><span class="sxs-lookup"><span data-stu-id="ae38b-167">No</span></span> | <span data-ttu-id="ae38b-p105">**Nota:** Esto solo se aplica en elementos **Rule** dentro de elementos **ExtensionPoint**. Especifica cómo debe resaltar el cliente las entidades coincidentes. Puede ser uno de los siguientes: `all` o `none`. Si no se especifica, el valor predeterminado es `all`.</span><span class="sxs-lookup"><span data-stu-id="ae38b-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="ae38b-172">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ae38b-172">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="ae38b-173">Regla ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="ae38b-173">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="ae38b-174">Define una regla que evalúa en verdadero si se encuentra una coincidencia para la expresión regular especificada en la propiedad indicada del elemento.</span><span class="sxs-lookup"><span data-stu-id="ae38b-174">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="ae38b-175">Atributos</span><span class="sxs-lookup"><span data-stu-id="ae38b-175">Attributes</span></span>

| <span data-ttu-id="ae38b-176">Atributo</span><span class="sxs-lookup"><span data-stu-id="ae38b-176">Attribute</span></span> | <span data-ttu-id="ae38b-177">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="ae38b-177">Required</span></span> | <span data-ttu-id="ae38b-178">Descripción</span><span class="sxs-lookup"><span data-stu-id="ae38b-178">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="ae38b-179">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="ae38b-179">**RegExName**</span></span> | <span data-ttu-id="ae38b-180">Sí</span><span class="sxs-lookup"><span data-stu-id="ae38b-180">Yes</span></span> | <span data-ttu-id="ae38b-181">Especifica el nombre de una expresión regular para que pueda hacer referencia a dicha expresión en el código de su complemento.</span><span class="sxs-lookup"><span data-stu-id="ae38b-181">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="ae38b-182">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="ae38b-182">**RegExValue**</span></span> | <span data-ttu-id="ae38b-183">Sí</span><span class="sxs-lookup"><span data-stu-id="ae38b-183">Yes</span></span> | <span data-ttu-id="ae38b-184">Especifica la expresión regular que se evaluará para determinar si se debe mostrar el complemento de correo.</span><span class="sxs-lookup"><span data-stu-id="ae38b-184">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="ae38b-185">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="ae38b-185">**PropertyName**</span></span> | <span data-ttu-id="ae38b-186">Sí</span><span class="sxs-lookup"><span data-stu-id="ae38b-186">Yes</span></span> | <span data-ttu-id="ae38b-p106">Especifica el nombre de la propiedad contra la que se evaluará la expresión regular. Puede ser uno de los siguientes: `Subject`, `BodyAsPlaintext`, `BodyAsHtml` o `SenderSTMPAddress`.</span><span class="sxs-lookup"><span data-stu-id="ae38b-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHtml`, or `SenderSTMPAddress`.</span></span> |
| <span data-ttu-id="ae38b-189">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="ae38b-189">**IgnoreCase**</span></span> | <span data-ttu-id="ae38b-190">No</span><span class="sxs-lookup"><span data-stu-id="ae38b-190">No</span></span> | <span data-ttu-id="ae38b-191">Especifica que se ignoren las mayúsculas y minúsculas cuando se ejecute la expresión regular.</span><span class="sxs-lookup"><span data-stu-id="ae38b-191">Specifies to ignore the case when executing the regular expression.</span></span> |
| <span data-ttu-id="ae38b-192">**Resaltar**</span><span class="sxs-lookup"><span data-stu-id="ae38b-192">**Highlight**</span></span> | <span data-ttu-id="ae38b-193">No</span><span class="sxs-lookup"><span data-stu-id="ae38b-193">No</span></span> | <span data-ttu-id="ae38b-p107">**Nota:** Esto solo se aplica en elementos **Rule** dentro de elementos **ExtensionPoint**. Especifica cómo debe resaltar el cliente el texto coincidente. Puede ser uno de los siguientes: `all` o `none`. Si no se especifica, el valor predeterminado es `all`.</span><span class="sxs-lookup"><span data-stu-id="ae38b-p107">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching text. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="ae38b-198">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ae38b-198">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHtml" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="ae38b-199">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="ae38b-199">RuleCollection</span></span>

<span data-ttu-id="ae38b-200">Define una colección de reglas y el operador lógico que se debe usar cuando se evalúen.</span><span class="sxs-lookup"><span data-stu-id="ae38b-200">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="ae38b-201">Atributos</span><span class="sxs-lookup"><span data-stu-id="ae38b-201">Attributes</span></span>

| <span data-ttu-id="ae38b-202">Atributo</span><span class="sxs-lookup"><span data-stu-id="ae38b-202">Attribute</span></span> | <span data-ttu-id="ae38b-203">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="ae38b-203">Required</span></span> | <span data-ttu-id="ae38b-204">Descripción</span><span class="sxs-lookup"><span data-stu-id="ae38b-204">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="ae38b-205">**Mode**</span><span class="sxs-lookup"><span data-stu-id="ae38b-205">**Mode**</span></span> | <span data-ttu-id="ae38b-206">Sí</span><span class="sxs-lookup"><span data-stu-id="ae38b-206">Yes</span></span> | <span data-ttu-id="ae38b-p108">Especifica el operador lógico que se usará al evaluar esta colección de reglas. Puede ser: `And` o `Or`.</span><span class="sxs-lookup"><span data-stu-id="ae38b-p108">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="ae38b-209">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="ae38b-209">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="ae38b-210">Vea también</span><span class="sxs-lookup"><span data-stu-id="ae38b-210">See also</span></span>

- [<span data-ttu-id="ae38b-211">Definir reglas para mostrar una aplicación de correo en Outlook 2013 Preview</span><span class="sxs-lookup"><span data-stu-id="ae38b-211">Activation rules for Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [<span data-ttu-id="ae38b-212">Coincidencia de cadenas en un elemento de Outlook como entidades conocidas</span><span class="sxs-lookup"><span data-stu-id="ae38b-212">Match strings in an Outlook item as well-known entities</span></span>](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [<span data-ttu-id="ae38b-213">Usar expresiones regulares para mostrar una aplicación de correo en Outlook 2013 Preview</span><span class="sxs-lookup"><span data-stu-id="ae38b-213">Use regular expression activation rules to show an Outlook add-in</span></span>](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)