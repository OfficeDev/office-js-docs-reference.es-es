# <a name="word-javascript-api-requirement-sets"></a>Conjuntos de requisitos de la API de JavaScript de Word

Los conjuntos de requisitos son grupos de miembros de la API con nombre. Complementos de Office utilizan conjuntos de requisitos especificados en el manifiesto o una comprobación en tiempo de ejecución para determinar si un host de Office admite las API que un complemento necesita. Para obtener más información, vea [conjuntos de versiones de Office y los requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Complementos de Word que se ejecutará a través de varias versiones de Office, incluidos Office 2016 o posterior para Windows, Office para iPad, Office para Mac y Office Online. En la siguiente tabla se enumera los conjuntos de requisitos de Word, las aplicaciones de host de Office que admiten que los números de conjunto y la compilación o la versión de requisito para esas aplicaciones.

> [!NOTE]
> Conjuntos de requisitos que están marcados como versión Beta, use la versión especificada (o posterior) del software de Office y usar la biblioteca de la versión Beta de la CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> 
> Entradas no aparece como Beta están generalmente disponibles y se puede seguir usando la biblioteca de producción CDN:https://appsforoffice.microsoft.com/lib/1/hosted/office.js

|  Conjunto de requisitos  |   Office 365 para profesionales de Windows\*  |  Office 365 para iPad  |  Office 365 para Mac  | Office Online  | Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| WordApi 1.3 | Versión 1612 (compilación 7668.1000) o posterior| Marzo de 2017, 2.22 o posterior | Marzo de 2017, 15.32 o posterior| Marzo de 2017 ||
| WordApi 1.2  | Actualización de diciembre de 2015, versión 1601 (compilación 6568.1000) o posterior | Enero de 2016, 1.18 o posterior | Enero de 2016, 15.19 o posterior| Septiembre de 2016 | |
| WordApi 1.1  | Versión 1509 (compilación 4266.1001) o posterior| Enero de 2016, 1.18 o posterior | Enero de 2016, 15.19 o posterior| Septiembre de 2016 | |

> [!NOTE]
> El número de compilación de 2016 Office instalado a través de MSI es 16.0.4266.1001. Esta versión sólo contiene el conjunto de requisitos de WordApi 1.1.

Para obtener más información sobre las versiones, los números de compilación y Office Online Server, vea:

- [Números de versión y compilación de las versiones del canal de actualización para los clientes de Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [¿Qué versión de Office estoy usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Dónde puede encontrar el número de versión y de compilación de una aplicación de cliente de Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- 
  [Información general de Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos comunes de la API de Office

Para obtener información sobre los conjuntos de requisitos comunes de la API, consulte [Office common API requirement sets (Conjuntos de requisitos comunes de la API de Office)](office-add-in-requirement-sets.md).

## <a name="whats-new-in-word-javascript-api-13"></a>Novedades de la API de JavaScript de Word 1.3 

Las siguientes son las nuevas incorporaciones a las API de JavaScript de Word en el conjunto de requisitos 1.3. 

|Objeto| Novedades| Descripción|Conjunto de requisitos| 
|:-----|-----|:----|:----| 
|[application](/javascript/api/word/word.application)|_Método_ > createDocument(base64File: string) | Crea un nuevo documento usando un archivo .docx codificado en Base64. Solo lectura.|1.3|
|[body](/javascript/api/word/word.body)|_Relación_ > lists|Obtiene la colección de objetos de lista en el cuerpo. Solo lectura.|1.3|
|[body](/javascript/api/word/word.body)|_Relación_ > parentBody|Obtiene el cuerpo primario del cuerpo. Por ejemplo, un cuerpo primario de un cuerpo de celda de tabla podría ser un encabezado. Solo lectura.|1.3|
|[body](/javascript/api/word/word.body)|_Relación_ > parentSection|Obtiene la sección primaria del cuerpo. Solo lectura.|1.3|
|[body](/javascript/api/word/word.body)|_Relación_ > styleBuiltIn|Obtiene o establece el nombre del estilo integrado del cuerpo. Use esta propiedad para los estilos integrados que son portátiles entre configuraciones regionales. Para usar estilos personalizados o nombres de estilo localizados, consulte la propiedad "style".|1.3|
|[body](/javascript/api/word/word.body)|_Relación_ > tables|Obtiene la colección de objetos de tabla en el cuerpo. Solo lectura.|1.3|
|[body](/javascript/api/word/word.body)|_Relación_ > type|Obtiene el tipo del cuerpo. El tipo puede ser 'MainDoc', 'Section', 'Header', 'Footer' o 'TableCell'. Solo lectura.|1.3|
|[body](/javascript/api/word/word.body)|_Método_ > getRange(rangeLocation: RangeLocation)|Obtiene el cuerpo completo, o el punto de inicio o final del cuerpo, como un intervalo.|1.3|
|[body](/javascript/api/word/word.body)|_Método_ > insertTable (rowCount: número, columnCount: insertLocation número,: InsertLocation, los valores: cadena)|Inserta una tabla con el número especificado de filas y columnas. El valor insertLocation puede ser 'Start' o 'End'.|1.3|
|[breaktype](/javascript/api/word/word.breaktype)|_Relación_ > saltos|Especifica el formato de un salto: tipo de sección, página o línea. Solo lectura.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relación_ > lists|Obtiene la colección de objetos de lista en el control de contenido. Solo lectura.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relación_ > parentBody|Obtiene el cuerpo primario del control de contenido. Solo lectura.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relación_ > parentTable|Obtiene la tabla que contiene el control de contenido. Devuelve un objeto NULL si este no se incluye en una tabla. Solo lectura.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relación_ > parentTableCell|Obtiene la celda de tabla que contiene el control de contenido. Devuelve un objeto NULL si no se incluye en una celda de tabla. Solo lectura.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relación_ > styleBuiltIn|Obtiene o establece el nombre del estilo integrado del control de contenido. Use esta propiedad para los estilos integrados que son portátiles entre configuraciones regionales. Para usar estilos personalizados o nombres de estilo localizados, consulte la propiedad "style".|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relación_ > subtype|Obtiene el subtipo de control de contenido. El subtipo puede ser 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' y 'RichTextTable' para los controles de contenido de texto enriquecido. Solo lectura.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relación_ > tables|Obtiene la colección de objetos de tabla en el control de contenido. Solo lectura.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Método_ > getRange(rangeLocation: RangeLocation)|Obtiene el control de contenido completo, o el punto inicial o final del control de contenido, como un intervalo.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Método_ > getTextRanges (endingMarks: string, trimSpacing: bool)|Obtiene los intervalos de texto del control de contenido mediante los signos de puntuación u otras marcas finales.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Método_ > insertTable (rowCount: número, columnCount: insertLocation número,: InsertLocation, los valores: cadena)|Inserta una tabla con el número especificado de filas y columnas en, o junto a, un control de contenido. El valor insertLocation puede ser 'Start', 'End', 'Before' o 'After'.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Método_ > Dividir (delimitadores: cadena, multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|Divide el control de contenido en intervalos secundarios mediante delimitadores.|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_Método_ > getByTypes(types: ContentControlType)|Obtiene los controles de contenido que tienen los tipos o subtipos especificados.|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_Método_ > getFirst()|Obtiene el primer control de contenido de esta colección.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Propiedad_ > key|Obtiene la clave de la propiedad personalizada. Solo lectura. |1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Propiedad_ > value|Obtiene o establece el valor de la propiedad personalizada.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Relación_ > type|Obtiene el tipo de valor de la propiedad personalizada. Solo lectura.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Método_ > delete()|Elimina la propiedad personalizada.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Propiedad_ > items|Una colección de objetos customProperty. Solo lectura.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Método_ > deleteAll()|Elimina todas las propiedades personalizadas de la colección.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Método_ > getCount()|Obtiene el recuento de las propiedades personalizadas.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Método_ > getItem(key: string)|Obtiene un objeto de propiedad personalizada mediante su clave, que no distingue mayúsculas de minúsculas.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Método_ > establecer (clave: valor de la cadena,: objeto)|Crea o establece una propiedad personalizada.|1.3|
|[document](/javascript/api/word/word.document)|_Relación_ > properties|Obtiene las propiedades del documento actual. Solo lectura.|1.3|
|[document](/javascript/api/word/word.document)|_Método_ > open()|Abrir el documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propiedad_ > applicationName|Obtiene el nombre de la aplicación del documento. Solo lectura.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propiedad_ > author|Obtiene o establece el autor del documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propiedad_ > category|Obtiene o establece la categoría del documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propiedad_ > comments|Obtiene o establece los comentarios del documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propiedad_ > company|Obtiene o establece la compañía del documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propiedad_ > format|Obtiene o establece el formato del documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propiedad_ > keywords|Obtiene o establece las palabras clave del documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propiedad_ > lastAuthor|Obtiene o establece el último autor del documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propiedad_ > manager|Obtiene o establece el administrador del documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propiedad_ > revisionNumber|Obtiene el número de revisión del documento. Solo lectura.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propiedad_ > security|Obtiene la seguridad del documento. Solo lectura.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propiedad_ > subject|Obtiene o establece el asunto del documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propiedad_ > template|Obtiene la plantilla del documento. Solo lectura.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propiedad_ > title|Obtiene o establece el título del documento.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relación_ > creationDate|Obtiene la fecha de creación del documento. Solo lectura.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relación_ > customProperties|Obtiene la colección de propiedades personalizadas del documento de sólo lectura.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relación_ > lastPrintDate|Obtiene la última fecha de impresión del documento. Solo lectura.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relación_ > lastSaveTime|Obtiene la última ahorrar tiempo del documento. Solo lectura.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Relación_ > parentTable|Obtiene la tabla que contiene la imagen incorporada. Devuelve un objeto NULL si este no se incluye en una tabla. Solo lectura.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Relación_ > parentTableCell|Obtiene la celda de tabla que contiene la imagen incorporada. Devuelve un objeto NULL si no se incluye en una celda de tabla. Solo lectura.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_ > getNext()|Obtiene la siguiente imagen incorporada.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_ > getRange(rangeLocation: RangeLocation)|Obtiene la imagen, o el punto de inicio o final de la imagen, como un intervalo.|1.3|
|[inlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|_Método_ > getFirst()|Obtiene la primera imagen incorporada de esta colección.|1.3|
|[list](/javascript/api/word/word.list)|_Propiedad_ > id|Obtiene el identificador de la lista. Solo lectura.|1.3|
|[list](/javascript/api/word/word.list)|_Propiedad_ > levelExistences|Comprueba si cada uno de los 9 niveles existe en la lista. Un valor True indica que el nivel existe, lo que significa que existe al menos un elemento de lista en ese nivel. Solo lectura.|1.3|
|[list](/javascript/api/word/word.list)|_Relación_ > levelTypes|Obtiene todos los tipos de los 9 niveles de la lista. Cada tipo puede ser 'Bullet', 'Number' o 'Picture'. Solo lectura.|1.3|
|[list](/javascript/api/word/word.list)|_Relación_ > paragraphs|Obtiene los párrafos de la lista. Solo lectura.|1.3|
|[list](/javascript/api/word/word.list)|_Método_ > getLevelParagraphs(level: number)|Obtiene los párrafos que se producen en el nivel especificado de la lista.|1.3|
|[list](/javascript/api/word/word.list)|_Método_ > getLevelString(level: number)|Obtiene la viñeta, el número o la imagen en el nivel especificado como una cadena.|1.3|
|[list](/javascript/api/word/word.list)|_Método_ > insertParagraph (paragraphText: string, insertLocation: InsertLocation)|Inserta un párrafo en la ubicación especificada. El valor insertLocation puede ser 'Start', 'End', 'Before' o 'After'.|1.3|
|[list](/javascript/api/word/word.list)|_Método_ > setLevelAlignment (nivel: número, alineación: alineación)|Obtiene la alineación de la viñeta, el número o la imagen en el nivel especificado de la lista.|1.3|
|[list](/javascript/api/word/word.list)|_Método_ > setLevelBullet (nivel: número, listBullet: ListBullet, charCode: fontName número,: cadena)|Establece el formato de viñeta en el nivel especificado de la lista. Si la viñeta es 'Custom', se requiere charCode.|1.3|
|[list](/javascript/api/word/word.list)|_Método_ > setLevelIndents (nivel: número, textIndent: float, textIndent: float)|Establece las dos sangrías del nivel especificado de la lista.|1.3|
|[list](/javascript/api/word/word.list)|_Método_ > setLevelNumbering (nivel: número, listNumbering: ListNumbering, formatString: objeto)|Establece el formato de numeración en el nivel especificado de la lista.|1.3|
|[list](/javascript/api/word/word.list)|_Método_ > setLevelStartingNumber (nivel: número, startingNumber: número)|Establece el número de inicio en el nivel especificado de la lista. El valor predeterminado es 1.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Propiedad_ > items|Una colección de objetos de lista. Solo lectura.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Método_ > getById(id: number)|Obtiene una lista mediante su identificador.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Método_ > getFirst()|Obtiene la primera lista de esta colección.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Método_ > getItem(index: number)|Obtiene un objeto de lista por su índice en la colección.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Propiedad_ > level|Obtiene o establece el nivel del elemento de la lista.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Propiedad_ > listString|Obtiene la viñeta del elemento de lista, el número o la imagen como una cadena. Solo lectura.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Propiedad_ > siblingIndex|Obtiene el número de ordenación del elemento de lista en relación a los de su mismo nivel. Solo lectura.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Método_ > getAncestor(parentOnly: bool)|Obtiene el elemento primario de lista, o el antecesor más cercano si el primario no existe.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Método_ > getDescendants(directChildrenOnly: bool)|Obtiene todos los elementos de lista descendientes del elemento de lista.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Propiedad_ > isLastParagraph|Indica el párrafo que es el último dentro de su cuerpo primario. Solo lectura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Propiedad_ > isListItem|Comprueba si el párrafo es un elemento de lista. Solo lectura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Propiedad_ > tableNestingLevel|Obtiene el nivel de la tabla del párrafo. Devuelve 0 si el párrafo no está en una tabla. Solo lectura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relación_ > list|Obtiene la lista a la que pertenece este párrafo. Devuelve un objeto NULL si el párrafo no está en una lista. Solo lectura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relación_ > listItem|Obtiene ListItem para el párrafo. Devuelve un objeto NULL si el párrafo no forma parte de una lista. Solo lectura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relación_ > parentBody|Obtiene el cuerpo primario del párrafo. Solo lectura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relación_ > parentTable|Obtiene la tabla que contiene el párrafo. Devuelve un objeto NULL si este no se incluye en una tabla. Solo lectura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relación_ > parentTableCell|Obtiene la celda de tabla que contiene el párrafo. Devuelve un objeto NULL si no se incluye en una celda de tabla. Solo lectura.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relación_ > styleBuiltIn|Obtiene o establece el nombre del estilo integrado del párrafo. Use esta propiedad para los estilos integrados que son portátiles entre configuraciones regionales. Para usar estilos personalizados o nombres de estilo localizados, consulte la propiedad "style".|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > attachToList (listId: número, nivel: número)|Permite que el párrafo se una a una lista existente en el nivel especificado. Se produce un error si el párrafo no puede unirse a la lista o si este ya es un elemento de lista.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > detachFromList()|Mueve este párrafo fuera de la lista, si este es un elemento de lista.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > getNext()|Obtiene el párrafo siguiente.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > getPrevious()|Obtiene el párrafo anterior.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > getRange(rangeLocation: RangeLocation)|Obtiene el párrafo completo, o el punto de inicio o final del párrafo, como un intervalo.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > getTextRanges (endingMarks: string, trimSpacing: bool)|Obtiene los intervalos de texto del párrafo mediante los signos de puntuación u otras marcas finales.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > insertTable (rowCount: número, columnCount: insertLocation número,: InsertLocation, los valores: cadena)|Inserta una tabla con el número especificado de filas y columnas. El valor de insertLocation puede ser 'Before' o 'After'.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > Dividir (delimitadores: cadena, trimDelimiters: bool, trimSpacing: bool)|Divide el párrafo en intervalos secundarios mediante delimitadores.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Método_ > startNewList()|Inicia una nueva lista con este párrafo. Se produce un error si el párrafo ya es un elemento de lista.|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_Método_ > getFirst()|Obtiene el primer párrafo de esta colección.|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_Método_ > getLast()|Obtiene el último párrafo de esta colección.|1.3|
|[range](/javascript/api/word/word.range)|_Propiedad_ > hyperlink|Obtiene el primer hipervínculo del intervalo, o establece un hipervínculo en el intervalo. Todos los hipervínculos del intervalo se eliminan cuando establece un hipervínculo nuevo en el intervalo. Use un carácter de nueva línea ('\n') para separar la parte de la dirección de la parte opcional de la ubicación.|1.3|
|[range](/javascript/api/word/word.range)|_Propiedad_ > isEmpty|Comprueba si la longitud del intervalo es cero. Solo lectura.|1.3|
|[range](/javascript/api/word/word.range)|_Relación_ > lists|Obtiene la colección de objetos de lista en el intervalo. Solo lectura.|1.3|
|[range](/javascript/api/word/word.range)|_Relación_ > parentBody|Obtiene el cuerpo primario del intervalo. Solo lectura.|1.3|
|[range](/javascript/api/word/word.range)|_Relación_ > parentTable|Obtiene la tabla que contiene el intervalo. Devuelve NULL si este no se incluye en una tabla. Solo lectura.|1.3|
|[range](/javascript/api/word/word.range)|_Relación_ > parentTableCell|Obtiene la celda de tabla que contiene el intervalo. Devuelve un objeto NULL si no se incluye en una celda de tabla. Solo lectura.|1.3|
|[range](/javascript/api/word/word.range)|_Relación_ > styleBuiltIn|Obtiene o establece el nombre del estilo integrado del intervalo. Use esta propiedad para los estilos integrados que son portátiles entre configuraciones regionales. Para usar estilos personalizados o nombres de estilo localizados, consulte la propiedad "style".|1.3|
|[range](/javascript/api/word/word.range)|_Relación_ > tables|Obtiene la colección de objetos de tabla en el intervalo. Solo lectura.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > compareLocationWith(range: Range)|Compara esta ubicación del intervalo con otra ubicación de este.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > expandTo(range: Range)|Devuelve un nuevo intervalo que se extiende desde este intervalo en cualquier dirección para cubrir otro intervalo. Este intervalo no se cambia.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > getHyperlinkRanges()|Obtiene intervalos secundarios de hipervínculo dentro del intervalo.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > getNextTextRange (endingMarks: string, trimSpacing: bool)|Obtiene el siguiente intervalo de texto mediante los signos de puntuación u otras marcas finales.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > getRange(rangeLocation: RangeLocation)|Clona el intervalo u obtiene el punto inicial o final del intervalo como un intervalo nuevo.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > getTextRanges (endingMarks: string, trimSpacing: bool)|Obtiene los intervalos secundarios de texto del intervalo mediante los signos de puntuación u otras marcas finales.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > insertTable (rowCount: número, columnCount: insertLocation número,: InsertLocation, los valores: cadena)|Inserta una tabla con el número especificado de filas y columnas. El valor de insertLocation puede ser 'Before' o 'After'.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > intersectWith(range: Range)|Devuelve un intervalo nuevo como la intersección de este intervalo con otro. Este intervalo no se cambia.|1.3|
|[range](/javascript/api/word/word.range)|_Método_ > Dividir (delimitadores: cadena, multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|Divide el intervalo en intervalos secundarios mediante delimitadores.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Propiedad_ > items|Una colección de objetos de intervalo. Solo lectura.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Método_ > getFirst()|Obtiene el primer intervalo de esta colección.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Método_ > getItem(index: number)|Obtiene un objeto de intervalo por su índice en la colección.|1.3|
|[requestContext](/javascript/api/word/word.requestcontext)|_Método_ > cargar (objeto: objetos, opción: objeto)|Rellena el objeto proxy creado en la capa de JavaScript con la propiedad y las opciones especificadas en el parámetro. |1.3|
|[requestContext](/javascript/api/word/word.requestcontext)|_Método_ > sync()|Envía la cola de solicitudes a Word y devuelve un objeto Promise, que puede usarse para encadenar más acciones.|1.3|
|[section](/javascript/api/word/word.section)|_Método_ > getNext()|Obtiene la próxima sección.|1.3|
|[sectionCollection](/javascript/api/word/word.sectioncollection)|_Método_ > getFirst()|Obtiene la primera sección de esta colección.|1.3|
|[table](/javascript/api/word/word.table)|_Propiedad_ > headerRowCount|Obtiene y establece el número de filas de encabezado.|1.3|
|[table](/javascript/api/word/word.table)|_Propiedad_ > height|Obtiene el alto de la tabla en puntos. Solo lectura.|1.3|
|[table](/javascript/api/word/word.table)|_Propiedad_ > isUniform|Indica si todas las filas de la tabla son uniformes. Solo lectura.|1.3|
|[table](/javascript/api/word/word.table)|_Propiedad_ > nestingLevel|Obtiene el nivel de anidamiento de la tabla. Las tablas de nivel superior tienen el nivel 1. Solo lectura.|1.3|
|[table](/javascript/api/word/word.table)|_Propiedad_ > rowCount|Obtiene el número de filas de la tabla. Solo lectura.|1.3|
|[table](/javascript/api/word/word.table)|_Propiedad_ > shadingColor|Obtiene y establece el color de sombreado.|1.3|
|[table](/javascript/api/word/word.table)|_Propiedad_ > style|Obtiene o establece el nombre de estilo de la tabla. Use esta propiedad para los estilos personalizados y los nombres de estilo localizados. Para usar los estilos integrados portátiles entre configuraciones regionales, consulte la propiedad "styleBuiltIn".|1.3|
|[table](/javascript/api/word/word.table)|_Propiedad_ > styleBandedColumns|Obtiene y establece si la tabla tiene columnas con bandas.|1.3|
|[table](/javascript/api/word/word.table)|_Propiedad_ > styleBandedRows|Obtiene y establece si la tabla tiene filas con bandas.|1.3|
|[table](/javascript/api/word/word.table)|_Propiedad_ > styleFirstColumn|Obtiene y establece si la tabla tiene una primera columna con un estilo especial.|1.3|
|[table](/javascript/api/word/word.table)|_Propiedad_ > styleLastColumn|Obtiene y establece si la tabla tiene una última columna con un estilo especial.|1.3|
|[table](/javascript/api/word/word.table)|_Propiedad_ > styleTotalRow|Obtiene y establece si la tabla tiene una (última) fila total con un estilo especial.|1.3|
|[table](/javascript/api/word/word.table)|_Propiedad_ > values|Obtiene y establece los valores de texto de la tabla, como una matriz 2D de Javascript.|1.3|
|[table](/javascript/api/word/word.table)|_Propiedad_ > width|Obtiene y establece el ancho de la tabla en puntos.|1.3|
|[table](/javascript/api/word/word.table)|_Relación_ > font|Obtiene la fuente. Úselo para obtener y establecer el nombre de la fuente, el tamaño, el color y otras propiedades. Solo lectura.|1.3|
|[table](/javascript/api/word/word.table)|_Relación_ > horizontalAlignment|Obtiene y establece la alineación horizontal de cada celda de la tabla. El valor puede ser 'left', 'centered', 'right' o 'justified'.|1.3|
|[table](/javascript/api/word/word.table)|_Relación_ > paragraphAfter|Obtiene el párrafo después de la tabla. Solo lectura.|1.3|
|[table](/javascript/api/word/word.table)|_Relación_ > paragraphBefore|Obtiene el párrafo antes de la tabla. Solo lectura.|1.3|
|[table](/javascript/api/word/word.table)|_Relación_ > parentBody|Obtiene el cuerpo primario de la tabla. Solo lectura.|1.3|
|[table](/javascript/api/word/word.table)|_Relación_ > parentContentControl|Obtiene el control de contenido que contiene la tabla. Solo lectura.|1.3|
|[table](/javascript/api/word/word.table)|_Relación_ > parentTable|Obtiene la tabla que contiene esta tabla. Devuelve un objeto NULL si este no se incluye en una tabla. Solo lectura.|1.3|
|[table](/javascript/api/word/word.table)|_Relación_ > parentTableCell|Obtiene la celda de tabla que contiene esta tabla. Devuelve un objeto NULL si no se incluye en una celda de tabla. Solo lectura.|1.3|
|[table](/javascript/api/word/word.table)|_Relación_ > rows|Obtiene todas las filas de la tabla. Solo lectura.|1.3|
|[table](/javascript/api/word/word.table)|_Relación_ > styleBuiltIn|Obtiene o establece el nombre del estilo integrado de la tabla. Use esta propiedad para los estilos integrados que son portátiles entre configuraciones regionales. Para usar estilos personalizados o nombres de estilo localizados, consulte la propiedad "style".|1.3|
|[table](/javascript/api/word/word.table)|_Relación_ > tables|Obtiene las tablas secundarias anidadas un nivel más profundo. Solo lectura.|1.3|
|[table](/javascript/api/word/word.table)|_Relación_ > verticalAlignment|Obtiene y establece la alineación vertical de cada celda de la tabla. El valor puede ser 'top', 'center' o 'bottom'.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > addColumns (insertLocation: InsertLocation, columnCount: valores numéricos: cadena)|Agrega columnas al inicio o al final de la tabla, mediante la primera o la última columna existente como una plantilla. Esto se aplica a las tablas uniformes. Los valores de cadena, si se especifican, se establecen en las filas recién insertadas.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > addRows (insertLocation: InsertLocation, rowCount: valores numéricos: cadena)|Agrega filas al inicio o al final de la tabla, mediante la primera o la última fila existente como una plantilla. Los valores de cadena, si se especifican, se establecen en las filas recién insertadas.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > autoFitContents()|Ajusta automáticamente las columnas de la tabla al ancho de su contenido.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > autoFitWindow()|Ajusta automáticamente las columnas de la tabla al ancho de la ventana.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > clear()|Borra el contenido de la tabla.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > delete()|Elimina toda la tabla.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > deleteColumns (columnIndex: número, columnCount: número)|Elimina las columnas especificadas. Esto se aplica a las tablas uniformes.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > deleteRows (rowIndex: número, rowCount: número)|Elimina las filas especificadas.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > distributeColumns()|Distribuye los anchos de columna de manera uniforme.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > distributeRows()|Distribuye los altos de fila de manera uniforme.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > getBorder(borderLocation: BorderLocation)|Obtiene el estilo del borde para el borde especificado.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > getCell (rowIndex: número, cellIndex: número)|Obtiene la celda de tabla de una fila y columna especificadas.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|Obtiene el espaciado entre borde y texto en puntos.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > getNext()|Obtiene la próxima tabla.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > getRange(rangeLocation: RangeLocation)|Obtiene el intervalo que contiene esta tabla, o el intervalo al inicio o al final de la tabla.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > insertContentControl()|Inserta un control de contenido en la tabla.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > insertParagraph (paragraphText: string, insertLocation: InsertLocation)|Inserta un párrafo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > insertTable (rowCount: número, columnCount: insertLocation número,: InsertLocation, los valores: cadena)|Inserta una tabla con el número especificado de filas y columnas. El valor de insertLocation puede ser 'Before' o 'After'.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > búsqueda (searchText: cadena, searchOptions: ParamTypeStrings.SearchOptions)|Realiza una búsqueda con el valor searchOptions especificado en el ámbito del objeto de tabla. Los resultados de la búsqueda son una colección de objetos de intervalo.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > select(selectionMode: SelectionMode)|Selecciona la tabla, o la posición al inicio o al final de la tabla, y dirige ahí la interfaz de usuario de Word.|1.3|
|[table](/javascript/api/word/word.table)|_Método_ > setCellPadding (cellPaddingLocation: CellPaddingLocation, cellPadding: float)|Establece el espaciado entre borde y texto en puntos.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Propiedad_ > color|Obtiene o establece el color del borde la tabla, como un valor hexadecimal o un nombre.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Propiedad_ > width|Obtiene o establece el ancho, en puntos, del borde de la tabla. No se aplica a los tipos de borde de la tabla que tengan anchos fijos.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Relación_ > type|Obtiene o establece el tipo de borde de la tabla.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propiedad_ > cellIndex|Obtiene el índice de la celda de la fila. Solo lectura.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propiedad_ > columnWidth|Obtiene y establece el ancho de la columna de la celda en puntos. Esto se aplica a las tablas uniformes.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propiedad_ > rowIndex|Obtiene el índice de fila de la celda en la tabla. Solo lectura.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propiedad_ > shadingColor|Obtiene o establece el color de sombreado de la celda. El color se especifica en el formato "#RRGGBB" o mediante el nombre del color.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propiedad_ > value|Obtiene y establece el texto de la celda.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propiedad_ > width|Obtiene el ancho de la celda en puntos. Solo lectura.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relación_ > body|Obtiene el objeto de cuerpo de la celda. Solo lectura.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relación_ > horizontalAlignment|Obtiene y establece la alineación horizontal de la celda. El valor puede ser 'left', 'centered', 'right' o 'justified'.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relación_ > parentRow|Obtiene la fila primaria de la celda. Solo lectura.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relación_ > parentTable|Obtiene la tabla primaria de la celda. Solo lectura.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relación_ > verticalAlignment|Obtiene y establece la alineación vertical de la celda. El valor puede ser 'top', 'center' o 'bottom'.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_ > deleteColumn()|Elimina la columna que contiene esta celda. Esto se aplica a las tablas uniformes.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_ > deleteRow()|Elimina la fila que contiene esta celda.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_ > getBorder(borderLocation: BorderLocation)|Obtiene el estilo del borde para el borde especificado.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|Obtiene el espaciado entre borde y texto en puntos.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_ > getNext()|Obtiene la próxima celda.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_ > insertColumns (insertLocation: InsertLocation, columnCount: valores numéricos: cadena)|Agrega columnas a la izquierda o derecha de la celda, mediante la columna de la celda como una plantilla. Esto se aplica a las tablas uniformes. Los valores de cadena, si se especifican, se establecen en las filas recién insertadas.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_ > insertRows (insertLocation: InsertLocation, rowCount: valores numéricos: cadena)|Inserta filas por encima o debajo de la celda, mediante la fila de la celda como una plantilla. Los valores de cadena, si se especifican, se establecen en las filas recién insertadas.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Método_ > setCellPadding (cellPaddingLocation: CellPaddingLocation, cellPadding: float)|Establece el espaciado entre borde y texto en puntos.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Propiedad_ > items|Una colección de objetos tableCell. Solo lectura.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Método_ > getFirst()|Obtiene la primera celda de tabla de esta colección.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Método_ > getItem(index: number)|Obtiene un objeto de celda de tabla por su índice en la colección.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Propiedad_ > items|Una colección de objetos de tabla. Solo lectura.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Método_ > getFirst()|Obtiene la primera tabla de esta colección.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Método_ > getItem(index: number)|Obtiene un objeto de tabla por su índice en la colección.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propiedad_ > cellCount|Obtiene el número de celdas de la fila. Solo lectura.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propiedad_ > isHeader|Comprueba si la fila es una fila de encabezado. Solo lectura. Para establecer el número de filas de encabezado, use HeaderRowCount en el objeto de tabla. Solo lectura.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propiedad_ > preferredHeight|Obtiene y establece el alto preferido de la fila en puntos.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propiedad_ > rowIndex|Obtiene el índice de la fila en la tabla primaria. Solo lectura.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propiedad_ > shadingColor|Obtiene y establece el color de sombreado.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propiedad_ > values|Obtiene y establece los valores de texto de la fila, como una matriz 1D de Javascript.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relación_ > cells|Obtiene celdas. Solo lectura.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relación_ > font|Obtiene la fuente. Úselo para obtener y establecer el nombre de la fuente, el tamaño, el color y otras propiedades. Solo lectura.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relación_ > horizontalAlignment|Obtiene y establece la alineación horizontal de cada celda de la fila. El valor puede ser 'left', 'centered', 'right' o 'justified'.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relación_ > parentTable|Obtiene la tabla primaria. Solo lectura.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relación_ > verticalAlignment|Obtiene y establece la alineación vertical de las celdas de la fila. El valor puede ser 'top', 'center' o 'bottom'.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > clear()|Borra el contenido de la fila.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > delete()|Elimina toda la fila.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > getBorder(borderLocation: BorderLocation)|Obtiene el estilo del borde de las celdas de la fila.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|Obtiene el espaciado entre borde y texto en puntos.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > getNext()|Obtiene la próxima fila.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > insertRows (insertLocation: InsertLocation, rowCount: valores numéricos: cadena)|Inserta filas mediante esta fila como una plantilla. Si los valores se especifican, inserta los valores en las nuevas filas.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > búsqueda (searchText: cadena, searchOptions: ParamTypeStrings.SearchOptions)|Realiza una búsqueda con el valor searchOptions especificado en el ámbito de la fila. Los resultados de la búsqueda son una colección de objetos de intervalo.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > select(selectionMode: SelectionMode)|Selecciona la fila y se desplaza por la interfaz de usuario de Word hasta ella.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Método_ > setCellPadding (cellPaddingLocation: CellPaddingLocation, cellPadding: float)|Establece el espaciado entre borde y texto en puntos.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Propiedad_ > items|Una colección de objetos tableRow. Solo lectura.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Método_ > getFirst()|Obtiene la primera fila de esta colección.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Método_ > getItem(index: number)|Obtiene un objeto de fila de tabla por su índice en la colección.|1.3|


## <a name="whats-new-in-word-javascript-api-12"></a>Novedades de la API de JavaScript de Word 1.2

Las siguientes son las nuevas incorporaciones a las API de JavaScript de Word en el conjunto de requisitos 1.2. 

|Objeto| Novedades| Descripción|Conjunto de requisitos|
|:-----|-----|:----|:----|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Método_ > insertInlinePictureFromBase64 (base64EncodedImage: string, insertLocation: InsertLocation)|Inserta una imagen incorporada en el control de contenido en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start' o 'End'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Relación_ > paragraph|Obtiene el párrafo primario que contiene la imagen incorporada. Solo lectura.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_ > delete()|Elimina la imagen incorporada del documento.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_ > insertBreak (breakType: BreakType, insertLocation: InsertLocation)|Inserta un salto en la ubicación especificada del documento principal. El valor de insertLocation puede ser 'Before' o 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_ > insertFileFromBase64 (base64File: string, insertLocation: InsertLocation)|Inserta un documento en la ubicación especificada. El valor de insertLocation puede ser 'Before' o 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_ > insertHtml (html: string, insertLocation: InsertLocation)|Inserta HTML en la ubicación especificada. El valor de insertLocation puede ser 'Before' o 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_ > insertInlinePictureFromBase64 (base64EncodedImage: string, insertLocation: InsertLocation|Inserta una imagen incorporada en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Before' o 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_ > insertOoxml (ooxml: string, insertLocation: InsertLocation)|Inserta OOXML en la ubicación especificada.  El valor de insertLocation puede ser 'Before' o 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_ > insertParagraph (paragraphText: string, insertLocation: InsertLocation)|Inserta un párrafo en la ubicación especificada. El valor insertLocation puede ser 'Before' o 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_ > insertText (texto: string, insertLocation: InsertLocation)|Inserta texto en la ubicación especificada. El valor de insertLocation puede ser 'Before' o 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Método_ > select(selectionMode: SelectionMode)|Selecciona la imagen incorporada. Esto hace que Word se desplace hasta la selección.|1.2|
|[range](/javascript/api/word/word.range)|_Relación_ > inlinePictures|Obtiene la colección de objetos de imagen incorporada del intervalo. Solo lectura.|1.2|
|[range](/javascript/api/word/word.range)|_Método_ > insertInlinePictureFromBase64 (base64EncodedImage: string, insertLocation: InsertLocation)|Inserta una imagen en la ubicación especificada. El valor insertLocation puede ser 'Replace', 'Start', 'End', 'Before' o 'After'.|1.2|

## <a name="word-javascript-api-11"></a>API de JavaScript de Word 1.1

API de JavaScript de Word 1.1 es la primera versión de la API. Para obtener más información sobre la API, consulte los temas de referencia de [API de JavaScript de Word](/javascript/api/word). 

## <a name="see-also"></a>Vea también

- [Versiones de Office y conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar los hosts de Office y los requisitos de la API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifiesto XML de complementos para Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
