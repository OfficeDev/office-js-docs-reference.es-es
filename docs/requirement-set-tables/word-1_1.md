| Class | Campos | Descripción |
|:---|:---|:---|
|[Cuerpo](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear--)|Borra el contenido del objeto de cuerpo.|
||[getHtml()](/javascript/api/word/word.body#gethtml--)|Obtiene una representación HTML del objeto Body.|
||[getOoxml()](/javascript/api/word/word.body#getooxml--)|Obtiene la representación OOXML (Office Open XML) del objeto de cuerpo.|
||[ignorePunct](/javascript/api/word/word.body#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.body#ignorespace)||
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertbreak-breaktype--insertlocation-)|Inserta un salto en la ubicación especificada del documento principal.|
||[insertContentControl()](/javascript/api/word/word.body#insertcontentcontrol--)|Ajusta el objeto de cuerpo con un control de contenido de texto enriquecido.|
||[insertFileFromBase64 (base64File: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertfilefrombase64-base64file--insertlocation-)|Inserta un documento en el cuerpo en la ubicación especificada.|
||[insertHtml (HTML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#inserthtml-html--insertlocation-)|Inserta HTML en la ubicación especificada.|
||[insertOoxml (OOXML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertooxml-ooxml--insertlocation-)|Inserta OOXML en la ubicación especificada.|
||[insertParagraph (paragraphText: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertparagraph-paragraphtext--insertlocation-)|Inserta un párrafo en la ubicación especificada.|
||[insertText (Text: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#inserttext-text--insertlocation-)|Inserta texto en el cuerpo en la ubicación especificada.|
||[matchCase](/javascript/api/word/word.body#matchcase)||
||[matchPrefix](/javascript/api/word/word.body#matchprefix)||
||[matchSuffix](/javascript/api/word/word.body#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.body#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.body#matchwildcards)||
||[contentControls](/javascript/api/word/word.body#contentcontrols)|Obtiene la colección de objetos de control de contenido de texto enriquecido en el cuerpo.|
||[font](/javascript/api/word/word.body#font)|Obtiene el formato de texto del cuerpo.|
||[inlinePictures](/javascript/api/word/word.body#inlinepictures)|Obtiene la colección de objetos InlinePicture en el cuerpo.|
||[paragraphs](/javascript/api/word/word.body#paragraphs)|Obtiene la colección de objetos Paragraph en el cuerpo.|
||[parentContentControl](/javascript/api/word/word.body#parentcontentcontrol)|Obtiene el control de contenido que contiene el cuerpo.|
||[text](/javascript/api/word/word.body#text)|Obtiene el texto del cuerpo.|
||[Search (Dentrotexto: String, searchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignoreSpace?: Boolean matchCase?: Boolean matchPrefix?: Boolean matchSuffix?: Boolean?: Boolean matchWholeWord?: Boolean AC?: Boolean})](/javascript/api/word/word.body#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Realiza una búsqueda con el SearchOptions especificado en el ámbito del objeto de cuerpo.|
||[Select (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.body#select-selectionmode-)|Selecciona el cuerpo y se desplaza por la interfaz de usuario de Word hasta él.|
||[estilo](/javascript/api/word/word.body#style)|Obtiene o establece el nombre de estilo del cuerpo.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[efecto](/javascript/api/word/word.contentcontrol#appearance)|Obtiene o establece el aspecto del control de contenido.|
||[cannotDelete](/javascript/api/word/word.contentcontrol#cannotdelete)|Obtiene o establece un valor que indica si el usuario puede eliminar el control de contenido.|
||[cannotEdit](/javascript/api/word/word.contentcontrol#cannotedit)|Obtiene o establece un valor que indica si el usuario puede editar el contenido del control de contenido.|
||[clear()](/javascript/api/word/word.contentcontrol#clear--)|Borra el contenido del control de contenido.|
||[color](/javascript/api/word/word.contentcontrol#color)|Obtiene o establece el color del control de contenido.|
||[Delete (keepContent: Boolean)](/javascript/api/word/word.contentcontrol#delete-keepcontent-)|Elimina el control de contenido y su contenido.|
||[getHtml()](/javascript/api/word/word.contentcontrol#gethtml--)|Obtiene una representación HTML del objeto de control de contenido.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getooxml--)|Obtiene la representación Office Open XML (OOXML) del objeto de control de contenido.|
||[ignorePunct](/javascript/api/word/word.contentcontrol#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#ignorespace)||
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertbreak-breaktype--insertlocation-)|Inserta un salto en la ubicación especificada del documento principal.|
||[insertFileFromBase64 (base64File: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertfilefrombase64-base64file--insertlocation-)|Inserta un documento en el control de contenido en la ubicación especificada.|
||[insertHtml (HTML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserthtml-html--insertlocation-)|Inserta HTML en el control de contenido en la ubicación especificada.|
||[insertOoxml (OOXML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertooxml-ooxml--insertlocation-)|Inserta OOXML en el control de contenido en la ubicación especificada.|
||[insertParagraph (paragraphText: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertparagraph-paragraphtext--insertlocation-)|Inserta un párrafo en la ubicación especificada.|
||[insertText (Text: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserttext-text--insertlocation-)|Inserta texto en el control de contenido en la ubicación especificada.|
||[matchCase](/javascript/api/word/word.contentcontrol#matchcase)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#matchprefix)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#matchwildcards)||
||[placeholderText](/javascript/api/word/word.contentcontrol#placeholdertext)|Obtiene o establece el texto de marcador de posición del control de contenido.|
||[contentControls](/javascript/api/word/word.contentcontrol#contentcontrols)|Obtiene la colección de objetos de control de contenido que se encuentran en el control de contenido.|
||[font](/javascript/api/word/word.contentcontrol#font)|Obtiene el formato de texto del control de contenido.|
||[id](/javascript/api/word/word.contentcontrol#id)|Obtiene un entero que representa el identificador del control de contenido.|
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinepictures)|Obtiene la colección de objetos inlinePicture que se encuentran en el control de contenido.|
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|Obtiene la colección de objetos de párrafo que se encuentran en el control de contenido.|
||[parentContentControl](/javascript/api/word/word.contentcontrol#parentcontentcontrol)|Obtiene el control de contenido que contiene el control de contenido.|
||[text](/javascript/api/word/word.contentcontrol#text)|Obtiene el texto del control de contenido.|
||[type](/javascript/api/word/word.contentcontrol#type)|Obtiene el tipo de control de contenido.|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#removewhenedited)|Obtiene o establece un valor que indica si el control de contenido se elimina después de su edición.|
||[Search (Dentrotexto: String, searchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignoreSpace?: Boolean matchCase?: Boolean matchPrefix?: Boolean matchSuffix?: Boolean?: Boolean matchWholeWord?: Boolean AC?: Boolean})](/javascript/api/word/word.contentcontrol#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Realiza una búsqueda con el SearchOptions especificado en el ámbito del objeto de control de contenido.|
||[Select (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.contentcontrol#select-selectionmode-)|Selecciona el control de contenido.|
||[estilo](/javascript/api/word/word.contentcontrol#style)|Obtiene o establece el nombre de estilo para el control de contenido.|
||[Etiquetas](/javascript/api/word/word.contentcontrol#tag)|Obtiene o establece una etiqueta para identificar un control de contenido.|
||[title](/javascript/api/word/word.contentcontrol#title)|Obtiene o establece el título de un control de contenido.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyid-id-)|Obtiene un control de contenido mediante su identificador.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getbytag-tag-)|Obtiene los controles de contenido que tienen la etiqueta especificada.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getbytitle-title-)|Obtiene los controles de contenido que tienen el título especificado.|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#getitem-index-)|Obtiene un control de contenido por su índice en la colección.|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[Document](/javascript/api/word/word.document)|[getSelection ()](/javascript/api/word/word.document#getselection--)|Obtiene la selección actual del documento.|
||[body](/javascript/api/word/word.document#body)|Obtiene el objeto de cuerpo del documento.|
||[contentControls](/javascript/api/word/word.document#contentcontrols)|Obtiene la colección de objetos de control de contenido en el documento.|
||[Guard](/javascript/api/word/word.document#saved)|Indica si se guardaron los cambios en el documento.|
||[sections](/javascript/api/word/word.document#sections)|Obtiene la colección de objetos Section del documento.|
||[save()](/javascript/api/word/word.document#save--)|Guarda el documento.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|Obtiene o establece un valor que indica si la fuente está en negrita.|
||[color](/javascript/api/word/word.font#color)|Obtiene o establece el color de la fuente especificada.|
||[doubleStrikeThrough](/javascript/api/word/word.font#doublestrikethrough)|Obtiene o establece un valor que indica si la fuente tiene un tachado doble.|
||[highlightColor](/javascript/api/word/word.font#highlightcolor)|Obtiene o establece el color de resaltado.|
||[italic](/javascript/api/word/word.font#italic)|Obtiene o establece un valor que indica si la fuente está en cursiva.|
||[name](/javascript/api/word/word.font#name)|Obtiene o establece un valor que representa el nombre de la fuente.|
||[size](/javascript/api/word/word.font#size)|Obtiene o establece un valor que representa el tamaño de la fuente en puntos.|
||[Aplique](/javascript/api/word/word.font#strikethrough)|Obtiene o establece un valor que indica si la fuente tiene tachado.|
||[subscript](/javascript/api/word/word.font#subscript)|Obtiene o establece un valor que indica si la fuente está en subíndice.|
||[superscript](/javascript/api/word/word.font#superscript)|Obtiene o establece un valor que indica si la fuente está en superíndice.|
||[underline](/javascript/api/word/word.font#underline)|Obtiene o establece un valor que indica el tipo de subrayado de la fuente.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#alttextdescription)|Obtiene o establece una cadena que representa el texto alternativo asociado a la imagen incorporada.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#alttexttitle)|Obtiene o establece una cadena que contiene el título de la imagen incorporada.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getbase64imagesrc--)|Obtiene la representación de cadena codificada en Base64 de la imagen incorporada.|
||[height](/javascript/api/word/word.inlinepicture#height)|Obtiene o establece un número que describe la altura de la imagen incorporada.|
||[hyperlink](/javascript/api/word/word.inlinepicture#hyperlink)|Obtiene o establece un hipervínculo en la imagen.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertcontentcontrol--)|Ajusta la imagen incorporada con un control de contenido de texto enriquecido.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockaspectratio)|Obtiene o establece un valor que indica si la imagen incorporada mantiene sus proporciones originales cuando se cambia su tamaño.|
||[parentContentControl](/javascript/api/word/word.inlinepicture#parentcontentcontrol)|Obtiene el control de contenido que contiene la imagen incorporada.|
||[width](/javascript/api/word/word.inlinepicture#width)|Obtiene o establece un número que describe el ancho de la imagen incorporada.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[Paragraph](/javascript/api/word/word.paragraph)|[orientación](/javascript/api/word/word.paragraph#alignment)|Obtiene o establece la alineación de un párrafo.|
||[clear()](/javascript/api/word/word.paragraph#clear--)|Borra el contenido del objeto de párrafo.|
||[delete()](/javascript/api/word/word.paragraph#delete--)|Elimina el párrafo y su contenido del documento.|
||[firstLineIndent](/javascript/api/word/word.paragraph#firstlineindent)|Obtiene o establece el valor (en puntos) para una sangría en la primera línea o francesa.|
||[getHtml()](/javascript/api/word/word.paragraph#gethtml--)|Obtiene una representación HTML del objeto Paragraph.|
||[getOoxml()](/javascript/api/word/word.paragraph#getooxml--)|Obtiene la representación Office Open XML (OOXML) del objeto de párrafo.|
||[ignorePunct](/javascript/api/word/word.paragraph#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.paragraph#ignorespace)||
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertbreak-breaktype--insertlocation-)|Inserta un salto en la ubicación especificada del documento principal.|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertcontentcontrol--)|Ajusta el objeto de párrafo con un control de contenido de texto enriquecido.|
||[insertFileFromBase64 (base64File: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertfilefrombase64-base64file--insertlocation-)|Inserta un documento en el párrafo en la ubicación especificada.|
||[insertHtml (HTML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#inserthtml-html--insertlocation-)|Inserta HTML en el párrafo en la ubicación especificada.|
||[insertInlinePictureFromBase64 (base64EncodedImage: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Inserta una imagen en el párrafo en la ubicación especificada.|
||[insertOoxml (OOXML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertooxml-ooxml--insertlocation-)|Inserta OOXML en el párrafo, en la ubicación especificada.|
||[insertParagraph (paragraphText: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertparagraph-paragraphtext--insertlocation-)|Inserta un párrafo en la ubicación especificada.|
||[insertText (Text: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#inserttext-text--insertlocation-)|Inserta texto en el párrafo en la ubicación especificada.|
||[Property](/javascript/api/word/word.paragraph#leftindent)|Obtiene o establece el valor de sangría izquierda (en puntos) del párrafo.|
||[lineSpacing](/javascript/api/word/word.paragraph#linespacing)|Obtiene o establece el espaciado de línea (en puntos) del párrafo especificado.|
||[lineUnitAfter](/javascript/api/word/word.paragraph#lineunitafter)|Obtiene o establece la cantidad de espaciado, en líneas de cuadrícula, detrás del párrafo.|
||[lineUnitBefore](/javascript/api/word/word.paragraph#lineunitbefore)|Obtiene o establece la cantidad de espaciado (en líneas de cuadrícula) antes del párrafo.|
||[matchCase](/javascript/api/word/word.paragraph#matchcase)||
||[matchPrefix](/javascript/api/word/word.paragraph#matchprefix)||
||[matchSuffix](/javascript/api/word/word.paragraph#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.paragraph#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.paragraph#matchwildcards)||
||[outlineLevel](/javascript/api/word/word.paragraph#outlinelevel)|Obtiene o establece el nivel de esquema del párrafo.|
||[contentControls](/javascript/api/word/word.paragraph#contentcontrols)|Obtiene la colección de objetos de control de contenido en el párrafo.|
||[font](/javascript/api/word/word.paragraph#font)|Obtiene el formato de texto del párrafo.|
||[inlinePictures](/javascript/api/word/word.paragraph#inlinepictures)|Obtiene la colección de objetos InlinePicture del párrafo.|
||[parentContentControl](/javascript/api/word/word.paragraph#parentcontentcontrol)|Obtiene el control de contenido que contiene el párrafo.|
||[text](/javascript/api/word/word.paragraph#text)|Obtiene el texto del párrafo.|
||[Property](/javascript/api/word/word.paragraph#rightindent)|Obtiene o establece el valor de sangría derecha (en puntos) del párrafo.|
||[Search (Dentrotexto: String, searchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignoreSpace?: Boolean matchCase?: Boolean matchPrefix?: Boolean matchSuffix?: Boolean?: Boolean matchWholeWord?: Boolean AC?: Boolean})](/javascript/api/word/word.paragraph#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Realiza una búsqueda con el SearchOptions especificado en el ámbito del objeto Paragraph.|
||[Select (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.paragraph#select-selectionmode-)|Selecciona y se desplaza por la interfaz de usuario de Word hasta el párrafo.|
||[Property](/javascript/api/word/word.paragraph#spaceafter)|Obtiene o establece el espaciado (en puntos) después del párrafo.|
||[spaceBefore](/javascript/api/word/word.paragraph#spacebefore)|Obtiene o establece el espaciado (en puntos) antes del párrafo.|
||[estilo](/javascript/api/word/word.paragraph#style)|Obtiene o establece el nombre de estilo del párrafo.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear--)|Borra el contenido del objeto de intervalo.|
||[delete()](/javascript/api/word/word.range#delete--)|Elimina el intervalo y su contenido del documento.|
||[getHtml()](/javascript/api/word/word.range#gethtml--)|Obtiene una representación HTML del objeto Range.|
||[getOoxml()](/javascript/api/word/word.range#getooxml--)|Obtiene la representación OOXML del objeto de intervalo.|
||[ignorePunct](/javascript/api/word/word.range#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.range#ignorespace)||
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertbreak-breaktype--insertlocation-)|Inserta un salto en la ubicación especificada del documento principal.|
||[insertContentControl()](/javascript/api/word/word.range#insertcontentcontrol--)|Ajusta el objeto de intervalo con un control de contenido de texto enriquecido.|
||[insertFileFromBase64 (base64File: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertfilefrombase64-base64file--insertlocation-)|Inserta un documento en la ubicación especificada.|
||[insertHtml (HTML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#inserthtml-html--insertlocation-)|Inserta HTML en la ubicación especificada.|
||[insertOoxml (OOXML: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertooxml-ooxml--insertlocation-)|Inserta OOXML en la ubicación especificada.|
||[insertParagraph (paragraphText: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertparagraph-paragraphtext--insertlocation-)|Inserta un párrafo en la ubicación especificada.|
||[insertText (Text: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#inserttext-text--insertlocation-)|Inserta texto en la ubicación especificada.|
||[matchCase](/javascript/api/word/word.range#matchcase)||
||[matchPrefix](/javascript/api/word/word.range#matchprefix)||
||[matchSuffix](/javascript/api/word/word.range#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.range#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.range#matchwildcards)||
||[contentControls](/javascript/api/word/word.range#contentcontrols)|Obtiene la colección de objetos de control de contenido en el intervalo.|
||[font](/javascript/api/word/word.range#font)|Obtiene el formato de texto del intervalo.|
||[paragraphs](/javascript/api/word/word.range#paragraphs)|Obtiene la colección de objetos Paragraph del intervalo.|
||[parentContentControl](/javascript/api/word/word.range#parentcontentcontrol)|Obtiene el control de contenido que contiene el intervalo.|
||[text](/javascript/api/word/word.range#text)|Obtiene el texto del intervalo.|
||[Search (Dentrotexto: String, searchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignoreSpace?: Boolean matchCase?: Boolean matchPrefix?: Boolean matchSuffix?: Boolean?: Boolean matchWholeWord?: Boolean AC?: Boolean})](/javascript/api/word/word.range#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Realiza una búsqueda con el SearchOptions especificado en el ámbito del objeto de intervalo.|
||[Select (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.range#select-selectionmode-)|Selecciona y se desplaza por la interfaz de usuario de Word hasta el intervalo.|
||[estilo](/javascript/api/word/word.range#style)|Obtiene o establece el nombre de estilo del rango.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#ignorepunct)|Obtiene o establece un valor que indica si se van a pasar por alto todos los caracteres de puntuación entre las palabras.|
||[ignoreSpace](/javascript/api/word/word.searchoptions#ignorespace)|Obtiene o establece un valor que indica si se van a omitir todos los espacios en blanco entre palabras.|
||[matchCase](/javascript/api/word/word.searchoptions#matchcase)|Obtiene o establece un valor que indica si se va a realizar una búsqueda distinguiendo entre mayúsculas y minúsculas.|
||[matchPrefix](/javascript/api/word/word.searchoptions#matchprefix)|Obtiene o establece un valor que indica si se van a buscar palabras que empiecen por la cadena de búsqueda.|
||[matchSuffix](/javascript/api/word/word.searchoptions#matchsuffix)|Obtiene o establece un valor que indica si se van a buscar palabras que finalicen por la cadena de búsqueda.|
||[matchWholeWord](/javascript/api/word/word.searchoptions#matchwholeword)|Obtiene o establece un valor que indica si se van a buscar solamente palabras completas y no texto que forme parte de una palabra más larga.|
||[matchWildcards](/javascript/api/word/word.searchoptions#matchwildcards)|Obtiene o establece un valor que indica si la búsqueda se realizará usando operadores de búsqueda especiales.|
|[Section](/javascript/api/word/word.section)|[getFooter (tipo: Word. HeaderFooterType)](/javascript/api/word/word.section#getfooter-type-)|Obtiene uno de los pies de página de la sección.|
||[getHeader (tipo: Word. HeaderFooterType)](/javascript/api/word/word.section#getheader-type-)|Obtiene uno de los encabezados de la sección.|
||[body](/javascript/api/word/word.section#body)|Obtiene el objeto de cuerpo de la sección.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
