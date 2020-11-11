| Class | Campos | Descripción |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument (base64File?: String)](/javascript/api/word/word.application#createdocument-base64file-)|Crea un nuevo documento con un archivo. docx codificado y Base64 opcional.|
|[Cuerpo](/javascript/api/word/word.body)|[getRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.body#getrange-rangelocation-)|Obtiene el cuerpo completo, o el punto de inicio o fin del cuerpo, como un intervalo.|
||[insertTable (rowCount: Number, columnCount: Number, insertLocation: Word. InsertLocation, Values?: String [] [])](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|Inserta una tabla con el número especificado de filas y columnas.|
||[lists](/javascript/api/word/word.body#lists)|Obtiene la colección de objetos de lista en el cuerpo.|
||[parentBody](/javascript/api/word/word.body#parentbody)|Obtiene el cuerpo primario del cuerpo.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|Obtiene el cuerpo primario del cuerpo.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|Obtiene el control de contenido que contiene el cuerpo.|
||[parentSection](/javascript/api/word/word.body#parentsection)|Obtiene la sección primaria del cuerpo.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|Obtiene la sección primaria del cuerpo.|
||[Table](/javascript/api/word/word.body#tables)|Obtiene la colección de objetos de tabla en el cuerpo.|
||[type](/javascript/api/word/word.body#type)|Obtiene el tipo del cuerpo.|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|Obtiene o establece el nombre del estilo integrado del cuerpo.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|Obtiene el control de contenido completo, o el punto inicial o final del control de contenido, como un intervalo.|
||[getTextRanges (endingMarks: String [], trimSpacing?: Boolean)](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|Obtiene los intervalos de texto del control de contenido mediante el uso de signos de puntuación u otras marcas de fin.|
||[insertTable (rowCount: Number, columnCount: Number, insertLocation: Word. InsertLocation, Values?: String [] [])](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|Inserta una tabla con el número especificado de filas y columnas en, o junto a, un control de contenido.|
||[lists](/javascript/api/word/word.contentcontrol#lists)|Obtiene la colección de objetos de lista en el control de contenido.|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|Obtiene el cuerpo primario del control de contenido.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|Obtiene el control de contenido que contiene el control de contenido.|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|Obtiene la tabla que contiene el control de contenido.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|Obtiene la celda de tabla que contiene el control de contenido.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|Obtiene la celda de tabla que contiene el control de contenido.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parenttableornullobject)|Obtiene la tabla que contiene el control de contenido.|
||[subtipo](/javascript/api/word/word.contentcontrol#subtype)|Obtiene el subtipo de control de contenido.|
||[Table](/javascript/api/word/word.contentcontrol#tables)|Obtiene la colección de objetos de tabla en el control de contenido.|
||[Split (Delimiters: String [], Paragraphs?: Boolean, trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Divide el control de contenido en intervalos secundarios mediante delimitadores.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|Obtiene o establece el nombre del estilo integrado del control de contenido.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject (ID: Number)](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|Obtiene un control de contenido mediante su identificador.|
||[getByTypes (Types: Word. ContentControlType [])](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|Obtiene los controles de contenido que tienen los tipos y subtipos especificados.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|Obtiene el primer control de contenido de esta colección.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|Obtiene el primer control de contenido de esta colección.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|Elimina la propiedad personalizada.|
||[key](/javascript/api/word/word.customproperty#key)|Obtiene la clave de la propiedad personalizada.|
||[type](/javascript/api/word/word.customproperty#type)|Obtiene el tipo de valor de la propiedad personalizada.|
||[value](/javascript/api/word/word.customproperty#value)|Obtiene o establece el valor de la propiedad personalizada.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[Add (Key: String, Value: any)](/javascript/api/word/word.custompropertycollection#add-key--value-)|Crea una nueva propiedad personalizada o establece una existente.|
||[deleteAll ()](/javascript/api/word/word.custompropertycollection#deleteall--)|Elimina todas las propiedades personalizadas de la colección.|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|Obtiene el recuento de las propiedades personalizadas.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|Obtiene un objeto de propiedad personalizada mediante su clave, que no distingue mayúsculas de minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|Obtiene un objeto de propiedad personalizada mediante su clave, que no distingue mayúsculas de minúsculas.|
||[items](/javascript/api/word/word.custompropertycollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|Obtiene las propiedades del documento.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[Open ()](/javascript/api/word/word.documentcreated#open--)|Abre el documento.|
||[body](/javascript/api/word/word.documentcreated#body)|Obtiene el objeto de cuerpo del documento.|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|Obtiene la colección de objetos de control de contenido en el documento.|
||[properties](/javascript/api/word/word.documentcreated#properties)|Obtiene las propiedades del documento.|
||[Guard](/javascript/api/word/word.documentcreated#saved)|Indica si se guardaron los cambios en el documento.|
||[sections](/javascript/api/word/word.documentcreated#sections)|Obtiene la colección de objetos Section del documento.|
||[save()](/javascript/api/word/word.documentcreated#save--)|Guarda el documento.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[crea](/javascript/api/word/word.documentproperties#author)|Obtiene o establece el autor del documento.|
||[Categoría](/javascript/api/word/word.documentproperties#category)|Obtiene o establece la categoría del documento.|
||[comments](/javascript/api/word/word.documentproperties#comments)|Obtiene o establece los comentarios del documento.|
||[company](/javascript/api/word/word.documentproperties#company)|Obtiene o establece la compañía del documento.|
||[format](/javascript/api/word/word.documentproperties#format)|Obtiene o establece el formato del documento.|
||[Palabras clave](/javascript/api/word/word.documentproperties#keywords)|Obtiene o establece las palabras clave del documento.|
||[manager](/javascript/api/word/word.documentproperties#manager)|Obtiene o establece el administrador del documento.|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|Obtiene el nombre de aplicación del documento.|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|Obtiene la fecha de creación del documento.|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|Obtiene la colección de propiedades personalizadas del documento.|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|Obtiene el último autor del documento.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastprintdate)|Obtiene la última fecha de impresión del documento.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|Obtiene la última fecha de modificación del documento.|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|Obtiene el número de revisión del documento.|
||[seguridad](/javascript/api/word/word.documentproperties#security)|Obtiene la configuración de seguridad del documento.|
||[template](/javascript/api/word/word.documentproperties#template)|Obtiene la plantilla del documento.|
||[subject](/javascript/api/word/word.documentproperties#subject)|Obtiene o establece el asunto del documento.|
||[title](/javascript/api/word/word.documentproperties#title)|Obtiene o establece el título del documento.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext ()](/javascript/api/word/word.inlinepicture#getnext--)|Obtiene la siguiente imagen incorporada.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#getnextornullobject--)|Obtiene la siguiente imagen incorporada.|
||[getRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|Obtiene la imagen, o el punto de inicio o final de la imagen, como un intervalo.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|Obtiene el control de contenido que contiene la imagen incorporada.|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|Obtiene la tabla que contiene la imagen incorporada.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|Obtiene la celda de tabla que contiene la imagen incorporada.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|Obtiene la celda de tabla que contiene la imagen incorporada.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parenttableornullobject)|Obtiene la tabla que contiene la imagen incorporada.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|Obtiene la primera imagen incorporada de esta colección.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|Obtiene la primera imagen incorporada de esta colección.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs (LEVEL: Number)](/javascript/api/word/word.list#getlevelparagraphs-level-)|Obtiene los párrafos que se producen en el nivel especificado de la lista.|
||[getLevelString (LEVEL: Number)](/javascript/api/word/word.list#getlevelstring-level-)|Obtiene la viñeta, el número o la imagen en el nivel especificado como una cadena.|
||[insertParagraph (paragraphText: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|Inserta un párrafo en la ubicación especificada.|
||[id](/javascript/api/word/word.list#id)|Obtiene el identificador de la lista.|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|Comprueba si cada uno de los 9 niveles existe en la lista.|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|Obtiene todos los tipos de los 9 niveles de la lista.|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|Obtiene los párrafos de la lista.|
||[setLevelAlignment (LEVEL: Number, Alignment: Word. Alignment)](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|Establece la alineación de la viñeta, el número o la imagen en el nivel especificado de la lista.|
||[setLevelBullet (LEVEL: Number, listBullet: Word. ListBullet, charCode?: Number, fontName?: String)](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|Establece el formato de viñeta en el nivel especificado de la lista.|
||[setLevelIndents (LEVEL: Number, textIndent: Number, bulletNumberPictureIndent: Number)](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|Establece las dos sangrías del nivel especificado de la lista.|
||[setLevelNumbering (LEVEL: Number, listNumbering: Word. ListNumbering, formatString?: Array<número de cadena \|>)](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|Establece el formato de numeración en el nivel especificado de la lista.|
||[setLevelStartingNumber (LEVEL: Number, startingNumber: Number)](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|Establece el número de inicio en el nivel especificado de la lista.|
|[listCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|Obtiene una lista mediante su identificador.|
||[getByIdOrNullObject (ID: Number)](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|Obtiene una lista mediante su identificador.|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|Obtiene la primera lista de esta colección.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getfirstornullobject--)|Obtiene la primera lista de esta colección.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getitem-index-)|Obtiene un objeto de lista por su índice en la colección.|
||[items](/javascript/api/word/word.listcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor (parentOnly?: Boolean)](/javascript/api/word/word.listitem#getancestor-parentonly-)|Obtiene el elemento primario de lista o el antecesor más cercano si el primario no existe.|
||[getAncestorOrNullObject (parentOnly?: Boolean)](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|Obtiene el elemento primario de lista o el antecesor más cercano si el primario no existe.|
||[getDescendants (directChildrenOnly?: Boolean)](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|Obtiene todos los elementos de lista descendientes del elemento de lista.|
||[level](/javascript/api/word/word.listitem#level)|Obtiene o establece el nivel del elemento de la lista.|
||[listString](/javascript/api/word/word.listitem#liststring)|Obtiene la viñeta, el número o la imagen del elemento de lista como una cadena.|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|Obtiene el número de ordenación del elemento de lista en relación a los de su mismo nivel.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList (listId: Number, Level: Number)](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|Permite que el párrafo se una a una lista existente en el nivel especificado.|
||[detachFromList()](/javascript/api/word/word.paragraph#detachfromlist--)|Mueve este párrafo fuera de la lista, si este es un elemento de lista.|
||[getNext ()](/javascript/api/word/word.paragraph#getnext--)|Obtiene el párrafo siguiente.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#getnextornullobject--)|Obtiene el párrafo siguiente.|
||[getPrevious ()](/javascript/api/word/word.paragraph#getprevious--)|Obtiene el párrafo anterior.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#getpreviousornullobject--)|Obtiene el párrafo anterior.|
||[getRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.paragraph#getrange-rangelocation-)|Obtiene el párrafo completo, o el punto de inicio o final del párrafo, como un intervalo.|
||[getTextRanges (endingMarks: String [], trimSpacing?: Boolean)](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|Obtiene los intervalos de texto del párrafo usando signos de puntuación u otras marcas de fin.|
||[insertTable (rowCount: Number, columnCount: Number, insertLocation: Word. InsertLocation, Values?: String [] [])](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|Inserta una tabla con el número especificado de filas y columnas.|
||[isLastParagraph](/javascript/api/word/word.paragraph#islastparagraph)|Indica el párrafo que es el último dentro de su cuerpo primario.|
||[isListItem](/javascript/api/word/word.paragraph#islistitem)|Comprueba si el párrafo es un elemento de lista.|
||[list](/javascript/api/word/word.paragraph#list)|Obtiene la lista a la que pertenece este párrafo.|
||[listItem](/javascript/api/word/word.paragraph#listitem)|Obtiene ListItem para el párrafo.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listitemornullobject)|Obtiene ListItem para el párrafo.|
||[listOrNullObject](/javascript/api/word/word.paragraph#listornullobject)|Obtiene la lista a la que pertenece este párrafo.|
||[parentBody](/javascript/api/word/word.paragraph#parentbody)|Obtiene el cuerpo primario del párrafo.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentcontentcontrolornullobject)|Obtiene el control de contenido que contiene el párrafo.|
||[parentTable](/javascript/api/word/word.paragraph#parenttable)|Obtiene la tabla que contiene el párrafo.|
||[parentTableCell](/javascript/api/word/word.paragraph#parenttablecell)|Obtiene la celda de tabla que contiene el párrafo.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parenttablecellornullobject)|Obtiene la celda de tabla que contiene el párrafo.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parenttableornullobject)|Obtiene la tabla que contiene el párrafo.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tablenestinglevel)|Obtiene el nivel de la tabla del párrafo.|
||[Split (Delimiters: String [], trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|Divide el párrafo en intervalos secundarios mediante delimitadores.|
||[startNewList()](/javascript/api/word/word.paragraph#startnewlist--)|Inicia una nueva lista con este párrafo.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|Obtiene o establece el nombre del estilo integrado del párrafo.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|Obtiene el primer párrafo de esta colección.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|Obtiene el primer párrafo de esta colección.|
||[getLast ()](/javascript/api/word/word.paragraphcollection#getlast--)|Obtiene el último párrafo de esta colección.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|Obtiene el último párrafo de esta colección.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith (Range: Word. Range)](/javascript/api/word/word.range#comparelocationwith-range-)|Compara esta ubicación del intervalo con otra ubicación de este.|
||[expandto (Range: Word. Range)](/javascript/api/word/word.range#expandto-range-)|Devuelve un nuevo intervalo que se extiende desde este intervalo en cualquier dirección para cubrir otro intervalo.|
||[expandToOrNullObject (Range: Word. Range)](/javascript/api/word/word.range#expandtoornullobject-range-)|Devuelve un nuevo intervalo que se extiende desde este intervalo en cualquier dirección para cubrir otro intervalo.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#gethyperlinkranges--)|Obtiene intervalos secundarios de hipervínculo dentro del intervalo.|
||[getNextTextRange (endingMarks: String [], trimSpacing?: Boolean)](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|Obtiene el siguiente intervalo de texto mediante signos de puntuación u otras marcas de fin.|
||[getNextTextRangeOrNullObject (endingMarks: String [], trimSpacing?: Boolean)](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|Obtiene el siguiente intervalo de texto mediante signos de puntuación u otras marcas de fin.|
||[getRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.range#getrange-rangelocation-)|Clona el intervalo u obtiene el punto inicial o final del intervalo como un intervalo nuevo.|
||[getTextRanges (endingMarks: String [], trimSpacing?: Boolean)](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|Obtiene los intervalos secundarios de texto del intervalo mediante signos de puntuación u otras marcas de fin.|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|Obtiene el primer hipervínculo del intervalo, o establece un hipervínculo en el intervalo.|
||[insertTable (rowCount: Number, columnCount: Number, insertLocation: Word. InsertLocation, Values?: String [] [])](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|Inserta una tabla con el número especificado de filas y columnas.|
||[intersectWith (Range: Word. Range)](/javascript/api/word/word.range#intersectwith-range-)|Devuelve un intervalo nuevo como la intersección de este intervalo con otro.|
||[intersectWithOrNullObject (Range: Word. Range)](/javascript/api/word/word.range#intersectwithornullobject-range-)|Devuelve un intervalo nuevo como la intersección de este intervalo con otro.|
||[Vacío](/javascript/api/word/word.range#isempty)|Comprueba si la longitud del intervalo es cero.|
||[lists](/javascript/api/word/word.range#lists)|Obtiene la colección de objetos de lista en el intervalo.|
||[parentBody](/javascript/api/word/word.range#parentbody)|Obtiene el cuerpo primario del intervalo.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|Obtiene el control de contenido que contiene el intervalo.|
||[parentTable](/javascript/api/word/word.range#parenttable)|Obtiene la tabla que contiene el intervalo.|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|Obtiene la celda de tabla que contiene el intervalo.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parenttablecellornullobject)|Obtiene la celda de tabla que contiene el intervalo.|
||[parentTableOrNullObject](/javascript/api/word/word.range#parenttableornullobject)|Obtiene la tabla que contiene el intervalo.|
||[Table](/javascript/api/word/word.range#tables)|Obtiene la colección de objetos de tabla en el intervalo.|
||[Split (Delimiters: String [], Paragraphs?: Boolean, trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Divide el intervalo en intervalos secundarios mediante delimitadores.|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|Obtiene o establece el nombre del estilo integrado del intervalo.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|Obtiene el primer intervalo de esta colección.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getfirstornullobject--)|Obtiene el primer intervalo de esta colección.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#application)|[Conjunto de API: WordApi 1,3] *|
|[Section](/javascript/api/word/word.section)|[getNext ()](/javascript/api/word/word.section#getnext--)|Obtiene la próxima sección.|
||[getNextOrNullObject()](/javascript/api/word/word.section#getnextornullobject--)|Obtiene la próxima sección.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|Obtiene la primera sección de esta colección.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|Obtiene la primera sección de esta colección.|
|[Table](/javascript/api/word/word.table)|[addColumns (insertLocation: Word. InsertLocation, columnCount: Number, Values?: String [] [])](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|Agrega columnas al inicio o al final de la tabla, mediante la primera o la última columna existente como una plantilla.|
||[addRows (insertLocation: Word. InsertLocation, rowCount: Number, Values?: String [] [])](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|Agrega filas al inicio o al final de la tabla, mediante la primera o la última fila existente como una plantilla.|
||[orientación](/javascript/api/word/word.table#alignment)|Obtiene o establece la alineación de la tabla con la columna de la página.|
||[autoFitWindow()](/javascript/api/word/word.table#autofitwindow--)|Ajusta automáticamente las columnas de la tabla al ancho de la ventana.|
||[clear()](/javascript/api/word/word.table#clear--)|Borra el contenido de la tabla.|
||[delete()](/javascript/api/word/word.table#delete--)|Elimina toda la tabla.|
||[deleteColumns (columnIndex: Number, columnCount?: Number)](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|Elimina las columnas especificadas.|
||[deleteRows (rowIndex: Number, rowCount?: Number)](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|Elimina las filas especificadas.|
||[distributeColumns()](/javascript/api/word/word.table#distributecolumns--)|Distribuye los anchos de columna de manera uniforme.|
||[getBorder (borderLocation: Word. BorderLocation)](/javascript/api/word/word.table#getborder-borderlocation-)|Obtiene el estilo del borde para el borde especificado.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|Obtiene la celda de tabla de una fila y columna especificadas.|
||[getCellOrNullObject (rowIndex: Number, cellIndex: Number)](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|Obtiene la celda de tabla de una fila y columna especificadas.|
||[getCellPadding (cellPaddingLocation: Word. CellPaddingLocation)](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|Obtiene el espaciado entre borde y texto en puntos.|
||[getNext ()](/javascript/api/word/word.table#getnext--)|Obtiene la próxima tabla.|
||[getNextOrNullObject()](/javascript/api/word/word.table#getnextornullobject--)|Obtiene la próxima tabla.|
||[getParagraphAfter()](/javascript/api/word/word.table#getparagraphafter--)|Obtiene el párrafo después de la tabla.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getparagraphafterornullobject--)|Obtiene el párrafo después de la tabla.|
||[getParagraphBefore()](/javascript/api/word/word.table#getparagraphbefore--)|Obtiene el párrafo antes de la tabla.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|Obtiene el párrafo antes de la tabla.|
||[getRange (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.table#getrange-rangelocation-)|Obtiene el intervalo que contiene esta tabla, o el intervalo al inicio o al final de la tabla.|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|Obtiene y establece el número de filas de encabezado.|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|Obtiene y establece la alineación horizontal de cada celda de la tabla.|
||[ignorePunct](/javascript/api/word/word.table#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.table#ignorespace)||
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|Inserta un control de contenido en la tabla.|
||[insertParagraph (paragraphText: String, insertLocation: Word. InsertLocation)](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|Inserta un párrafo en la ubicación especificada.|
||[insertTable (rowCount: Number, columnCount: Number, insertLocation: Word. InsertLocation, Values?: String [] [])](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|Inserta una tabla con el número especificado de filas y columnas.|
||[matchCase](/javascript/api/word/word.table#matchcase)||
||[matchPrefix](/javascript/api/word/word.table#matchprefix)||
||[matchSuffix](/javascript/api/word/word.table#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.table#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.table#matchwildcards)||
||[font](/javascript/api/word/word.table#font)|Obtiene la fuente.|
||[isUniform](/javascript/api/word/word.table#isuniform)|Indica si todas las filas de la tabla son uniformes.|
||[nestingLevel](/javascript/api/word/word.table#nestinglevel)|Obtiene el nivel de anidamiento de la tabla.|
||[parentBody](/javascript/api/word/word.table#parentbody)|Obtiene el cuerpo primario de la tabla.|
||[parentContentControl](/javascript/api/word/word.table#parentcontentcontrol)|Obtiene el control de contenido que contiene la tabla.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentcontentcontrolornullobject)|Obtiene el control de contenido que contiene la tabla.|
||[parentTable](/javascript/api/word/word.table#parenttable)|Obtiene la tabla que contiene esta tabla.|
||[parentTableCell](/javascript/api/word/word.table#parenttablecell)|Obtiene la celda de tabla que contiene esta tabla.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parenttablecellornullobject)|Obtiene la celda de tabla que contiene esta tabla.|
||[parentTableOrNullObject](/javascript/api/word/word.table#parenttableornullobject)|Obtiene la tabla que contiene esta tabla.|
||[Filas](/javascript/api/word/word.table#rowcount)|Obtiene el número de filas de la tabla.|
||[rows](/javascript/api/word/word.table#rows)|Obtiene todas las filas de la tabla.|
||[Table](/javascript/api/word/word.table#tables)|Obtiene las tablas secundarias anidadas un nivel más profundo.|
||[Search (Dentrotexto: String, searchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignoreSpace?: Boolean matchCase?: Boolean matchPrefix?: Boolean matchSuffix?: Boolean?: Boolean matchWholeWord?: Boolean AC?: Boolean})](/javascript/api/word/word.table#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Realiza una búsqueda con el SearchOptions especificado en el ámbito del objeto Table.|
||[Select (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.table#select-selectionmode-)|Selecciona la tabla, o la posición al inicio o al final de la tabla, y dirige ahí la interfaz de usuario de Word.|
||[setCellPadding (cellPaddingLocation: Word. CellPaddingLocation, cellPadding: Number)](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)|Establece el espaciado entre borde y texto en puntos.|
||[shadingColor](/javascript/api/word/word.table#shadingcolor)|Obtiene y establece el color de sombreado.|
||[estilo](/javascript/api/word/word.table#style)|Obtiene o establece el nombre de estilo de la tabla.|
||[styleBandedColumns](/javascript/api/word/word.table#stylebandedcolumns)|Obtiene y establece si la tabla tiene columnas con bandas.|
||[styleBandedRows](/javascript/api/word/word.table#stylebandedrows)|Obtiene y establece si la tabla tiene filas con bandas.|
||[styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)|Obtiene o establece el nombre del estilo integrado de la tabla.|
||[styleFirstColumn](/javascript/api/word/word.table#stylefirstcolumn)|Obtiene y establece si la tabla tiene una primera columna con un estilo especial.|
||[styleLastColumn](/javascript/api/word/word.table#stylelastcolumn)|Obtiene y establece si la tabla tiene una última columna con un estilo especial.|
||[styleTotalRow](/javascript/api/word/word.table#styletotalrow)|Obtiene y establece si la tabla tiene una (última) fila total con un estilo especial.|
||[values](/javascript/api/word/word.table#values)|Obtiene y establece los valores de texto de la tabla, como una matriz 2D de Javascript.|
||[verticalAlignment](/javascript/api/word/word.table#verticalalignment)|Obtiene y establece la alineación vertical de cada celda de la tabla.|
||[width](/javascript/api/word/word.table#width)|Obtiene y establece el ancho de la tabla en puntos.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|Obtiene o establece el color del borde de la tabla.|
||[type](/javascript/api/word/word.tableborder#type)|Obtiene o establece el tipo de borde de la tabla.|
||[width](/javascript/api/word/word.tableborder#width)|Obtiene o establece el ancho, en puntos, del borde de la tabla.|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnwidth)|Obtiene y establece el ancho de la columna de la celda en puntos.|
||[deleteColumn()](/javascript/api/word/word.tablecell#deletecolumn--)|Elimina la columna que contiene esta celda.|
||[deleteRow ()](/javascript/api/word/word.tablecell#deleterow--)|Elimina la fila que contiene esta celda.|
||[getBorder (borderLocation: Word. BorderLocation)](/javascript/api/word/word.tablecell#getborder-borderlocation-)|Obtiene el estilo del borde para el borde especificado.|
||[getCellPadding (cellPaddingLocation: Word. CellPaddingLocation)](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)|Obtiene el espaciado entre borde y texto en puntos.|
||[getNext ()](/javascript/api/word/word.tablecell#getnext--)|Obtiene la celda siguiente.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#getnextornullobject--)|Obtiene la celda siguiente.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)|Obtiene y establece la alineación horizontal de la celda.|
||[insertColumns (insertLocation: Word. InsertLocation, columnCount: Number, Values?: String [] [])](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)|Agrega columnas a la izquierda o derecha de la celda, mediante la columna de la celda como una plantilla.|
||[insertRows (insertLocation: Word. InsertLocation, rowCount: Number, Values?: String [] [])](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)|Inserta filas por encima o debajo de la celda, mediante la fila de la celda como una plantilla.|
||[body](/javascript/api/word/word.tablecell#body)|Obtiene el objeto de cuerpo de la celda.|
||[cellIndex](/javascript/api/word/word.tablecell#cellindex)|Obtiene el índice de la celda de la fila.|
||[parentRow](/javascript/api/word/word.tablecell#parentrow)|Obtiene la fila primaria de la celda.|
||[parentTable](/javascript/api/word/word.tablecell#parenttable)|Obtiene la tabla primaria de la celda.|
||[rowIndex](/javascript/api/word/word.tablecell#rowindex)|Obtiene el índice de fila de la celda en la tabla.|
||[width](/javascript/api/word/word.tablecell#width)|Obtiene el ancho de la celda en puntos.|
||[setCellPadding (cellPaddingLocation: Word. CellPaddingLocation, cellPadding: Number)](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)|Establece el espaciado entre borde y texto en puntos.|
||[shadingColor](/javascript/api/word/word.tablecell#shadingcolor)|Obtiene o establece el color de sombreado de la celda.|
||[value](/javascript/api/word/word.tablecell#value)|Obtiene y establece el texto de la celda.|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)|Obtiene y establece la alineación vertical de la celda.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)|Obtiene la primera celda de tabla de esta colección.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#getfirstornullobject--)|Obtiene la primera celda de tabla de esta colección.|
||[items](/javascript/api/word/word.tablecellcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getfirst--)|Obtiene la primera tabla de esta colección.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#getfirstornullobject--)|Obtiene la primera tabla de esta colección.|
||[items](/javascript/api/word/word.tablecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear--)|Borra el contenido de la fila.|
||[delete()](/javascript/api/word/word.tablerow#delete--)|Elimina toda la fila.|
||[getBorder (borderLocation: Word. BorderLocation)](/javascript/api/word/word.tablerow#getborder-borderlocation-)|Obtiene el estilo del borde de las celdas de la fila.|
||[getCellPadding (cellPaddingLocation: Word. CellPaddingLocation)](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)|Obtiene el espaciado entre borde y texto en puntos.|
||[getNext ()](/javascript/api/word/word.tablerow#getnext--)|Obtiene la fila siguiente.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#getnextornullobject--)|Obtiene la fila siguiente.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)|Obtiene y establece la alineación horizontal de cada celda de la fila.|
||[ignorePunct](/javascript/api/word/word.tablerow#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.tablerow#ignorespace)||
||[insertRows (insertLocation: Word. InsertLocation, rowCount: Number, Values?: String [] [])](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)|Inserta filas mediante esta fila como una plantilla.|
||[matchCase](/javascript/api/word/word.tablerow#matchcase)||
||[matchPrefix](/javascript/api/word/word.tablerow#matchprefix)||
||[matchSuffix](/javascript/api/word/word.tablerow#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.tablerow#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.tablerow#matchwildcards)||
||[preferredHeight](/javascript/api/word/word.tablerow#preferredheight)|Obtiene y establece el alto preferido de la fila en puntos.|
||[cellCount](/javascript/api/word/word.tablerow#cellcount)|Obtiene el número de celdas de la fila.|
||[células](/javascript/api/word/word.tablerow#cells)|Obtiene celdas.|
||[font](/javascript/api/word/word.tablerow#font)|Obtiene la fuente.|
||[isHeader](/javascript/api/word/word.tablerow#isheader)|Comprueba si la fila es una fila de encabezado.|
||[parentTable](/javascript/api/word/word.tablerow#parenttable)|Obtiene la tabla primaria.|
||[rowIndex](/javascript/api/word/word.tablerow#rowindex)|Obtiene el índice de la fila en la tabla primaria.|
||[Search (Dentrotexto: String, searchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignoreSpace?: Boolean matchCase?: Boolean matchPrefix?: Boolean matchSuffix?: Boolean?: Boolean matchWholeWord?: Boolean AC?: Boolean})](/javascript/api/word/word.tablerow#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Realiza una búsqueda con el SearchOptions especificado en el ámbito de la fila.|
||[Select (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.tablerow#select-selectionmode-)|Selecciona la fila y se desplaza por la interfaz de usuario de Word hasta ella.|
||[setCellPadding (cellPaddingLocation: Word. CellPaddingLocation, cellPadding: Number)](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)|Establece el espaciado entre borde y texto en puntos.|
||[shadingColor](/javascript/api/word/word.tablerow#shadingcolor)|Obtiene y establece el color de sombreado.|
||[values](/javascript/api/word/word.tablerow#values)|Obtiene y establece los valores de texto de la fila, como una matriz 2D de JavaScript.|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)|Obtiene y establece la alineación vertical de las celdas de la fila.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)|Obtiene la primera fila de esta colección.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#getfirstornullobject--)|Obtiene la primera fila de esta colección.|
||[items](/javascript/api/word/word.tablerowcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
