| Class | Campos | Descripción |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#ondatachanged)|Se produce cuando se cambian los datos del control de contenido.|
||[onDeleted](/javascript/api/word/word.contentcontrol#ondeleted)|Se produce cuando se elimina el control de contenido.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onselectionchanged)|Se produce cuando se cambia la selección en el control de contenido.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentcontrol)|El objeto que generó el evento.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventtype)|El tipo de evento.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete--)|Elimina el elemento XML personalizado.|
||[deleteAttribute (XPath: String, namespaceMappings: any, Name: String)](/javascript/api/word/word.customxmlpart#deleteattribute-xpath--namespacemappings--name-)|Elimina un atributo con el nombre especificado del elemento identificado por XPath.|
||[deleteElement (XPath: String, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteelement-xpath--namespacemappings-)|Elimina el elemento identificado por XPath.|
||[getXml ()](/javascript/api/word/word.customxmlpart#getxml--)|Obtiene el contenido XML completo del elemento XML personalizado.|
||[insertAttribute (XPath: String, namespaceMappings: any, Name: String, Value: String)](/javascript/api/word/word.customxmlpart#insertattribute-xpath--namespacemappings--name--value-)|Inserta un atributo con el nombre y el valor dados en el elemento identificado por XPath.|
||[insertElement (XPath: String, XML: String, namespaceMappings: any, index?: Number)](/javascript/api/word/word.customxmlpart#insertelement-xpath--xml--namespacemappings--index-)|Inserta el XML especificado en el elemento primario identificado por XPath en el índice de posición secundaria.|
||[consulta (XPath: String, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query-xpath--namespacemappings-)|Consulta el contenido XML del elemento XML personalizado.|
||[id](/javascript/api/word/word.customxmlpart#id)|Obtiene el identificador del elemento XML personalizado.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceuri)|Obtiene el URI del espacio de nombres del elemento XML personalizado.|
||[setXml (XML: String)](/javascript/api/word/word.customxmlpart#setxml-xml-)|Establece el contenido XML completo del elemento XML personalizado.|
||[updateAttribute (XPath: String, namespaceMappings: any, Name: String, Value: String)](/javascript/api/word/word.customxmlpart#updateattribute-xpath--namespacemappings--name--value-)|Actualiza el valor de un atributo con el nombre especificado del elemento identificado por XPath.|
||[updateElement (XPath: String, XML: String, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateelement-xpath--xml--namespacemappings-)|Actualiza el XML del elemento identificado por XPath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[Add (XML: String)](/javascript/api/word/word.customxmlpartcollection#add-xml-)|Agrega un nuevo elemento XML personalizado al documento.|
||[getByNamespace (namespaceUri: String)](/javascript/api/word/word.customxmlpartcollection#getbynamespace-namespaceuri-)|Obtiene una nueva colección con ámbito de elementos XML personalizados cuyos espacios de nombres coinciden con el espacio de nombres determinado.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getcount--)|Obtiene el número de objetos de la colección.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getitem-id-)|Obtiene un elemento XML personalizado a partir de su identificador.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getitemornullobject-id-)|Obtiene un elemento XML personalizado a partir de su identificador.|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getcount--)|Obtiene el número de objetos de la colección.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitem-id-)|Obtiene un elemento XML personalizado a partir de su identificador.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitemornullobject-id-)|Obtiene un elemento XML personalizado a partir de su identificador.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitem--)|Si la colección contiene exactamente un elemento, este método lo devuelve.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitemornullobject--)|Si la colección contiene exactamente un elemento, este método lo devuelve.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[Document](/javascript/api/word/word.document)|[deleteBookmark (Name: String)](/javascript/api/word/word.document#deletebookmark-name-)|Elimina del documento un marcador, si existe.|
||[getBookmarkRange (Name: String)](/javascript/api/word/word.document#getbookmarkrange-name-)|Obtiene el intervalo de un marcador.|
||[getBookmarkRangeOrNullObject (Name: String)](/javascript/api/word/word.document#getbookmarkrangeornullobject-name-)|Obtiene el intervalo de un marcador.|
||[customXmlParts](/javascript/api/word/word.document#customxmlparts)|Obtiene los elementos XML personalizados del documento.|
||[onContentControlAdded](/javascript/api/word/word.document#oncontentcontroladded)|Se produce cuando se agrega un control de contenido.|
||[settings](/javascript/api/word/word.document#settings)|Obtiene la configuración del complemento en el documento.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark (Name: String)](/javascript/api/word/word.documentcreated#deletebookmark-name-)|Elimina del documento un marcador, si existe.|
||[getBookmarkRange (Name: String)](/javascript/api/word/word.documentcreated#getbookmarkrange-name-)|Obtiene el intervalo de un marcador.|
||[getBookmarkRangeOrNullObject (Name: String)](/javascript/api/word/word.documentcreated#getbookmarkrangeornullobject-name-)|Obtiene el intervalo de un marcador.|
||[customXmlParts](/javascript/api/word/word.documentcreated#customxmlparts)|Obtiene los elementos XML personalizados del documento.|
||[settings](/javascript/api/word/word.documentcreated#settings)|Obtiene la configuración del complemento en el documento.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageformat)|Obtiene el formato de la imagen incorporada.|
|[List](/javascript/api/word/word.list)|[getLevelFont (LEVEL: Number)](/javascript/api/word/word.list#getlevelfont-level-)|Obtiene la fuente de la viñeta, el número o la imagen en el nivel especificado de la lista.|
||[getLevelPicture (LEVEL: Number)](/javascript/api/word/word.list#getlevelpicture-level-)|Obtiene la representación de cadena codificada en Base64 de la imagen en el nivel especificado de la lista.|
||[resetLevelFont (LEVEL: Number, resetFontName?: Boolean)](/javascript/api/word/word.list#resetlevelfont-level--resetfontname-)|Restablece la fuente de la viñeta, el número o la imagen en el nivel especificado de la lista.|
||[setLevelPicture (LEVEL: Number, base64EncodedImage?: String)](/javascript/api/word/word.list#setlevelpicture-level--base64encodedimage-)|Establece la imagen en el nivel especificado de la lista.|
|[Range](/javascript/api/word/word.range)|[getBookmarks (includeHidden?: Boolean, includeAdjacent?: Boolean)](/javascript/api/word/word.range#getbookmarks-includehidden--includeadjacent-)|Obtiene los nombres de todos los marcadores o superpuestos del intervalo.|
||[insertBookmark (Name: String)](/javascript/api/word/word.range#insertbookmark-name-)|Inserta un marcador en el rango.|
|[Setting](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete--)|Elimina la configuración.|
||[key](/javascript/api/word/word.setting#key)|Obtiene la clave de la configuración.|
||[value](/javascript/api/word/word.setting#value)|Obtiene o establece el valor de la configuración.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[Add (Key: String, Value: any)](/javascript/api/word/word.settingcollection#add-key--value-)|Crea una nueva configuración o establece una configuración existente.|
||[deleteAll ()](/javascript/api/word/word.settingcollection#deleteall--)|Elimina toda la configuración de este complemento.|
||[getCount()](/javascript/api/word/word.settingcollection#getcount--)|Obtiene el recuento de opciones de configuración.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getitem-key-)|Obtiene un objeto de configuración por su clave, que distingue mayúsculas de minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getitemornullobject-key-)|Obtiene un objeto de configuración por su clave, que distingue mayúsculas de minúsculas.|
||[items](/javascript/api/word/word.settingcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[Table](/javascript/api/word/word.table)|[mergeCells (topRow: Number, firstCell: Number, bottomRow: Number, lastCell: Number)](/javascript/api/word/word.table#mergecells-toprow--firstcell--bottomrow--lastcell-)|Combina las celdas limitadas por una primera y la última celda.|
|[TableCell](/javascript/api/word/word.tablecell)|[Split (rowCount: Number, columnCount: Number)](/javascript/api/word/word.tablecell#split-rowcount--columncount-)|Divide la celda en el número especificado de filas y columnas.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertcontentcontrol--)|Inserta un control de contenido en la fila.|
||[Merge ()](/javascript/api/word/word.tablerow#merge--)|Combina la fila en una celda.|
