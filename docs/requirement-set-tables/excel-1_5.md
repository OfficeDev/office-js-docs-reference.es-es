| Clase | Campos | Descripción |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete--)|Elimina el elemento XML personalizado.|
||[getXml()](/javascript/api/excel/excel.customxmlpart#getxml--)|Obtiene el contenido XML completo del elemento XML personalizado.|
||[id](/javascript/api/excel/excel.customxmlpart#id)|Identificador del elemento XML personalizado.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#namespaceuri)|URI del espacio de nombres del elemento XML personalizado.|
||[setXml(xml: string)](/javascript/api/excel/excel.customxmlpart#setxml-xml-)|Establece el contenido XML completo del elemento XML personalizado.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[add(xml: string)](/javascript/api/excel/excel.customxmlpartcollection#add-xml-)|Se agrega un nuevo elemento XML personalizado al libro.|
||[getByNamespace(namespaceUri: string)](/javascript/api/excel/excel.customxmlpartcollection#getbynamespace-namespaceuri-)|Obtiene una nueva colección con ámbito de elementos XML personalizados cuyos espacios de nombres coinciden con el espacio de nombres determinado.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getcount--)|Obtiene el número de elementos XML personalizados de la colección.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitem-id-)|Obtiene un elemento XML personalizado a partir de su identificador.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitemornullobject-id-)|Obtiene un elemento XML personalizado a partir de su identificador.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getcount--)|Obtiene el número de elementos CustomXML de esta colección.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitem-id-)|Obtiene un elemento XML personalizado a partir de su identificador.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitemornullobject-id-)|Obtiene un elemento XML personalizado a partir de su identificador.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitem--)|Si la colección contiene exactamente un elemento, este método lo devuelve.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitemornullobject--)|Si la colección contiene exactamente un elemento, este método lo devuelve.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|Id. de la tabla dinámica.|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[tiempo de ejecución](/javascript/api/excel/excel.requestcontext#runtime)|[Conjunto de api: ExcelApi 1.5]|
|[Tiempo de ejecución](/javascript/api/excel/excel.runtime)||[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customxmlparts)|Representa la colección de elementos XML personalizados contenidos en este libro.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getnext-visibleonly-)|Obtiene la hoja de cálculo que sigue a esta.|
||[getNextOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getnextornullobject-visibleonly-)|Obtiene la hoja de cálculo que sigue a esta.|
||[getPrevious(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getprevious-visibleonly-)|Obtiene la hoja de cálculo que precede a esta.|
||[getPreviousOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getpreviousornullobject-visibleonly-)|Obtiene la hoja de cálculo que precede a esta.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getfirst-visibleonly-)|Obtiene la primera hoja de cálculo de la colección.|
||[getLast(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getlast-visibleonly-)|Obtiene la última hoja de cálculo de la colección.|
