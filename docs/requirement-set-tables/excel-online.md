| Clase | Campos | Descripción |
|:---|:---|:---|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Activa esta vista de hoja.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Quita la vista de hoja de la hoja de cálculo.|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Crea una copia de esta vista de hoja.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Obtiene o establece el nombre de la vista de hoja.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Crea una nueva vista de hoja con el nombre especificado.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Crea y activa una nueva vista de hoja temporal.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Sale de la vista de hoja activa actualmente.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Obtiene la vista de hoja activa actualmente de la hoja de cálculo.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Obtiene el número de vistas de hoja en esta hoja de cálculo.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Obtiene una vista de hoja con su nombre.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Obtiene una vista de hoja por su índice en la colección.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[Range](/javascript/api/excel/excel.range)|[getExtendedRange(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|Devuelve un objeto range que incluye el intervalo actual y hasta el borde del intervalo, en función de la dirección proporcionada.|
||[getMergedAreas()](/javascript/api/excel/excel.range#getmergedareas--)|Devuelve un `RangeAreas` objeto que representa las áreas combinadas de este intervalo.|
||[getRangeEdge(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|Devuelve un objeto range que es la celda perimetral de la región de datos que corresponde a la dirección proporcionada.|
|[Table](/javascript/api/excel/excel.table)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#resize-newrange-)|Cambie el tamaño de la tabla al nuevo intervalo.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Devuelve una colección de vistas de hoja que están presentes en la hoja de cálculo.|
