| Class | Campos | Descripción |
|:---|:---|:---|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|El nombre del proveedor de datos para el tipo de datos vinculados.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|La fecha y la hora locales de la zona horaria local desde que se abrió el libro cuando se actualizó por última vez el tipo de datos vinculados.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|Nombre del tipo de datos vinculados.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|Frecuencia, en segundos, con la que se actualiza el tipo de datos vinculados si `refreshMode` se establece en "periódico".|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|Mecanismo por el cual se recuperan los datos para el tipo de datos vinculados.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|Identificador único del tipo de datos vinculados.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Devuelve una matriz con todos los modos de actualización admitidos por el tipo de datos vinculados.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Realiza una solicitud para actualizar el tipo de datos vinculados.|
||[requestSetRefreshMode (refreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Realiza una solicitud para cambiar el modo de actualización para este tipo de datos vinculados.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|Identificador único del nuevo tipo de datos vinculados.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Obtiene el origen del evento.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Obtiene el tipo del evento.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Obtiene el número de tipos de datos vinculados de la colección.|
||[getItem (clave: número)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Obtiene un tipo de datos vinculado por identificador de servicio.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Obtiene un tipo de datos vinculado por su índice en la colección.|
||[getItemOrNullObject (Key: Number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Obtiene un tipo de datos vinculado por ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Realiza una solicitud para actualizar todos los tipos de datos vinculados de la colección.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Activa esta vista de hoja.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Quita la vista de hoja de la hoja de cálculo.|
||[Duplicate (Name?: String)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Crea una copia de esta vista de hoja.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Obtiene o establece el nombre de la vista de hoja.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Crea una nueva vista de hoja con el nombre especificado.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Crea y activa una nueva vista de hoja temporal.|
||[Exit ()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Sale de la vista de hoja activa en ese momento.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Obtiene la vista de hoja actualmente activa de la hoja de cálculo.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Obtiene el número de vistas de hoja de esta hoja de cálculo.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Obtiene una vista de hoja mediante su nombre.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Obtiene una vista de hoja por su índice en la colección.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|La descripción del texto alternativo de la tabla dinámica.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|Título del texto alternativo de la tabla dinámica.|
||[displayBlankLineAfterEachItem (display: Boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Establece si se va a mostrar una línea en blanco después de cada elemento.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Texto que se rellena automáticamente en cualquier celda vacía de la tabla dinámica si `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Especifica si las celdas vacías de la tabla dinámica deben rellenarse con el `emptyCellText` .|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtiene una única celda de la tabla dinámica en función de una jerarquía de datos y de los elementos de fila y columna de sus respectivas jerarquías.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Estilo aplicado a la tabla dinámica.|
||[repeatAllItemLabels (repeatLabels: Boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Establece la configuración "Repetir todas las etiquetas de elementos" en todos los campos de la tabla dinámica.|
||[setStyle (Style: String \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Establece el estilo aplicado a la tabla dinámica.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Especifica si la tabla dinámica muestra los encabezados de campo (títulos de campo y listas desplegables de filtros).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Especifica si la tabla dinámica se actualiza cuando se abre el libro.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Devuelve un `WorkbookRangeAreas` objeto Object que representa el rango que contiene todas las celdas precedentes de una celda en la misma hoja de cálculo o en varias hojas de cálculo.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|El modo de actualización de tipo de datos vinculados.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|Identificador único del objeto cuyo modo de actualización se ha cambiado.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Obtiene el origen del evento.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Obtiene el tipo del evento.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[actualiza](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indica si la solicitud de actualización se realizó correctamente.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|Identificador único del objeto cuya solicitud de actualización se ha completado.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Obtiene el origen del evento.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Obtiene el tipo del evento.|
||[relativas](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Matriz que contiene las advertencias generadas a partir de la solicitud de actualización.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Crea un gráfico vectorial escalable (SVG) de una cadena XML y lo agrega a la hoja de cálculo.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Indica el nombre de la segmentación usado en la fórmula.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Estilo aplicado a la segmentación de la rebanada.|
||[setStyle (Style: String \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Establece el estilo aplicado a la segmentación de la rebanada.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Cambia la tabla para usar el estilo de tabla predeterminado.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Se produce cuando se aplica el filtro en una tabla en particular.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Estilo aplicado a la tabla.|
||[setStyle (estilo: String \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Establece el estilo aplicado a la tabla.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Se produce cuando se aplica el filtro en cualquier tabla en un libro o una hoja de cálculo.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Obtiene el identificador de la tabla en la que se aplica el filtro.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo que contiene la tabla.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Devuelve una colección de tipos de datos vinculados que forman parte del libro.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Especifica si se muestra el panel lista de campos de tabla dinámica en el nivel de libro.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True si el libro usa el sistema de fechas 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Devuelve una colección de vistas de hoja que están presentes en la hoja de cálculo.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Se produce cuando se aplica el filtro en una hoja de cálculo concreta.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Inserta las hojas de cálculo especificadas de un libro en el libro actual.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Se produce cuando se aplica cualquier filtro de hoja de cálculo al libro.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se aplica el filtro.|
