| Clase | Campos | Descripción |
|:---|:---|:---|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|Dirección de la celda que contiene la fórmula modificada.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|Representa la fórmula anterior, antes de cambiarla.|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#positiontype)|Posición de inserción, en el libro actual, de las nuevas hojas de cálculo.|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#relativeto)|Hoja de cálculo del libro actual al que se hace referencia para el `WorksheetPositionType` parámetro.|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetnamestoinsert)|Los nombres de las hojas de cálculo individuales que se insertarán.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|La descripción de texto alternativo de la tabla dinámica.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|El título de texto alternativo de la tabla dinámica.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Establece si se va a mostrar una línea en blanco después de cada elemento.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Texto que se rellena automáticamente en cualquier celda vacía de la tabla dinámica si `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Especifica si las celdas vacías de la tabla dinámica deben rellenarse con `emptyCellText` .|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Establece el valor "repetir todas las etiquetas de elementos" en todos los campos de la tabla dinámica.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Especifica si la tabla dinámica muestra encabezados de campo (títulos de campo y listas desplegables de filtro).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Especifica si la tabla dinámica se actualiza cuando se abre el libro.|
|[Range](/javascript/api/excel/excel.range)|[getDirectDependents()](/javascript/api/excel/excel.range#getdirectdependents--)|Devuelve un objeto que representa el rango que contiene todos los dependientes directos de una celda en la misma hoja de cálculo `WorkbookRangeAreas` o en varias hojas de cálculo.|
||[getExtendedRange(direction: Excel. KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|Devuelve un objeto range que incluye el intervalo actual y hasta el borde del intervalo, en función de la dirección proporcionada.|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#getmergedareasornullobject--)|Devuelve un objeto RangeAreas que representa las áreas combinadas de este rango.|
||[getRangeEdge(direction: Excel. KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|Devuelve un objeto range que es la celda perimetral de la región de datos que corresponde a la dirección proporcionada.|
|[Table](/javascript/api/excel/excel.table)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#resize-newrange-)|Cambie el tamaño de la tabla al nuevo intervalo.|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel. InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#insertworksheetsfrombase64-base64file--options-)|Inserta las hojas de cálculo especificadas de un libro de origen en el libro actual.|
||[onActivated](/javascript/api/excel/excel.workbook#onactivated)|Se produce cuando se activa el libro.|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#type)|Obtiene el tipo del evento.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|Se produce cuando se cambian una o más fórmulas en esta hoja de cálculo.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|Se produce cuando se cambian una o más fórmulas en cualquier hoja de cálculo de esta colección.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|Obtiene una matriz `FormulaChangedEventDetail` de objetos, que contienen los detalles sobre todas las fórmulas modificadas.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|Origen del evento.|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que cambió la fórmula.|
