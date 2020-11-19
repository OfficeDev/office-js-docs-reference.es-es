| Class | Campos | Descripción |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Especifica el ángulo al que está orientado el texto para el título del eje del gráfico.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (dimensión: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Obtiene los valores de una única dimensión de la serie del gráfico.|
|[Comentario](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Obtiene el tipo de contenido del comentario.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|Obtiene la `CommentDetail` matriz que contiene el identificador de comentario y los identificadores de sus respuestas relacionadas.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Especifica el origen del evento.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se produjo el evento.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|Obtiene el tipo de cambio que representa el modo en que se desencadena el evento Changed.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|Obtenga la matriz CommentDetail que contiene el identificador de comentario y los identificadores de sus respuestas relacionadas.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Especifica el origen del evento.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se produjo el evento.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|Se produce cuando se agregan los comentarios.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|Se produce cuando se cambian los comentarios o las respuestas de una colección de comentarios, incluso cuando se eliminan las respuestas.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|Se produce cuando se eliminan comentarios en la colección de comentarios.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|Obtiene la `CommentDetail` matriz que contiene el identificador de comentario y los identificadores de sus respuestas relacionadas.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Especifica el origen del evento.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se produjo el evento.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|Representa el identificador del comentario.|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|Representa los identificadores de las respuestas relacionadas que pertenecen al comentario.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|El tipo de contenido de la respuesta.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|Define el formato adecuado culturalmente para mostrar la fecha y la hora.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|Obtiene la cadena que se usa como separador de fecha.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|Obtiene la cadena de formato para un valor de fecha larga.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|Obtiene la cadena de formato para un valor de hora larga.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|Obtiene la cadena de formato para un valor de fecha corta.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|Obtiene la cadena que se usa como separador de hora.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparación](/javascript/api/excel/excel.pivotdatefilter#comparator)|El Comparator es el valor estático con el que se comparan otros valores.|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|Especifica la condición del filtro, que define los criterios de filtrado necesarios.|
||[exclusivas](/javascript/api/excel/excel.pivotdatefilter#exclusive)|Si es true, Filter *excluye* los elementos que cumplen los criterios.|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|Límite inferior del intervalo para la `Between` condición de filtro.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|Límite superior del intervalo para la `Between` condición de filtro.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|Para `Equals` `Before` `After` las condiciones de filtrado,, y `Between` , indica si las comparaciones deben realizarse como días completos.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter (Filter: Excel. PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|Establece uno o varios los filtros dinámicos actuales del campo y los aplica al campo.|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|Borra todos los criterios de todos los filtros del campo.|
||[clearFilter (filterType: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|Borra todos los criterios existentes del filtro del campo del tipo especificado (si se aplica uno actualmente).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getfilters--)|Obtiene todos los filtros aplicados actualmente en el campo.|
||[isFiltered (filterType?: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|Comprueba si hay filtros aplicados en el campo.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|El filtro de fecha aplicado actualmente del campo dinámico.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|El filtro de etiqueta aplicado actualmente al campo de campo.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|El filtro manual aplicado actualmente del campo dinámico.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|El filtro de valor aplicado actualmente del campo dinámico.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[comparación](/javascript/api/excel/excel.pivotlabelfilter#comparator)|El Comparator es el valor estático con el que se comparan otros valores.|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#condition)|Especifica la condición del filtro, que define los criterios de filtrado necesarios.|
||[exclusivas](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|Si es true, Filter *excluye* los elementos que cumplen los criterios.|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|Límite inferior del intervalo para la condición de filtro between.|
||[subcadena](/javascript/api/excel/excel.pivotlabelfilter#substring)|Subcadena que se usa `BeginsWith` para `EndsWith` `Contains` las condiciones de filtro, y.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|Límite superior del intervalo para la condición de filtro between.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Lista de elementos seleccionados que se va a filtrar manualmente.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Especifica si la tabla dinámica permite la aplicación de varios filtros dinámicos en un campo dinámico dado en la tabla.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Obtiene el número de tablas dinámicas de la colección.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Obtiene la primera tabla dinámica de la colección.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Obtiene una tabla dinámica por nombre.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Obtiene una tabla dinámica por nombre.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparación](/javascript/api/excel/excel.pivotvaluefilter#comparator)|El Comparator es el valor estático con el que se comparan otros valores.|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|Especifica la condición del filtro, que define los criterios de filtrado necesarios.|
||[exclusivas](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Si es true, Filter *excluye* los elementos que cumplen los criterios.|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|Límite inferior del intervalo para la `Between` condición de filtro.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Especifica si el filtro es para los N elementos superiores/inferiores, N% superior/inferior o suma N superior o inferior.|
||[límite](/javascript/api/excel/excel.pivotvaluefilter#threshold)|El número de umbral "N" de elementos, porcentaje o suma que se va a filtrar para una condición de filtro superior o inferior.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|Límite superior del intervalo para la `Between` condición de filtro.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Nombre del "valor" elegido en el campo por el que se va a filtrar.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getdirectprecedents--)|Devuelve un objeto WorkbookRangeAreas que representa el rango que contiene todas las celdas precedentes directas de una celda en la misma hoja de cálculo o en varias hojas de cálculo.|
||[getPivotTables (fullyContained?: Boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Obtiene una colección con ámbito de tablas dinámicas superpuestas con el intervalo.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtiene el intervalo que contiene la celda de delimitador de una celda en la que se derrama.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtiene el intervalo que contiene la celda de delimitador de una celda en la que se derrama.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtiene el objeto de intervalo que contiene el intervalo de desbordamiento al llamar en una celda de delimitador.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtiene el objeto de intervalo que contiene el intervalo de desbordamiento al llamar en una celda de delimitador.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Indica si todas las celdas tienen un borde de desbordamiento.|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|Representa la categoría del formato de número de cada celda.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Representa si todas las celdas se guardarían como una fórmula de matriz.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|Obtiene el número de objetos RangeAreas de esta colección.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|Devuelve el objeto RangeAreas basándose en la posición de la colección.|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet (Key: String)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|Devuelve el `RangeAreas` objeto basándose en el identificador de la hoja de cálculo o el nombre de la colección.|
||[getRangeAreasOrNullObjectBySheet (Key: String)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|Devuelve el `RangeAreas` objeto basándose en el nombre de la hoja de cálculo o el identificador de la colección.|
||[direcciones](/javascript/api/excel/excel.workbookrangeareas#addresses)|Devuelve una matriz de direcciones en estilo a1.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|Devuelve el `RangeAreasCollection` objeto.|
||[intervalos](/javascript/api/excel/excel.workbookrangeareas#ranges)|Devuelve los rangos que componen este objeto en un  `RangeCollection`   objeto.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Obtiene una colección de propiedades personalizadas en el nivel de hoja de cálculo.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Elimina la propiedad personalizada.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Obtiene la clave de la propiedad personalizada.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Obtiene o establece el valor de la propiedad personalizada.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[Add (Key: String, Value: String)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Agrega una nueva propiedad personalizada que se asigna a la clave proporcionada.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Obtiene el número de propiedades personalizadas de esta hoja de cálculo.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Obtiene un objeto de propiedad personalizada mediante su clave, que no distingue mayúsculas de minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Obtiene un objeto de propiedad personalizada mediante su clave, que no distingue mayúsculas de minúsculas.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
