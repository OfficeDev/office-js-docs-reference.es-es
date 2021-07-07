| Clase | Campos | Descripción |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Especifica el ángulo al que está orientado el texto para el título del eje del gráfico.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Obtiene los valores de una sola dimensión de la serie de gráficos.|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Obtiene el tipo de contenido del comentario.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|Obtiene la matriz que contiene el identificador de comentario y los identificadores `CommentDetail` de sus respuestas relacionadas.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Especifica el origen del evento.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se produjo el evento.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|Obtiene el tipo de cambio que representa cómo se desencadena el evento modificado.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|Obtiene la matriz que contiene el identificador de comentario y `CommentDetail` los identificadores de sus respuestas relacionadas.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Especifica el origen del evento.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se produjo el evento.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|Se produce cuando se agregan los comentarios.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|Se produce cuando se cambian los comentarios o las respuestas de una colección de comentarios, incluso cuando se eliminan las respuestas.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|Se produce cuando los comentarios se eliminan en la colección de comentarios.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|Obtiene la matriz que contiene el identificador de comentario y los identificadores `CommentDetail` de sus respuestas relacionadas.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Especifica el origen del evento.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se produjo el evento.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|Representa el identificador del comentario.|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|Representa los IDs de las respuestas relacionadas que pertenecen al comentario.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|Tipo de contenido de la respuesta.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|Define el formato culturalmente adecuado para mostrar fecha y hora.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|Obtiene la cadena usada como separador de fecha.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|Obtiene la cadena de formato de un valor de fecha larga.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|Obtiene la cadena de formato de un valor de tiempo largo.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|Obtiene la cadena de formato de un valor de fecha breve.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|Obtiene la cadena usada como separador de hora.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparador](/javascript/api/excel/excel.pivotdatefilter#comparator)|El comparador es el valor estático con el que se comparan otros valores.|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|Especifica la condición del filtro, que define los criterios de filtrado necesarios.|
||[exclusive](/javascript/api/excel/excel.pivotdatefilter#exclusive)|Si `true` , filter excluye los *elementos* que cumplen los criterios.|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|Límite inferior del intervalo para la `between` condición de filtro.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|Límite superior del intervalo para la `between` condición de filtro.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|For , y las condiciones de filtro, indica si se deben realizar comparaciones `equals` `before` como días `after` `between` enteros.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel. PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|Establece uno o varios de los filtros dinámicos actuales del campo y los aplica al campo.|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|Borra todos los criterios de todos los filtros del campo.|
||[clearFilter(filterType: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|Borra todos los criterios existentes del filtro del campo del tipo especificado (si se aplica uno actualmente).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getfilters--)|Obtiene todos los filtros aplicados actualmente en el campo.|
||[isFiltered(filterType?: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|Comprueba si hay filtros aplicados en el campo.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|El filtro de fecha aplicado actualmente del campo dinámico.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|El filtro de etiquetas aplicado actualmente del Campo dinámico.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|El filtro manual del campo dinámico aplicado actualmente.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|El filtro de valores aplicado actualmente del campo dinámico.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[comparador](/javascript/api/excel/excel.pivotlabelfilter#comparator)|El comparador es el valor estático con el que se comparan otros valores.|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#condition)|Especifica la condición del filtro, que define los criterios de filtrado necesarios.|
||[exclusive](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|Si `true` , filter excluye los *elementos* que cumplen los criterios.|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|Límite inferior del intervalo para la `between` condición de filtro.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|La subcadena usada para `beginsWith` , y las condiciones de `endsWith` `contains` filtro.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|Límite superior del intervalo para la `between` condición de filtro.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Una lista de elementos seleccionados para filtrar manualmente.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Especifica si la tabla dinámica permite la aplicación de varios filtros dinámicos en un campo dinámico determinado de la tabla.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Obtiene el número de tablas dinámicas de la colección.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Obtiene la primera tabla dinámica de la colección.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Obtiene una tabla dinámica por nombre.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Obtiene una tabla dinámica por nombre.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparador](/javascript/api/excel/excel.pivotvaluefilter#comparator)|El comparador es el valor estático con el que se comparan otros valores.|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|Especifica la condición del filtro, que define los criterios de filtrado necesarios.|
||[exclusive](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Si `true` , filter excluye los *elementos* que cumplen los criterios.|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|Límite inferior del intervalo para la `between` condición de filtro.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Especifica si el filtro es para los elementos N superior/inferior, el porcentaje N superior/inferior o la suma N superior/inferior.|
||[umbral](/javascript/api/excel/excel.pivotvaluefilter#threshold)|Número de umbral "N" de elementos, porcentaje o suma que se filtrarán para una condición de filtro superior/inferior.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|Límite superior del intervalo para la `between` condición de filtro.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Nombre del "valor" elegido en el campo por el que se va a filtrar.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getdirectprecedents--)|Devuelve un objeto que representa el rango que contiene todos los precedentes directos de una celda en la misma hoja de cálculo o `WorkbookRangeAreas` en varias hojas de cálculo.|
||[getPivotTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Obtiene una colección con ámbito de tablas dinámicas que se superponen con el intervalo.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtiene el intervalo que contiene la celda de delimitador de una celda en la que se derrama.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtiene el objeto de intervalo que contiene la celda delimitadora de la celda en la que se va a desbordar.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtiene el objeto de intervalo que contiene el intervalo de desbordamiento al llamar en una celda de delimitador.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtiene el objeto de intervalo que contiene el intervalo de desbordamiento al llamar en una celda de delimitador.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Indica si todas las celdas tienen un borde de desbordamiento.|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|Representa la categoría del formato de número de cada celda.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Representa si todas las celdas se guardarían como una fórmula de matriz.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|Obtiene el número de `RangeAreas` objetos de esta colección.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|Devuelve el `RangeAreas` objeto en función de la posición de la colección.|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|Devuelve el `RangeAreas` objeto en función del identificador o el nombre de la hoja de cálculo de la colección.|
||[getRangeAreasOrNullObjectBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|Devuelve el `RangeAreas` objeto basado en el nombre de la hoja de cálculo o el identificador de la colección.|
||[direcciones](/javascript/api/excel/excel.workbookrangeareas#addresses)|Devuelve una matriz de direcciones al estilo A1.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|Devuelve el `RangeAreasCollection` objeto.|
||[intervalos](/javascript/api/excel/excel.workbookrangeareas#ranges)|Devuelve intervalos que componen este objeto en un `RangeCollection` objeto.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Obtiene una colección de propiedades personalizadas a nivel de hoja de cálculo.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Elimina la propiedad personalizada.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Obtiene la clave de la propiedad personalizada.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Obtiene o establece el valor de la propiedad personalizada.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Agrega una nueva propiedad personalizada que se asigna a la clave proporcionada.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Obtiene el número de propiedades personalizadas de esta hoja de cálculo.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Obtiene un objeto de propiedad personalizada mediante su clave, que no distingue mayúsculas de minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Obtiene un objeto de propiedad personalizada mediante su clave, que no distingue mayúsculas de minúsculas.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
