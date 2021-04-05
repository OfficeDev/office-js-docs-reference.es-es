| Clase | Campos | Descripción |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getcount--)|Obtiene el número de enlaces de la colección.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#getitemornullobject-id-)|Obtiene un objeto de enlace por identificador.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getcount--)|Devuelve el número de gráficos de la hoja de cálculo.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#getitemornullobject-name-)|Obtiene un gráfico mediante su nombre.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getcount--)|Devuelve el número de puntos del gráfico de la serie.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getcount--)|Devuelve el número de series incluidas en la colección.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#comment)|Especifica el comentario asociado con este nombre.|
||[delete()](/javascript/api/excel/excel.nameditem#delete--)|Elimina el nombre especificado.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getrangeornullobject--)|Devuelve el objeto de rango asociado al nombre.|
||[ámbito](/javascript/api/excel/excel.nameditem#scope)|Especifica si el nombre está en el ámbito del libro o de una hoja de cálculo específica.|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|Devuelve la hoja de cálculo que tiene como ámbito el elemento con nombre.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetornullobject)|Devuelve la hoja de cálculo a la que está ámbito el elemento con nombre.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add(name: string, reference: Range \| string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#add-name--reference--comment-)|Agrega un nuevo nombre a la colección del ámbito especificado.|
||[addFormulaLocal(name: string, formula: string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#addformulalocal-name--formula--comment-)|Agrega un nuevo nombre a la colección del ámbito especificado, empleando la configuración regional del usuario para la fórmula.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getcount--)|Obtiene el número de elementos con nombre de la colección.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#getitemornullobject-name-)|Obtiene un `NamedItem` objeto con su nombre.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getcount--)|Obtiene el número de tablas dinámicas de una colección.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#getitemornullobject-name-)|Obtiene una tabla dinámica por nombre.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getintersectionornullobject-anotherrange-)|Obtiene el objeto de rango que representa la intersección rectangular de los rangos especificados.|
||[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.range#getusedrangeornullobject-valuesonly-)|Devuelve el intervalo usado del objeto de rango especificado.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getcount--)|Obtiene el número de `RangeView` objetos de la colección.|
|[Valor](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete--)|Elimina la configuración.|
||[key](/javascript/api/excel/excel.setting#key)|La clave que representa el identificador de la configuración.|
||[value](/javascript/api/excel/excel.setting#value)|Representa el valor almacenado para esta configuración.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add(key: string, value: string \| number boolean Date Array \| \| \| <any> \| any)](/javascript/api/excel/excel.settingcollection#add-key--value-)|Establece o agrega la configuración especificada en el libro.|
||[getCount()](/javascript/api/excel/excel.settingcollection#getcount--)|Obtiene el número de opciones de configuración de la colección.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getitem-key-)|Obtiene una entrada de configuración a través de la clave.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getitemornullobject-key-)|Obtiene una entrada de configuración a través de la clave.|
||[items](/javascript/api/excel/excel.settingcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onsettingschanged)|Se produce cuando se cambia la configuración del documento.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[settings](/javascript/api/excel/excel.settingschangedeventargs#settings)|Obtiene el `Setting` objeto que representa el enlace que presentó el evento de cambio de configuración|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getcount--)|Obtiene el número de tablas de la colección.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getitemornullobject-key-)|Obtener una tabla por nombre o identificador.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getcount--)|Obtiene el número de columnas de la tabla.|
||[getItemOrNullObject(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#getitemornullobject-key-)|Obtiene un objeto de columna por nombre o identificador.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getcount--)|Obtiene el número de filas de la tabla.|
|[Workbook](/javascript/api/excel/excel.workbook)|[settings](/javascript/api/excel/excel.workbook#settings)|Representa una colección de opciones asociadas al libro.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.worksheet#getusedrangeornullobject-valuesonly-)|El rango usado es el rango más pequeño que abarque todas las celdas que tengan asignado un valor o un formato.|
||[names](/javascript/api/excel/excel.worksheet#names)|Colección de nombres en el ámbito de la hoja de cálculo actual.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getcount-visibleonly-)|Obtiene el número de hojas de cálculo de la colección.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getitemornullobject-key-)|Obtiene un objeto de hoja de cálculo mediante su nombre o identificador.|
