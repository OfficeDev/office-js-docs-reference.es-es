| Clase | Campos | Descripción |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculate(calculationType: Excel.CalculationType)](/javascript/api/excel/excel.application#calculate-calculationtype-)|Recalcula todos los libros abiertos actualmente en Excel.|
||[calculationMode](/javascript/api/excel/excel.application#calculationmode)|Devuelve el modo de cálculo usado en el libro, tal como lo definen las constantes de `Excel.CalculationMode` .|
|[Enlace](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getrange--)|Devuelve el intervalo representado por el enlace.|
||[getTable()](/javascript/api/excel/excel.binding#gettable--)|Devuelve la tabla representada por el enlace.|
||[getText()](/javascript/api/excel/excel.binding#gettext--)|Devuelve el texto representado por el enlace.|
||[id](/javascript/api/excel/excel.binding#id)|Representa el identificador de enlace.|
||[type](/javascript/api/excel/excel.binding#type)|Devuelve el tipo de enlace.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getitem-id-)|Obtiene un objeto de enlace por identificador.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getitemat-index-)|Obtiene un objeto de enlace según su posición en la matriz de elementos.|
||[count](/javascript/api/excel/excel.bindingcollection#count)|Devuelve el número de enlaces incluidos en la colección.|
||[items](/javascript/api/excel/excel.bindingcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete--)|Elimina el objeto de gráfico.|
||[height](/javascript/api/excel/excel.chart#height)|Especifica el alto, en puntos, del objeto de gráfico.|
||[left](/javascript/api/excel/excel.chart#left)|La distancia, en puntos, desde el lado izquierdo del gráfico hasta el origen de la hoja de cálculo.|
||[name](/javascript/api/excel/excel.chart#name)|Especifica el nombre de un objeto de gráfico.|
||[ejes](/javascript/api/excel/excel.chart#axes)|Representa los ejes del gráfico.|
||[dataLabels](/javascript/api/excel/excel.chart#datalabels)|Representa la clase DataLabels del gráfico.|
||[format](/javascript/api/excel/excel.chart#format)|Encapsula las propiedades de formato del área del gráfico.|
||[leyenda](/javascript/api/excel/excel.chart#legend)|Representa la leyenda del gráfico.|
||[series](/javascript/api/excel/excel.chart#series)|Representa una sola serie o una colección de series del gráfico.|
||[title](/javascript/api/excel/excel.chart#title)|Representa el título del gráfico especificado, incluido el texto, la visibilidad, la posición y el formato del título.|
||[setData(sourceData: Range, seriesBy?: Excel.ChartSeriesBy)](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|Configura los datos de origen para el gráfico.|
||[setPosition(startCell: Range \| string, endCell?: Range \| string)](/javascript/api/excel/excel.chart#setposition-startcell--endcell-)|Coloca el gráfico con respecto a las celdas de la hoja de cálculo.|
||[top](/javascript/api/excel/excel.chart#top)|Especifica la distancia, en puntos, desde el borde superior del objeto hasta la parte superior de la fila 1 (en una hoja de cálculo) o la parte superior del área del gráfico (en un gráfico).|
||[width](/javascript/api/excel/excel.chart#width)|Especifica el ancho, en puntos, del objeto de gráfico.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|Representa el formato de relleno de un objeto, que incluye información del formato de fondo.|
||[font](/javascript/api/excel/excel.chartareaformat#font)|Representa los atributos de fuente (nombre de fuente, tamaño de fuente, color, etc.) del objeto actual.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryaxis)|Representa el eje de categorías de un gráfico.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#seriesaxis)|Representa el eje de serie de un gráfico 3D.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#valueaxis)|Representa el eje de valores de un eje.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorunit)|Representa el intervalo entre dos marcas de graduación principales.|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|Representa el valor máximo del eje de valores.|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|Representa el valor mínimo del eje de valores.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorunit)|Representa el rango entre dos marcas de graduación secundarias.|
||[format](/javascript/api/excel/excel.chartaxis#format)|Representa el formato de un objeto de gráfico, que incluye el formato de línea y de fuente.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorgridlines)|Devuelve un objeto que representa las líneas de cuadrícula principales del eje especificado.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorgridlines)|Devuelve un objeto que representa las líneas de cuadrícula secundarias del eje especificado.|
||[title](/javascript/api/excel/excel.chartaxis#title)|Representa el título del eje.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|Especifica los atributos de fuente (nombre de fuente, tamaño de fuente, color, etc.) de un elemento de eje de gráfico.|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|Especifica el formato de línea del gráfico.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|Especifica el formato del título del eje del gráfico.|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|Especifica el título del eje.|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|Especifica si el título del eje es visibile.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|Especifica los atributos de fuente del título del eje del gráfico, como el nombre de fuente, el tamaño de fuente o el color, del objeto de título del eje del gráfico.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add(type: Excel.ChartType, sourceData: Range, seriesBy?: Excel.ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|Crea un nuevo gráfico.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getitem-name-)|Obtiene un gráfico mediante su nombre.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getitemat-index-)|Obtiene un gráfico en función de su posición en la colección.|
||[count](/javascript/api/excel/excel.chartcollection#count)|Devuelve el número de gráficos de la hoja de cálculo.|
||[items](/javascript/api/excel/excel.chartcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|Representa el formato de relleno de la etiqueta de datos del gráfico actual.|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|Representa los atributos de fuente (como el nombre de fuente, el tamaño de fuente y el color) de una etiqueta de datos del gráfico.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|Valor que representa la posición de la etiqueta de datos.|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|Especifica el formato de las etiquetas de datos del gráfico, que incluye el formato de relleno y fuente.|
||[separador](/javascript/api/excel/excel.chartdatalabels#separator)|Cadena que representa el separador empleado para las etiquetas de datos de un gráfico.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showbubblesize)|Especifica si el tamaño de burbuja de la etiqueta de datos está visible.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showcategoryname)|Especifica si el nombre de categoría de la etiqueta de datos está visible.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showlegendkey)|Especifica si la clave de leyenda de la etiqueta de datos está visible.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showpercentage)|Especifica si el porcentaje de etiqueta de datos está visible.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showseriesname)|Especifica si el nombre de la serie de etiquetas de datos está visible.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showvalue)|Especifica si el valor de la etiqueta de datos está visible.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear--)|Borra el color de relleno de un elemento de gráfico.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setsolidcolor-color-)|Establece el formato de relleno de un elemento de gráfico en un color uniforme.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|Indica el estado de negrita de la fuente.|
||[color](/javascript/api/excel/excel.chartfont#color)|Representación del código de color HTML del color del texto (por ejemplo, #FF0000 representa El rojo).|
||[italic](/javascript/api/excel/excel.chartfont#italic)|Representa el estado de cursiva de la fuente.|
||[name](/javascript/api/excel/excel.chartfont#name)|Nombre de fuente (por ejemplo, "Calibri")|
||[size](/javascript/api/excel/excel.chartfont#size)|Tamaño de la fuente (por ejemplo, 11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|Tipo de subrayado aplicado a la fuente.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|Representa el formato de las líneas de cuadrícula del gráfico.|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|Especifica si las líneas de cuadrícula del eje están visibles.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|Indica el formato de línea de gráfico.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[overlay](/javascript/api/excel/excel.chartlegend#overlay)|Especifica si la leyenda del gráfico debe superponerse con el cuerpo principal del gráfico.|
||[position](/javascript/api/excel/excel.chartlegend#position)|Especifica la posición de la leyenda en el gráfico.|
||[format](/javascript/api/excel/excel.chartlegend#format)|Representa el formato de una leyenda del gráfico, que incluye el formato de relleno y de fuente.|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|Especifica si la leyenda del gráfico está visible.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|Representa el formato de relleno de un objeto, que incluye información del formato de fondo.|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|Representa los atributos de fuente, como el nombre de fuente, el tamaño de fuente y el color de una leyenda del gráfico.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear--)|Borra el formato de línea de un elemento de gráfico.|
||[color](/javascript/api/excel/excel.chartlineformat#color)|Código de color HTML que representa el color de las líneas del gráfico.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|Encapsula las propiedades de formato del punto del gráfico.|
||[value](/javascript/api/excel/excel.chartpoint#value)|Devuelve el valor de un punto del gráfico.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|Representa el formato de relleno de un gráfico, que incluye información de formato de fondo.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getitemat-index-)|Recupera un punto basándose en su posición dentro de la serie.|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|Devuelve el número de puntos del gráfico de la serie.|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|Especifica el nombre de una serie en un gráfico.|
||[format](/javascript/api/excel/excel.chartseries#format)|Representa el formato de una serie del gráfico, que incluye el formato de relleno y de línea.|
||[puntos](/javascript/api/excel/excel.chartseries#points)|Devuelve una colección de todos los puntos de la serie.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getitemat-index-)|Recupera una serie basada en su posición en la colección.|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|Devuelve el número de series incluidas en la colección.|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|Representa el formato de relleno de una serie del gráfico, que incluye información del formato de fondo.|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|Representa el formato de línea.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[overlay](/javascript/api/excel/excel.charttitle#overlay)|Especifica si el título del gráfico se superponerá al gráfico.|
||[format](/javascript/api/excel/excel.charttitle#format)|Representa el formato de un título del gráfico, que incluye el formato de relleno y de fuente.|
||[text](/javascript/api/excel/excel.charttitle#text)|Especifica el texto del título del gráfico.|
||[visible](/javascript/api/excel/excel.charttitle#visible)|Especifica si el título del gráfico es visibile.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|Representa el formato de relleno de un objeto, que incluye información del formato de fondo.|
||[font](/javascript/api/excel/excel.charttitleformat#font)|Representa los atributos de fuente (como el nombre de fuente, el tamaño de fuente y el color) de un objeto.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getrange--)|Devuelve el objeto de rango asociado al nombre.|
||[name](/javascript/api/excel/excel.nameditem#name)|Nombre del objeto.|
||[type](/javascript/api/excel/excel.nameditem#type)|Especifica el tipo del valor devuelto por la fórmula del nombre.|
||[value](/javascript/api/excel/excel.nameditem#value)|Representa el valor calculado por la fórmula del nombre.|
||[visible](/javascript/api/excel/excel.nameditem#visible)|Especifica si el objeto está visible.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getitem-name-)|Obtiene un `NamedItem` objeto con su nombre.|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear-applyto-)|Borra valores de rango, formato, relleno, borde, etc.|
||[delete(shift: Excel.DeleteShiftDirection)](/javascript/api/excel/excel.range#delete-shift-)|Elimina las celdas asociadas al rango.|
||[formulas](/javascript/api/excel/excel.range#formulas)|Representa la fórmula en notación de estilo A1.|
||[formulasLocal](/javascript/api/excel/excel.range#formulaslocal)|Representa la fórmula en notación de estilo A1, en el idioma del usuario y en la configuración regional del formato numérico.|
||[getBoundingRect(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getboundingrect-anotherrange-)|Obtiene el objeto de intervalo más pequeño que abarca los intervalos especificados.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getcell-row--column-)|Obtiene el objeto de intervalo que contiene la celda en función de los números de fila y columna.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getcolumn-column-)|Obtiene una columna contenida en el intervalo.|
||[getEntireColumn()](/javascript/api/excel/excel.range#getentirecolumn--)|Obtiene un objeto que representa toda la columna del rango (por ejemplo, si el rango actual representa celdas "B4:E11", es un rango que representa las columnas `getEntireColumn` "B:E").|
||[getEntireRow()](/javascript/api/excel/excel.range#getentirerow--)|Obtiene un objeto que representa toda la fila del rango (por ejemplo, si el rango actual representa celdas "B4:E11", es un rango que representa las filas `GetEntireRow` "4:11").|
||[getIntersection(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getintersection-anotherrange-)|Obtiene el objeto de intervalo que representa la intersección rectangular de los intervalos especificados.|
||[getLastCell()](/javascript/api/excel/excel.range#getlastcell--)|Obtiene la última celda del intervalo.|
||[getLastColumn()](/javascript/api/excel/excel.range#getlastcolumn--)|Obtiene la última columna del intervalo.|
||[getLastRow()](/javascript/api/excel/excel.range#getlastrow--)|Obtiene la última fila del intervalo.|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-)|Obtiene un objeto que representa un intervalo desplazado con respecto al intervalo especificado.|
||[getRow(row: number)](/javascript/api/excel/excel.range#getrow-row-)|Obtiene una fila contenida en el intervalo.|
||[insert(shift: Excel.InsertShiftDirection)](/javascript/api/excel/excel.range#insert-shift-)|Inserta una celda o un intervalo de celdas en la hoja de cálculo en lugar de este intervalo y desplaza las demás celdas para crear espacio.|
||[numberFormat](/javascript/api/excel/excel.range#numberformat)|Representa el código de formato de número de Excel para el intervalo especificado.|
||[address](/javascript/api/excel/excel.range#address)|Especifica la referencia de intervalo en estilo A1.|
||[addressLocal](/javascript/api/excel/excel.range#addresslocal)|Representa la referencia de intervalo para el intervalo especificado en el idioma del usuario.|
||[cellCount](/javascript/api/excel/excel.range#cellcount)|Especifica el número de celdas del rango.|
||[columnCount](/javascript/api/excel/excel.range#columncount)|Especifica el número total de columnas del intervalo.|
||[columnIndex](/javascript/api/excel/excel.range#columnindex)|Especifica el número de columna de la primera celda del rango.|
||[format](/javascript/api/excel/excel.range#format)|Devuelve un objeto de formato que encapsula la fuente, el relleno, los bordes, la alineación y otras propiedades del rango.|
||[rowCount](/javascript/api/excel/excel.range#rowcount)|Devuelve el número total de filas del intervalo.|
||[rowIndex](/javascript/api/excel/excel.range#rowindex)|Devuelve el número de fila de la primera celda del intervalo.|
||[text](/javascript/api/excel/excel.range#text)|Valores de texto del rango especificado.|
||[valueTypes](/javascript/api/excel/excel.range#valuetypes)|Especifica el tipo de datos de cada celda.|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|Hoja de cálculo que contiene el rango actual.|
||[select()](/javascript/api/excel/excel.range#select--)|Selecciona el intervalo especificado en la interfaz de usuario de Excel.|
||[values](/javascript/api/excel/excel.range#values)|Representa los valores sin formato del rango especificado.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|Código de color HTML que representa el color de la línea de borde con el formato #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideindex)|Valor constante que indica el lado específico del borde.|
||[estilo](/javascript/api/excel/excel.rangeborder#style)|Una de las constantes de estilo de línea que especifica el estilo de línea del borde.|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|Especifica el grosor del borde que rodea un rango.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem(index: Excel.BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|Obtiene un objeto de borde mediante su nombre.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getitemat-index-)|Obtiene un objeto de borde mediante su índice.|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|Número de objetos de borde de la colección.|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear--)|Restablece el fondo del intervalo.|
||[color](/javascript/api/excel/excel.rangefill#color)|Código de color HTML que representa el color del fondo, con el formato #RRGGBB (por ejemplo, "FFA500"), o como un color HTML con nombre (por ejemplo, "naranja")|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|Representa el estado en negrita de la fuente.|
||[color](/javascript/api/excel/excel.rangefont#color)|Representación del código de color HTML del color del texto (por ejemplo, #FF0000 representa El rojo).|
||[italic](/javascript/api/excel/excel.rangefont#italic)|Especifica el estado de cursiva de la fuente.|
||[name](/javascript/api/excel/excel.rangefont#name)|Nombre de fuente (por ejemplo, "Calibri").|
||[size](/javascript/api/excel/excel.rangefont#size)|Tamaño de fuente.|
||[underline](/javascript/api/excel/excel.rangefont#underline)|Tipo de subrayado aplicado a la fuente.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalalignment)|Representa la alineación horizontal del objeto especificado.|
||[borders](/javascript/api/excel/excel.rangeformat#borders)|Colección de objetos de borde que se aplica al rango global seleccionado.|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|Devuelve el objeto de relleno definido en el intervalo global.|
||[font](/javascript/api/excel/excel.rangeformat#font)|Devuelve el objeto de fuente definido en el rango global.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalalignment)|Representa la alineación vertical del objeto especificado.|
||[wrapText](/javascript/api/excel/excel.rangeformat#wraptext)|Especifica si Excel ajusta el texto del objeto.|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete--)|Elimina la tabla.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#getdatabodyrange--)|Obtiene el objeto de rango asociado al cuerpo de datos de la tabla.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getheaderrowrange--)|Obtiene el objeto de intervalo asociado a la fila de encabezado de la tabla.|
||[getRange()](/javascript/api/excel/excel.table#getrange--)|Obtiene el objeto de rango asociado a toda la tabla.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#gettotalrowrange--)|Obtiene el objeto de intervalo asociado a la fila de totales de la tabla.|
||[name](/javascript/api/excel/excel.table#name)|Nombre de la tabla.|
||[columns](/javascript/api/excel/excel.table#columns)|Representa una colección de todas las columnas de la tabla.|
||[id](/javascript/api/excel/excel.table#id)|Devuelve un valor que identifica de forma única la tabla de un libro determinado.|
||[rows](/javascript/api/excel/excel.table#rows)|Representa una colección de todas las filas de la tabla.|
||[showHeaders](/javascript/api/excel/excel.table#showheaders)|Especifica si la fila de encabezado está visible.|
||[showTotals](/javascript/api/excel/excel.table#showtotals)|Especifica si la fila total está visible.|
||[estilo](/javascript/api/excel/excel.table#style)|Valor constante que representa el estilo de tabla.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add(address: Range \| string, hasHeaders: boolean)](/javascript/api/excel/excel.tablecollection#add-address--hasheaders-)|Crear una tabla.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getitem-key-)|Obtener una tabla por nombre o identificador.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getitemat-index-)|Obtiene una tabla basada en su posición en la colección.|
||[count](/javascript/api/excel/excel.tablecollection#count)|Devuelve el número de tablas del libro.|
||[items](/javascript/api/excel/excel.tablecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete--)|Elimina la columna de la tabla.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#getdatabodyrange--)|Obtiene el objeto de intervalo asociado al cuerpo de datos de la columna.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getheaderrowrange--)|Obtiene el objeto de intervalo asociado a la fila de encabezado de la columna.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getrange--)|Obtiene el objeto de intervalo asociado a toda la columna.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#gettotalrowrange--)|Obtiene el objeto de rango asociado a la fila de totales de la columna.|
||[name](/javascript/api/excel/excel.tablecolumn#name)|Especifica el nombre de la columna de tabla.|
||[id](/javascript/api/excel/excel.tablecolumn#id)|Devuelve una clave única que identifica la columna de la tabla.|
||[index](/javascript/api/excel/excel.tablecolumn#index)|Devuelve el número de índice de la columna dentro de la colección de columnas de la tabla.|
||[values](/javascript/api/excel/excel.tablecolumn#values)|Representa los valores sin formato del rango especificado.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add(index?: number, values?: Array<Array<boolean \| string number>> boolean string \| \| \| \| number, name?: string)](/javascript/api/excel/excel.tablecolumncollection#add-index--values--name-)|Agrega una nueva columna a la tabla.|
||[getItem(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#getitem-key-)|Obtiene un objeto de columna por nombre o identificador.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getitemat-index-)|Obtiene una columna en función de su posición en la colección.|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|Devuelve el número de columnas de la tabla.|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete--)|Elimina la fila de la tabla.|
||[getRange()](/javascript/api/excel/excel.tablerow#getrange--)|Devuelve el objeto de rango asociado a toda la fila.|
||[index](/javascript/api/excel/excel.tablerow#index)|Devuelve el número de índice de la fila dentro de la colección de filas de la tabla.|
||[values](/javascript/api/excel/excel.tablerow#values)|Representa los valores sin formato del rango especificado.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add(index?: number, values?: Array<Array<boolean \| string number>> boolean string \| \| \| \| number)](/javascript/api/excel/excel.tablerowcollection#add-index--values-)|Agrega una o más filas a la tabla.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getitemat-index-)|Obtiene una fila en función de su posición en la colección.|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|Devuelve el número de filas de la tabla.|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange()](/javascript/api/excel/excel.workbook#getselectedrange--)|Obtiene el rango único seleccionado actualmente del libro.|
||[application](/javascript/api/excel/excel.workbook#application)|Representa la instancia de aplicación de Excel que contiene este libro.|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|Representa una colección de enlaces que forman parte del libro.|
||[names](/javascript/api/excel/excel.workbook#names)|Representa una colección de elementos con nombre con ámbito de libro (intervalos con nombre y constantes).|
||[tablas](/javascript/api/excel/excel.workbook#tables)|Representa una colección de tablas asociadas con el libro.|
||[hojas de cálculo](/javascript/api/excel/excel.workbook#worksheets)|Representa una colección de hojas de cálculo asociadas con el libro.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate--)|Activa la hoja de cálculo en la interfaz de usuario de Excel.|
||[delete()](/javascript/api/excel/excel.worksheet#delete--)|Elimina la hoja de cálculo del libro.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getcell-row--column-)|Obtiene el `Range` objeto que contiene la celda única en función de los números de fila y columna.|
||[getRange(address?: string)](/javascript/api/excel/excel.worksheet#getrange-address-)|Obtiene el `Range` objeto, que representa un único bloque rectangular de celdas, especificado por la dirección o el nombre.|
||[name](/javascript/api/excel/excel.worksheet#name)|Nombre para mostrar de la hoja de cálculo.|
||[position](/javascript/api/excel/excel.worksheet#position)|Posición de base cero de la hoja de cálculo dentro del libro.|
||[gráficos](/javascript/api/excel/excel.worksheet#charts)|Devuelve una colección de gráficos que forman parte de la hoja de cálculo.|
||[id](/javascript/api/excel/excel.worksheet#id)|Devuelve un valor que identifica de forma única la hoja de cálculo de un libro determinado.|
||[tablas](/javascript/api/excel/excel.worksheet#tables)|Colección de tablas que forman parte de la hoja de cálculo.|
||[visibility](/javascript/api/excel/excel.worksheet#visibility)|Visibilidad de la hoja de cálculo.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add(name?: string)](/javascript/api/excel/excel.worksheetcollection#add-name-)|Agrega una nueva hoja al libro.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#getactiveworksheet--)|Obtiene la hoja de cálculo activa del libro.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getitem-key-)|Obtiene un objeto de hoja de cálculo mediante su nombre o identificador.|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
