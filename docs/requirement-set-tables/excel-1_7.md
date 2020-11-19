| Class | Campos | Descripción |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|Especifica el tipo de gráfico.|
||[id](/javascript/api/excel/excel.chart#id)|Identificador único del gráfico.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showallfieldbuttons)|Especifica si se van a mostrar todos los botones de campo en un gráfico dinámico.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[border](/javascript/api/excel/excel.chartareaformat#border)|Representa el formato de borde del área del gráfico, que incluye color, LineStyle y weight.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem (tipo: Excel. ChartAxisType, Group?: Excel. ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Devuelve el eje específico identificado por tipo y grupo.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#basetimeunit)|Especifica la unidad base del eje de categorías especificado.|
||[categoryType](/javascript/api/excel/excel.chartaxis#categorytype)|Especifica el tipo de eje de categorías.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayunit)|Representa la unidad de visualización del eje.|
||[logBase](/javascript/api/excel/excel.chartaxis#logbase)|Especifica la base del logaritmo al usar escalas logarítmicas.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majortickmark)|Especifica el tipo de marca de graduación principal del eje especificado.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majortimeunitscale)|Especifica el valor de la escala de unidades principales del eje de categorías cuando la propiedad CategoryType está establecida en escala temporal.|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minortickmark)|Especifica el tipo de marca de graduación secundaria del eje especificado.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minortimeunitscale)|Especifica el valor de la escala de unidades secundaria del eje de categorías cuando la propiedad CategoryType está establecida en escala temporal.|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisgroup)|Especifica el grupo del eje especificado.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customdisplayunit)|Especifica el valor de la unidad de visualización del eje personalizado.|
||[height](/javascript/api/excel/excel.chartaxis#height)|Especifica el alto, en puntos, del eje del gráfico.|
||[left](/javascript/api/excel/excel.chartaxis#left)|Especifica la distancia, en puntos, desde el borde izquierdo del eje a la izquierda del área del gráfico.|
||[top](/javascript/api/excel/excel.chartaxis#top)|Especifica la distancia, en puntos, desde el borde superior del eje hasta la parte superior del área del gráfico.|
||[type](/javascript/api/excel/excel.chartaxis#type)|Especifica el tipo de eje.|
||[width](/javascript/api/excel/excel.chartaxis#width)|Especifica el ancho, en puntos, del eje del gráfico.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reverseplotorder)|Especifica si Excel traza los puntos de datos desde el último hasta el primero.|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaletype)|Especifica el tipo de escala del eje de valores.|
||[setCategoryNames (sourceData: Range)](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|Establece todos los nombres de categoría del eje especificado.|
||[setCustomDisplayUnit (Value: Number)](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|Establece la unidad de visualización de ejes en un valor personalizado.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|Especifica si la etiqueta de la unidad de visualización del eje es visible.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|Especifica la posición de los rótulos de marcas de graduación en el eje especificado.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|Especifica el número de categorías o series entre rótulos de marcas de graduación.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|Especifica el número de categorías o series entre las marcas de graduación.|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|Especifica si el eje es visible.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|Código de color HTML que representa el color de los bordes del gráfico.|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|Representa el estilo de línea del borde.|
||[weight](/javascript/api/excel/excel.chartborder#weight)|Representa el grosor del borde, en puntos.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|Valor DataLabelPosition que representa la posición de la etiqueta de datos.|
||[separador](/javascript/api/excel/excel.chartdatalabel#separator)|Cadena que representa el separador empleado para la etiqueta de datos de un gráfico.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|Especifica si el tamaño de la burbuja de la etiqueta de datos es visible.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|Especifica si el nombre de categoría de la etiqueta de datos es visible.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|Especifica si la clave de la leyenda de la etiqueta de datos está visible.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|Especifica si el porcentaje de la etiqueta de datos es visible.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|Especifica si el nombre de serie de la etiqueta de datos es visible.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|Especifica si el valor de la etiqueta de datos es visible.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|Representa los atributos de fuente, como el nombre de fuente, el tamaño de fuente, el color, etc.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Especifica el alto, en puntos, de la leyenda en el gráfico.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Especifica la izquierda, en puntos, de la leyenda en el gráfico.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|Representa una colección de legendEntries en la leyenda.|
||[showShadow](/javascript/api/excel/excel.chartlegend#showshadow)|Especifica si la leyenda tiene una sombra en el gráfico.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Especifica la parte superior de la leyenda de un gráfico.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Especifica el ancho, en puntos, de la leyenda en el gráfico.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Representa la parte visible de la entrada de la leyenda de un gráfico.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|Devuelve el número de legendEntry de la colección.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|Devuelve una legendEntry en el índice especificado.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|Representa el estilo de línea.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Indica el grosor de la línea, en puntos.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|Representa si un punto de datos tiene una etiqueta de datos.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|Representación del código de color HTML del color de fondo del marcador del punto de datos (por ejemplo, #FF0000 representa el rojo).|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|Representación del código de color HTML del color de primer plano del marcador del punto de datos (por ejemplo, #FF0000 representa el rojo).|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|Representa el tamaño del marcador de punto de datos.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerstyle)|Representa el estilo de marcador de un punto de datos del gráfico.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|Devuelve la etiqueta de datos de un punto del gráfico.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[border](/javascript/api/excel/excel.chartpointformat#border)|Representa el formato de borde de un punto de datos del gráfico, que incluye información sobre el color, el estilo y el grosor.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#charttype)|Representa el tipo de gráfico de una serie.|
||[delete()](/javascript/api/excel/excel.chartseries#delete--)|Elimina la serie del gráfico.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutholesize)|Representa el tamaño del agujero de los anillos de una serie de gráfico.|
||[filtrado](/javascript/api/excel/excel.chartseries#filtered)|Especifica si se ha filtrado la serie.|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapwidth)|Representa el ancho del rango de una serie de gráfico.|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasdatalabels)|Especifica si la serie tiene etiquetas de datos.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerbackgroundcolor)|Especifica el color de fondo de los marcadores de una serie de gráficos.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerforegroundcolor)|Especifica el color de primer plano de los marcadores de una serie del gráfico.|
||[markerSize](/javascript/api/excel/excel.chartseries#markersize)|Especifica el tamaño del marcador de una serie de gráficos.|
||[markerStyle](/javascript/api/excel/excel.chartseries#markerstyle)|Especifica el estilo de marcador de una serie de gráficos.|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotorder)|Especifica el orden de trazado de una serie de gráficos dentro del grupo de gráficos.|
||[Trendlines](/javascript/api/excel/excel.chartseries#trendlines)|La colección de líneas de tendencia de la serie.|
||[setBubbleSizes (sourceData: Range)](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|Establece el tamaño de las burbujas de una serie de gráficos.|
||[setValues (sourceData: Range)](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|Establece los valores de una serie de gráficos.|
||[setXAxisValues (sourceData: Range)](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|Establece los valores del eje X para una serie de gráficos.|
||[showShadow](/javascript/api/excel/excel.chartseries#showshadow)|Especifica si la serie tiene una sombra.|
||[eficaz](/javascript/api/excel/excel.chartseries#smooth)|Especifica si la serie es fluida.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[Add (Name?: String, index?: Number)](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|Agrega una nueva serie a la colección.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring (Inicio: número, longitud: número)](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|Obtener la subcadena de un título de gráfico.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|Especifica la alineación horizontal del título del gráfico.|
||[left](/javascript/api/excel/excel.charttitle#left)|Especifica la distancia, en puntos, desde el borde izquierdo del título del gráfico hasta el borde izquierdo del área del gráfico.|
||[position](/javascript/api/excel/excel.charttitle#position)|Representa la posición del título del gráfico.|
||[height](/javascript/api/excel/excel.charttitle#height)|Devuelve el alto, en puntos, del título del gráfico.|
||[width](/javascript/api/excel/excel.charttitle#width)|Especifica el ancho, en puntos, del título del gráfico.|
||[setFormula (Formula: String)](/javascript/api/excel/excel.charttitle#setformula-formula-)|Establece un valor de cadena que representa la fórmula del título del gráfico con la notación de estilo A1.|
||[showShadow](/javascript/api/excel/excel.charttitle#showshadow)|Representa un valor booleano que determina si el título del gráfico tiene una propiedad reemplazada.|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|Especifica el ángulo al que está orientado el texto para el título del gráfico.|
||[top](/javascript/api/excel/excel.charttitle#top)|Especifica la distancia, en puntos, desde el borde superior del título del gráfico hasta la parte superior del área del gráfico.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|Especifica la alineación vertical del título del gráfico.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[border](/javascript/api/excel/excel.charttitleformat#border)|Representa el formato de borde del título del gráfico, que incluye color, LineStyle y weight.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete--)|Elimina el objeto de la línea de tendencia.|
||[CEPT](/javascript/api/excel/excel.charttrendline#intercept)|Representa el valor de la intersección de la línea de tendencia.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|Representa el período de una tendencia de un gráfico.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Representa el nombre de la línea de tendencia.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|Representa el orden de una tendencia de gráfico.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Representa el formato de línea de tendencia de un gráfico.|
||[type](/javascript/api/excel/excel.charttrendline#type)|Representa el tipo de línea de tendencia de un gráfico.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[Add (Type?: Excel. ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Agrega una nueva línea de tendencia a la colección de líneas de tendencia.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|Devuelve el número de líneas de tendencia de la colección.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|Obtiene el objeto de línea de tendencia por el índice, que es el orden de introducción en la matriz de elementos.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Indica el formato de línea de gráfico.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|Elimina la propiedad personalizada.|
||[key](/javascript/api/excel/excel.customproperty#key)|Clave de la propiedad personalizada.|
||[type](/javascript/api/excel/excel.customproperty#type)|Tipo del valor que se usa para la propiedad personalizada.|
||[value](/javascript/api/excel/excel.customproperty#value)|Valor de la propiedad personalizada.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[Add (Key: String, Value: any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|Crea una nueva propiedad personalizada o establece una existente.|
||[deleteAll ()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|Elimina todas las propiedades personalizadas de la colección.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|Obtiene el recuento de las propiedades personalizadas.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|Obtiene un objeto de propiedad personalizada mediante su clave, que no distingue mayúsculas de minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|Obtiene un objeto de propiedad personalizada mediante su clave, que no distingue mayúsculas de minúsculas.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll ()](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|Actualiza todas las conexiones de datos de la colección.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[crea](/javascript/api/excel/excel.documentproperties#author)|El autor del libro.|
||[Categoría](/javascript/api/excel/excel.documentproperties#category)|Categoría del libro.|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|Comentarios del libro.|
||[company](/javascript/api/excel/excel.documentproperties#company)|La empresa del libro.|
||[Palabras clave](/javascript/api/excel/excel.documentproperties#keywords)|Las palabras clave del libro.|
||[manager](/javascript/api/excel/excel.documentproperties#manager)|El administrador del libro.|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|Obtiene la fecha de creación del libro.|
||[customiza](/javascript/api/excel/excel.documentproperties#custom)|Obtiene la colección de propiedades personalizadas del libro.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|Obtiene el último autor del libro.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|Obtiene el número de revisión del libro.|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|El asunto del libro.|
||[title](/javascript/api/excel/excel.documentproperties#title)|Título del libro.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|Fórmula del elemento con nombre.|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|Devuelve un objeto que contiene los valores y tipos del elemento con nombre.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Representa los tipos de cada elemento de la matriz de elementos con nombre.|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Representa los valores de cada elemento de la matriz de elementos con nombre.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange (numRows: Number, numColumns: Number)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|Obtiene un objeto Range con la misma celda superior izquierda que el objeto Range actual, pero con los números de filas y columnas especificados.|
||[getImage ()](/javascript/api/excel/excel.range#getimage--)|Representa el intervalo como una imagen PNG codificada en Base64.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getsurroundingregion--)|Devuelve un objeto Range que representa la región circundante para la celda superior izquierda de este intervalo.|
||[hyperlink](/javascript/api/excel/excel.range#hyperlink)|Representa el hipervínculo para el intervalo actual.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|Representa el código de formato numérico de Excel para el intervalo determinado, en función de la configuración de idioma del usuario.|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|Representa si el intervalo actual es una columna completa.|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|Representa si el intervalo actual es una fila completa.|
||[showCard ()](/javascript/api/excel/excel.range#showcard--)|Muestra la tarjeta de una celda activa si tiene contenido de valor enriquecido.|
||[estilo](/javascript/api/excel/excel.range#style)|Representa el estilo del rango actual.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|Orientación del texto de todas las celdas dentro del rango.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Determina si la altura de la fila del objeto Range es igual a la altura estándar de la hoja.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Especifica si el ancho de columna del objeto Range es igual al ancho estándar de la hoja.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|Representa la dirección url de destino del hipervínculo.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|Representa el destino de referencia de documento para el hipervínculo.|
||[Pantalla](/javascript/api/excel/excel.rangehyperlink#screentip)|Representa la cadena que se muestra al mantener el puntero sobre el hipervínculo.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|Representa la cadena que se muestra en la celda superior izquierda del rango.|
|[Estilo](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete--)|Elimina este estilo.|
||[formulaHidden](/javascript/api/excel/excel.style#formulahidden)|Especifica si la fórmula se ocultará cuando la hoja de cálculo esté protegida.|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalalignment)|Representa la alineación horizontal del estilo.|
||[includeAlignment](/javascript/api/excel/excel.style#includealignment)|Especifica si el estilo incluye las propiedades autodent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel y TextOrientation.|
||[includeBorder](/javascript/api/excel/excel.style#includeborder)|Especifica si el estilo incluye las propiedades color, ColorIndex, LineStyle y Weight Border.|
||[includeFont](/javascript/api/excel/excel.style#includefont)|Especifica si el estilo incluye las propiedades Background, Bold, color, ColorIndex, FontStyle, Italic, Name, size, StrikeThrough, subscript, Superscript y subrayado de la fuente.|
||[includeNumber](/javascript/api/excel/excel.style#includenumber)|Especifica si el estilo incluye la propiedad NumberFormat.|
||[includePatterns](/javascript/api/excel/excel.style#includepatterns)|Especifica si el estilo incluye las propiedades interiores color, ColorIndex, InvertIfNegative, Pattern, PatternColor y PatternColorIndex.|
||[includeProtection](/javascript/api/excel/excel.style#includeprotection)|Especifica si el estilo incluye las propiedades de protección FormulaHidden y locked.|
||[indentLevel](/javascript/api/excel/excel.style#indentlevel)|Un número entero entre 0 y 250 que indica el nivel de sangría para el estilo.|
||[locked](/javascript/api/excel/excel.style#locked)|Especifica si el objeto está bloqueado cuando la hoja de cálculo está protegida.|
||[numberFormat](/javascript/api/excel/excel.style#numberformat)|El código de formato del formato de número para el estilo.|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberformatlocal)|El código de formato localizado del formato de número para el estilo.|
||[readingOrder](/javascript/api/excel/excel.style#readingorder)|El orden de lectura para el estilo.|
||[borders](/javascript/api/excel/excel.style#borders)|Una colección de bordes de cuatro objetos Border que representan el estilo de los cuatro bordes.|
||[builtIn](/javascript/api/excel/excel.style#builtin)|Especifica si el estilo es un estilo integrado.|
||[fill](/javascript/api/excel/excel.style#fill)|El relleno del estilo.|
||[font](/javascript/api/excel/excel.style#font)|Un objeto de fuente que representa la fuente del estilo.|
||[name](/javascript/api/excel/excel.style#name)|El nombre del estilo.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|Especifica si el texto se reduce automáticamente para ajustarse al ancho de columna disponible.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|Especifica la alineación vertical del estilo.|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|Especifica si Excel ajusta el texto del objeto.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|Agrega un nuevo estilo a la colección.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|Obtiene un estilo por nombre.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|Se produce cuando cambian los datos de las celdas de una tabla concreta.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|Se produce cuando la selección cambia en una tabla concreta.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|Obtiene la dirección que representa el área de cambio de una tabla en una hoja de cálculo concreta.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Obtiene el tipo de cambio que representa cómo se activa el evento cambiado.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|Obtiene el origen del evento.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableid)|Obtiene el identificador de la tabla en la que se cambian los datos.|
||[type](/javascript/api/excel/excel.tablechangedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se cambian los datos.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onchanged)|Se produce cuando cambian los datos en una tabla de un libro o en una hoja de cálculo.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|Obtiene la dirección del intervalo que representa el área seleccionada de la tabla en una hoja de cálculo específica.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isinsidetable)|Especifica si la selección está dentro de una tabla, la dirección será inútil si IsInsideTable es false.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableid)|Obtiene el identificador de la tabla en la que se cambia la selección.|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se cambia la selección.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveCell()](/javascript/api/excel/excel.workbook#getactivecell--)|Obtiene la celda activa del libro.|
||[dataConnections](/javascript/api/excel/excel.workbook#dataconnections)|Representa todas las conexiones de datos en el libro.|
||[name](/javascript/api/excel/excel.workbook#name)|Obtiene el nombre del libro.|
||[properties](/javascript/api/excel/excel.workbook#properties)|Obtiene las propiedades del libro.|
||[protection](/javascript/api/excel/excel.workbook#protection)|Devuelve el objeto Protection de un libro.|
||[Stil](/javascript/api/excel/excel.workbook#styles)|Representa una colección de estilos asociada con el libro.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[Protect (contraseña?: String)](/javascript/api/excel/excel.workbookprotection#protect-password-)|Protege un libro.|
||[protegidas](/javascript/api/excel/excel.workbookprotection#protected)|Especifica si el libro está protegido.|
||[Unprotect (password?: String)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|Desprotege un libro.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Copy (positionType?: Excel. WorksheetPositionType, relativeto?: Excel. Worksheet)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Copia una hoja de cálculo y la coloca en la posición especificada.|
||[getRangeByIndexes (startRow: Number, Columnainicio: Number, rowCount: Number, columnCount: Number)](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|Obtiene el objeto de intervalo comenzando en un índice de columna y fila determinado, y que abarca un determinado número de filas y columnas.|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezepanes)|Obtiene un objeto que se puede usar para manipular paneles inmovilizados en la hoja de cálculo.|
||[onActivated](/javascript/api/excel/excel.worksheet#onactivated)|Se produce cuando se activa la hoja de cálculo.|
||[onChanged](/javascript/api/excel/excel.worksheet#onchanged)|Se produce cuando cambian los datos en una hoja de cálculo específica.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#ondeactivated)|Se produce cuando se desactiva la hoja de cálculo.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onselectionchanged)|Se produce cuando la selección cambia en una hoja de cálculo específica.|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardheight)|Devuelve el ancho estándar (predeterminado) de todas las filas de la hoja de cálculo, en puntos.|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardwidth)|Especifica el ancho estándar (predeterminado) de todas las columnas de la hoja de cálculo.|
||[tabColor](/javascript/api/excel/excel.worksheet#tabcolor)|Color de la pestaña de la hoja de cálculo.|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo que se ha activado.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|Obtiene el origen del evento.|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo que se ha añadido al libro.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#address)|Obtiene la dirección del intervalo que representa el área que ha cambiado en una hoja de cálculo específica.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changetype)|Obtiene el tipo de cambio que representa cómo se activa el evento cambiado.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|Obtiene el origen del evento.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se cambian los datos.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|Este evento se produce cuando se activa cualquier hoja de cálculo del libro.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|Este evento se produce cuando se agrega una nueva hoja de cálculo al libro.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|Este evento se produce cuando se desactiva cualquier hoja de cálculo del libro.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|Este evento se produce cuando se elimina una hoja de cálculo del libro.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo que se ha desactivado.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|Obtiene el origen del evento.|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo que se ha eliminado del libro.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt (frozenRange: cadena de rango \| )](/javascript/api/excel/excel.worksheetfreezepanes#freezeat-frozenrange-)|Establece las celdas inmovilizadas en la vista de la hoja de cálculo activa.|
||[freezeColumns (Count?: Number)](/javascript/api/excel/excel.worksheetfreezepanes#freezecolumns-count-)|Inmovilizar la primera columna o columnas de la hoja de cálculo en su lugar.|
||[freezeRows (Count?: Number)](/javascript/api/excel/excel.worksheetfreezepanes#freezerows-count-)|Inmovilizar la fila superior o filas de la hoja de cálculo en su lugar.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getlocation--)|Obtiene un rango que describe las celdas bloqueadas en la vista de hoja de cálculo activa.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#getlocationornullobject--)|Obtiene un rango que describe las celdas bloqueadas en la vista de hoja de cálculo activa.|
||[unfreeze ()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze--)|Elimina todos los paneles inmovilizados de la hoja de cálculo.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[Unprotect (password?: String)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|Desprotege una hoja de cálculo.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|Representa la opción de protección de la hoja de cálculo que permite editar objetos.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|Representa la opción de protección de la hoja de cálculo que permite editar escenarios.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|Representa la opción de protección de la hoja de cálculo del modo de selección.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|Obtiene la dirección del intervalo que representa el área seleccionada de una hoja de cálculo específica.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se cambia la selección.|
