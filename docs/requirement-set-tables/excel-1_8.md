| Class | Campos | Descripción |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[Formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|Especifica el operando derecho cuando la propiedad Operator se establece en un operador binario como GreaterThan (el operando izquierdo es el valor que el usuario intenta escribir en la celda).|
||[Formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|Con los operadores ternarios between y NotBetween, especifica el operando de límite superior.|
||[operator](/javascript/api/excel/excel.basicdatavalidation#operator)|El operador para validar los datos.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categorylabellevel)|Especifica una constante de enumeración ChartCategoryLabelLevel que hace referencia a|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayblanksas)|Especifica la forma en que se trazan las celdas en blanco en un gráfico.|
||[plotBy](/javascript/api/excel/excel.chart#plotby)|Especifica la manera en que las columnas o las filas se usan como series de datos en el gráfico.|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotvisibleonly)|True si solo se trazan las celdas visibles.False si se trazan tanto las celdas visibles como las ocultas.|
||[onActivated](/javascript/api/excel/excel.chart#onactivated)|Se produce cuando se activa el gráfico.|
||[onDeactivated](/javascript/api/excel/excel.chart#ondeactivated)|Se produce cuando se desactiva el gráfico.|
||[Trazado](/javascript/api/excel/excel.chart#plotarea)|Representa el plotArea del gráfico.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesnamelevel)|Especifica una constante de enumeración ChartSeriesNameLevel que hace referencia a|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showdatalabelsovermaximum)|Especifica si se muestran los rótulos de datos cuando el valor es mayor que el valor máximo del eje de valores.|
||[estilo](/javascript/api/excel/excel.chart#style)|Especifica el estilo de gráfico para el gráfico.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#chartid)|Obtiene el identificador del gráfico que está activado.|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se activa el gráfico.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#chartid)|Obtiene el identificador del gráfico que se agrega a la hoja de cálculo.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|Obtiene el origen del evento.|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo a la que se agrega el gráfico.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[orientación](/javascript/api/excel/excel.chartaxis#alignment)|Especifica la alineación de la etiqueta de la marca del eje especificado.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isbetweencategories)|Especifica si el eje de valores cruza el eje de categorías entre categorías.|
||[contraíble](/javascript/api/excel/excel.chartaxis#multilevel)|Especifica si un eje es multinivel.|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberformat)|Especifica el código de formato para la etiqueta de la marca del eje.|
||[desplaza](/javascript/api/excel/excel.chartaxis#offset)|Especifica la distancia entre los niveles de las etiquetas y la distancia entre el primer nivel y la línea del eje.|
||[position](/javascript/api/excel/excel.chartaxis#position)|Especifica la posición del eje especificado donde se cruza el otro eje.|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionat)|Especifica la posición del eje especificado donde se cruza el otro eje.|
||[setPositionAt (Value: Number)](/javascript/api/excel/excel.chartaxis#setpositionat-value-)|Establece la posición del eje especificado donde se cruza el otro eje.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textorientation)|Especifica el ángulo al que está orientado el texto para la etiqueta de marca del eje del gráfico.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|Especifica el formato de relleno del gráfico.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula (Formula: String)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|Un valor de cadena que representa la fórmula del título del eje del gráfico mediante la notación de estilo A1.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[border](/javascript/api/excel/excel.chartaxistitleformat#border)|Especifica el formato de borde del título del eje del gráfico, que incluye color, LineStyle y weight.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|Especifica el formato de relleno del título del eje del gráfico.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|Borra el formato del borde de un elemento del gráfico.	|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|Se produce cuando se activa un gráfico.|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|Se produce cuando se agrega un nuevo gráfico a la hoja de cálculo.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|Se produce cuando se desactiva un gráfico.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|Se produce cuando se elimina un gráfico.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[Autotexto](/javascript/api/excel/excel.chartdatalabel#autotext)|Especifica si la etiqueta de datos genera automáticamente un texto apropiado basado en el contexto.|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|Valor de cadena que representa la fórmula de la etiqueta de datos del gráfico mediante la notación de estilo A1.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|Representa la alineación horizontal de la etiqueta de datos del gráfico.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Representa la distancia, en puntos, desde el borde izquierdo de la etiqueta de datos del gráfico hasta el borde izquierdo del área del gráfico.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|Valor de cadena que representa el código de formato de la etiqueta de datos.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Representa el formato de la etiqueta de datos del gráfico.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Devuelve la altura, en puntos, de la etiqueta de datos del gráfico.|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Devuelve la anchura, en puntos, de la etiqueta de datos del gráfico.|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|Cadena que representa el texto de la etiqueta de datos en un gráfico.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|Representa el ángulo al que está orientado el texto para la etiqueta de datos del gráfico.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Representa la distancia, en puntos, desde el borde superior de la etiqueta de datos del gráfico hasta el borde superior del área del gráfico. |
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|Representa la alineación vertical de la etiqueta de datos del gráfico.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[border](/javascript/api/excel/excel.chartdatalabelformat#border)|Representa el formato de borde, que incluye el color, grosor y estilo de línea.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[Autotexto](/javascript/api/excel/excel.chartdatalabels#autotext)|Especifica si las etiquetas de datos generan automáticamente el texto apropiado basado en el contexto.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|Especifica la alineación horizontal de la etiqueta de datos del gráfico.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|Especifica el código de formato para los rótulos de datos.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|Representa el ángulo al que está orientado el texto para los rótulos de datos.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|Representa la alineación vertical de la etiqueta de datos del gráfico.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#chartid)|Obtiene el identificador del gráfico que está desactivado.|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se desactiva el gráfico.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#chartid)|Obtiene el identificador del gráfico que se ha eliminado de la hoja de cálculo.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|Obtiene el origen del evento.|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se ha eliminado el gráfico.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Especifica el alto de legendEntry en la leyenda del gráfico.|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|Especifica el índice de legendEntry en la leyenda del gráfico.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Especifica la izquierda de un legendEntry de gráfico.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Especifica la parte superior de un legendEntry de un gráfico.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Representa la anchura de legendEntry en la leyenda del gráfico. |
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[border](/javascript/api/excel/excel.chartlegendformat#border)|Representa el formato de borde, que incluye el color, grosor y estilo de línea.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|Especifica el valor del alto de plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|Especifica el valor de insideHeight de plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|Especifica el valor insideLeft de plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|Especifica el valor insideTop de plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|Especifica el valor de insideWidth de plotArea.|
||[left](/javascript/api/excel/excel.chartplotarea#left)|Especifica el valor de la izquierda de plotArea.|
||[position](/javascript/api/excel/excel.chartplotarea#position)|Especifica la posición de plotArea.|
||[format](/javascript/api/excel/excel.chartplotarea#format)|Especifica el formato del plotArea de un gráfico.|
||[top](/javascript/api/excel/excel.chartplotarea#top)|Especifica el valor superior de plotArea.|
||[width](/javascript/api/excel/excel.chartplotarea#width)|Especifica el valor del ancho de plotArea.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[border](/javascript/api/excel/excel.chartplotareaformat#border)|Especifica los atributos de borde de un gráfico plotArea.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|Especifica el formato de relleno de un objeto, que incluye información de formato de fondo.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|Especifica el grupo de la serie especificada.|
||[Expansion](/javascript/api/excel/excel.chartseries#explosion)|Especifica el valor de explosión de un sector del gráfico circular o de anillos.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|Especifica el ángulo del primer sector del gráfico circular o de anillos, en grados (en el sentido de las agujas del reloj desde vertical).|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|True si Excel invierte el patrón en el elemento cuando corresponde a un número negativo.|
||[se](/javascript/api/excel/excel.chartseries#overlap)|Especifica cómo se colocan las barras y columnas.|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|Representa una colección de todos las dataLabels de la serie.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|Especifica el tamaño de la sección secundaria de un gráfico circular con subgráfico circular o con subgráfico de barras, como un porcentaje del tamaño del gráfico circular principal.|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|Especifica la forma en que se dividen las dos secciones de un gráfico circular con subgráfico circular o de barras.|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|True si Excel asigna un color o trama diferente a cada marcador de datos.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|Representa el número de periodos que la línea de tendencia se extiende hacia atrás.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|Representa el número de periodos que la línea de tendencia se extiende hacia delante.|
||[rótulo](/javascript/api/excel/excel.charttrendline#label)|Representa la etiqueta de una línea de tendencia del gráfico.|
||[showEquation](/javascript/api/excel/excel.charttrendline#showequation)|True si la ecuación de la línea de tendencia se muestra en el gráfico.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|True si la R cuadrada de la línea de tendencia se muestra en el gráfico.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[Autotexto](/javascript/api/excel/excel.charttrendlinelabel#autotext)|Especifica si la etiqueta de la tendencia genera automáticamente el texto apropiado basado en el contexto.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|Valor de cadena que representa la fórmula de la etiqueta de línea de tendencia del gráfico mediante la notación de estilo A1.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|Representa la alineación horizontal de la etiqueta de línea de tendencia del gráfico.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|Representa la distancia, en puntos, desde el borde izquierdo de la etiqueta de línea de tendencia del gráfico hasta el borde izquierdo del área del gráfico.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|Valor de cadena que representa el código de formato de la etiqueta de línea de tendencia.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|El formato de la etiqueta de la tendencia del gráfico.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|Devuelve la altura, en puntos, de la etiqueta de línea de tendencia del gráfico.|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|Devuelve la anchura, en puntos, de la etiqueta de línea de tendencia del gráfico.|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|Cadena que representa el texto de la etiqueta de línea de tendencia en un gráfico.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|Representa el ángulo al que se orienta el texto para la etiqueta de la tendencia del gráfico.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|Representa la distancia, en puntos, desde el borde superior de la etiqueta de línea de tendencia del gráfico hasta el borde superior del área del gráfico.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|Representa la alineación vertical de la etiqueta de línea de tendencia del gráfico.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[border](/javascript/api/excel/excel.charttrendlinelabelformat#border)|Especifica el formato de borde, que incluye color, LineStyle y weight.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|Especifica el formato de relleno de la etiqueta de tendencia del gráfico actual.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|Especifica los atributos de fuente (nombre de fuente, tamaño de fuente, color, etc.) de una etiqueta de tendencia de gráfico.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|Una fórmula de validación de datos personalizados.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|Nombre de DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|Formato de número de DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|Posición de la DataPivotHierarchy.|
||[campo](/javascript/api/excel/excel.datapivothierarchy#field)|Devuelve los PivotFields asociados con DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|Identificador de DataPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Restablece DataPivotHierarchy a sus valores predeterminados.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|Especifica si los datos se deben mostrar como un cálculo de resumen específico.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|Especifica si se muestran todos los elementos de la DataPivotHierarchy.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[Add (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|Agrega PivotHierarchy al eje actual.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|Obtiene el número de jerarquías dinámicas en la colección.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|Obtiene una DataPivotHierarchy por su nombre o identificador.	|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|Obtiene una DataPivotHierarchy por su nombre.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[Remove (DataPivotHierarchy: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|Elimina PivotHierarchy del eje actual.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|Borra la validación de datos del rango actual.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|Mensaje de error cuando el usuario escribe datos no válidos.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|Especifica si la validación de datos se realizará en celdas en blanco; el valor predeterminado es true.|
||[prompt](/javascript/api/excel/excel.datavalidation#prompt)|Pregunta cuando los usuarios seleccionan una celda.|
||[type](/javascript/api/excel/excel.datavalidation#type)|Tipo de la validación de datos, vea Excel.DataValidationType para obtener más información.|
||[validación](/javascript/api/excel/excel.datavalidation#valid)|Representa si todos los valores de celda son válidos de acuerdo con las reglas de validación de datos.|
||[Rule](/javascript/api/excel/excel.datavalidation#rule)|Regla de validación de datos que contiene distintos tipos de criterios de validación de datos.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[mensaje](/javascript/api/excel/excel.datavalidationerroralert#message)|Representa el mensaje de alerta de error.|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showalert)|Especifica si se muestra un cuadro de diálogo de alerta de error cuando un usuario escribe datos no válidos.|
||[estilo](/javascript/api/excel/excel.datavalidationerroralert#style)|El tipo de alerta de validación de datos, consulte Excel. DataValidationAlertStyle para obtener información detallada.|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|Representa el título del cuadro de diálogo de alerta de error.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[mensaje](/javascript/api/excel/excel.datavalidationprompt#message)|Especifica el mensaje del aviso.|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#showprompt)|Especifica si se muestra un mensaje cuando un usuario selecciona una celda con validación de datos.|
||[title](/javascript/api/excel/excel.datavalidationprompt#title)|Especifica el título del mensaje.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[customiza](/javascript/api/excel/excel.datavalidationrule#custom)|Criterios de validación de datos personalizados.|
||[date](/javascript/api/excel/excel.datavalidationrule#date)|Criterios de validación de datos de fecha.|
||[decimal](/javascript/api/excel/excel.datavalidationrule#decimal)|Criterios de validación de datos decimales.|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|Criterios de validación de datos de lista.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textlength)|Criterios de validación de datos TextLength.|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|Criterios de validación de datos de tiempo.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholenumber)|Criterios de validación de datos de WholeNumber.|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[Formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Especifica el operando derecho cuando la propiedad Operator se establece en un operador binario como GreaterThan (el operando izquierdo es el valor que el usuario intenta escribir en la celda).|
||[Formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|Con los operadores ternarios between y NotBetween, especifica el operando de límite superior.|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#operator)|El operador para validar los datos.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|Determina si se permiten varios elementos de filtro.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|Nombre de FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|Posición de la FilterPivotHierarchy.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|Devuelve los PivotFields asociados con FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|Identificador de FilterPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|Restablece FilterPivotHierarchy a sus valores predeterminados.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[Add (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|Agrega PivotHierarchy al eje actual.|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|Obtiene el número de jerarquías dinámicas en la colección.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|Obtiene un FilterPivotHierarchy por su nombre o identificador.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|Obtiene una FilterPivotHierarchy por su nombre.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[Remove (filterPivotHierarchy: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|Elimina PivotHierarchy del eje actual.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|Muestra la lista en la celda desplegable o no, el valor predeterminado es true.|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|Origen de la lista de validación de datos|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|Nombre de PivotField.|
||[id](/javascript/api/excel/excel.pivotfield#id)|Identificador de PivotField.|
||[items](/javascript/api/excel/excel.pivotfield#items)|Devuelve los PivotFields asociados con el PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|Determina si se muestran todos los elementos de PivotField.|
||[sortByLabels (sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|Ordena el PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|Subtotales del PivotField.	|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|Obtiene el número de campos dinámicos de la colección.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|Obtiene un campo dinámico por su nombre o identificador.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|Obtiene un PivotField por nombre.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|Nombre de la PivotHierarchy.|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|Devuelve los PivotFields asociados con PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|Id. de la PivotHierarchy.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|Obtiene el número de jerarquías dinámicas en la colección.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|Obtiene un PivotHierarchy por su nombre o identificador.	|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|Obtiene un PivotHierarchy por su nombre.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|Determina si el elemento se expande para mostrar los elementos secundarios o si está contraído y se ocultan los elementos secundarios.|
||[name](/javascript/api/excel/excel.pivotitem#name)|Nombre del PivotItem.|
||[id](/javascript/api/excel/excel.pivotitem#id)|Identificador de PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|Especifica si el PivotItem está visible.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|Obtiene el número de PivotItems de la colección.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|Obtiene un PivotItem por su nombre o identificador.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|Obtiene un PivotItem por nombre.|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|Devuelve el intervalo donde residen las etiquetas de columna de la tabla dinámica.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|Devuelve el intervalo donde residen los valores de datos de tabla dinámica.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|Devuelve el intervalo del área de filtro de la tabla dinámica.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|Devuelve el intervalo en el que existe la tabla dinámica, excluyendo el área de filtro.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|Devuelve el intervalo donde residen las etiquetas de fila de la tabla dinámica.|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|Esta propiedad indica el PivotLayoutType de todos los campos de la tabla dinámica.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|Especifica si el informe de tabla dinámica muestra el total general de las columnas.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|Especifica si el informe de tabla dinámica muestra el total general de las filas.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|Esta propiedad indica el SubtotalLocationType de todos los campos de la tabla dinámica.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|Elimina la tabla dinámica.|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|Las jerarquías dinámicas de columna de la tabla dinámica.|
||[Hierarchies](/javascript/api/excel/excel.pivottable#datahierarchies)|Las jerarquías dinámicas de datos de la tabla dinámica.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|Las jerarquías dinámicas de filtro de la tabla dinámica.|
||[jerarquías](/javascript/api/excel/excel.pivottable#hierarchies)|Las jerarquías dinámicas de la tabla dinámica.|
||[diseñar](/javascript/api/excel/excel.pivottable#layout)|El PivotLayout que describe el diseño y la estructura visual de la tabla dinámica.|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowhierarchies)|Las jerarquías dinámicas de fila de la tabla dinámica.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[Add (Name: String, Source: Range \| String \| TABLE, Destination: Range \| String)](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|Agrega una tabla dinámica basada en los datos de origen especificados e insértela en la celda superior izquierda del rango de destino.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|Devuelve un objeto de validación de datos.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|Nombre de la RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|Posición de la RowColumnPivotHierarchy.|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|Devuelve los PivotFields asociados con la RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|Identificador de la RowColumnPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|Restablece la RowColumnPivotHierarchy a sus valores predeterminados.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[Add (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|Agrega PivotHierarchy al eje actual.|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|Obtiene el número de jerarquías dinámicas en la colección.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|Obtiene una RowColumnPivotHierarchy por su nombre o identificador.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|Obtiene una RowColumnPivotHierarchy por su nombre.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[Remove (rowColumnPivotHierarchy: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|Elimina PivotHierarchy del eje actual.|
|[Tiempo de ejecución](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|Alternar eventos de JavaScript en el panel de tareas o el complemento de contenido actual.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|El PivotField de base para el cálculo de ShowAs, si procede, será en función del tipo de ShowAsCalculation, en el resto de casos será null.|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|El elemento base para el cálculo de ShowAs, si procede, será en función del tipo de ShowAsCalculation, en el resto de casos será null.|
||[cubo](/javascript/api/excel/excel.showasrule#calculation)|El cálculo de ShowAs que se usará para el PivotField de datos.|
|[Estilo](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|Especifica si se aplica sangría al texto automáticamente cuando la alineación del texto de una celda se establece en una distribución igualada.|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|Orientación del texto para el estilo.|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|Si Automatic se establece en true, se omitirán el resto de valores al establecer los subtotales.|
||[normal](/javascript/api/excel/excel.subtotals#average)||
||[count](/javascript/api/excel/excel.subtotals#count)||
||[countNumbers](/javascript/api/excel/excel.subtotals#countnumbers)||
||[Nº](/javascript/api/excel/excel.subtotals#max)||
||[min](/javascript/api/excel/excel.subtotals#min)||
||[Product](/javascript/api/excel/excel.subtotals#product)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#standarddeviation)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#standarddeviationp)||
||[total](/javascript/api/excel/excel.subtotals#sum)||
||[diferencias](/javascript/api/excel/excel.subtotals#variance)||
||[varianceP](/javascript/api/excel/excel.subtotals#variancep)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyid)|Devuelve un identificador numérico.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|Obtiene el rango que representa el área modificada de una tabla en una hoja de cálculo específica.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|Obtiene el rango que representa el área modificada de una tabla en una hoja de cálculo específica.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|True si el libro está abierto en modo de solo lectura.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[Worksheet](/javascript/api/excel/excel.worksheet)|[he calculado](/javascript/api/excel/excel.worksheet#oncalculated)|Se produce cuando se calcula la hoja de cálculo.|
||[showGridlines](/javascript/api/excel/excel.worksheet#showgridlines)|Especifica si la cuadrícula es visible para el usuario.|
||[showHeadings](/javascript/api/excel/excel.worksheet#showheadings)|Especifica si los títulos están visibles para el usuario.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se produjo el cálculo.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|Obtiene el intervalo que representa el área que ha cambiado en una hoja de cálculo específica.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|Obtiene el intervalo que representa el área que ha cambiado en una hoja de cálculo específica.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[he calculado](/javascript/api/excel/excel.worksheetcollection#oncalculated)|Este evento se produce cuando se calcula cualquier hoja de cálculo del libro.|
