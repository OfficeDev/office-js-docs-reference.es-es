| Clase | Campos | Descripción |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|Devuelve la versión del motor de cálculo de Excel usada para la última actualización completa.|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|Devuelve el estado del cálculo de la aplicación.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|Devuelve la configuración de cálculo iterativo.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|Suspende la actualización de pantalla hasta que se `context.sync()` llame a la siguiente.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|Aplica el objecto AutoFilter a un rango.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|Borra los criterios de filtro de AutoFilter.|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|Devuelve un objeto Range que representa el intervalo al que se aplica el objeto AutoFilter.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|Devuelve un objeto Range que representa el intervalo al que se aplica el objeto AutoFilter.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|Matriz que contiene todos los criterios de filtro de un intervalo autofiltrado.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Especifica si el autofiltro está habilitado.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|Especifica si el Autofiltro tiene criterios de filtro.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|Aplica el objeto Autofilter especificado actualmente en el intervalo.|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|Quita el AutoFilter para el intervalo.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|Representa la propiedad `color` de un único borde.|
||[style](/javascript/api/excel/excel.cellborder#style)|Representa la propiedad `style` de un único borde.|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintandshade)|Representa la propiedad `tintAndShade` de un único borde.|
||[weight](/javascript/api/excel/excel.cellborder#weight)|Representa la propiedad `weight` de un único borde.|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)|Representa la propiedad `format.borders.bottom`.|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonaldown)|Representa la propiedad `format.borders.diagonalDown`.|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalup)|Representa la propiedad `format.borders.diagonalUp`.|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|Representa la propiedad `format.borders.horizontal`.|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|Representa la propiedad `format.borders.left`.|
||[right](/javascript/api/excel/excel.cellbordercollection#right)|Representa la propiedad `format.borders.right`.|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|Representa la propiedad `format.borders.top`.|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|Representa la propiedad `format.borders.vertical`.|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)|Representa la propiedad `address`.|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addresslocal)|Representa la propiedad `addressLocal`.|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|Representa la propiedad `hidden`.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|Representa la propiedad `format.fill.color`.|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|Representa la propiedad `format.fill.pattern`.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)|Representa la propiedad `format.fill.patternColor`.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)|Representa la propiedad `format.fill.patternTintAndShade`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)|Representa la propiedad `format.fill.tintAndShade`.|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|Representa la propiedad `format.font.bold`.|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|Representa la propiedad `format.font.color`.|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|Representa la propiedad `format.font.italic`.|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|Representa la propiedad `format.font.name`.|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|Representa la propiedad `format.font.size`.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|Representa la propiedad `format.font.strikethrough`.|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|Representa la propiedad `format.font.subscript`.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|Representa la propiedad `format.font.superscript`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintandshade)|Representa la propiedad `format.font.tintAndShade`.|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|Representa la propiedad `format.font.underline`.|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoindent)|Representa la propiedad `autoIndent`.|
||[borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|Representa la propiedad `borders`.|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|Representa la propiedad `fill`.|
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)|Representa la propiedad `font`.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalalignment)|Representa la propiedad `horizontalAlignment`.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentlevel)|Representa la propiedad `indentLevel`.|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|Representa la propiedad `protection`.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingorder)|Representa la propiedad `readingOrder`.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinktofit)|Representa la propiedad `shrinkToFit`.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textorientation)|Representa la propiedad `textOrientation`.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#usestandardheight)|Representa la propiedad `useStandardHeight`.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#usestandardwidth)|Representa la propiedad `useStandardWidth`.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalalignment)|Representa la propiedad `verticalAlignment`.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wraptext)|Representa la propiedad `wrapText`.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulahidden)|Representa la propiedad `format.protection.formulaHidden`.|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)|Representa la propiedad `format.protection.locked`.|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueafter)|Representa el valor después de cambiar.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valuebefore)|Representa el valor antes de cambiar.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valuetypeafter)|Representa el tipo de valor después de cambiar.|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valuetypebefore)|Representa el tipo de valor antes de cambiar.|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|Activa el gráfico en la interfaz de usuario de Excel.|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|Contiene las opciones del gráfico dinámico.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|Especifica la combinación de colores del gráfico.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|Especifica si el área del gráfico tiene esquinas redondeadas.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|Especifica si el formato de número está vinculado a las celdas.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|Especifica si el desbordamiento de bin está habilitado en un gráfico de histograma o un gráfico de pareto.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|Especifica si bin underflow está habilitado en un gráfico de histograma o un gráfico de pareto.|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|Especifica el recuento de ubicaciones de un gráfico de histograma o un gráfico de pareto.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|Especifica el valor de desbordamiento de bandeja de un gráfico de histograma o gráfico de pareto.|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|Especifica el tipo de la bandeja para un gráfico de histograma o un gráfico de pareto.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|Especifica el valor de subflujo bin de un gráfico de histograma o gráfico de pareto.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Especifica el valor de ancho de bin de un gráfico de histograma o gráfico de pareto.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|Especifica si el tipo de cálculo de cuartil de un cuadro y un gráfico de bigotes.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|Especifica si los puntos internos se muestran en un cuadro y un gráfico de bigotes.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|Especifica si la línea media se muestra en un cuadro y un gráfico de bigotes.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|Especifica si el marcador medio se muestra en un cuadro y un gráfico de bigotes.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|Especifica si los puntos atípicos se muestran en un cuadro y un gráfico de bigotes.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|Especifica si el formato de número está vinculado a las celdas (para que el formato de número cambie en las etiquetas cuando cambie en las celdas).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|Especifica si el formato de número está vinculado a las celdas.|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|Especifica si las barras de error tienen un límite de estilo final.|
||[include](/javascript/api/excel/excel.charterrorbars#include)|Especifica las partes de las barras de error a incluir.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Especifica el tipo de formato de las barras de error.|
||[type](/javascript/api/excel/excel.charterrorbars#type)|El tipo de rango marcado por las barras de error.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Especifica si se muestran las barras de error.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Representa el formato de línea de gráfico.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|Especifica la estrategia de etiquetas de mapa de serie de un gráfico de mapa de región.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Especifica el nivel de asignación de serie de un gráfico de mapa de región.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|Especifica el tipo de proyección de serie de un gráfico de mapa de región.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|Especifica si se muestran los botones de campo del eje en un gráfico dinámico.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|Especifica si se muestran los botones de campo de leyenda en un gráfico dinámico.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|Especifica si se muestran los botones de campo de filtro de informe en un gráfico dinámico.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|Especifica si se muestran los botones del campo mostrar valor en un gráfico dinámico.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|Esto puede ser un valor entero entre 0 (cero) y 300, que representa el porcentaje del tamaño predeterminado.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|Especifica el color para el valor máximo de una serie de gráficos de mapa de región.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|Especifica el tipo para el valor máximo de una serie de gráficos de mapa de región.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|Especifica el valor máximo de una serie de gráficos de mapa de región.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|Especifica el color del valor de punto medio de una serie de gráficos de mapa de región.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|Especifica el tipo para el valor de punto medio de una serie de gráficos de mapa de región.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|Especifica el valor de punto medio de una serie de gráficos de mapa de región.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|Especifica el color para el valor mínimo de una serie de gráficos de mapa de región.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|Especifica el tipo para el valor mínimo de una serie de gráficos de mapa de región.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|Especifica el valor mínimo de una serie de gráficos de mapa de región.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|Especifica el estilo de degradado de serie de un gráfico de mapa de región.|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|Especifica el color de relleno de los puntos de datos negativos de una serie.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|Especifica el área de estrategia de etiqueta principal de serie para un gráfico de mapas de árbol.|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|Contiene las opciones de intervalo para gráficos de histograma y diagramas de Pareto.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|Contiene las opciones para el gráfico de cajas y bigotes.|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|Contiene las opciones para un gráfico de mapa de región.|
||[xerrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|Indica el objeto de la barra de error de una serie de gráficos.|
||[yerrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|Indica el objeto de la barra de error de una serie de gráficos.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|Especifica si las líneas de conector se muestran en gráficos de cascada.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|Especifica si se muestran líneas de directriz para cada etiqueta de datos de la serie.|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|Especifica el valor de umbral que separa dos secciones de un gráfico circular o de una barra circular.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|Especifica si el formato de número está vinculado a las celdas (para que el formato de número cambie en las etiquetas cuando cambie en las celdas).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|Representa la propiedad `address`.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|Representa la propiedad `addressLocal`.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|Representa la propiedad `columnIndex`.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|Devuelve el objeto RangeAreas, que incluye uno o más intervalos rectangulares, al que se aplica el formato condicional.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|Devuelve un objeto RangeAreas, que incluye uno o más intervalos rectangulares, con valores de celda no válidos.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|Devuelve un objeto RangeAreas, que incluye uno o más intervalos rectangulares, con valores de celda no válidos.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|La propiedad que usa el filtro para realizar filtrado enriquecido en richvalues.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Devuelve el identificador de la forma.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Devuelve el objeto de forma de la forma geométrica.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|Devuelve el número de formas del grupo de formas.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|Obtiene una forma mediante su nombre o identificador.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|Obtiene una forma basada en su posición en la colección.|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|Pie de página central de la hoja de cálculo.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|Encabezado central de la hoja de cálculo.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|Pie de página izquierdo de la hoja de cálculo.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|Encabezado izquierdo de la hoja de cálculo.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|Pie de página derecho de la hoja de cálculo.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|Encabezado derecho de la hoja de cálculo.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|El encabezado o pie general, utilizado para todas las páginas, salvo que se indique par, impar o primera página.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|El encabezado o pie de página para páginas pares, debe especificarse el encabezado o pie de página para páginas impares.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|El encabezado o pie de página de la primera página, para las demás páginas se usa el encabezado o pie de página general, par o impar.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|El encabezado o pie de página para páginas impares, debe especificarse el encabezado o pie de página para páginas pares.|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|Estado por el que se establecen los encabezados o pies de página.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|Obtiene o establece una marca que indica si los encabezados o pies de página están alineados con los márgenes de página establecidos en las opciones de diseño de página de la hoja de cálculo.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|Obtiene o establece una marca que indica si los encabezados o pies de página deben escalarse según la escala de porcentaje de página establecida en las opciones de diseño de página de la hoja de cálculo.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Devuelve el formato de la imagen.|
||[id](/javascript/api/excel/excel.image#id)|Especifica el identificador de forma del objeto image.|
||[shape](/javascript/api/excel/excel.image#shape)|Devuelve el objeto de forma asociado con la imagen.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|True si Excel usa iteración para resolver referencias circulares.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|Especifica la cantidad máxima de cambios entre cada iteración a medida que Excel resuelve las referencias circulares.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Especifica el número máximo de iteraciones que Excel puede usar para resolver una referencia circular.|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginarrowheadlength)|Representa la longitud de la punta de flecha al comienzo de la línea especificada.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginarrowheadstyle)|Representa el estilo de la punta de flecha al principio de la línea especificada.|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#beginarrowheadwidth)|Representa el ancho de la punta de flecha al comienzo de la línea especificada.|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectbeginshape-shape--connectionsite-)|Une el principio del conector especificado a una forma específica.|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectendshape-shape--connectionsite-)|Une el final del conector especificado a una forma específica.|
||[connectorType](/javascript/api/excel/excel.line#connectortype)|Indica el tipo de conector de la línea.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectbeginshape--)|Separa el principio del conector especificado de una forma.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectendshape--)|Separa el final del conector especificado de una forma.|
||[endArrowheadLength](/javascript/api/excel/excel.line#endarrowheadlength)|Representa la longitud de la punta de flecha al final de la línea especificada.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endarrowheadstyle)|Representa el estilo de la punta de flecha al final de la línea especificada.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endarrowheadwidth)|Representa el ancho de la punta de flecha al final de la línea especificada.|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginconnectedshape)|Representa la forma a la que está unido el principio de la línea especificada.|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginconnectedsite)|Representa el sitio de conexión al que está conectado el principio de un conector.|
||[endConnectedShape](/javascript/api/excel/excel.line#endconnectedshape)|Representa la forma a la que está unido el final de la línea especificada.|
||[endConnectedSite](/javascript/api/excel/excel.line#endconnectedsite)|Representa el sitio de conexión al que está conectado el extremo de un conector.|
||[id](/javascript/api/excel/excel.line#id)|Especifica el identificador de forma.|
||[isBeginConnected](/javascript/api/excel/excel.line#isbeginconnected)|Especifica si el principio de la línea especificada está conectado a una forma.|
||[isEndConnected](/javascript/api/excel/excel.line#isendconnected)|Especifica si el final de la línea especificada está conectado a una forma.|
||[shape](/javascript/api/excel/excel.line#shape)|Devuelve el objeto de la forma asociado a la línea.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|Elimina un objeto de salto de página.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|Obtiene la primera celda después del salto de página.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|Especifica el índice de columna para el salto de página|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|Especifica el índice de fila para el salto de página|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|Agrega un salto de página antes de la celda superior izquierda del intervalo especificado.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|Obtiene el número de saltos de página de una colección.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|Obtiene un objeto de salto de página a través del índice.|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|Restablece todos los saltos de página manuales de la colección.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackandwhite)|Opción de impresión en blanco y negro de la hoja de cálculo.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottommargin)|Margen de página inferior de la hoja de cálculo que se usará para imprimir en puntos.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerhorizontally)|Marca horizontal del centro de la hoja de cálculo.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centervertically)|Marca vertical del centro de la hoja de cálculo.|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftmode)|Opción de modo borrador de la hoja de cálculo.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstpagenumber)|El primer número de página de la hoja de cálculo que se va a imprimir.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footermargin)|Margen de pie de página de la hoja de cálculo, en puntos, para su uso al imprimir.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getprintarea--)|Obtiene el objeto RangeAreas, que incluye uno o más intervalos rectangulares, que representa el área de impresión de la hoja de cálculo.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getprintareaornullobject--)|Obtiene el objeto RangeAreas, que incluye uno o más intervalos rectangulares, que representa el área de impresión de la hoja de cálculo.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumns--)|Obtiene el objeto de intervalo que representa las columnas de título.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumnsornullobject--)|Obtiene el objeto de intervalo que representa las columnas de título.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getprinttitlerows--)|Obtiene el objeto de intervalo que representa las filas de título.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlerowsornullobject--)|Obtiene el objeto de intervalo que representa las filas de título.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headermargin)|Margen de encabezado de la hoja de cálculo, en puntos, para su uso al imprimir.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftmargin)|Margen izquierdo de la hoja de cálculo, en puntos, para su uso al imprimir.|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|La orientación de la hoja de cálculo de la página.|
||[paperSize](/javascript/api/excel/excel.pagelayout#papersize)|Tamaño de papel de la hoja de cálculo de la página.|
||[printComments](/javascript/api/excel/excel.pagelayout#printcomments)|Especifica si los comentarios de la hoja de cálculo deben mostrarse al imprimir.|
||[printErrors](/javascript/api/excel/excel.pagelayout#printerrors)|Opción de errores de impresión de la hoja de cálculo.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printgridlines)|Especifica si se imprimirán las líneas de cuadrícula de la hoja de cálculo.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printheadings)|Especifica si se imprimirán los encabezados de la hoja de cálculo.|
||[printOrder](/javascript/api/excel/excel.pagelayout#printorder)|Opción de orden de impresión de página de la hoja de cálculo.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersfooters)|Configuración de encabezado y pie de página de la hoja de cálculo.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightmargin)|Margen derecho de la hoja de cálculo, en puntos, para su uso al imprimir.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|Establece el área de impresión de la hoja de cálculo.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Establece los márgenes de página de la hoja de cálculo con unidades.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|Establece las columnas que contienen las celdas que se repetirán a la izquierda de cada página de la hoja de cálculo que se va a imprimir.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|Establece las filas que contienen las celdas que se repetirán en la parte superior de cada página de la hoja de cálculo que se va a imprimir.|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|Margen superior de la hoja de cálculo, en puntos, para su uso al imprimir.|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|Opciones de zoom de impresión de la hoja de cálculo.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|Especifica el margen inferior del diseño de página en la unidad especificada para la impresión.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|Especifica el margen de pie de página del diseño de página en la unidad especificada para la impresión.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|Especifica el margen de encabezado de diseño de página en la unidad especificada para la impresión.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|Especifica el margen izquierdo del diseño de página en la unidad especificada para la impresión.|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|Especifica el margen derecho del diseño de página en la unidad especificada para la impresión.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|Especifica el margen superior del diseño de página en la unidad especificada para la impresión.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|Número de páginas para ajustar horizontalmente.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|El valor de la escala de página de impresión puede estar entre 10 y 400.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|Número de páginas para ajustar verticalmente.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortby: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Ordena el campo de tabla dinámica por los valores especificados en un ámbito determinado.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|Especifica si el formato se formateará automáticamente cuando se actualiza o cuando se mueven los campos.|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|Obtiene la DataHierarchy que se usa para calcular el valor de un intervalo especificado en la tabla dinámica.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Obtiene los elementos dinámicos de un eje que componen el valor de un intervalo especificado en la tabla dinámica.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|Especifica si el formato se conserva cuando el informe se actualiza o se vuelve a calcular mediante operaciones como la rotación, ordenación o cambio de elementos de campo de página.|
||[setAutosortOnCell(cell: Range \| string, sortby: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Establece la tabla dinámica para la ordenación automática mediante la celda especificada para seleccionar automáticamente el contexto y todos los criterios necesarios.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|Especifica si la tabla dinámica permite que el usuario edite los valores del cuerpo de los datos.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|Especifica si la tabla dinámica usa listas personalizadas al ordenar.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange?: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|Rellena el intervalo desde el intervalo actual hasta el intervalo de destino mediante la lógica de autorrelleno especificada.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|Convierte el intervalo de celdas con tipos de datos en texto.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|Convierte el intervalo de celdas en un tipo de datos vinculado en la hoja de cálculo.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copia el formato o los datos de la celda del intervalo de origen o RangeAreas al intervalo actual.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|Busca la cadena especificada, según los criterios especificados.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|Busca la cadena especificada, según los criterios especificados.|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|Aplica el relleno rápido en el rango actual. Relleno rápido rellena automáticamente los datos cuando detecta un patrón, por lo que el rango debe ser de una única columna y tener datos en los laterales para obtener el patrón.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|Devuelve una matriz 2D que encapsula los datos para la fuente, el relleno, los bordes, la alineación y otras propiedades de la celda.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|Devuelve una matriz de una sola dimensión que encapsula los datos para la fuente, el relleno, los bordes, la alineación y otras propiedades de la columna.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|Devuelve una matriz de una sola dimensión que encapsula los datos para la fuente, el relleno, los bordes, la alineación y otras propiedades de la fila.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Obtiene el objeto RangeAreas, que incluye uno o más intervalos rectangulares y representa todas las celdas que coinciden con el tipo y valor especificados.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Obtiene el objeto RangeAreas, que incluye uno o más intervalos y representa todas las celdas que coinciden con el tipo y valor especificados.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|Obtiene una colección con ámbito de tablas que se superpone con el intervalo.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|Indica el estado del tipo de datos de cada celda.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)|Quita los valores duplicados del intervalo especificado por las columnas.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceall-text--replacement--criteria-)|Busca y reemplaza la cadena especificada, según los criterios especificados dentro del intervalo actual.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setcellproperties-cellpropertiesdata-)|Actualiza el intervalo con una matriz 2D de propiedades de celda, encapsulando elementos como la fuente, el relleno, los bordes, la alineación, y así sucesivamente.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setcolumnproperties-columnpropertiesdata-)|Actualiza el intervalo con una matriz de una sola dimensión de propiedades de columna, encapsulando elementos como la fuente, el relleno, los bordes, la alineación, y así sucesivamente.|
||[setDirty()](/javascript/api/excel/excel.range#setdirty--)|Establece un intervalo que se deberá actualizar cuando se realice la próxima actualización.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setrowproperties-rowpropertiesdata-)|Actualiza el intervalo con una matriz de una sola dimensión de propiedades de fila, encapsulando elementos como la fuente, el relleno, los bordes, la alineación, y así sucesivamente.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate--)|Calcula todas las celdas del objeto RangeAreas.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Borra valores, formato, relleno, bordes, etc., en cada una de las áreas que conforman este objeto RangeAreas.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|Convierte todas las celdas de la RangeAreas con tipos de datos en texto.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|Convierte todas las celdas de la RangeAreas en tipos de datos vinculados.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copia el formato o los datos de la celda del intervalo de origen o RangeAreas al objeto RangeAreas actual.|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|Devuelve un objeto RangeAreas que representa toda las columnas de RangeAreas (por ejemplo, si el RangeAreas actual representa las celdas "B4:E11, H2", devuelve un RangeAreas que representa las columnas "B:E, H:H").|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|Devuelve un objeto RangeAreas que representa toda las filas de RangeAreas (por ejemplo, si el RangeAreas actual representa las celdas "B4:E11", devuelve un RangeAreas que representa las filas "4:11").|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|Devuelve el objeto RangeAreas que representa la intersección de los rangos o RangeAreas especificados.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|Devuelve el objeto RangeAreas que representa la intersección de los rangos o RangeAreas especificados.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|Devuelve un objeto RangeAreas que se desplaza por el desplazamiento de filas y columnas específico.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Devuelve un objeto RangeAreas que representa todas las celdas que coinciden con el tipo y valor especificados.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|Devuelve un objeto RangeAreas que representa todas las celdas que coinciden con el tipo y valor especificados.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#gettables-fullycontained-)|Devuelve una colección con ámbito de tablas que se superpone con cualquier rango en este objeto RangeAreas.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareas-valuesonly-)|Devuelve el RangeAreas usado que incluye todas las áreas usadas de intervalos rectangulares individuales en el objeto RangeAreas.|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareasornullobject-valuesonly-)|Devuelve el RangeAreas usado que incluye todas las áreas usadas de intervalos rectangulares individuales en el objeto RangeAreas.|
||[address](/javascript/api/excel/excel.rangeareas#address)|Devuelve la referencia RangeAreas en estilo A1.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addresslocal)|Devuelve la referencia RangeAreas en la configuración regional del usuario.|
||[areaCount](/javascript/api/excel/excel.rangeareas#areacount)|Devuelve el número de intervalos rectangulares que conforman este objeto RangeAreas.|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|Devuelve una colección de intervalos rectangulares que conforman este objeto RangeAreas.|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellcount)|Devuelve el número de celdas en el objeto RangeAreas, sumando los recuentos de celda de todos los rangos individuales rectangulares.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalformats)|Devuelve una colección de ConditionalFormats que forman una intersección con cualquier celda en este objeto RangeAreas.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#datavalidation)|Devuelve un objeto dataValidation para todos los rangos del RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareas#format)|Devuelve un objeto RangeFormat, encapsulando la fuente, el relleno, los bordes, la alineación y otras propiedades de todos los rangos del objeto RangeAreas.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isentirecolumn)|Especifica si todos los rangos de este objeto RangeAreas representan columnas enteras (por ejemplo, "A:C, Q:Z").|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isentirerow)|Especifica si todos los intervalos de este objeto RangeAreas representan filas completas (por ejemplo, "1:3, 5:7").|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|Devuelve la hoja de cálculo de la RangeAreas actual.|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|Establece el RangeAreas que se deberá actualizar cuando se realice la próxima actualización.|
||[style](/javascript/api/excel/excel.rangeareas#style)|Indica el estilo de todos los intervalos con nombre de este objeto RangeAreas.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|Especifica un doble que aligera u oscurece un color para borde de intervalo, el valor está entre -1 (más oscuro) y 1 (más brillante), con 0 para el color original.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|Especifica un doble que aligera u oscurece un color para bordes de intervalo, el valor está entre -1 (más oscuro) y 1 (más brillante), con 0 para el color original.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|Devuelve el número de intervalos del RangeCollection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|Devuelve el objeto de intervalo según su posición en el RangeCollection.|
||[items](/javascript/api/excel/excel.rangecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|El patrón de un intervalo.|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|Código de color HTML que representa el color del patrón de intervalo, del formulario #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|Especifica un doble que aligera u oscurece un color de patrón para relleno de rango, el valor está entre -1 (más oscuro) y 1 (más brillante), con 0 para el color original.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|Especifica un doble que aligera u oscurece un color para el relleno de rango.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|Especifica el estado tachado de fuente.|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|Especifica el estado subíndice de la fuente.|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|Especifica el estado de Superíndice de fuente.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|Especifica un doble que aligera u oscurece un color para fuente de rango, el valor está entre -1 (más oscuro) y 1 (más brillante), con 0 para el color original.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|Especifica si el texto se aplica sangría automáticamente cuando la alineación de texto se establece en distribución igual.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|Un número entero entre 0 y 250 que indica el nivel de sangría.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|El orden de lectura para el intervalo.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|Especifica si el texto se reduce automáticamente para ajustarse al ancho de columna disponible.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|Número de filas duplicadas quitadas por la operación.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|Número de filas únicas restantes presentes en el intervalo resultante.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|Especifica si la coincidencia debe ser completa o parcial.|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|Especifica si la coincidencia distingue mayúsculas de minúsculas.|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|Representa la propiedad `address`.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|Representa la propiedad `addressLocal`.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|Representa la propiedad `rowIndex`.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|Especifica si la coincidencia debe ser completa o parcial.|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|Especifica si la coincidencia distingue mayúsculas de minúsculas.|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|Especifica la dirección de la búsqueda|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|Representa la propiedad `format`.|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|Representa la propiedad `hyperlink`.|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|Representa la propiedad `style`.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)|Representa la propiedad `columnHidden`.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnwidth)||
||[format: Excel.CellPropertiesFormat & {
            columnWidth?] (/javascript/api/excel/excel.settablecolumnproperties#format)|Representa la propiedad `format`.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format: Excel.CellPropertiesFormat & {
            rowHeight?] (/javascript/api/excel/excel.settablerowproperties#format)|Representa la propiedad `format`.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowheight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)|Representa la propiedad `rowHidden`.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|Especifica el texto de descripción alternativo para un objeto Shape.|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|Especifica el texto de título alternativo para un objeto Shape.|
||[delete()](/javascript/api/excel/excel.shape#delete--)|Quita la forma de la hoja de cálculo.|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|Especifica el tipo de forma geométrica de esta forma geométrica.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getasimage-format-)|Convierte la forma a una imagen y devuelve la imagen como una cadena con codificación base 64.|
||[height](/javascript/api/excel/excel.shape#height)|Especifica el alto, en puntos, de la forma.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|Mueve la forma horizontalmente el número de puntos especificado.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|Gira la forma en el sentido de las agujas del reloj alrededor del eje Z según el número de grados.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|Mueve la forma verticalmente el número de puntos.|
||[left](/javascript/api/excel/excel.shape#left)|La distancia, en puntos, desde el lado izquierdo de la forma hasta el lado izquierdo de la hoja de cálculo.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|Especifica si la relación de aspecto de esta forma está bloqueada.|
||[name](/javascript/api/excel/excel.shape#name)|Especifica el nombre de la forma.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionsitecount)|Devuelve el número de sitios de conexión en esta forma.|
||[fill](/javascript/api/excel/excel.shape#fill)|Devuelve el formato de relleno de esta forma.|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|Devuelve la forma geométrica asociada con la forma.|
||[group](/javascript/api/excel/excel.shape#group)|Devuelve el grupo de forma asociado con la forma.|
||[id](/javascript/api/excel/excel.shape#id)|Especifica el identificador de forma.|
||[image](/javascript/api/excel/excel.shape#image)|Devuelve la imagen asociada con la forma.|
||[level](/javascript/api/excel/excel.shape#level)|Especifica el nivel de la forma especificada.|
||[line](/javascript/api/excel/excel.shape#line)|Devuelve la línea asociada con la forma.|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|Devuelve el formato de línea de esta forma.|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|Se produce cuando se activa la forma.|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|Se produce cuando se desactiva la forma.|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|Especifica el grupo primario de esta forma.|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|Devuelve el objeto de marco de texto de esta forma.|
||[type](/javascript/api/excel/excel.shape#type)|Devuelve el tipo de esta forma.|
||[zOrderPosition](/javascript/api/excel/excel.shape#zorderposition)|Devuelve la posición de la forma especificada en el orden z, siendo 0 la parte inferior de la pila del orden.|
||[rotation](/javascript/api/excel/excel.shape#rotation)|Especifica el giro, en grados, de la forma.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Cambia el alto de la forma en un factor especificado.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Cambia el ancho de la forma en un factor especificado.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|Mueve la forma especificada hacia arriba o hacia abajo en el orden z de la colección, que se desplaza delante o detrás de otras formas.|
||[top](/javascript/api/excel/excel.shape#top)|La distancia, en puntos, desde el borde superior de la forma al borde superior de la hoja de cálculo.|
||[visible](/javascript/api/excel/excel.shape#visible)|Especifica si la forma está visible.|
||[width](/javascript/api/excel/excel.shape#width)|Especifica el ancho, en puntos, de la forma.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|Obtiene el identificador de la forma activada.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se activa la forma.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|Agrega una forma geométrica a la hoja de cálculo.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|Agrupa un subconjunto de formas en la hoja de cálculo de esta colección.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|Crea una imagen de una cadena con codificación base64 y la agrega a la hoja de cálculo.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Agrega una línea a la hoja de cálculo.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|Agrega un cuadro de texto a la hoja de cálculo con el texto proporcionado como contenido.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|Devuelve el número de formas de la hoja de cálculo.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|Obtiene una forma mediante su nombre o identificador.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|Obtiene una forma utilizando su posición en la colección.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|Obtiene el identificador de la forma desactivada.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se desactiva la forma.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|Limpia el formato de relleno de esta forma.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|Representa el color de primer plano de relleno de forma en formato de color HTML, del formulario #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja")|
||[type](/javascript/api/excel/excel.shapefill#type)|Devuelve el tipo de relleno de la forma.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|Establece el formato de relleno de la forma en un color uniforme.|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|Especifica el porcentaje de transparencia del relleno como un valor de 0,0 (opaco) a 1,0 (claro).|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Indica el estado de negrita de la fuente.|
||[color](/javascript/api/excel/excel.shapefont#color)|Representación del código de color HTML del color del texto (por ejemplo, "#FF0000" representa el rojo).|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Indica el estado de cursiva de la fuente.|
||[name](/javascript/api/excel/excel.shapefont#name)|Representa el nombre de fuente (por ejemplo, "Calibri").|
||[size](/javascript/api/excel/excel.shapefont#size)|Representa el tamaño de fuente en puntos (por ejemplo, 11).|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Tipo de subrayado aplicado a la fuente.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Especifica el identificador de forma.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Devuelve el objeto de forma asociado con el grupo.|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Devuelve una colección de objetos de forma.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|Desagrupa las formas agrupadas en el grupo de formas especificado.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Representa el color de línea en formato de color HTML, con el formato #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|Indica el estilo de línea de la forma.|
||[estilo](/javascript/api/excel/excel.shapelineformat#style)|Indica el estilo de línea de la forma.|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Indica o establece el grado de transparencia de la línea especificada como valor entre 0,0 (opaco) y 1,0 (transparente).|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Especifica si el formato de línea de un elemento de forma está visible.|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Indica el grosor de la línea, en puntos.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|Especifica el subcampo que es el nombre de la propiedad de destino de un valor enriquecido para ordenar.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|Obtiene el número de estilos de la colección.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|Obtiene un estilo basándose en su posición en la colección.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autofilter)|Indica el objeto AutoFilter de la tabla.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|Obtiene el origen del evento.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|Obtiene el identificador de la tabla que se agrega.|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo a la que se agrega la tabla.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#details)|Obtiene la información sobre el detalle del cambio.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|Se produce cuando se agrega una nueva tabla en un libro.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|Se produce cuando se elimina la tabla especificada en un libro.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Obtiene el origen del evento.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|Obtiene el identificador de la tabla que se elimina.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|Obtiene el nombre de la tabla que se elimina.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se elimina la tabla.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|Obtiene el número de tablas de la colección.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|Obtiene la primera tabla de la colección.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|Obtiene una tabla por nombre o identificador.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autosizesetting)|La configuración automática de tamaño del marco de texto.|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|Indica el margen inferior, en puntos, del marco de texto.|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|Elimina todo el texto en el marco de texto.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|Indica la alineación horizontal del marco de texto.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|Indica el comportamiento de desbordamiento horizontal del marco de texto.|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|Indica el margen izquierdo, en puntos, del marco de texto.|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|Representa el ángulo al que está orientado el texto para el marco de texto.|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|Representa el orden de lectura del marco de texto, ya sea de izquierda a derecha o de derecha a izquierda.|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|Especifica si el marco de texto contiene texto.|
||[textRange](/javascript/api/excel/excel.textframe#textrange)|Representa el texto que hay unido a una forma en el marco de texto y las propiedades y los métodos de manipulación del texto.|
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|Indica el margen derecho, en puntos, del marco de texto.|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|Indica el margen superior, en puntos, del marco de texto.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|Indica la alineación vertical del marco de texto.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|Representa el comportamiento de desbordamiento vertical del marco de texto.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|Devuelve un objeto TextRange para la subcadena en el rango especificado.|
||[font](/javascript/api/excel/excel.textrange#font)|Devuelve un objeto ShapeFont que representa los atributos de fuente para el intervalo de texto.|
||[text](/javascript/api/excel/excel.textrange#text)|Indica el contenido de texto sin formato del intervalo de texto.|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|True si todos los gráficos en el libro están siguiendo los puntos de datos reales a los que están conectados.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|Obtiene el gráfico activo del libro.|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|Obtiene el gráfico activo del libro.|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|True si varios usuarios están editando el libro (coautoría).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|Obtiene los intervalos seleccionados actualmente en el libro.|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|Especifica si se han realizado cambios desde la última vez que se guardó el libro.|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|Especifica si el libro está en modo de guardado automático.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|Devuelve un número acerca de la versión del motor de cálculo de Excel.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|Se produce cuando se cambia la configuración de guardado automático en el libro.|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|Especifica si el libro se ha guardado localmente o en línea.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|True si los cálculos de este libro se llevan a cabo con la misma precisión con que se muestran los números.|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|Obtiene el tipo del evento.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|Determina si Excel debe volver a calcular la hoja de cálculo cuando sea necesario.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|Busca todas las repeticiones de la cadena especificada, según los criterios especificados y las devuelve como un objeto RangeAreas, que incluye uno o más intervalos rectangulares.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|Busca todas las repeticiones de la cadena especificada, según los criterios especificados y las devuelve como un objeto RangeAreas, que incluye uno o más intervalos rectangulares.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|Obtiene el objeto RangeAreas, que representa uno o varios bloques de intervalos rectangulares especificadas por la dirección o el nombre.|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|Indica el objeto AutoFilter de la hoja.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|Obtiene la colección de saltos de página horizontales de la hoja de cálculo.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|Se produce cuando se cambia el formato una hoja de cálculo concreta.|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|Obtiene el objeto PageLayout de la hoja de cálculo.|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|Devuelve la colección de todos los objetos Shape en la hoja de cálculo.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|Obtiene la colección de saltos de página verticales de la hoja de cálculo.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|Busca y reemplaza la cadena especificada, según los criterios especificados dentro de la hoja de cálculo actual.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#details)|Representa la información sobre el detalle del cambio.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|Se produce cuando se cambia una hoja de cálculo del libro.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|Se produce cuando cambia el formato de una hoja de cálculo del libro.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|Se produce cuando cambia la selección de cualquier hoja de cálculo.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Obtiene la dirección del intervalo que representa el área que ha cambiado en una hoja de cálculo específica.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|Obtiene el intervalo que representa el área que ha cambiado en una hoja de cálculo específica.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|Obtiene el intervalo que representa el área que ha cambiado en una hoja de cálculo específica.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Obtiene el origen del evento.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se cambian los datos.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|Especifica si la coincidencia debe ser completa o parcial.|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|Especifica si la coincidencia distingue mayúsculas de minúsculas.|
