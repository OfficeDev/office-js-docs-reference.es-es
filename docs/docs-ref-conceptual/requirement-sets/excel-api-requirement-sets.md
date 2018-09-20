# <a name="excel-javascript-api-requirement-sets"></a>Conjuntos de requisitos de la API de JavaScript de Excel

Los conjuntos de requisitos son grupos de miembros de la API con nombre. Complementos de Office utilizan conjuntos de requisitos especificados en el manifiesto o una comprobación en tiempo de ejecución para determinar si un host de Office admite las API que un complemento necesita. Para obtener más información, vea [conjuntos de versiones de Office y los requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Complementos de Excel se ejecutará a través de varias versiones de Office, incluidas 2016 de Office para Windows, Office para iPad, Office para Mac y Office Online. En la siguiente tabla se enumera los conjuntos de requisitos de Excel, las aplicaciones de host de Office que admiten cada conjunto de requisitos y las versiones de compilación o de número para esas aplicaciones.

> [!NOTE]
> Cualquier API que se marca como **Beta** no está preparada para producción para el usuario final. Se pone a disposición para desarrolladores tratar de desprotección en entornos de desarrollo y prueba. No están diseñados para usarse frente a documentos críticos para el negocio y producción.
> 
> Para los conjuntos de requisito que están marcados como **versión Beta**, utilice la versión especificada (o posterior) del software de Office y utilizar la biblioteca de Beta en la CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Las entradas no se marca como **versión Beta** están generalmente disponibles y se puede usar biblioteca de producción en la CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.

|  Conjunto de requisitos  |  Office 365 para profesionales de Windows\*  |  Office 365 para iPad  |  Office 365 para Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| Beta  | ¡Por favor, [visite nuestra página de especificación abierta de la API de JavaScript de Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)! |
| ExcelApi1.7  | Versión 1801 (compilación 9001.2171) o posterior| 2,9 o posterior | 16,9 o posterior | Abril de 2018 | Próximamente |
| ExcelApi1.6  | Versión 1704 (compilación 8201.2001) o posterior| 2.2 o posterior |15.36 o posterior| Abril de 2017 | Próximamente|
| ExcelApi1.5  | Versión 1703 (compilación 8067.2070) o posterior| 2.2 o posterior |15.36 o posterior| Marzo de 2017 | Próximamente|
| ExcelApi1.4 | Versión 1701 (compilación 7870.2024) o posterior| 2.2 o posterior |15.36 o posterior| Enero de 2017 | Próximamente|
| ExcelApi1.3  | Versión 1608 (compilación 7369.2055) o posterior| 1.27 o posterior |  15.27 o posterior| Septiembre de 2016 | Versión 1608 (compilación 7601.6800) o posterior|
| ExcelApi1.2  | Versión 1601 (compilación 6741.2088) o posterior | 1.21 o posterior | 15.22 o posterior| Enero de 2016 ||
| ExcelApi1.1  | Versión 1509 (compilación 4266.1001) o posterior | 1.19 o posterior | 15.20 o posterior| Enero de 2016 ||

> [!NOTE]
> El número de compilación de 2016 Office instalado a través de MSI es 16.0.4266.1001. Esta versión sólo contiene el conjunto de requisitos de ExcelApi 1.1.

Para obtener más información acerca de las versiones, números de compilación y el servidor de Office Online, vea:

- [Números de versión y compilación de las versiones del canal de actualización para los clientes de Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [¿Qué versión de Office estoy usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Dónde puede encontrar el número de versión y de compilación de una aplicación de cliente de Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- 
  [Información general de Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="whats-new-in-excel-javascript-api-17"></a>Novedades en Excel JavaScript API 1.7

Las características de conjunto 1.7 de requisito de API de JavaScript de Excel incluyen las API para gráficos, eventos, validación de datos, hojas de cálculo, rangos, propiedades del documento, denominados elementos, opciones de protección y estilos.

### <a name="customize-charts"></a>Personalizar los gráficos

Con el nuevo gráfico API, puede crear tipos de gráficos adicionales, agregar una serie de datos a un gráfico, establecer el título del gráfico, agregar un título de eje, agregar la unidad de presentación, agregue una línea de tendencia con Media móvil, cambiar una línea de tendencia lineal y mucho más. Los siguientes son algunos ejemplos:

* Eje de gráfico - obtener, establecer, dar formato y quitar unidades del eje, etiqueta y el título de un gráfico.
* Serie de gráficos - agregar, establecer y eliminar una serie de un gráfico.  Cambiar los marcadores de series, trazado y ajuste de tamaño.
* Las líneas de tendencia de gráfico - agregar, obtener y dar formato a las líneas de tendencia en un gráfico.
* Leyenda de gráfico - formato de la fuente de la leyenda de un gráfico.
* Punto del gráfico - establecer el color de punto de gráfico.
* Gráfico de subcadenas de título - obtener y establecer subcadena de título de un gráfico.
* Tipo de gráfico: opción para crear más tipos de gráficos.

### <a name="events"></a>Eventos

Se produce eventos de Excel API proporcionan una gran variedad de controladores de eventos que permiten el complemento para que se ejecute automáticamente una función designada cuando un evento específico. Puede dise?ar esa funci?n para realizar la acci?n que necesite el escenario. Para obtener una lista de eventos que están actualmente disponibles, vea [trabajar con eventos de uso de la API de JavaScript de Excel](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events).

### <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Personalizar la apariencia de las hojas de cálculo y rangos

Mediante las nuevas API, puede personalizar la apariencia de las hojas de cálculo de varias maneras:

* Inmovilizar paneles para mantener visibles filas o columnas específicas cuando se desplaza por la hoja de cálculo. Por ejemplo, si la primera fila de la hoja de cálculo contiene encabezados, es posible que inmovilizar esa fila para que los encabezados de columna permanecerá visibles mientras se desplaza hacia abajo de la hoja de cálculo.
* Modificar el color de la ficha de hoja de cálculo.
* Agregar encabezados de hoja de cálculo.


Puede personalizar la apariencia de rangos de varias maneras:

* Establecer el estilo de celda de un rango para asegurarse de asegurarse de que todas las celdas del rango tienen un formato coherente. Un estilo de celda es un conjunto definido de características, como las fuentes y tamaños de fuente, formatos de número, bordes de celda y sombreado de celda de formato. Usar cualquiera de los estilos de celda integrados de Excel o crear sus propios estilos de celda personalizado.
* Establecer la orientación del texto para un rango.
* Agregar o modificar un hipervínculo en un rango que se vincula a otra ubicación en el libro o a una ubicación externa.

### <a name="manage-document-properties"></a>Administrar las propiedades de documento

Mediante las API de propiedades de documento, puede obtener acceso a propiedades de documento integradas y también crear y administrar las propiedades de documento personalizadas para almacenar el estado del libro y la unidad de flujo de trabajo y la lógica empresarial.

### <a name="copy-worksheets"></a>Copie las hojas de cálculo

Con las API de copia de hoja de cálculo, puede copiar los datos y el formato de una hoja de cálculo a una hoja de cálculo nueva dentro del mismo libro y reducir la cantidad de transferencia de datos necesitada.

### <a name="handle-ranges-with-ease"></a>Controlar los intervalos con facilidad

Con el rango diversas API, puede hacer cosas como get en la región que lo rodea y obtener un intervalo cuyo tamaño ha cambiado y mucho más. Estas API deben realizar tareas como la manipulación de rango y hacer frente a mucho más eficaz.

Además:

* Opciones de protección de libro y hoja de cálculo: use estas API para proteger los datos en una hoja de cálculo y la estructura del libro.
* Actualizar un elemento con nombre: utilizar esta API para actualizar un elemento con nombre.
* Obtener celda activa: utilizar esta API para obtener la celda activa de un libro.

|Objeto| Novedades| Descripción|Conjunto de requisitos|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Propiedad_ > chartType|Representa el tipo del gráfico. Los valores posibles son: ColumnClustered, columnas, Apiladas100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etcetera...|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Propiedad_ > id|Identificador único del gráfico. Solo lectura.|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Propiedad_ > showAllFieldButtons|Representa si se debe mostrar todos los botones de campo de un gráfico dinámico.|1.7|
|[chartAreaFormat](/javascript/api/excel/excel.chartareaformat)|_Relación_ > borde|Representa el formato de borde del área del gráfico, que incluye el color, linestyle y weight. Solo lectura.|1.7|
|[chartAxes](/javascript/api/excel/excel.chartaxes)|_Método_ > getItem (tipo: string, de grupo: cadena)|Devuelve el eje específico identificado por tipo y grupo.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > axisBetweenCategories|Representa si el eje de valores cruza el eje de categorías entre categorías.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > axisGroup|Representa el grupo del eje especificado. Solo lectura. Los valores posibles son: principal, secundario.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > categoryType|Devuelve o establece el tipo del eje de categorías. Los valores posibles son: automático, TextAxis, DateAxis.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > cruza|Representa el eje especificado donde se cruza el otro eje. Los valores posibles son: automática, máximo, mínimo, personalizado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > crossesAt|Representa el eje especificado donde se cruza el otro eje en. Solo lectura. Conjunto a esta propiedad debe utilizar el método de SetCrossesAt(double). Solo lectura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > customDisplayUnit|Representa el valor de unidad de visualización de eje personalizado. Solo lectura. Para establecer esta propiedad, utilice el método SetCustomDisplayUnit(double). Solo lectura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > displayUnit|Representa la unidad de presentación del eje. Los valores posibles son: ninguno, cientos, miles, TenThousands, HundredThousands, millones, TenMillions, HundredMillions, miles de millones, trillones, personalizado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > height|Representa el alto, en puntos, del eje del gráfico. Null si el eje no s visible. Solo lectura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > izquierdo|Representa la distancia, en puntos, desde el borde izquierdo del eje a la izquierda del área del gráfico. Null si el eje no s visible. Solo lectura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > logBase|Representa la base del logaritmo cuando se usa una escala logarítmica.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > reversePlotOrder|Indica si Microsoft Excel traza los puntos de datos desde el último al primero.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > scaleType|Representa el tipo de escala del eje de valor. Los valores posibles son: logarítmica, lineal.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > showDisplayUnitLabel|Representa si está visible el rótulo de unidades de presentación del eje.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > tickLabelSpacing|Representa el número de categorías o series entre rótulos de marcas de graduación. Puede ser un valor de 1 31999 a o una cadena vacía para la configuración automática. El valor devuelto siempre es un número.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > tickMarkSpacing|Representa el número de categorías o series entre marcas de graduación.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > parte superior|Representa la distancia, en puntos, desde el borde superior del eje a la parte superior del área del gráfico. Null si el eje no s visible. Solo lectura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > tipo|Representa el tipo de eje. Solo lectura. Los valores posibles son: válida, categoría, valor, serie.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > visible|Un valor de tipo boolean representa la visibilidad del eje.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propiedad_ > width|Representa el ancho, en puntos, del eje del gráfico. Null si el eje no s visible. Solo lectura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relación_ > baseTimeUnit|Devuelve o establece la unidad base del eje de categorías especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relación_ > majorTickMark|Representa el tipo de marcas de graduación principales del eje especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relación_ > majorTimeUnitScale|Devuelve o establece el valor de la escala de unidades principales del eje de categorías cuando la propiedad CategoryType está establecida en escala temporal.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relación_ > minorTickMark|Representa el tipo de marcas de graduación secundarias del eje especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relación_ > minorTimeUnitScale|Devuelve o establece el valor de la escala de unidades secundarias del eje de categorías cuando la propiedad CategoryType está establecida en escala temporal.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relación_ > tickLabelPosition|Representa la posición de los rótulos de marcas de graduación en el eje especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Método_ > setCategoryNames(sourceData: Range)|Establece todos los nombres de categoría para el eje especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Método_ > setCrossesAt(value: double)|Establezca el eje especificado donde se cruza el otro eje en.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Método_ > setCustomDisplayUnit(value: double)|Establece la unidad de presentación del eje en un valor personalizado.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Propiedad_ > color|Código de color HTML que representa el color de los bordes del gráfico.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Propiedad_ > peso|Representa el grosor del borde, en puntos.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Relación_ > lineStyle|Representa el estilo de línea del borde.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propiedad_ > posición|Valor DataLabelPosition que representa la posición de la etiqueta de datos. Los valores posibles son: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propiedad_ > separador|Cadena que representa el separador utilizado para el rótulo de datos en un gráfico.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propiedad_ > showBubbleSize|Valor booleano que representa si el tamaño de la burbuja de la etiqueta de datos es visible o no.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propiedad_ > showCategoryName|Valor booleano que representa si el nombre de categoría de la etiqueta de datos es visible o no.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propiedad_ > showLegendKey|Valor booleano que representa si la clave de leyenda de la etiqueta de datos es visible o no.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propiedad_ > showPercentage|Valor booleano que representa si el porcentaje de la etiqueta de datos es visible o no.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propiedad_ > showSeriesName|Valor booleano que representa si el nombre de serie de la etiqueta de datos es visible o no.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propiedad_ > showValue|Valor booleano que representa si el valor de la etiqueta de datos es visible o no.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propiedad_ > height|Representa el alto de la leyenda del gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propiedad_ > izquierdo|Representa la izquierda de la leyenda de un gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propiedad_ > showShadow|Representa si la leyenda tiene sombra en el gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propiedad_ > parte superior|Representa la parte superior de la leyenda de un gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propiedad_ > width|Representa el ancho de la leyenda del gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Relación_ > legendEntries|Representa una colección de legendEntries en la leyenda. Solo lectura.|1.7|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propiedad_ > visible|Representa la visible de un elemento de leyenda del gráfico.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Propiedad_ > items|Una colección de objetos chartLegendEntry. Solo lectura.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Método_ > getCount()|Devuelve el número de legendEntry de la colección.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Método_ > getItemAt(index: number)|Devuelve un legendEntry en el índice especificado.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Propiedad_ > hasDataLabel|Representa si un punto de datos tiene datalabel. No es aplicable para los gráficos de superficie.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Propiedad_ > markerBackgroundColor|Punto de representación de código de color HTML del color de fondo de marcador de datos. Por ejemplo, #FF0000 representa el rojo.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Propiedad_ > markerForegroundColor|Punto de representación de código de color HTML del color de primer plano del marcador de datos. Por ejemplo, #FF0000 representa el rojo.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Propiedad_ > markerSize|Representa el tamaño de los marcadores de punto de datos.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Propiedad_ > markerStyle|Representa el estilo del marcador de un punto de datos del gráfico. Los valores posibles son: Invalid, automático, ninguno, cuadrado, rombo, triángulo, X, estrellas, puntos, guión, Circle, además, de imagen.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Relación_ > dataLabel|Devuelve la etiqueta de datos de un punto del gráfico. Solo lectura.|1.7|
|[chartPointFormat](/javascript/api/excel/excel.chartpointformat)|_Relación_ > borde|Representa el formato de borde de un punto de datos del gráfico, que incluye información de color, estilo y el grosor. Solo lectura.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propiedad_ > chartType|Representa el tipo de gráfico de una serie. Los valores posibles son: ColumnClustered, columnas, Apiladas100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etcetera...|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propiedad_ > doughnutHoleSize|Representa el tamaño del agujero anillos de una serie de gráficos.  Sólo es válido en los gráficos de anillos y doughnutExploded.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propiedad_ > filtrados|Valor de tipo Boolean que representa si la serie se ha filtrado o no. No es aplicable para los gráficos de superficie.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propiedad_ > gapWidth|Representa el ancho del rango de una serie de gráficos.  Sólo es válido en gráficos de barras y columna, así como|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propiedad_ > hasDataLabels|Valor de tipo Boolean que representa si la serie tiene rótulos de datos o no.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propiedad_ > markerBackgroundColor|Representa el color de fondo de los marcadores de una serie de gráficos.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propiedad_ > markerForegroundColor|Representa el color de primer plano de los marcadores de una serie de gráficos.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propiedad_ > markerSize|Representa el tamaño de los marcadores de una serie de gráficos.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propiedad_ > markerStyle|Representa el estilo del marcador de una serie de gráficos. Los valores posibles son: Invalid, automático, ninguno, cuadrado, rombo, triángulo, X, estrellas, puntos, guión, Circle, además, de imagen.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propiedad_ > plotOrder|Representa el orden de trazado de una serie de gráficos en el grupo de gráficos.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propiedad_ > showShadow|Valor de tipo Boolean que representa si la serie tiene sombra o no.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propiedad_ > suave|Valor de tipo Boolean que representa si la serie es suave o no. Sólo para los gráficos de líneas y de dispersión.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relación_ > dataLabels|Representa una colección de todos los dataLabels de la serie. Solo lectura.|ApiSet.InProgressFeatures.ChartingAPI|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relación_ > trendlines|Representa una colección de líneas de tendencia de la serie. Solo lectura.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > delete()|Elimina la serie del gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > setBubbleSizes(sourceData: Range)|Establecer los tamaños de burbuja para una serie de gráficos. Sólo funciona para los gráficos de burbujas.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > setValues(sourceData: Range)|Establezca los valores de una serie de gráficos. Gráfico de dispersión, significa que los valores del eje Y.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > setXAxisValues(sourceData: Range)|Establecer los valores de X del eje de una serie de gráficos. Sólo funciona para los gráficos de dispersión.|1.7|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Método_ > Agregar (nombre: string, index: número)|Agregar una nueva serie a la colección.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propiedad_ > height|Devuelve el alto, en puntos, del título del gráfico. Solo lectura. Null si no s visible el título del gráfico. Solo lectura.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propiedad_ > horizontalAlignment|Representa la alineación horizontal para el título del gráfico. Los valores posibles son: centro, izquierda, justificar, distribuida, a la derecha.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propiedad_ > izquierdo|Representa la distancia, en puntos, desde el borde izquierdo del título del gráfico hasta el borde izquierdo del área del gráfico. Null si no s visible el título del gráfico.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propiedad_ > posición|Representa la posición del título del gráfico. Los valores posibles son: superior, automático, inferior, izquierda, derecha.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propiedad_ > showShadow|Representa un valor boolean que determina si el título del gráfico tiene una sombra.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propiedad_ > textOrientation|Representa la orientación del texto del título del gráfico. El valor debe ser un número entero, ya sea desde -90 a 90 o 180 orientados verticalmente a texto.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propiedad_ > parte superior|Representa la distancia, en puntos, desde el borde superior del título del gráfico a la parte superior del área del gráfico. Null si no s visible el título del gráfico.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propiedad_ > verticalAlignment|Representa la alineación vertical del título del gráfico. Los valores posibles son: centro, parte inferior, superior, Justify, distribuida.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propiedad_ > width|Devuelve el ancho, en puntos, del título del gráfico. Solo lectura. Null si no s visible el título del gráfico. Solo lectura.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Método_ > setFormula(formula: string)|Establece un valor de tipo string que representa la fórmula del título del gráfico mediante la notación de estilo A1.|1.7|
|[chartTitleFormat](/javascript/api/excel/excel.charttitleformat)|_Relación_ > borde|Representa el formato de borde del título de gráfico, que incluye el color, linestyle y weight. Solo lectura.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propiedad_ > con versiones anteriores|Representa el número de períodos que la línea de tendencia se extiende hacia atrás.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propiedad_ > displayEquation|True si la ecuación de la línea de tendencia se muestra en el gráfico.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propiedad_ > displayRSquared|True si la R cuadrado de la línea de tendencia se muestra en el gráfico.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propiedad_ > hacia delante|Representa el número de períodos que la línea de tendencia se extiende hacia delante.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propiedad_ > intersección|Representa el valor de la intersección de la línea de tendencia. Se puede establecer en un valor numérico o una cadena vacía (para valores automáticas). El valor devuelto siempre es un número.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propiedad_ > movingAveragePeriod|Representa el período de una línea de tendencia de gráfico, sólo de línea de tendencia con tipo de Media móvil.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propiedad_ > name|Representa el nombre de la línea de tendencia. Se puede establecer en un valor de tipo string o se puede establecer en los valores de valor nulo representa automática. El valor devuelto siempre es una cadena|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propiedad_ > polynomialOrder|Representa el orden de una línea de tendencia de gráfico, sólo de la línea de tendencia con tipo polinomial.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propiedad_ > tipo|Representa el tipo de una línea de tendencia de gráfico. Los valores posibles son: lineal, exponencial, logarítmica, Media móvil, polinomial, Power.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Relación_ > formato|Representa el formato de una línea de tendencia de gráfico. Solo lectura.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Método_ > delete()|Eliminar el objeto de línea de tendencia.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Propiedad_ > items|Una colección de objetos chartTrendline. Solo lectura.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Método_ > add(type: string)|Agrega una nueva línea de tendencia a la colección de línea de tendencia.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Método_ > getCount()|Devuelve el número de líneas de tendencia de la colección.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Método_ > getItem(index: number)|Obtenga el objeto trendline por index, que es el orden de inserción en la matriz de elementos.|1.7|
|[chartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|_Relación_ > línea|Representa el formato de línea de gráfico. Solo lectura.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Propiedad_ > key|Obtiene la clave de la propiedad personalizada. Solo lectura. Solo lectura.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Propiedad_ > tipo|Obtiene el tipo de valor de la propiedad personalizada. Solo lectura. Solo lectura. Los valores posibles son: número, Boolean, Date, String, Float.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Propiedad_ > value|Obtiene o establece el valor de la propiedad personalizada.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Método_ > delete()|Elimina la propiedad personalizada.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Propiedad_ > items|Una colección de objetos customProperty. Solo lectura.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > Agregar (clave: valor de la cadena,: objeto)|Crea una nueva propiedad personalizada o establece una existente.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > deleteAll()|Elimina todas las propiedades personalizadas de la colección.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > getCount()|Obtiene el recuento de las propiedades personalizadas.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > getItem(key: string)|Obtiene un objeto de propiedad personalizada mediante su clave, que no distingue mayúsculas de minúsculas. Se genera si la propiedad personalizada no existe.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > getItemOrNullObject(key: string)|Obtiene un objeto de propiedad personalizada mediante su clave, que no distingue mayúsculas de minúsculas. Devuelve un objeto null si la propiedad personalizada no existe.|1.7|
|[Datasourcecollection](/javascript/api/excel/excel.dataconnectioncollection)|_Propiedad_ > items|Una colección de objetos de dataConnection. Solo lectura.|1.7|
|[Datasourcecollection](/javascript/api/excel/excel.dataconnectioncollection)|_Método_ > refreshAll()|Actualiza todas las conexiones de datos en la colección.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propiedad_ > author|Obtiene o establece al autor del libro.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propiedad_ > category|Obtiene o establece la categoría del libro.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propiedad_ > comments|Obtiene o establece los comentarios del libro.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propiedad_ > company|Obtiene o establece la compañía del libro.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propiedad_ > keywords|Obtiene o establece las palabras clave del libro.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propiedad_ > lastAuthor|Obtiene al último autor del libro. Solo lectura. Solo lectura.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propiedad_ > manager|Obtiene o establece el administrador del libro.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propiedad_ > revisionNumber|Obtiene el número de revisión del libro. Solo lectura.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propiedad_ > subject|Obtiene o establece al asunto del libro.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propiedad_ > title|Obtiene o establece el título del libro.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Relación_ > creationDate|Obtiene la fecha de creación del libro. Solo lectura. Solo lectura.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Relación_ > personalizado|Obtiene la colección de propiedades personalizadas del libro. Solo lectura. Solo lectura.|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Propiedad_ > fórmula|Obtiene o establece la fórmula del elemento con nombre.  Fórmula siempre comienza con un signo '='.|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relación_ > arrayValues|Devuelve un objeto que contiene los valores y los tipos de elemento con nombre. Solo lectura.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Propiedad_ > tipos|Representa los tipos para cada elemento de la matriz de elemento con nombre de sólo lectura. Los valores posibles son: desconocido, vacío, String, Integer, Double, Boolean, Error.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Propiedad_ > values|Representa los valores de cada elemento de la matriz de elemento con nombre. Solo lectura.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propiedad_ > isEntireColumn|Representa si el intervalo actual es una columna completa. Solo lectura.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propiedad_ > isEntireRow|Representa si el intervalo actual es una fila completa. Solo lectura.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propiedad_ > numberFormatLocal|Representa el código de formato numérico de Excel para el rango determinado como una cadena en el idioma del usuario.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propiedad_ > style|Representa el estilo del rango actual. Esto devuelve null o una cadena.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > getAbsoluteResizedRange (numRows: número, numColumns: número)|Obtiene un objeto Range con la misma celda superior izquierda como el objeto Range actual, pero con los números de filas y columnas especificados.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > getImage()|El intervalo se representa como una imagen con codificación base64.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > getSurroundingRegion()|Devuelve un objeto Range que representa la región que lo rodea de la celda superior izquierda de este intervalo. Una región que lo rodea es un rango limitado por cualquier combinación de filas en blanco y columnas en blanco con respecto a este intervalo.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > showCard()|Muestra la tarjeta de una celda activa si tiene contenido enriquecido valor.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propiedad_ > textOrientation|Obtiene o establece la orientación del texto de todas las celdas dentro del intervalo.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propiedad_ > useStandardHeight|Determina si el alto de fila del objeto Range es igual al alto estándar de la hoja.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propiedad_ > useStandardWidth|Determina si columnwidth del objeto Range es igual al ancho estándar de la hoja.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propiedad_ > address|Representa la dirección url de destino para el hipervínculo.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propiedad_ > documento..|Representa el documento.. destino para el hipervínculo.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propiedad_ > información en pantalla|Representa la cadena que se muestra al colocar el puntero sobre el hipervínculo.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propiedad_ > textToDisplay|Representa la cadena que se muestra en la parte superior izquierda de la mayoría de celda del rango.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > addIndent|Indica si se aplica sangría al texto automáticamente cuando la alineación del texto en una celda se establece en una distribución igual.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > sangría automática|Indica si se aplica sangría al texto automáticamente cuando la alineación del texto en una celda se establece en una distribución igual.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > builtIn|Indica si el estilo es un estilo integrado. Solo lectura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > formulaHidden|Indica si la fórmula se ocultará cuando la hoja de cálculo está protegida.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > horizontalAlignment|Representa la alineación horizontal para el estilo. Los valores posibles son: General, izquierda, centro, derecha, relleno, Justify, CenterAcrossSelection, distribuida.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > includeAlignment|Indica si el estilo incluye las propiedades de sangría automática, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel y TextOrientation.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > includeBorder|Indica si el estilo incluye las propiedades de borde Color, ColorIndex, LineStyle y Weight.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > includeFont|Indica si el estilo incluye las propiedades de fuente de fondo, Bold, Color, ColorIndex, FontStyle, Italic, nombre, tamaño, tachado, subíndice, superíndice y subrayado.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > includeNumber|Indica si el estilo incluye la propiedad NumberFormat.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > includePatterns|Indica si el estilo incluye las propiedades de interior Color, ColorIndex, InvertIfNegative, patrón, PatternColor y PatternColorIndex.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > includeProtection|Indica si el estilo incluye las propiedades de protección FormulaHidden y Locked.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > indentLevel|Entero de 0 a 250 que indica el nivel de sangría del estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > locked|Indica si el objeto está bloqueado cuando la hoja de cálculo está protegida.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > name|El nombre del estilo. Solo lectura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > numberFormat|El código de formato del formato de número para el estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > numberFormatLocal|El código de formato localizado el formato de número para el estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > orientación|La orientación del texto para el estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > readingOrder|El orden de lectura para el estilo. Los valores posibles son: contexto, LeftToRight, RightToLeft.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > shrinkToFit|Indica si el texto se reduce automáticamente para ajustarse al ancho de columna disponible.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > textOrientation|La orientación del texto para el estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > verticalAlignment|Representa la alineación vertical para el estilo. Los valores posibles son: superior, centro, inferior, Justify, distribuida.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propiedad_ > wrapText|Indica si Microsoft Excel ajusta el texto en el objeto.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relación_ > bordes|Una colección de borde de cuatro objetos Border que representan el estilo de los cuatro bordes. Solo lectura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relación_ > relleno|El relleno del estilo. Solo lectura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relación_ > font|Un objeto Font que representa la fuente del estilo. Solo lectura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Método_ > delete()|Elimina este estilo.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Propiedad_ > items|Una colección de objetos de estilo. Solo lectura.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Método_ > add(name: string)]|Agrega un nuevo estilo a la colección.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Método_ > getItem(name: string)|Obtiene un estilo por su nombre.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propiedad_ > address|Obtiene la dirección que representa el área que ha cambiado de una tabla en una hoja de cálculo específica.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propiedad_ > changeType|Obtiene el tipo de cambio que representa cómo se desencadena el evento Changed. Los valores posibles son: otras personas, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propiedad_ > origen|Obtiene el origen del evento. Los valores posibles son: Local y remota.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propiedad_ > tableId|Obtiene el identificador de la tabla en la que ha cambiado los datos.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propiedad_ > tipo|Obtiene el tipo del evento. Los valores posibles son: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propiedad_ > worksheetId|Obtiene el identificador de la hoja de cálculo en la que ha cambiado los datos.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propiedad_ > address|Obtiene la dirección del rango que representa el área seleccionada de la tabla en una hoja de cálculo específica.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propiedad_ > isInsideTable|Indica si la selección se encuentra dentro de una tabla, dirección será inútil si IsInsideTable es false.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propiedad_ > tableId|Obtiene el identificador de la tabla en la que ha cambiado la selección.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propiedad_ > tipo|Obtiene el tipo del evento. Los valores posibles son: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propiedad_ > worksheetId|Obtiene el identificador de la hoja de cálculo en la que ha cambiado la selección.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Propiedad_ > name|Obtiene el nombre del libro. Solo lectura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relación_ > dataConnections|Actualiza todas las conexiones de datos en el libro. Solo lectura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relación_ > properties|Obtiene las propiedades de libro. Solo lectura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relación_ > protection|Devuelve el objeto de la protección de libro de un libro. Solo lectura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relación_ > estilos|Representa una colección de estilos asociada con el libro. Solo lectura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Método_ > getActiveCell()|Obtiene la celda actualmente activa del libro.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Propiedad_ > protected|Indica si el libro está protegido. Solo lectura. Solo lectura.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Método_ > protect(password: string)|Protege un libro. Se produce un error si el libro se ha protegido.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Método_ > unprotect(password: string)|Desprotege un libro.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propiedad_ > líneas de cuadrícula|Obtiene o establece la marca de las líneas de cuadrícula de la hoja de cálculo.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propiedad_ > encabezados|Obtiene o establece la marca de los títulos de la hoja de cálculo.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propiedad_ > showHeadings|Obtiene o establece la marca de los títulos de la hoja de cálculo.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propiedad_ > standardHeight|Devuelve el alto estándar (valor predeterminado) de todas las filas de la hoja de cálculo, en puntos. Solo lectura.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propiedad_ > standardWidth|Devuelve o establece el ancho estándar (valor predeterminado) de todas las columnas de la hoja de cálculo.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propiedad_ > tabColor|Obtiene o establece el color de la ficha de hoja de cálculo.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relación_ > freezePanes|Obtiene un objeto que se puede usar para manipular los paneles inmovilizados en la hoja de cálculo de sólo lectura.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > copia (positionType: WorksheetPositionType, relativeTo: hoja de cálculo)|Copiar una hoja de cálculo y colocarlo en la posición especificada. Devolver la hoja de cálculo copiada.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getRangeByIndexes (startRow: número, startColumn: número, rowCount: columnCount número,: número)| Obtiene el objeto de intervalo comenzando en un índice de columna y fila determinado, y que abarca un determinado número de filas y columnas.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Propiedad_ > tipo|Obtiene el tipo del evento. Los valores posibles son: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Propiedad_ > worksheetId|Obtiene el identificador de la hoja de cálculo que se activa.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Propiedad_ > origen|Obtiene el origen del evento. Los valores posibles son: Local y remota.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Propiedad_ > tipo|Obtiene el tipo del evento. Los valores posibles son: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Propiedad_ > worksheetId|Obtiene el identificador de la hoja de cálculo que se agrega al libro.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propiedad_ > address|Obtiene la dirección del rango que representa el área que ha cambiado de una hoja de cálculo específica.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propiedad_ > changeType|Obtiene el tipo de cambio que representa cómo se desencadena el evento Changed. Los valores posibles son: otras personas, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propiedad_ > origen|Obtiene el origen del evento. Los valores posibles son: Local y remota.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propiedad_ > tipo|Obtiene el tipo del evento. Los valores posibles son: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propiedad_ > worksheetId|Obtiene el identificador de la hoja de cálculo en la que ha cambiado los datos.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Propiedad_ > tipo|Obtiene el tipo del evento. Los valores posibles son: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Propiedad_ > worksheetId|Obtiene el identificador de la hoja de cálculo que se desactiva.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Propiedad_ > origen|Obtiene el origen del evento. Los valores posibles son: Local y remota.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Propiedad_ > tipo|Obtiene el tipo del evento. Los valores posibles son: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Propiedad_ > worksheetId|Obtiene el identificador de la hoja de cálculo que se elimina del libro.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > freezeAt (frozenRange: rango o cadena)|Establece las celdas inmovilizadas en la vista de hoja de cálculo activa.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > freezeColumns(count: number)|Inmovilizar las columnas de la primera de la hoja de cálculo en su lugar.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > freezeRows(count: number)|Inmovilizar las filas superiores de la hoja de cálculo en su lugar.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > getLocation()|Obtiene un rango que describa las celdas inmovilizadas en la vista de hoja de cálculo activa.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > getLocationOrNullObject()|Obtiene un rango que describa las celdas inmovilizadas en la vista de hoja de cálculo activa.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > unfreeze()|Quita todos los paneles inmovilizados en la hoja de cálculo.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propiedad_ > allowEditObjects|Representa la opción de permitir la edición de objetos de protección de hoja de cálculo.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propiedad_ > allowEditScenarios|Representa la opción de permitir la edición de escenarios de protección de hoja de cálculo.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Relación_ > selectionMode|Representa la opción de protección de la hoja de cálculo del modo de selección.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Propiedad_ > address|Obtiene la dirección del rango que representa el área seleccionada de una hoja de cálculo específica.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Propiedad_ > tipo|Obtiene el tipo del evento. Los valores posibles son: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Propiedad_ > worksheetId|Obtiene el identificador de la hoja de cálculo en la que ha cambiado la selección.|1.7|


## <a name="whats-new-in-excel-javascript-api-16"></a>Novedades en Excel JavaScript API 1.6 

### <a name="conditional-formatting"></a>Formato condicional

Presenta un formato condicional de un intervalo. Permite a los siguientes tipos de formato condicional:

* Escala de colores
* Barra de datos
* Conjunto de iconos
* Personalizado

Además:

* Devuelve el rango a que el formato condicional se aplica. 
* Eliminación del formato condicional. 
* Proporciona la capacidad de prioridad y stopifTrue. 
* Obtiene la colección de todo el formato condicional de un intervalo determinado. 
* Borra todos los formatos condicionales activos en el intervalo actual especificado. 

|Objeto| Novedades| Descripción|Conjunto de requisitos|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Método_ > suspendApiCalculationUntilNextSync()|Suspende el cálculo hasta que se llama al siguiente "context.sync()". Una vez establecido, será responsabilidad del desarrollador actualizar el libro para asegurarse de que se propaguen las dependencias.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Relación_ > formato|Devuelve un objeto de formato que encapsula la fuente de formatos condicionales, el relleno, los bordes y otras propiedades. Solo lectura.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Relación_ > regla|Representa el objeto Rule en este formato condicional.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Propiedad_ > threeColorScale|Si es true, la escala de colores tendrá tres puntos (mínimo, punto medio y máximo); de lo contrario, tendrá dos (mínimo y máximo). Solo lectura.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Relación_ > criteria|Criterios de la escala de colores. El punto medio es opcional cuando se usa una escala de colores de dos puntos.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Propiedad_ > formula1|La fórmula, si es necesaria, en la que se debe evaluar la regla de formato condicional.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Propiedad_ > formula2|La fórmula, si es necesaria, en la que se debe evaluar la regla de formato condicional.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Propiedad_ > operator|Operador del formato condicional de texto. Los valores posibles son Invalid, Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqual y LessThanOrEqual.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relación_ > máximo|Criterio de escala de colores de punto máximo.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relación_ > punto medio|Criterio de escala de colores de punto medio si la escala de colores es una escala de 3 colores.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relación_ > mínimo|Criterio de escala de colores de punto mínimo.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Propiedad_ > color|Representación del código de color HTML del color de la escala de colores. Por ejemplo, #FF0000 representa el rojo.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Propiedad_ > fórmula|Número, fórmula o valor null (si el tipo es LowestValue).|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Propiedad_ > tipo|Elemento en el que se debe basar la fórmula condicional de icono. Los valores posibles son Invalid, LowestValue, HighestValue, Number, Percent, Formula y Percentile.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propiedad_ > borderColor|Código de color HTML que representa el color de la línea de borde con el formato #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propiedad_ > fillColor|Código de color HTML que representa el color de relleno, con el formato #RRGGBB (p. ej. "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propiedad_ > matchPositiveBorderColor|Representación booleana de si la barra de datos negativa tiene o no el mismo color de borde que la barra de datos positiva.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propiedad_ > matchPositiveFillColor|Representación booleana de si la barra de datos negativa tiene o no el mismo color de relleno que la barra de datos positiva.|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Propiedad_ > borderColor|Código de color HTML que representa el color de la línea de borde con el formato #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Propiedad_ > fillColor|Código de color HTML que representa el color de relleno, con el formato #RRGGBB (p. ej. "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Propiedad_ > gradientFill|Representación booleana de si la barra de datos tiene o no un degradado.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Propiedad_ > fórmula|La fórmula, si es necesaria, en la que se debe evaluar la regla databar.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Propiedad_ > tipo|Tipo de regla para la barra de datos. Los valores posibles son LowestValue, HighestValue, Number, Percent, Formula, Percentile y Automatic.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propiedad_ > id|La prioridad del formato condicional dentro de la ConditionalFormatCollection actual. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propiedad_ > prioridad|La prioridad (o índice) dentro de la colección de formato condicional en la que este formato condicional existe actualmente. Al cambiar esto, también|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propiedad_ > stopIfTrue|Si se cumplen las condiciones de este formato condicional, los formatos de menor prioridad no surtirán efecto en esa celda.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propiedad_ > tipo|Un tipo de formato condicional. Solo se puede establecer al mismo tiempo. Solo lectura. Solo lectura. Los valores posibles son: Custom, DataBar, ColorScale, IconSet.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > cellValue|Devuelve las propiedades del formato condicional del valor de la celda si el formato condicional actual es un tipo de CellValue. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > cellValueOrNullObject|Devuelve las propiedades del formato condicional del valor de la celda si el formato condicional actual es un tipo de CellValue. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > colorScale|Devuelve las propiedades del formato condicional ColorScale si el formato condicional actual es un tipo de ColorScale. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > colorScaleOrNullObject|Devuelve las propiedades del formato condicional ColorScale si el formato condicional actual es un tipo de ColorScale. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > personalizado|Devuelve las propiedades del formato condicional personalizadas si el formato condicional actual es un tipo personalizado. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > customOrNullObject|Devuelve las propiedades del formato condicional personalizadas si el formato condicional actual es un tipo personalizado. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > dataBar|Devuelve las propiedades de barra de datos si el formato condicional actual es una barra de datos. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > dataBarOrNullObject|Devuelve las propiedades de barra de datos si el formato condicional actual es una barra de datos. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > iconSet|Devuelve las propiedades del formato condicional IconSet si el formato condicional actual es un tipo IconSet. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > iconSetOrNullObject|Devuelve las propiedades del formato condicional IconSet si el formato condicional actual es un tipo IconSet. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > preestablecida|Devuelve el formato condicional de criterios preestablecidos como las propiedades mencionadas anteriormente averagebelow averageunique valuescontains blanknonblankerrornoerror. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > presetOrNullObject|Devuelve el formato condicional de criterios preestablecidos como las propiedades mencionadas anteriormente averagebelow averageunique valuescontains blanknonblankerrornoerror. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > textComparison|Devuelve las propiedades del formato condicional del texto especificado si el formato condicional actual es un tipo de texto. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > textComparisonOrNullObject|Devuelve las propiedades del formato condicional del texto especificado si el formato condicional actual es un tipo de texto. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > arriba o abajo|Devuelve las propiedades del formato condicional TopBottom si el formato condicional actual es un tipo TopBottom. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relación_ > topBottomOrNullObject|Devuelve las propiedades del formato condicional TopBottom si el formato condicional actual es un tipo TopBottom. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Método_ > delete()|Elimina este formato condicional.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Método_ > getRange()|Devuelve el intervalo al que se aplica el formato condicional o un objeto NULL si el rango es discontinuo. Solo lectura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Método_ > getRangeOrNullObject()|Devuelve el intervalo al que se aplica el formato condicional o un objeto NULL si el rango es discontinuo. Solo lectura.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Propiedad_ > items|Colección de objetos conditionalFormat. Solo lectura.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > add(type: string)|Agrega un nuevo formato condicional a la colección en la prioridad firsttop.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > clearAll()|Borra todos los formatos condicionales activos en el intervalo actual especificado.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > getCount()|Devuelve el número de formatos condicionales del libro. Solo lectura.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > getItem(id: string)|Devuelve un formato condicional para el identificador determinado.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > getItemAt(index: number)|Devuelve un formato condicional en el índice especificado.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Propiedad_ > fórmula|La fórmula, si es necesaria, en la que se debe evaluar la regla de formato condicional.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Propiedad_ > formulaLocal|La fórmula, si es necesaria, en la que se debe evaluar la regla de formato condicional en el idioma del usuario.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Propiedad_ > formulaR1C1|La fórmula, si es necesaria, en la que se debe evaluar la regla de formato condicional en la notación de estilo R1C1.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Propiedad_ > fórmula|Número o fórmula en función del tipo.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Propiedad_ > operator|Operador GreaterThan o GreaterThanOrEqual para cada tipo de regla para el formato condicional de icono. Los valores posibles son Invalid, GreaterThan y GreaterThanOrEqual.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Relación_ > customIcon|Icono personalizado para el criterio actual si es diferente del IconSet predeterminado; si no, se devolverá null.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Relación_ > type|Elemento en el que se debe basar la fórmula condicional de icono.|1.6|
|[conditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|_Propiedad_ > criterio|criterio del formato condicional. Los valores posibles son Invalid, Blanks, NonBlanks, Errors, NonErrors, Yesterday, Today, Tomorrow, LastSevenDays, LastWeek, ThisWeek, NextWeek, LastMonth, ThisMonth, NextMonth, AboveAverage, BelowAverage, EqualOrAboveAverage, EqualOrBelowAverage, OneStdDevAboveAverage, OneStdDevBelowAverage, TwoStdDevAboveAverage, TwoStdDevBelowAverage, ThreeStdDevAboveAverage, ThreeStdDevBelowAverage, UniqueValues y DuplicateValues.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propiedad_ > color|Código de color HTML que representa el color de la línea de borde con el formato #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propiedad_ > id|Representa el identificador de borde. Solo lectura. Los valores posibles son EdgeTop, EdgeBottom, EdgeLeft y EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propiedad_ > sideIndex|Valor constante que indica el lado específico del borde. Solo lectura. Los valores posibles son EdgeTop, EdgeBottom, EdgeLeft y EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propiedad_ > style|Una de las constantes de estilo de línea que especifica el estilo de línea del borde. Los valores posibles son: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Propiedad_ > recuento|Número de objetos de borde de la colección. Solo lectura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Propiedad_ > items|Colección de objetos conditionalRangeBorder. Solo lectura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relación_ > inferior|Obtiene el borde superior. Solo lectura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relación_ > izquierdo|Obtiene el borde superior. Solo lectura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relación_ > derecho|Obtiene el borde superior. Solo lectura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relación_ > parte superior|Obtiene el borde superior. Solo lectura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Método_ > getItem(index: string)|Obtiene un objeto de borde mediante su nombre.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Método_ > getItemAt(index: number)|Obtiene un objeto de borde mediante su índice.|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Propiedad_ > color|Código de color HTML que representa el color del relleno, con el formato #RRGGBB (p. ej. "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Método_ > clear()|Restablece el relleno.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propiedad_ > negrita|Representa el estado de negrita de la fuente.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propiedad_ > color|Representación del código de color HTML del color del texto. Por ejemplo, #FF0000 representa el rojo.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propiedad_ > cursiva|Representa el estado de cursiva de la fuente.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propiedad_ > tachado|Representa el estado de tachado de la fuente.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propiedad_ > subrayado|Tipo de subrayado aplicado a la fuente. Los valores posibles son None, Single y Double.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Método_ > clear()|Restablece los formatos de fuente.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Propiedad_ > numberFormat|Representa el código de formato numérico de Excel para el rango especificado. Si se pasa un valor nulo, aparece borrada.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relación_ > bordes|Colección de objetos border que se aplican al rango global de formato condicional. Solo lectura.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relación_ > relleno|Devuelve el objeto de relleno definido en el rango global de formato condicional. Solo lectura.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relación_ > font|Devuelve el objeto de fuente definido en el rango global de formato condicional. Solo lectura.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Propiedad_ > operator|Operador del formato condicional de texto. Los valores posibles son Invalid, Contains, NotContains, BeginsWith y EndsWith.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Propiedad_ > text|Valor de texto del formato condicional.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Propiedad_ > clasificación|El rango entre 1 y 1000 para los rangos numéricos o entre 1 y 100 para los rangos porcentuales.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Propiedad_ > tipo|Valores de formato en función del rango superior o inferior. Los valores posibles son Invalid, TopItems, TopPercent, BottomItems y BottomPercent.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Relación_ > formato|Devuelve un objeto de formato que encapsula la fuente de formatos condicionales, el relleno, los bordes y otras propiedades. Solo lectura.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Relación_ > regla|Representa el objeto Rule en este formato condicional. Solo lectura.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propiedad_ > axisColor|Código de color HTML que representa el color de la línea de eje con el formato #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propiedad_ > axisFormat|Representación de cómo se determina el eje para una barra de datos de Excel. Los valores posibles son Automatic, None y CellMidPoint.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propiedad_ > barDirection|Representa la dirección en la que se debe basar el gráfico de barras de datos. Los valores posibles son Context, LeftToRight y RightToLeft.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propiedad_ > showDataBarOnly|Si es true, oculta los valores de las celdas donde se aplica la barra de datos.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relación_ > lowerBoundRule|Regla de lo que constituye el límite inferior (y cómo calcularlo, si es aplicable) de una barra de datos.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relación_ > negativeFormat|Representación de todos los valores situados a la izquierda del eje en una barra de datos de Excel. Solo lectura.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relación_ > positiveFormat|Representación de todos los valores situados a la derecha del eje en una barra de datos de Excel. Solo lectura.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relación_ > upperBoundRule|Regla de lo que constituye el límite superior (y cómo calcularlo, si es aplicable) de una barra de datos.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Propiedad_ > reverseIconOrder|Si es true, invierte el orden de los iconos del IconSet. Tenga en cuenta que no se puede establecer si se usan iconos personalizados.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Propiedad_ > showIconOnly|Si es true, oculta los valores y solo muestra los iconos.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Propiedad_ > style|Si se establece, muestra la opción IconSet para el formato condicional. Los valores posibles son Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Relación_ > criteria|Matriz de Criteria y IconSets para las reglas y posibles iconos personalizados para los iconos condicionales. Tenga en cuenta que, para el primer criterio, solo se puede modificar el icono personalizado, mientras que el tipo, la fórmula y el operador se omitirán cuando se establezcan.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Relación_ > formato|Devuelve un objeto de formato que encapsula la fuente de formatos condicionales, el relleno, los bordes y otras propiedades. Solo lectura.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Relación_ > regla|Regla del formato condicional.|1.6|
|[range](/javascript/api/excel/excel.range)|_Relación_ > conditionalFormats|Colección de ConditionalFormats que forman una intersección en el intervalo. Solo lectura.|1.6|
|[range](/javascript/api/excel/excel.range)|_Método_ > calculate()|Calcula un rango de celdas en una hoja de cálculo.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Relación_ > formato|Devuelve un objeto de formato que encapsula la fuente de formatos condicionales, el relleno, los bordes y otras propiedades. Solo lectura.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Relación_ > regla|Regla del formato condicional.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Relación_ > formato|Devuelve un objeto de formato que encapsula la fuente de formatos condicionales, el relleno, los bordes y otras propiedades. Solo lectura.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Relación_ > regla|Los criterios del formato condicional TopBottom.|1.6|
|[workbook](/javascript/api/excel/excel.workbook)|_Relación_ > internalTest|Solo para uso interno. Solo lectura.|1.6|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > calculate(markAllDirty: bool)|Calcula todas las celdas de una hoja de cálculo.|1.6|

##  <a name="whats-new-in-excel-javascript-api-15"></a>¿Qué es nuevo en 1,5 de API de JavaScript de Excel

### <a name="custom-xml-part"></a>Elementos XML personalizados

* Adición de la colección de elementos XML personalizados al objeto de libro.
* Obtiene elementos XML personalizados con el identificador.
* Obtiene una nueva colección con ámbito de elementos XML personalizados cuyos espacios de nombres coinciden con el espacio de nombres determinado.
* Obtiene la cadena XML asociada con un elemento.
* Proporciona el id. y el espacio de nombres de un elemento.
* Se agrega un nuevo elemento XML personalizado al libro.
* Establece un elemento XML completo.
* Elimina un elemento XML personalizado.
* Elimina un atributo con el nombre determinado del elemento que XPath ha identificado.
* Consulta el contenido XML con XPath.
* Inserta, actualiza y elimina el atributo.

**Implementación de referencia:** Consulte [aquí](https://github.com/mandren/Excel-CustomXMLPart-Demo) para obtener una implementación de referencia que muestra cómo los elementos XML personalizados pueden usarse en un complemento.

### <a name="others"></a>Otros
* `range.getSurroundingRegion()` Devuelve un objeto Range que representa la región circundante de este intervalo. Una región circundante es un intervalo limitado por cualquier combinación de filas y columnas en blanco en relación a este intervalo.
* `getNextColumn()` y `getPreviousColumn()`, `getLast() en la columna de la tabla.
* `getActiveWorksheet()` en el libro.
* `getRange(address: string)` fuera del libro.
* `getBoundingRange(ranges: )` Obtiene el objeto de intervalo más pequeño que abarca los intervalos proporcionados. Por ejemplo, el intervalo delimitador entre "B2:C5" y "D10:E15" es "B2:E15".
* `getCount()` en varias colecciones como elemento con nombre, hoja de cálculo, tabla, etc. para obtener el número de elementos de una colección. `workbook.worksheets.getCount()`
* `getFirst()` y `getLast()` en varias colecciones como hoja de cálculo, columna de tabla, puntos de gráfico, colección de la vista de intervalo.
* `getNext()` y `getPrevious()` en la hoja de cálculo, colección de columna de tabla.
* `getRangeR1C1()` Obtiene el objeto de intervalo comenzando en un índice de columna y fila determinado, y que abarca un determinado número de filas y columnas.

|Objeto| Novedades| Descripción|Conjunto de requisitos|
|:----|:----|:----|:----|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Propiedad_ > id|Id. del elemento XML personalizado. Solo lectura.|1,5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Propiedad_ > namespaceUri|URI de espacio de nombres del elemento XML personalizado. Solo lectura.|1,5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Método_ > delete()|Elimina el elemento XML personalizado.|1,5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Método_ > getXml()|Obtiene el contenido XML completo del elemento XML personalizado.|1,5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Método_ > setXml(xml: string)|Establece el contenido XML completo del elemento XML personalizado.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Propiedad_ > items|Colección de objetos customXmlPart. Solo lectura.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > add(xml: string)|Se agrega un nuevo elemento XML personalizado al libro.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getByNamespace(namespaceUri: string)|Obtiene una nueva colección con ámbito de elementos XML personalizados cuyos espacios de nombres coinciden con el espacio de nombres determinado.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getCount()|Obtiene el número de elementos CustomXml de la colección.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getItem(id: string)|Obtiene un elemento XML personalizado a partir de su identificador.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getItemOrNullObject(id: string)|Obtiene un elemento XML personalizado a partir de su identificador.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Propiedad_ > items|Colección de objetos customXmlPartScoped. Solo lectura.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getCount()|Obtiene el número de elementos CustomXML de esta colección.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getItem(id: string)|Obtiene un elemento XML personalizado a partir de su identificador.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getItemOrNullObject(id: string)|Obtiene un elemento XML personalizado a partir de su identificador.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getOnlyItem()|Si la colección contiene exactamente un elemento, este método lo devuelve.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getOnlyItemOrNullObject()|Si la colección contiene exactamente un elemento, este método lo devuelve.|1,5|
|[workbook](/javascript/api/excel/excel.workbook)|_Relación_ > customXmlParts|Representa la colección de elementos XML personalizados incluidos en este libro. Solo lectura.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getNext(visibleOnly: bool)|Obtiene la siguiente hoja de cálculo. Si no hay ninguna hoja de cálculo que siga a esta, este método producirá un error.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getNextOrNullObject(visibleOnly: bool)|Obtiene la siguiente hoja de cálculo. Si no hay ninguna hoja de cálculo que siga a esta, este método devolverá un objeto null.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getPrevious(visibleOnly: bool)|Obtiene la hoja de cálculo anterior. Si no hay ninguna hoja de cálculo anterior, este método producirá un error.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getPreviousOrNullObject(visibleOnly: bool)|Obtiene la hoja de cálculo anterior. Si no hay ninguna hoja de cálculo anterior, este método devolverá un objeto null.|1,5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getFirst(visibleOnly: bool)|Obtiene la primera hoja de cálculo de la colección.|1,5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getLast(visibleOnly: bool)|Obtiene la última hoja de cálculo de la colección.|1,5|

## <a name="whats-new-in-excel-javascript-api-14"></a>Novedades de la API de JavaScript de Excel 1.4
Los siguientes son que las nuevas adiciones a las APIs de JavaScript de Excel en requisito establecer 1.4.

### <a name="named-item-add-and-new-properties"></a>Agregar elementos con nombre y nuevas propiedades

Nuevas propiedades:

* `comment`
* `scope`: elementos con ámbito de hoja de cálculo o libro.
* `worksheet`: devuelve la hoja de cálculo que tiene como ámbito el elemento con nombre.

Nuevos métodos:

* `add(name: string, reference: Range or string, comment: string)`: agrega un nuevo nombre a la colección del ámbito especificado.
* `addFormulaLocal(name: string, formula: string, comment: string)`: agrega un nuevo nombre a la colección del ámbito especificado con la configuración regional del usuario para la fórmula.

### <a name="settings-api-in-in-excel-namespace"></a>API de configuración en el espacio de nombres de Excel

El objeto [Setting](/javascript/api/excel/excel.setting) representa un par clave-valor de una configuración que se conserva en el documento. Ahora hemos agregado unas API relacionadas con la configuración en el espacio de nombres de Excel. De esta forma, aunque la funcionalidad no es estrictamente nueva, resulta más fácil permanecer en la sintaxis de API por lotes basada en compromisos y se reduce la dependencia de la API común para tareas relacionadas con Excel.

Las API incluyen `getItem()` para obtener acceso a la configuración mediante la clave `add()` para agregar al libro el par clave-valor de configuración especificado.

### <a name="others"></a>Otros

* Establecer el nombre de la columna de tabla (la versión anterior solo permite la lectura).
* Agregar una columna al final de la tabla (la versión anterior permite cualquier lugar excepto el último).
* Agregar varias filas a una tabla de una sola vez (la versión anterior solo permite agregarlas de una en una).
* `range.getColumnsAfter(count: number)` y `range.getColumnsBefore(count: number)` para obtener un número determinado de columnas a la derecha o izquierda del objeto Range actual.
* Función para obtener un elemento o un objeto NULL: esta funcionalidad permite obtener el objeto mediante una clave. Si el objeto no existe, la propiedad isNullObject del objeto devuelto será "true". Esto permite a los desarrolladores comprobar si existe un objeto sin tener que utilizar el control de excepciones para controlarlo. Disponible para hojas de cálculo, elementos con nombre, enlaces, series de gráfico, etc.

    ```javascript
    worksheet.GetItemOrNullObject()
    ```

|Objeto| Novedades| Descripción|Conjunto de requisitos|
|:----|:----|:----|:----|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > getCount()|Obtiene el número de enlaces de la colección.|1.4|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > getItemOrNullObject(id: string)|Obtiene un objeto de enlace por identificador. Si no existe el objeto de enlace, devolverá un objeto nulo.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Método_ > getCount()|Devuelve el número de gráficos de la hoja de cálculo.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Método_ > getItemOrNullObject(name: string)|Obtiene un gráfico mediante su nombre. Si hay varias tablas con el mismo nombre, se devolverá la primera.|1.4|
|[chartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|_Método_ > getCount()|Devuelve el número de puntos del gráfico de la serie.|1.4|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Método_ > getCount()|Devuelve el número de series incluidas en la colección.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Propiedad_ > comment|Representa el comentario asociado a este nombre.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Propiedad_ > scope|Indica si el nombre está en el ámbito del libro o de una hoja de cálculo específica. Solo lectura. Los valores posibles son: Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relación_ > worksheet|Devuelve la hoja de cálculo que tiene como ámbito el elemento con nombre. Se produce un error si el ámbito del elemento es el libro. Solo lectura.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relación_ > worksheetOrNullObject|Devuelve la hoja de cálculo que tiene como ámbito el elemento con nombre. Devuelve un objeto NULL si el ámbito del elemento es el libro. Solo lectura.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Método_ > delete()|Elimina el nombre especificado.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Método_ > getRangeOrNullObject()|Devuelve el objeto de rango asociado al nombre. Devuelve un objeto NULL si el tipo de elemento con nombre no es un rango.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > Agregar (nombre: cadena, referencia: rango o cadena, comentario: cadena)|Agrega un nuevo nombre a la colección del ámbito especificado.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > addFormulaLocal (nombre: cadena, fórmula: cadena, comentario: cadena)|Agrega un nuevo nombre a la colección del ámbito especificado, empleando la configuración regional del usuario para la fórmula.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > getCount()|Obtiene el número de elementos con nombre de la colección.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > getItemOrNullObject(name: string)|Obtiene un objeto NamedItem mediante su nombre. Si no existe el objeto NamedItem, devolverá un objeto NULL.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getCount()|Obtiene el número de tablas dinámicas de una colección.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getItemOrNullObject(name: string)|Obtiene una tabla dinámica por nombre. Si no existe la tabla dinámica, devolverá un objeto NULL.|1.4|
|[range](/javascript/api/excel/excel.range)|_Método_ > getIntersectionOrNullObject (anotherRange: rango o cadena)|Obtiene el objeto de intervalo que representa la intersección rectangular de los intervalos especificados. Si no se encuentra ninguna intersección, se devolverá un objeto NULL.|1.4|
|[range](/javascript/api/excel/excel.range)|_Método_ > getUsedRangeOrNullObject(valuesOnly: bool)|Devuelve el rango usado del objeto de rango especificado. Si no hay ninguna celda usada dentro del rango, esta función devolverá un objeto NULL.|1.4|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Método_ > getCount()|Obtiene el número de objetos RangeView de la colección.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Propiedad_ > key|Devuelve una clave que representa el identificador de la configuración. Solo lectura.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Propiedad_ > value|Representa el valor almacenado para esta configuración.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Método_ > delete()|Elimina la configuración.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Propiedad_ > items|Una colección de objetos de configuración. Solo lectura.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > Agregar (clave: valor de la cadena,: (cualquier))|Establece o agrega la configuración especificada en el libro.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getCount()|Obtiene el número de configuraciones de una colección.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItem(key: string)|Obtiene una entrada de configuración mediante la clave.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItemOrNullObject(key: string)|Obtiene una entrada de configuración mediante la clave. Si el valor no existe, devolverá un objeto NULL.|1.4|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Relación_ > settings|Obtiene el objeto Setting que representa el enlace que ha generado el evento SettingsChanged.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Método_ > getCount()]|Obtiene el número de tablas de la colección.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Método_ > getItemOrNullObject (clave: número o cadena)|Obtiene una tabla por nombre o identificador. Si la tabla no existe, devolverá un objeto NULL.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Método_ > getCount()|Obtiene el número de columnas de la tabla.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Método_ > getItemOrNullObject (clave: número o cadena)|Obtiene un objeto de columna por nombre o identificador. Si la columna no existe, devolverá un objeto NULL.|1.4|
|[tableRowCollection](/javascript/api/excel/excel.tablerowcollection)|_Método_ > getCount()|Obtiene el número de filas de la tabla.|1.4|
|[workbook](/javascript/api/excel/excel.workbook)|_Relación_ > settings|Representa una colección de configuraciones asociadas con el libro. Solo lectura.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relación_ > names|Colección de nombres en el ámbito de la hoja de cálculo actual. Solo lectura.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getUsedRangeOrNullObject(valuesOnly: bool)|El rango usado es el rango más pequeño que abarque todas las celdas que tengan asignado un valor o un formato. Si toda la hoja está en blanco, esta función devolverá un objeto NULL.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getCount(visibleOnly: bool)|Obtiene el número de hojas de cálculo de la colección.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getItemOrNullObject(key: string)|Obtiene un objeto de hoja de cálculo mediante su nombre o identificador. Si la hoja de cálculo no existe, devolverá un objeto NULL.|1.4|

## <a name="whats-new-in-excel-javascript-api-13"></a>Novedades de la API de JavaScript de Excel 1.3

Las siguientes son las nuevas incorporaciones a las API de JavaScript de Excel en el conjunto de requisitos 1.3.

|Objeto| Novedades| Descripción|Conjunto de requisitos|
|:----|:----|:----|:----|
|[binding](/javascript/api/excel/excel.binding)|_Método_ > delete()|Elimina el enlace.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > Agregar (rango: rango o cadena, bindingType: string, identificador de: cadena)|Agregar un enlace nuevo a un intervalo determinado.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > addFromNamedItem (nombre: string, bindingType: string, identificador: cadena)|Agregar un enlace nuevo basándose en un elemento con nombre del libro.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > addFromSelection (bindingType: string, identificador: cadena)|Agregar un enlace nuevo basándose en la selección actual.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > getItemOrNull(id: string)|Obtiene un objeto de enlace por identificador. Si el objeto de enlace no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Método_ > getItemOrNull(name: string)|Obtiene un gráfico mediante su nombre. Si hay varias tablas con el mismo nombre, se devolverá la primera.|1.3|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > getItemOrNull(name: string)|Obtiene un objeto NamedItem mediante su nombre. Si el objeto NamedItem no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Propiedad_ > name|Nombre la tabla dinámica.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relación_ > worksheet|La hoja de cálculo que contiene la tabla dinámica actual. Solo lectura.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Método_ > Actualizar()|Actualiza la tabla dinámica.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Propiedad_ > items|Una colección de objetos de tabla dinámica. Solo lectura.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getItem(name: string)|Obtiene una tabla dinámica por nombre.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getItemOrNull(name: string)|Obtiene una tabla dinámica por nombre. Si la tabla dinámica no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[range](/javascript/api/excel/excel.range)|_Método_ > getIntersectionOrNull (anotherRange: rango o cadena)|Obtiene el objeto de intervalo que representa la intersección rectangular de los intervalos especificados. Si no se encuentra ninguna intersección, se devolverá un objeto NULL.|1.3|
|[range](/javascript/api/excel/excel.range)|_Método_ > getVisibleView()|Representa las filas visibles del intervalo actual.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propiedad_ > cellAddresses|Representa las direcciones de celda de RangeView. Solo lectura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propiedad_ > columnCount|Devuelve el número de columnas visibles. Solo lectura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propiedad_ > formulas|Representa la fórmula en notación de estilo A1.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propiedad_ > formulasLocal|Representa la fórmula en notación de estilo A1, en el idioma del usuario y en la configuración regional del formato numérico.  Por ejemplo, la fórmula "=SUM(A1, presentada en 1.5)" en inglés se convertiría en "=SUMME(A1; 1,5)" en alemán.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propiedad_ > formulasR1C1|Representa la fórmula en notación de estilo R1C1.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propiedad_ > index|Devuelve un valor que representa el índice de RangeView. Solo lectura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propiedad_ > numberFormat|Representa el código de formato numérico de Excel para la celda especificada.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propiedad_ > rowCount|Devuelve el número de filas visibles. Solo lectura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propiedad_ > text|Valores de texto del rango especificado. El valor Text no dependerá del ancho de la celda. La sustitución del signo # que tiene lugar en la interfaz de usuario de Excel no afectará al valor de texto devuelto por la API. Solo lectura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propiedad_ > valueTypes|Representa el tipo de datos de cada celda. Solo lectura. Los valores posibles son: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propiedad_ > values|Representa los valores sin formato de la vista del intervalo especificado. Los datos devueltos pueden ser de tipo cadena, número o booleano. La celda que contenga un error devolverá la cadena de error.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Relación_ > rows|Representa una colección de vistas de intervalo asociadas a este. Solo lectura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Método_ > getRange()|Obtiene el intervalo primario asociado al RangeView actual.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Propiedad_ > items|Una colección de objetos RangeView. Solo lectura.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Método_ > getItemAt(index: number)|Obtiene una fila RangeView mediante su índice. Indexado con cero.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Propiedad_ > key|Devuelve una clave que representa el identificador de la configuración. Solo lectura.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Método_ > delete()|Elimina la configuración.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Propiedad_ > items|Una colección de objetos de configuración. Solo lectura.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItem(key: string)|Obtiene una entrada de configuración mediante la clave.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItemOrNull(key: string)|Obtiene una entrada de configuración mediante la clave. Si la configuración no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > establecer (clave: valor de la cadena,: cadena)|Establece o agrega la configuración especificada en el libro.|1.3|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Relación_ > settingCollection|Obtiene el objeto Setting que representa el enlace que ha generado el evento SettingsChanged.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propiedad_ > highlightFirstColumn|Indica si la primera columna contiene un formato especial.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propiedad_ > highlightLastColumn|Indica si la última columna contiene un formato especial.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propiedad_ > showBandedColumns|Indica si las columnas muestran un formato con bandas en el que las columnas impares están resaltadas de manera diferente que las pares para facilitar la lectura de la tabla.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propiedad_ > showBandedRows|Indica si las filas muestran un formato con bandas en el que las filas impares están resaltadas de manera diferente que las pares para facilitar la lectura de la tabla.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propiedad_ > showFilterButton|Indica si los botones de filtro son visibles en la parte superior de cada encabezado de columna. Esta configuración solo se permite si la tabla contiene una fila de encabezado.|1.3|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Método_ > getItemOrNull (clave: número o cadena)|Obtiene una tabla por nombre o identificador. Si la tabla no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Método_ > getItemOrNull (clave: número o cadena)|Obtiene un objeto de columna por nombre o identificador. Si la columna no existe, la propiedad isNull del objeto devuelto será True.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Relación_ > pivotTables|Representa una colección de tablas dinámicas asociadas con el libro. Solo lectura.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Relación_ > settings|Representa una colección de configuraciones asociadas con el libro. Solo lectura.|1.3|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relación_ > pivotTables|Colección de tablas dinámicas que forman parte de la hoja de cálculo. Solo lectura.|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Novedades de la API de JavaScript de Excel 1.2

Las siguientes son las nuevas incorporaciones a las API de JavaScript de Excel en el conjunto de requisitos 1.2.

|Objeto| Novedades| Descripción|Conjunto de requisitos|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Propiedad_ > id|Obtiene un gráfico en función de su posición en la colección. Solo lectura.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Relación_ > worksheet|La hoja de cálculo que contiene el gráfico actual. Solo lectura.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Método_ > getImage (alto: número, ancho: número, fittingMode: cadena)|Representa el gráfico como una imagen con codificación Base64 al escalar el gráfico a las dimensiones especificadas.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Relación_ > criteria|Filtro aplicado actualmente en la columna especificada. Solo lectura.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > apply(criteria: FilterCriteria)|Aplicar los criterios de filtro especificados en la columna especificada.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyBottomItemsFilter(count: number)|Aplicar un filtro de "Elemento inferior" a la columna para el número de elementos especificado.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyBottomPercentFilter(percent: number)]|Aplicar un filtro de "Porcentaje inferior" a la columna para el porcentaje de elementos especificado.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyCellColorFilter(color: string)|Aplicar un filtro de "Color de celda" a la columna para el color especificado.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyCustomFilter (criteria1: string, criterio 2: string, oper: cadena)|Aplicar un filtro de "Icono" a la columna para las cadenas de criterios especificadas.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyDynamicFilter(criteria: string)|Aplicar un filtro "Dinámico" a la columna.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyFontColorFilter(color: string)|Aplicar un filtro de "Color de fuente" a la columna para el color especificado.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyIconFilter(icon: Icon)|Aplicar un filtro de "Icono" a la columna para el icono especificado.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyTopItemsFilter(count: number)|Aplicar un filtro de "Elemento superior" a la columna para el número de elementos especificado.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyTopPercentFilter(percent: number)|Aplicar un filtro de "Porcentaje superior" a la columna para el porcentaje de elementos especificado.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyValuesFilter (valores: ())|Aplicar un filtro de "Valores" a la columna para los valores especificados.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > clear()|Borrar el filtro de la columna especificada.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propiedad_ > color|Cadena de color HTML que se usa para filtrar las celdas. Se usa con el filtrado de "cellColor" y "fontColor".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propiedad_ > criterion1|Primer criterio usado para filtrar los datos. Se usa como un operador en el caso del filtrado "personalizado".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propiedad_ > criterion2|Segundo criterio usado para filtrar los datos. Solo se usa como un operador en el caso del filtrado "personalizado".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propiedad_ > dynamicCriteria|Criterios dinámicos del conjunto Excel.DynamicFilterCriteria que se van a aplicar a esta columna. Se usa con el filtrado "dinámico". Los valores posibles son: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propiedad_ > filterOn|Propiedad usada por el filtro para determinar si los valores deben permanecer visibles. Los valores posibles son: BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propiedad_ > operator|Operador usado para combinar el criterio 1 y 2 cuando se usa el filtrado "personalizado". Los valores posibles son: And, Or.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propiedad_ > values|Conjunto de valores que se van a usar como parte del filtrado de "valores".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Relación_ > icon|Icono usado para filtrar las celdas. Se usa con el filtrado de "icono".|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Propiedad_ > date|La fecha en formato ISO8601 usada para filtrar los datos.|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Propiedad_ > specificity|El grado de especificidad de la fecha que se usará para mantener datos. Por ejemplo, si la fecha es 02-04-2005 y la especificidad se establece en "mes", la operación de filtrado conservará todas las filas con fecha de abril de 2005. Los valores posibles son: Year, Monday, Day, Hour, Minute, Second.|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Propiedad_ > formulaHidden|Indica si Excel oculta la fórmula de las celdas del rango. Un valor null indica que el rango no tiene una configuración de fórmula oculta uniforme.|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Propiedad_ > locked|Indica si Excel bloquea las celdas del objeto. Un valor nulo indica que todo el rango no tiene una configuración de bloqueo uniforme.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Propiedad_ > index|Representa el índice del icono en el conjunto concreto.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Propiedad_ > set|Representa el conjunto al que pertenece el icono. Los valores posibles son: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propiedad_ > columnHidden|Representa si todas las columnas del intervalo actual están ocultas.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propiedad_ > formulasR1C1|Representa la fórmula en notación de estilo R1C1.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propiedad_ > hidden|Representa si todas las celdas del rango actual están ocultas. Solo lectura.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propiedad_ > rowHidden|Representa si todas las filas del intervalo actual están ocultas.|1.2|
|[range](/javascript/api/excel/excel.range)|_Relación_ > sort|Representa la ordenación del intervalo del intervalo actual. Solo lectura.|1.2|
|[range](/javascript/api/excel/excel.range)|_Método_ > merge(across: bool)|Combina las celdas del intervalo en una región de la hoja de cálculo.|1.2|
|[range](/javascript/api/excel/excel.range)|_Método_ > unmerge()|Separa las celdas del intervalo en celdas independientes.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propiedad_ > columnWidth|Obtiene o establece el ancho de todas las columnas del rango. Si los anchos de columna no son uniformes, se devolverá null.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propiedad_ > rowHeight|Obtiene o establece el alto de todas las filas del rango. Si los altos de fila no son uniformes, se devolverá null.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Relación_ > protection|Devuelve el objeto de protección de formato de un rango. Solo lectura.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Método_ > autofitColumns()|Cambia el ancho de las columnas del intervalo actual para obtener el ajuste perfecto (según los datos actuales de las columnas).|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Método_ > autofitRows()|Cambia el alto de las filas del intervalo actual para obtener el ajuste perfecto (según los datos actuales de las columnas).|1.2|
|[rangeReference](/javascript/api/excel/excel.rangereference)|_Propiedad_ > address|Representa las filas visibles del intervalo actual.|1.2|
|[rangeSort](/javascript/api/excel/excel.rangesort)|_Método_ > Aplicar (campos: SortField, matchCase: bool, hasHeaders: bool, orientación: string, método: cadena)|Realiza una operación de ordenación.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propiedad_ > ascending|Representa si la ordenación se realiza en orden ascendente.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propiedad_ > color|Representa el color que es el destino de la condición si la ordenación se realiza según la fuente o el color de celda.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propiedad_ > dataOption|Representa opciones de ordenación adicionales para este campo. Los valores posibles son: Normal, TextAsNumber.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propiedad_ > key|Representa la columna (o fila, según la orientación de ordenación) en que se encuentra la condición. Se representa como un desplazamiento de la primera columna (o fila).|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propiedad_ > sortOn|Representa el tipo de ordenación de esta condición. Los valores posibles son: Value, CellColor, FontColor, Icon.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Relación_ > icon|Representa el icono que es el destino de la condición si la ordenación se realiza según el icono de la celda.|1.2|
|[table](/javascript/api/excel/excel.table)|_Relación_ > sort|Representa la ordenación de la tabla. Solo lectura.|1.2|
|[table](/javascript/api/excel/excel.table)|_Relación_ > worksheet|Hoja de cálculo que contiene la tabla actual. Solo lectura.|1.2|
|[table](/javascript/api/excel/excel.table)|_Método_ > clearFilters()|Borra todos los filtros aplicados actualmente en la tabla.|1.2|
|[table](/javascript/api/excel/excel.table)|_Método_ > convertToRange()|Convierte la tabla en un rango de celdas normal. Se conservan todos los datos.|1.2|
|[table](/javascript/api/excel/excel.table)|_Método_ > reapplyFilters()|Vuelve a aplicar todos los filtros aplicados actualmente en la tabla.|1.2|
|[tableColumn](/javascript/api/excel/excel.tablecolumn)|_Relación_ > filter|Recupera el filtro aplicado a la columna. Solo lectura.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Propiedad_ > matchCase|Indica si última ordenación de la tabla distinguía mayúsculas de minúsculas. Solo lectura.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Propiedad_ > method|Representa el método de ordenación de caracteres chinos usado por última vez para ordenar la tabla. Solo lectura. Los valores posibles son: PinYin, StrokeCount.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Relación_ > fields|Representa las condiciones actuales que se usaron por última vez para ordenar la tabla. Solo lectura.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Método_ > Aplicar (campos: SortField, matchCase: bool, método: cadena)|Realiza una operación de ordenación.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Método_ > clear()|Borra la ordenación que se aplica actualmente en la tabla. Aunque esto no modifica la ordenación de la tabla, borra el estado de los botones de encabezado.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Método_ > reapply()|Vuelve a aplicar los parámetros de ordenación actuales a la tabla.|1.2|
|[workbook](/javascript/api/excel/excel.workbook)|_Relación_ > functions|Representa una instancia de aplicación de Excel que contiene este libro. Solo lectura.|1.2|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relación_ > protection|Devuelve el objeto de protección de hoja de una hoja de cálculo. Solo lectura.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Propiedad_ > protected|Indica si la hoja de cálculo está protegida. Solo lectura. Solo lectura.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Relación_ > options|Opciones de protección de la hoja. Solo lectura.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Método_ > protect(options: WorksheetProtectionOptions)|Protege una hoja de cálculo. Produce un error si se ha protegido la hoja de cálculo.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Método_ > unprotect()|Desprotege una hoja de cálculo.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propiedad_ > allowAutoFilter|Representa la opción de protección de la hoja de cálculo que permite usar la característica de filtro automático.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propiedad_ > allowDeleteColumns|Representa la opción de protección de la hoja de cálculo que permite eliminar columnas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propiedad_ > allowDeleteRows|Representa la opción de protección de la hoja de cálculo que permite eliminar filas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propiedad_ > allowFormatCells|Representa la opción de protección de la hoja de cálculo que permite aplicar formato a celdas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propiedad_ > allowFormatColumns|Representa la opción de protección de la hoja de cálculo que permite aplicar formato a columnas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propiedad_ > allowFormatRows|Representa la opción de protección de la hoja de cálculo que permite aplicar formato a filas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propiedad_ > allowInsertColumns|Representa la opción de protección de la hoja de cálculo que permite insertar columnas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propiedad_ > allowInsertHyperlinks|Representa la opción de protección de la hoja de cálculo que permite insertar hipervínculos.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propiedad_ > allowInsertRows|Representa la opción de protección de la hoja de cálculo que permite insertar filas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propiedad_ > allowPivotTables|Representa la opción de protección de la hoja de cálculo que permite usar la característica de tabla dinámica.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propiedad_ > allowSort|Representa la opción de protección de la hoja de cálculo que permite usar la característica de ordenación.|1.2|

## <a name="excel-javascript-api-11"></a>API de JavaScript de Excel 1.1

Excel JavaScript API 1.1 es la primera versión de la API. Para obtener información detallada acerca de la API, vea los temas de referencia de la [API de JavaScript de Excel](/javascript/api/excel) .

## <a name="see-also"></a>Vea también

- [Versiones de Office y conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar los hosts de Office y los requisitos de la API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifiesto XML de complementos para Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
