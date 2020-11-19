| Class | Campos | Descripción |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Suspende el cálculo hasta que se llama al siguiente "context.sync()".|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Devuelve un objeto Format, encapsulando la fuente de los formatos condicionales, el relleno, los bordes y otras propiedades.|
||[Rule](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Especifica el objeto Rule en este formato condicional.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|Criterios de la escala de colores.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|Si es true, la escala de colores tendrá tres puntos (mínimo, punto medio, máximo); de lo contrario, tendrá dos (mínimo, máximo).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[Formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|La fórmula, si es necesaria, en la que se debe evaluar la regla de formato condicional.|
||[Formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|La fórmula, si es necesaria, en la que se debe evaluar la regla de formato condicional.|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|Operador del formato condicional del valor de celda.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|Criterio de escala de colores de punto máximo.|
||[punto medio](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|Criterio de escala de colores de punto medio si la escala de colores es una escala de 3 colores.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|Criterio de escala de colores de punto mínimo.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|Representación del código de color HTML del color de la escala de colores (por ejemplo, #FF0000 representa el rojo).|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|Número, fórmula o valor null (si el tipo es LowestValue).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|En qué se debe basar la fórmula condicional criterio.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|Código de color HTML que representa el color de la línea de borde con el formato #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|Código de color HTML que representa el color de relleno, de la forma #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|Especifica si la barra de la barra de la barra de la barra de la barra de la barra de la barra positiva es igual.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|Especifica si la barra de la barra de números negativo tiene el mismo color de relleno que la barra de la barra de números positivo.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|Código de color HTML que representa el color de la línea de borde con el formato #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|Código de color HTML que representa el color de relleno, de la forma #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Especifica si DataBar tiene un degradado.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|La fórmula, si es necesaria, en la que se debe evaluar la regla databar.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|Tipo de regla de la barra de la barra de texto.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|Elimina este formato condicional.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|Devuelve el intervalo en el que se aplica el formato condicional.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|Devuelve el intervalo al que se aplica el formato condicional o un objeto NULL si el formato condicional se aplica a varios rangos.|
||[asigna](/javascript/api/excel/excel.conditionalformat#priority)|Prioridad (o índice) dentro de la colección de formato condicional en la que este formato condicional existe actualmente.|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|Devuelve las propiedades de formato condicional del valor de celda si el formato condicional actual es un tipo CellValue.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|Devuelve las propiedades de formato condicional del valor de celda si el formato condicional actual es un tipo CellValue.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorscale)|Devuelve las propiedades del formato condicional de ColorScale si el formato condicional actual es un tipo ColorScale.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|Devuelve las propiedades del formato condicional de ColorScale si el formato condicional actual es un tipo ColorScale.|
||[customiza](/javascript/api/excel/excel.conditionalformat#custom)|Devuelve las propiedades de formato condicional personalizado si el formato condicional actual es un tipo personalizado.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customornullobject)|Devuelve las propiedades de formato condicional personalizado si el formato condicional actual es un tipo personalizado.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#databar)|Devuelve las propiedades de la barra de datos si el formato condicional actual es una barra de datos.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#databarornullobject)|Devuelve las propiedades de la barra de datos si el formato condicional actual es una barra de datos.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|Devuelve las propiedades del formato condicional IconSet si el formato condicional actual es un tipo IconSet.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|Devuelve las propiedades del formato condicional IconSet si el formato condicional actual es un tipo IconSet.|
||[id](/javascript/api/excel/excel.conditionalformat#id)|La prioridad del formato condicional dentro de la ConditionalFormatCollection actual.|
||[PRESET](/javascript/api/excel/excel.conditionalformat#preset)|Devuelve el formato condicional de criterios preestablecidos.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetornullobject)|Devuelve el formato condicional de criterios preestablecidos.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textcomparison)|Devuelve las propiedades de formato condicional de texto específico si el formato condicional actual es un tipo de texto.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|Devuelve las propiedades de formato condicional de texto específico si el formato condicional actual es un tipo de texto.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|Devuelve las propiedades de formato condicional superior o inferior si el formato condicional actual es un tipo de parte inferior.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|Devuelve las propiedades de formato condicional superior o inferior si el formato condicional actual es un tipo de parte inferior.|
||[type](/javascript/api/excel/excel.conditionalformat#type)|Un tipo de formato condicional.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|Si se cumplen las condiciones de este formato condicional, los formatos de menor prioridad no surtirán efecto en esa celda.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[Add (tipo: Excel. ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Agrega un nuevo formato condicional a la colección en la primera o máxima prioridad.|
||[clearAll ()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Borra todos los formatos condicionales activos en el intervalo actual especificado.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Devuelve el número de formatos condicionales del libro.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Devuelve un formato condicional para el identificador especificado.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Devuelve un formato condicional en el índice especificado.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|La fórmula, si es necesaria, en la que se debe evaluar la regla de formato condicional.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|La fórmula, si es necesaria, en la que se debe evaluar la regla de formato condicional en el idioma del usuario.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|La fórmula, si es necesaria, en la que se debe evaluar la regla de formato condicional en la notación de estilo R1C1.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|Icono personalizado para el criterio actual si es diferente del IconSet predeterminado; si no, se devolverá null.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|Número o fórmula en función del tipo.|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#operator)|GreaterThan o GreaterThanOrEqual para cada tipo de regla para el formato condicional de icono.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|Elemento en el que se debe basar la fórmula condicional de icono.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[criterio](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|El criterio del formato condicional.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|Código de color HTML que representa el color de la línea de borde con el formato #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Valor constante que indica el lado específico del borde.|
||[estilo](/javascript/api/excel/excel.conditionalrangeborder#style)|Una de las constantes de estilo de línea que especifica el estilo de línea del borde.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem (índice: Excel. ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Obtiene un objeto de borde mediante su nombre.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Obtiene un objeto de borde mediante su índice.|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Obtiene el borde inferior.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Número de objetos de borde de la colección.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Obtiene el borde izquierdo.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Obtiene el borde derecho.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Obtiene el borde superior.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Restablece el relleno.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|Código de color HTML que representa el color del relleno, con el formato #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Especifica si la fuente está en negrita.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Restablece los formatos de fuente.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|Representación del código de color HTML del color del texto (por ejemplo, #FF0000 representa el rojo).|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Especifica si la fuente está en cursiva.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Especifica el estado de tachado de la fuente.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Tipo de subrayado aplicado a la fuente.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Representa el código de formato numérico de Excel para el intervalo especificado.|
||[borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Colección de objetos Border que se aplican al intervalo de formato condicional general.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Devuelve el objeto de relleno definido en el intervalo de formato condicional general.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|Devuelve el objeto Font definido en el intervalo de formatos condicionales general.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|Operador del formato condicional de texto.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|Valor de texto del formato condicional.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|El rango entre 1 y 1000 para los rangos numéricos o entre 1 y 100 para los rangos porcentuales.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Da formato a valores basados en el rango superior o inferior.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Devuelve un objeto Format, encapsulando la fuente de los formatos condicionales, el relleno, los bordes y otras propiedades.|
||[Rule](/javascript/api/excel/excel.customconditionalformat#rule)|Especifica el objeto Rule en este formato condicional.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|Código de color HTML que representa el color de la línea del eje, del formulario #RRGGBB (por ejemplo, "FFA500") o como un color HTML con nombre (por ejemplo, "naranja").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Representación de cómo se determina el eje para una barra de datos de Excel.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Especifica la dirección en la que debe basarse el gráfico de la barra de datos.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|Regla de lo que constituye el límite inferior (y cómo calcularlo, si es aplicable) de una barra de datos.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Representación de todos los valores a la izquierda del eje en una barra de datos de Excel.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Representación de todos los valores a la derecha del eje en una barra de datos de Excel.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|Si es true, oculta los valores de las celdas donde se aplica la barra de datos.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|Regla de lo que constituye el límite superior (y cómo calcularlo, si es aplicable) de una barra de datos.|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|Una matriz de criterios y IconSets para las reglas y los posibles iconos personalizados para los iconos condicionales.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|Si es true, invierte el orden de los iconos del IconSet.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|Si es true, oculta los valores y solo muestra los iconos.|
||[estilo](/javascript/api/excel/excel.iconsetconditionalformat#style)|Si se establece, muestra la opción IconSet del formato condicional.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Devuelve un objeto Format, encapsulando la fuente de los formatos condicionales, el relleno, los bordes y otras propiedades.|
||[Rule](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|Regla del formato condicional.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Calcula un rango de celdas en una hoja de cálculo.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|La colección de ConditionalFormats que intersecta el rango.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Devuelve un objeto Format que encapsula la fuente, el relleno, los bordes y otras propiedades del formato condicional.|
||[Rule](/javascript/api/excel/excel.textconditionalformat#rule)|Regla del formato condicional.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Devuelve un objeto Format que encapsula la fuente, el relleno, los bordes y otras propiedades del formato condicional.|
||[Rule](/javascript/api/excel/excel.topbottomconditionalformat#rule)|Criterios del formato condicional superior o inferior.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Calculate (markAllDirty: Boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Calcula todas las celdas de una hoja de cálculo.|
