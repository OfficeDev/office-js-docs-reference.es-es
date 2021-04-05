| Clase | Campos | Descripción |
|:---|:---|:---|
|[Enlace](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|Elimina el enlace.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add(range: Range \| string, bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Agregar un enlace nuevo a un intervalo determinado.|
||[addFromNamedItem(name: string, bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Agregar un enlace nuevo basándose en un elemento con nombre del libro.|
||[addFromSelection(bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Agregar un enlace nuevo basándose en la selección actual.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Nombre la tabla dinámica.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|La hoja de cálculo que contiene la tabla dinámica actual.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|Actualiza la tabla dinámica.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|Obtiene una tabla dinámica por nombre.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[refreshAll()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|Actualiza todas las tablas dinámicas de la colección.|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView()](/javascript/api/excel/excel.range#getvisibleview--)|Representa las filas visibles del intervalo actual.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|Representa la fórmula en notación de estilo A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulaslocal)|Representa la fórmula en notación de estilo A1, en el idioma del usuario y en la configuración regional del formato numérico.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasr1c1)|Representa la fórmula en notación de estilo R1C1.|
||[getRange()](/javascript/api/excel/excel.rangeview#getrange--)|Obtiene el intervalo primario asociado con el `RangeView` actual .|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberformat)|Representa el código de formato numérico de Excel para la celda especificada.|
||[cellAddresses](/javascript/api/excel/excel.rangeview#celladdresses)|Representa las direcciones de celda de `RangeView` .|
||[columnCount](/javascript/api/excel/excel.rangeview#columncount)|Número de columnas visibles.|
||[index](/javascript/api/excel/excel.rangeview#index)|Devuelve un valor que representa el índice de `RangeView` .|
||[rowCount](/javascript/api/excel/excel.rangeview#rowcount)|Número de filas visibles.|
||[rows](/javascript/api/excel/excel.rangeview#rows)|Representa una colección de vistas de intervalo asociadas a este.|
||[text](/javascript/api/excel/excel.rangeview#text)|Valores de texto del rango especificado.|
||[valueTypes](/javascript/api/excel/excel.rangeview#valuetypes)|Representa el tipo de datos de cada celda.|
||[values](/javascript/api/excel/excel.rangeview#values)|Representa los valores sin formato de la vista del intervalo especificado.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|Obtiene una `RangeView` fila a través de su índice.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|Especifica si la primera columna contiene formato especial.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|Especifica si la última columna contiene formato especial.|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|Especifica si las columnas muestran un formato con bandas en el que las columnas impares se resaltan de forma diferente a incluso las columnas, para facilitar la lectura de la tabla.|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|Especifica si las filas muestran un formato con bandas en el que las filas impares se resaltan de forma diferente a las pares, para facilitar la lectura de la tabla.|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|Especifica si los botones de filtro están visibles en la parte superior de cada encabezado de columna.|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#pivottables)|Representa una colección de tablas dinámicas asociadas con el libro.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#pivottables)|Colección de tablas dinámicas que forman parte de la hoja de cálculo.|
