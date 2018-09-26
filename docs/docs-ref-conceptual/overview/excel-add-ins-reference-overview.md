# <a name="excel-javascript-api-overview"></a>Introducción a la API de JavaScript de Excel

Puede usar la API de JavaScript de Excel para crear complementos para Excel 2016 o posterior. En la lista siguiente se muestran los objetos de Excel de alto nivel que están disponibles en la API. Cada vínculo de objeto de página contiene una descripción de las propiedades, eventos y métodos que están disponibles en el objeto. Explore los vínculos del menú para más información.

Para su comodidad, a continuación se indican algunos de los objetos principales de Excel: 

- [Workbook](/javascript/api/excel/excel.workbook): objeto de nivel superior que contiene los objetos de libro relacionados, como hojas de cálculo, tablas, intervalos, etc. También puede usarse para enumerar las referencias relacionadas.

- [Worksheet](/javascript/api/excel/excel.worksheet) representa una hoja de cálculo en un libro. 
    - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection): una colección de objetos **Worksheet** de un libro.

- [Range](/javascript/api/excel/excel.range): representa una celda, una fila, una columna o una selección de celdas con uno o más bloques contiguos de celdas.

- [Table](/javascript/api/excel/excel.table): representa una colección de celdas organizadas diseñada para facilitar la administración de los datos.
    - [TableCollection](/javascript/api/excel/excel.tablecollection): colección de tablas de un libro o una hoja de cálculo.
    - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection): colección de todas las columnas de una tabla.
    - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection): colección de todas las filas de una tabla.

- [Chart](/javascript/api/excel/excel.chart): representa un objeto Chart de una hoja de cálculo, que es una representación visual de los datos subyacentes.
    - [ChartCollection](/javascript/api/excel/excel.chartcollection): una colección de gráficos en una hoja de cálculo.

- [TableSort](/javascript/api/excel/excel.tablesort): representa un objeto que administra las operaciones de ordenación en objetos **Table**.

- [RangeSort](/javascript/api/excel/excel.rangesort): representa un objeto que administra las operaciones de ordenación en objetos **Range**.

- [Filter](/javascript/api/excel/excel.filter): representa un objeto que administra el filtrado de la columna de una tabla.

- [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection): representa la protección de un objeto **Worksheet**.

- [NamedItem](/javascript/api/excel/excel.nameditem): representa un nombre definido para un rango de celdas o un valor. 
    - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection): una colección de los objetos **NamedItem** de un libro.

- [Binding](/javascript/api/excel/excel.binding): clase abstracta que representa un enlace a una sección del libro.
    - [BindingCollection](/javascript/api/excel/excel.bindingcollection): una colección de objetos **Binding** de un libro.

## <a name="excel-javascript-api-open-specifications"></a>Especificaciones abiertas de API de JavaScript de Excel

Diseñar y desarrollar nuevas API para los complementos de Excel, se podrán que estén disponibles para enviar sus comentarios en nuestra página de [especificaciones de la API de Open](../openspec.md) . Averigüe qué nuevas características en la canalización para las APIs de JavaScript de Excel y proporcionan su información en nuestras especificaciones de diseño.

## <a name="excel-javascript-api-reference"></a>Referencia de la API de JavaScript para Excel

Para obtener información detallada acerca de la API de JavaScript de Excel, consulte la [documentación de referencia de la API de JavaScript de Excel](/javascript/api/excel).

## <a name="see-also"></a>Vea también

- [Información general sobre los complementos de Excel](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-overview)
- [Informaci?n general sobre la plataforma de complementos de Office](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Ejemplos de complementos en depósito de Excel](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
