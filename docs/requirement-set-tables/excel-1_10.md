| Clase | Campos | Descripción |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|El contenido del comentario.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Elimina el comentario y todas las respuestas conectadas.|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|Obtiene la celda donde se encuentra este comentario.|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|Obtiene el correo electrónico del autor del comentario.|
||[authorName](/javascript/api/excel/excel.comment#authorname)|Obtiene el nombre del autor del comentario.|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|Obtiene la hora de creación del comentario.|
||[id](/javascript/api/excel/excel.comment#id)|Especifica el identificador de comentario.|
||[replies](/javascript/api/excel/excel.comment#replies)|Indica una colección de objetos de respuesta asociados con el comentario.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Crea un nuevo comentario con el contenido específico de la celda.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Obtiene el número de comentarios de la colección.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Obtiene un comentario de la colección en función de su identificador.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Obtiene un comentario de la colección en función de su posición.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Obtiene el comentario de la celda especificada.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Obtiene el comentario al que está conectada la respuesta dada.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Contenido de la respuesta del comentario.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Elimina la respuesta del comentario.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Obtiene la celda donde se encuentra esta respuesta de comentario.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Obtiene el comentario primario de esta respuesta.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Obtiene el correo electrónico del autor de la respuesta del comentario.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Obtiene el nombre del autor de la respuesta del comentario.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Obtiene la hora de creación de la respuesta del comentario.|
||[id](/javascript/api/excel/excel.commentreply#id)|Especifica el identificador de respuesta de comentario.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Crea una respuesta de comentario para un comentario.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Obtiene el número de respuestas de comentarios de la colección.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Devuelve una respuesta de comentario identificada por su Id.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Obtiene una respuesta comentario en función de su posición en la colección.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|Especifica si la lista de campos se puede mostrar en la interfaz de usuario.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|Elimina el estilo de tabla dinámica.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|Crea un duplicado de este estilo de tabla dinámica con copias de todos los elementos de estilo.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Obtiene el nombre del estilo de tabla dinámica.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|Especifica si este `PivotTableStyle` objeto es de solo lectura.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|Crea un espacio `PivotTableStyle` en blanco con el nombre especificado.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|Obtiene el número de estilos de tabla dinámica en la colección.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|Obtiene el estilo de tabla dinámica predeterminado para el ámbito del objeto primario.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|Obtiene un `PivotTableStyle` por su nombre.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|Obtiene un `PivotTableStyle` por su nombre.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|Establece el estilo de tabla dinámica predeterminado para su uso en el ámbito del objeto primario.|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#group-groupoption-)|Agrupa columnas y filas para un esquema.|
||[hideGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#hidegroupdetails-groupoption-)|Oculta los detalles del grupo de filas o columnas.|
||[height](/javascript/api/excel/excel.range#height)|Devuelve la distancia en puntos, para el zoom 100%, desde el borde superior del rango hasta el borde inferior del intervalo.|
||[left](/javascript/api/excel/excel.range#left)|Devuelve la distancia en puntos, para el zoom 100%, desde el borde izquierdo de la hoja de cálculo hasta el borde izquierdo del rango.|
||[top](/javascript/api/excel/excel.range#top)|Devuelve la distancia en puntos, para el zoom 100%, desde el borde superior de la hoja de cálculo hasta el borde superior del rango.|
||[width](/javascript/api/excel/excel.range#width)|Devuelve la distancia en puntos, para el zoom 100%, desde el borde izquierdo del rango hasta el borde derecho del intervalo.|
||[showGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#showgroupdetails-groupoption-)|Muestra los detalles del grupo de filas o columnas.|
||[ungroup(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#ungroup-groupoption-)|Desagrupa columnas y filas para un esquema.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Copia y pega un `Shape` objeto.|
||[placement](/javascript/api/excel/excel.shape#placement)|Representa cómo está asociado el objeto a las celdas inferiores.|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Representa el título de la segmentación de datos.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|Borra todos los filtros aplicados actualmente en la segmentación.|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|Elimina la segmentación.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|Devuelve una matriz de claves de los elementos seleccionados.|
||[height](/javascript/api/excel/excel.slicer#height)|Indica el alto, en puntos, de la segmentación.|
||[left](/javascript/api/excel/excel.slicer#left)|La distancia, en puntos, desde el lado izquierdo de la segmentación hasta el izquierdo de la hoja de cálculo.|
||[name](/javascript/api/excel/excel.slicer#name)|Representa el nombre de la segmentación de datos.|
||[id](/javascript/api/excel/excel.slicer#id)|Representa el identificador único de la segmentación de datos.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|El valor es si se borran todos los filtros aplicados actualmente en la `true` segmentación de datos.|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|Representa la colección de elementos de segmentación de datos que forman parte de la segmentación de datos.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Indica la hoja de cálculo que contiene la segmentación.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Selecciona los elementos de segmentación de datos en función de sus claves.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Indica el orden de los elementos de la segmentación.|
||[style](/javascript/api/excel/excel.slicer#style)|Valor constante que representa el estilo de segmentación de datos.|
||[top](/javascript/api/excel/excel.slicer#top)|La distancia, en puntos, desde el borde superior de la segmentación hasta la parte superior de la hoja de cálculo.|
||[width](/javascript/api/excel/excel.slicer#width)|Indica el ancho, en puntos, de la segmentación.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|Agrega una nueva segmentación al libro.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|Devuelve el número de segmentaciones incluidas en la colección.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|Obtiene un objeto slicer con su nombre o identificador.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|Obtiene una segmentación basándose en su posición en la colección.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|Obtiene una segmentación de datos con su nombre o identificador.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|El valor `true` es si el elemento de segmentación de datos está seleccionado.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|El valor `true` es si el elemento de segmentación de datos tiene datos.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Indica el valor único que representa el elemento de segmentación.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Representa el título que se muestra en la interfaz de usuario de Excel.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Indica el número de elementos de segmentación en la segmentación.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Obtiene una objeto de elemento de segmentación con su nombre o clave.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Obtiene un elemento de segmentación basándose en su posición en la colección.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Obtiene un elemento de segmentación mediante su nombre o la clave.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|Elimina el estilo de segmentación de datos.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|Crea un duplicado de este estilo de segmentación de datos con copias de todos los elementos de estilo.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Obtiene el nombre del estilo de segmentación de datos.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|Especifica si este `SlicerStyle` objeto es de solo lectura.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|Crea un estilo de segmentación de datos en blanco con el nombre especificado.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|Obtiene el número de estilos de segmentación en la colección.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|Obtiene el valor `SlicerStyle` predeterminado del ámbito del objeto primario.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|Obtiene un `SlicerStyle` por su nombre.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|Obtiene un `SlicerStyle` por su nombre.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|Establece el estilo de segmentación de datos predeterminado para su uso en el ámbito del objeto primario.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|Elimina el estilo de tabla.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|Crea un duplicado de este estilo de tabla con copias de todos los elementos de estilo.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Obtiene el nombre del estilo de tabla.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|Especifica si este `TableStyle` objeto es de solo lectura.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|Crea un espacio `TableStyle` en blanco con el nombre especificado.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|Obtiene el número de estilos de tabla en la colección.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|Obtiene el estilo de tabla predeterminado para el ámbito del objeto primario.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|Obtiene un `TableStyle` por su nombre.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|Obtiene un `TableStyle` por su nombre.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|Establece el estilo de tabla predeterminado para su uso en el ámbito del objeto primario.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|Elimina el estilo de tabla.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|Crea un duplicado de este estilo de escala de tiempo con copias de todos los elementos de estilo.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Obtiene el nombre del estilo de escala de tiempo.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|Especifica si este `TimelineStyle` objeto es de solo lectura.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|Crea un espacio `TimelineStyle` en blanco con el nombre especificado.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|Obtiene el número de estilos de escala de tiempo de la colección.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|Obtiene el estilo de escala de tiempo predeterminado del ámbito del objeto primario.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|Obtiene un `TimelineStyle` por su nombre.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|Obtiene un `TimelineStyle` por su nombre.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|Establece el estilo de escala de tiempo predeterminado para su uso en el ámbito del objeto primario.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Obtiene la segmentación activa del libro.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Obtiene la segmentación activa del libro.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Representa una colección de comentarios asociados con el libro.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Representa una colección de PivotTableStyles asociados con el libro.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Representa una colección de SlicerStyles asociados con el libro.|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|Representa una colección de segmentaciones de datos asociadas al libro.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Representa una colección de TableStyles asociados con el libro.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Representa una colección de TimelineStyles asociados con el libro.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Devuelve una colección de todos los objetos Comments en la hoja de cálculo.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|Se produce cuando se han ordenado una o más columnas.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|Se produce cuando se han ordenado una o más filas.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|Se produce cuando se produce una acción con clic izquierdo o pulsada en la hoja de cálculo.|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|Devuelve una colección de segmentaciones de datos que forman parte de la hoja de cálculo.|
||[showOutlineLevels(rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)|Muestra los grupos de filas o columnas por sus niveles de esquema.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|Se produce cuando se han ordenado una o más columnas.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|Se produce cuando se han ordenado una o más filas.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|Se produce cuando se produce una operación con clic izquierdo o pulsada en la colección de hojas de cálculo.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Obtiene la dirección del rango que representa el área seleccionada de una hoja de cálculo específica.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Obtiene el origen del evento.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo donde se produjo la ordenación.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Obtiene la dirección del rango que representa el área seleccionada de una hoja de cálculo específica.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Obtiene el origen del evento.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo donde se produjo la ordenación.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Obtiene la dirección que representa la celda que se pulsó o en la que se hizo clic izquierdo de una hoja de cálculo específica.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|Distancia, en puntos, desde el punto con clic izquierdo o pulsado hasta el borde de la cuadrícula izquierda (o derecha para los idiomas de derecha a izquierda) de la celda con clic izquierdo o pulsada.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|La distancia, en puntos, desde el punto que se pulsó o en el que se hizo clic izquierdo hasta el borde superior de la cuadrícula de la celda que se pulsó o en la que se hizo clic izquierdo.|
||[tipo](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se hizo clic con el botón izquierdo o se punteó la celda.|
