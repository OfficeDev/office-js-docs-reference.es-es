| Clase | Campos | Descripción |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(email: string)](/javascript/api/excel/excel.comment#assigntask-email-)|Asigna la tarea adjunta al comentario al usuario determinado como único usuario al que se asigna.|
||[getTask()](/javascript/api/excel/excel.comment#gettask--)|Obtiene la tarea asociada a este comentario.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Obtiene la tarea asociada a este comentario.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(email: string)](/javascript/api/excel/excel.commentreply#assigntask-email-)|Asigna la tarea adjunta al comentario al usuario determinado como único usuario al que se asigna.|
||[getTask()](/javascript/api/excel/excel.commentreply#gettask--)|Obtiene la tarea asociada a este comentario.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Obtiene la tarea asociada a este comentario.|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|Dirección de la celda que contiene la fórmula modificada.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|Representa la fórmula anterior, antes de cambiarla.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|Nombre del proveedor de datos para el tipo de datos vinculado.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|Fecha y hora de la zona horaria local desde que se abrió el libro cuando se actualizó el tipo de datos vinculado por última vez.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|Nombre del tipo de datos vinculado.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|Frecuencia, en segundos, con la que se actualiza el tipo de datos vinculado `refreshMode` si se establece en "Periódico".|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|Mecanismo mediante el cual se recuperan los datos del tipo de datos vinculado.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|Identificador único del tipo de datos vinculado.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Devuelve una matriz con todos los modos de actualización admitidos por el tipo de datos vinculado.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Realiza una solicitud para actualizar el tipo de datos vinculado.|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Realiza una solicitud para cambiar el modo de actualización de este tipo de datos vinculados.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|Identificador único del nuevo tipo de datos vinculado.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Obtiene el origen del evento.|
||[tipo](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Obtiene el tipo del evento.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Obtiene el número de tipos de datos vinculados de la colección.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Obtiene un tipo de datos vinculado por identificador de servicio.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Obtiene un tipo de datos vinculado por su índice en la colección.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Obtiene un tipo de datos vinculados por identificador.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Realiza una solicitud para actualizar todos los tipos de datos vinculados de la colección.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Activa esta vista de hoja.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Quita la vista de hoja de la hoja de cálculo.|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Crea una copia de esta vista de hoja.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Obtiene o establece el nombre de la vista de hoja.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Crea una nueva vista de hoja con el nombre especificado.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Crea y activa una nueva vista de hoja temporal.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Sale de la vista de hoja activa actualmente.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Obtiene la vista de hoja activa actualmente de la hoja de cálculo.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Obtiene el número de vistas de hoja de esta hoja de cálculo.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Obtiene una vista de hoja mediante su nombre.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Obtiene una vista de hoja por su índice en la colección.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|Descripción del texto alternativo de la tabla dinámica.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|Título del texto alternativo de la tabla dinámica.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Establece si se va a mostrar una línea en blanco después de cada elemento.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Texto que se rellena automáticamente en cualquier celda vacía de la tabla dinámica si `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Especifica si las celdas vacías de la tabla dinámica deben rellenarse con el `emptyCellText` archivo .|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtiene una única celda de la tabla dinámica en función de una jerarquía de datos y de los elementos de fila y columna de sus respectivas jerarquías.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Estilo aplicado a la tabla dinámica.|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Establece la opción "repetir todas las etiquetas de elementos" en todos los campos de la tabla dinámica.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Establece el estilo aplicado a la tabla dinámica.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Especifica si la tabla dinámica muestra encabezados de campo (títulos de campo y menús desplegables de filtro).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Especifica si la tabla dinámica se actualiza cuando se abre el libro.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Devuelve un objeto que representa el rango que contiene todos los precedentes de una celda en la misma hoja de cálculo `WorkbookRangeAreas` o en varias hojas de cálculo.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|Modo de actualización del tipo de datos vinculado.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|Identificador único del objeto cuyo modo de actualización se ha cambiado.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Obtiene el origen del evento.|
||[tipo](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Obtiene el tipo del evento.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[actualizado](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indica si la solicitud de actualización se ha realizado correctamente.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|Identificador único del objeto cuya solicitud de actualización se completó.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Obtiene el origen del evento.|
||[tipo](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Obtiene el tipo del evento.|
||[advertencias](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Matriz que contiene las advertencias generadas a partir de la solicitud de actualización.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Crea un gráfico vectorial escalable (SVG) de una cadena XML y lo agrega a la hoja de cálculo.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Indica el nombre de la segmentación usado en la fórmula.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Estilo aplicado a la segmentación de datos.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Establece el estilo aplicado a la segmentación de datos.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Cambia la tabla para usar el estilo de tabla predeterminado.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Se produce cuando se aplica el filtro en una tabla en particular.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Estilo aplicado a la tabla.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Establece el estilo aplicado a la tabla.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Se produce cuando se aplica el filtro en cualquier tabla en un libro o una hoja de cálculo.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Obtiene el identificador de la tabla en la que se aplica el filtro.|
||[tipo](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo que contiene la tabla.|
|[Task](/javascript/api/excel/excel.task)|[addAssignee(email: string)](/javascript/api/excel/excel.task#addassignee-email-)|Agrega un usuario al que se asigna a la tarea.|
||[applyChanges(taskChanges: Excel.TaskChanges)](/javascript/api/excel/excel.task#applychanges-taskchanges-)|Aplica los cambios dados a la tarea.|
||[assignees](/javascript/api/excel/excel.task#assignees)|Obtiene los usuarios a los que se asigna la tarea.|
||[comment](/javascript/api/excel/excel.task#comment)|Obtiene el comentario asociado a la tarea.|
||[dueDate](/javascript/api/excel/excel.task#duedate)|Obtiene la fecha y hora en que vence la tarea.|
||[historyRecords](/javascript/api/excel/excel.task#historyrecords)|Obtiene los registros de historial de la tarea.|
||[id](/javascript/api/excel/excel.task#id)|Obtiene el identificador de la tarea.|
||[percentComplete](/javascript/api/excel/excel.task#percentcomplete)|Obtiene el porcentaje de finalización de la tarea.|
||[priority](/javascript/api/excel/excel.task#priority)|Obtiene la prioridad de la tarea.|
||[startDate](/javascript/api/excel/excel.task#startdate)|Obtiene la fecha y hora de inicio de la tarea.|
||[title](/javascript/api/excel/excel.task#title)|Obtiene el título de la tarea.|
||[removeAllAssignees()](/javascript/api/excel/excel.task#removeallassignees--)|Quita todos los usuarios asignados de la tarea.|
||[removeAssignee(email: string)](/javascript/api/excel/excel.task#removeassignee-email-)|Quita un usuario asignado de la tarea.|
||[setPercentComplete(percentComplete: number)](/javascript/api/excel/excel.task#setpercentcomplete-percentcomplete-)|Cambia la finalización de la tarea.|
||[setPriority(priority: number)](/javascript/api/excel/excel.task#setpriority-priority-)|Cambia la prioridad de la tarea.|
||[setStartDateAndDueDate(startDate: Date, dueDate: Date)](/javascript/api/excel/excel.task#setstartdateandduedate-startdate--duedate-)|Cambia el inicio y las fechas de vencimiento de la tarea.|
||[setTitle(title: string)](/javascript/api/excel/excel.task#settitle-title-)|Cambia el título de la tarea.|
|[TaskChanges](/javascript/api/excel/excel.taskchanges)|[dueDate](/javascript/api/excel/excel.taskchanges#duedate)|Establece una nueva fecha de vencimiento para la tarea, en la zona horaria UTC.|
||[emailsToAssign](/javascript/api/excel/excel.taskchanges#emailstoassign)|Establece las direcciones de correo electrónico de los usuarios que se asignarán a la tarea.|
||[emailsToUnassign](/javascript/api/excel/excel.taskchanges#emailstounassign)|Establece las direcciones de correo electrónico de los usuarios para que se desasigne de la tarea.|
||[percentComplete](/javascript/api/excel/excel.taskchanges#percentcomplete)|Establece un nuevo porcentaje de finalización para la tarea.|
||[priority](/javascript/api/excel/excel.taskchanges#priority)|Establece una nueva prioridad para la tarea.|
||[removeAllPreviousAssignees](/javascript/api/excel/excel.taskchanges#removeallpreviousassignees)|Establece si el cambio debe quitar todos los usuarios asignados anteriores de la tarea.|
||[startDate](/javascript/api/excel/excel.taskchanges#startdate)|Establece una nueva fecha de inicio para la tarea, en la zona horaria UTC.|
||[title](/javascript/api/excel/excel.taskchanges#title)|Establece un nuevo título para la tarea.|
|[TaskCollection](/javascript/api/excel/excel.taskcollection)|[getCount()](/javascript/api/excel/excel.taskcollection#getcount--)|Obtiene el número de tareas de la colección.|
||[getItem(key: string)](/javascript/api/excel/excel.taskcollection#getitem-key-)|Obtiene una tarea mediante su identificador.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskcollection#getitemat-index-)|Obtiene una tarea por su índice en la colección.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.taskcollection#getitemornullobject-key-)|Obtiene una tarea mediante su identificador.|
||[items](/javascript/api/excel/excel.taskcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[TaskHistoryRecord](/javascript/api/excel/excel.taskhistoryrecord)|[anchorId](/javascript/api/excel/excel.taskhistoryrecord#anchorid)|Representa el identificador del objeto al que está anclada la tarea (por ejemplo, commentId para tareas adjuntas a comentarios).|
||[assignee](/javascript/api/excel/excel.taskhistoryrecord#assignee)|Representa el usuario asignado a la tarea para un tipo de registro de historial "Asignar" o el usuario que se desasigna de la tarea para un tipo de registro de historial "Unassign".|
||[attributionUser](/javascript/api/excel/excel.taskhistoryrecord#attributionuser)|Representa el usuario que creó o modificó la tarea.|
||[dueDate](/javascript/api/excel/excel.taskhistoryrecord#duedate)|Representa la fecha de vencimiento de la tarea.|
||[historyRecordCreatedDate](/javascript/api/excel/excel.taskhistoryrecord#historyrecordcreateddate)|Representa la fecha de creación del registro del historial de tareas.|
||[id](/javascript/api/excel/excel.taskhistoryrecord#id)|Identificador del registro de historial.|
||[percentComplete](/javascript/api/excel/excel.taskhistoryrecord#percentcomplete)|Representa el porcentaje de finalización de la tarea.|
||[priority](/javascript/api/excel/excel.taskhistoryrecord#priority)|Representa la prioridad de la tarea.|
||[startDate](/javascript/api/excel/excel.taskhistoryrecord#startdate)|Representa la fecha de inicio de la tarea.|
||[title](/javascript/api/excel/excel.taskhistoryrecord#title)|Representa el título de la tarea.|
||[tipo](/javascript/api/excel/excel.taskhistoryrecord#type)|Representa el tipo de registro del historial de tareas.|
||[undoHistoryId](/javascript/api/excel/excel.taskhistoryrecord#undohistoryid)|Representa la TaskHistoryRecord.id que se ha deshecho para el tipo de registro de historial "Deshacer".|
|[TaskHistoryRecordCollection](/javascript/api/excel/excel.taskhistoryrecordcollection)|[getCount()](/javascript/api/excel/excel.taskhistoryrecordcollection#getcount--)|Obtiene el número de registros de historial de la colección para la tarea.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskhistoryrecordcollection#getitemat-index-)|Obtiene un registro de historial de tareas mediante su índice en la colección.|
||[items](/javascript/api/excel/excel.taskhistoryrecordcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[Usuario](/javascript/api/excel/excel.user)|[displayName](/javascript/api/excel/excel.user#displayname)|Representa el nombre para mostrar del usuario.|
||[email](/javascript/api/excel/excel.user#email)|Representa la dirección de correo del usuario|
||[uid](/javascript/api/excel/excel.user#uid)|Representa el identificador único del usuario.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Devuelve una colección de tipos de datos vinculados que forman parte del libro.|
||[tareas](/javascript/api/excel/excel.workbook#tasks)|Devuelve una colección de tareas que están presentes en el libro.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Especifica si el panel de lista de campos de la tabla dinámica se muestra en el nivel del libro.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True si el libro usa el sistema de fechas 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Devuelve una colección de vistas de hoja que están presentes en la hoja de cálculo.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Se produce cuando se aplica el filtro en una hoja de cálculo concreta.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|Se produce cuando se cambian una o más fórmulas en esta hoja de cálculo.|
||[tareas](/javascript/api/excel/excel.worksheet#tasks)|Devuelve una colección de tareas que están presentes en la hoja de cálculo.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Inserta las hojas de cálculo especificadas de un libro en el libro actual.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Se produce cuando se aplica cualquier filtro de hoja de cálculo al libro.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|Se produce cuando se cambian una o más fórmulas en cualquier hoja de cálculo de esta colección.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se aplica el filtro.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|Obtiene una matriz de objetos FormulaChangedEventDetail, que contienen los detalles sobre todas las fórmulas modificadas.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|Origen del evento.|
||[tipo](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se cambió la fórmula.|
