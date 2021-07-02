| Clase | Campos | Descripción |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearcolumncriteria-columnindex-)|Borra los criterios de filtro de AutoFilter.|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteshiftdirection)|Representa la dirección (como hacia arriba o hacia la izquierda) que las celdas restantes cambiarán cuando se elimine una celda o celda.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertshiftdirection)|Representa la dirección (por ejemplo, hacia abajo o hacia la derecha) que las celdas existentes cambiarán cuando se inserte una nueva celda o celda.|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assigntask-assignee-)|Asigna la tarea adjunta al comentario al usuario dado como usuario asignado.|
||[getTask()](/javascript/api/excel/excel.comment#gettask--)|Obtiene la tarea asociada a este comentario.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Obtiene la tarea asociada a este comentario.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getitemornullobject-commentid-)|Obtiene un comentario de la colección en función de su identificador.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assigntask-assignee-)|Asigna la tarea adjunta al comentario al usuario dado como único destinatario.|
||[getTask()](/javascript/api/excel/excel.commentreply#gettask--)|Obtiene la tarea asociada con el subproceso de la respuesta de este comentario.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Obtiene la tarea asociada con el subproceso de la respuesta de este comentario.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitemornullobject-commentreplyid-)|Devuelve una respuesta de comentario identificada por su Id.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitemornullobject-id-)|Devuelve un formato condicional identificado por su identificador.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[percentComplete](/javascript/api/excel/excel.documenttask#percentcomplete)|Especifica el porcentaje de finalización de la tarea.|
||[prioridad](/javascript/api/excel/excel.documenttask#priority)|Especifica la prioridad de la tarea.|
||[assignees](/javascript/api/excel/excel.documenttask#assignees)|Devuelve una colección de usuarios asignados de la tarea.|
||[cambios](/javascript/api/excel/excel.documenttask#changes)|Obtiene los registros de cambios de la tarea.|
||[comment](/javascript/api/excel/excel.documenttask#comment)|Obtiene el comentario asociado a la tarea.|
||[completedBy](/javascript/api/excel/excel.documenttask#completedby)|Obtiene al usuario más reciente que ha completado la tarea.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completeddatetime)|Obtiene la fecha y hora en que se completó la tarea.|
||[createdBy](/javascript/api/excel/excel.documenttask#createdby)|Obtiene al usuario que creó la tarea.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createddatetime)|Obtiene la fecha y hora en que se creó la tarea.|
||[id](/javascript/api/excel/excel.documenttask#id)|Obtiene el identificador de la tarea.|
||[setStartAndDueDateTime(startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#setstartandduedatetime-startdatetime--duedatetime-)|Cambia el inicio y las fechas de vencimiento de la tarea.|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#startandduedatetime)|Obtiene o establece la fecha y hora en que debe iniciarse la tarea y vence.|
||[title](/javascript/api/excel/excel.documenttask#title)|Especifica el título de la tarea.|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[assignee](/javascript/api/excel/excel.documenttaskchange#assignee)|Representa el usuario asignado a la tarea para un tipo de registro de cambio o el usuario sin asignar de la tarea `assign` para un tipo de registro de `unassign` cambio.|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#changedby)|Representa el usuario que creó o cambió la tarea.|
||[commentId](/javascript/api/excel/excel.documenttaskchange#commentid)|Representa el identificador del cambio de tarea o al que está `Comment` `CommentReply` anclado.|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#createddatetime)|Representa la fecha y hora de creación del registro de cambio de tarea.|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#duedatetime)|Representa la fecha y hora de vencimiento de la tarea, en la zona horaria UTC.|
||[id](/javascript/api/excel/excel.documenttaskchange#id)|Id. del registro de cambio de tarea.|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#percentcomplete)|Representa el porcentaje de finalización de la tarea.|
||[prioridad](/javascript/api/excel/excel.documenttaskchange#priority)|Representa la prioridad de la tarea.|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#startdatetime)|Representa la fecha y hora de inicio de la tarea, en la zona horaria UTC.|
||[title](/javascript/api/excel/excel.documenttaskchange#title)|Representa el título de la tarea.|
||[type](/javascript/api/excel/excel.documenttaskchange#type)|Representa el tipo de acción del registro de cambio de tarea.|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#undohistoryid)|Representa la `DocumentTaskChange.id` propiedad que se deshizo para el `undo` tipo de registro de cambio.|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#getcount--)|Obtiene el número de registros de cambios de la colección para la tarea.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#getitemat-index-)|Obtiene un registro de cambio de tarea mediante su índice en la colección.|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#getcount--)|Obtiene el número de tareas de la colección.|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitem-key-)|Obtiene una tarea con su identificador.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#getitemat-index-)|Obtiene una tarea por su índice en la colección.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitemornullobject-key-)|Obtiene una tarea con su identificador.|
||[items](/javascript/api/excel/excel.documenttaskcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#duedatetime)|Obtiene la fecha y hora en que vence la tarea.|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#startdatetime)|Obtiene la fecha y hora en que debe iniciarse la tarea.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getitemornullobject-key-)|Obtiene una forma con su nombre o identificador.|
|[Identity](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayname)|Representa el nombre para mostrar del usuario.|
||[email](/javascript/api/excel/excel.identity#email)|Representa la dirección de correo del usuario|
||[id](/javascript/api/excel/excel.identity#id)|Representa el identificador único del usuario.|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#add-assignee-)|Agrega una identidad de usuario a la colección.|
||[clear()](/javascript/api/excel/excel.identitycollection#clear--)|Quita todas las identidades de usuario de la colección.|
||[getCount()](/javascript/api/excel/excel.identitycollection#getcount--)|Obtiene el número de objetos de la colección.|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getitemat-index-)|Obtiene una identidad de usuario de documento mediante su índice en la colección.|
||[items](/javascript/api/excel/excel.identitycollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#remove-assignee-)|Quita una identidad de usuario de la colección.|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#displayname)|Representa el nombre para mostrar del usuario.|
||[email](/javascript/api/excel/excel.identityentity#email)|Representa la dirección de correo del usuario|
||[id](/javascript/api/excel/excel.identityentity#id)|Representa el identificador único del usuario.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|Nombre del proveedor de datos para el tipo de datos vinculado.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|Fecha y hora de la zona horaria local desde que se abrió el libro cuando se actualizó por última vez el tipo de datos vinculado.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|Nombre del tipo de datos vinculado.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|Frecuencia, en segundos, a la que se actualiza el tipo de datos vinculado si `refreshMode` se establece en "Periódico".|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|Mecanismo por el que se recuperan los datos del tipo de datos vinculado.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|Identificador único del tipo de datos vinculado.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Devuelve una matriz con todos los modos de actualización admitidos por el tipo de datos vinculado.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Realiza una solicitud para actualizar el tipo de datos vinculado.|
||[requestSetRefreshMode(refreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Realiza una solicitud para cambiar el modo de actualización de este tipo de datos vinculado.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|El identificador único del nuevo tipo de datos vinculado.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Obtiene el origen del evento.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Obtiene el tipo del evento.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Obtiene el número de tipos de datos vinculados de la colección.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Obtiene un tipo de datos vinculado por identificador de servicio.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Obtiene un tipo de datos vinculado por su índice en la colección.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Obtiene un tipo de datos vinculado por identificador.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Realiza una solicitud para actualizar todos los tipos de datos vinculados de la colección.|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#breaklinks--)|Realiza una solicitud para romper los vínculos que apuntan al libro vinculado.|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|Dirección URL original que apunta al libro vinculado.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh--)|Realiza una solicitud para actualizar los datos recuperados del libro vinculado.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#breakalllinks--)|Rompe todos los vínculos a los libros vinculados.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getitem-key-)|Obtiene información sobre un libro vinculado por su dirección URL.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getitemornullobject-key-)|Obtiene información sobre un libro vinculado por su dirección URL.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#refreshall--)|Realiza una solicitud para actualizar todos los vínculos del libro.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbooklinksrefreshmode)|Representa el modo de actualización de los vínculos del libro.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitemornullobject-key-)|Obtiene una vista de hoja con su nombre.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtiene una única celda de la tabla dinámica en función de una jerarquía de datos y de los elementos de fila y columna de sus respectivas jerarquías.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Estilo aplicado a la tabla dinámica.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Establece el estilo aplicado a la tabla dinámica.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getfirstornullobject--)|Obtiene la primera tabla dinámica de la colección.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getdependents--)|Devuelve un objeto que representa el rango que contiene todos los dependientes de una celda en la misma hoja de cálculo `WorkbookRangeAreas` o en varias hojas de cálculo.|
||[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Devuelve un objeto que representa el rango que contiene todos los precedentes de una celda en la misma hoja de cálculo `WorkbookRangeAreas` o en varias hojas de cálculo.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|Modo de actualización del tipo de datos vinculado.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|Identificador único del objeto cuyo modo de actualización se ha cambiado.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Obtiene el origen del evento.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Obtiene el tipo del evento.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[actualizado](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indica si la solicitud de actualización se ha realizado correctamente.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|Identificador único del objeto cuya solicitud de actualización se completó.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Obtiene el origen del evento.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Obtiene el tipo del evento.|
||[advertencias](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Matriz que contiene las advertencias generadas a partir de la solicitud de actualización.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Crea un gráfico vectorial escalable (SVG) de una cadena XML y lo agrega a la hoja de cálculo.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getitemornullobject-key-)|Obtiene una forma con su nombre o identificador.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Indica el nombre de la segmentación usado en la fórmula.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Estilo aplicado a la segmentación de datos.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Establece el estilo aplicado a la segmentación de datos.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getitemornullobject-name-)|Obtiene un estilo por nombre.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Cambia la tabla para usar el estilo de tabla predeterminado.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Se produce cuando se aplica un filtro en una tabla específica.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Estilo aplicado a la tabla.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Establece el estilo aplicado a la tabla.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Se produce cuando se aplica un filtro en cualquier tabla de un libro o una hoja de cálculo.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Obtiene el identificador de la tabla en la que se aplica el filtro.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo que contiene la tabla.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitemornullobject-key-)|Obtener una tabla por nombre o identificador.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Devuelve una colección de tipos de datos vinculados que forman parte del libro.|
||[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedworkbooks)|Devuelve una colección de libros vinculados.|
||[tareas](/javascript/api/excel/excel.workbook#tasks)|Devuelve una colección de tareas que están presentes en el libro.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Especifica si el panel de lista de campos de la tabla dinámica se muestra en el nivel del libro.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True si el libro usa el sistema de fechas 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Se produce cuando se aplica un filtro en una hoja de cálculo específica.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheet#onprotectionchanged)|Se produce cuando se cambia el estado de protección de la hoja de cálculo.|
||[tabId](/javascript/api/excel/excel.worksheet#tabid)|Devuelve un valor que representa esta hoja de cálculo que puede leer Open Office XML.|
||[tareas](/javascript/api/excel/excel.worksheet#tasks)|Devuelve una colección de tareas que están presentes en la hoja de cálculo.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changedirectionstate)|Representa un cambio en la dirección en que las celdas de una hoja de cálculo cambiarán cuando se elimine o inserte una celda o celda.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggersource)|Representa el origen del desencadenador del evento.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Inserta las hojas de cálculo especificadas de un libro en el libro actual.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Se produce cuando se aplica cualquier filtro de hoja de cálculo al libro.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onprotectionchanged)|Se produce cuando se cambia el estado de protección de la hoja de cálculo.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se aplica el filtro.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isprotected)|Obtiene el estado de protección actual de la hoja de cálculo.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|Origen del evento.|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que se cambia el estado de protección.|
