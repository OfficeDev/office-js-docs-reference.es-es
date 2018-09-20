# <a name="javascript-api-for-office-error-codes"></a>Códigos de error de la API de JavaScript de Office

Este artículo muestra los mensajes de error que se puede encontrar al usar la API de JavaScript para Office (Office.js).

**Se aplica a:** Complementos de Office | Complementos de SharePoint | Excel | Outlook | PowerPoint | Project | Word

## <a name="error-codes"></a>Códigos de error

La tabla siguiente enumera los códigos, nombres y mensajes de error mostrados, así como las condiciones que indican.

|**Error.code**|**Error.name**|**Error.message**|**Condición**|
|:-----|:-----|:-----|:-----|
|1000|Tipo de conversión no válido|No se admite el tipo de conversión especificado.|No se admite el tipo de coerción en la aplicación host. (Por ejemplo, no se admiten los tipos de coerción OOXML y HTML en Excel).|
|1001|Error de lectura de datos|No se admite la selección actual.|No se admite la selección actual del usuario. (Es decir, hay algo diferente a los tipos de coerción admitidos).|
|1002|Tipo de conversión no válido|El tipo de conversión especificado no es compatible con este tipo de enlace.|El desarrollador de soluciones proporcionó una combinación incompatible de tipo de coerción y tipo de enlace.|
|1003|Error de lectura de datos|Los valores de rowCount o columnCount especificados no son válidos.|El usuario indica totales de filas o columnas no válidos.|
|1004|Error de lectura de datos|La selección actual no es compatible con el tipo de conversión especificado.|La selección actual no es compatible con el tipo de coerción que especifica esta aplicación.|
|1005|Error de lectura de datos|Los valores de startRow o startColumn no son válidos.|El usuario indica valores de startRow o startCol que no son válidos.|
|1006|Error de lectura de datos|No se pueden utilizar parámetros de coordenadas con el tipo de coerción "Tabla" cuando la tabla contiene celdas combinadas.|El usuario intenta obtener datos parciales de una tabla no uniforme (es decir, una tabla con celdas combinadas). |
|1007|Error de lectura de datos|El tamaño del documento es demasiado grande.|El usuario intenta obtener un documento que supera el tamaño compatible actualmente.|
|1008|Error de lectura de datos|El conjunto de datos solicitado es demasiado grande.|El usuario solicita leer datos que sobrepasan los límites de datos que definen lo complementos host.|
|1009|Error de lectura de datos|El tipo de archivo especificado no es compatible.|El usuario envía un tipo de archivo no válido.|
|2000|Error de escritura de datos|No se admite el tipo de objeto de datos proporcionado. |Se indicó un objeto de datos no admitido.|
|2001|Error de escritura de datos|No se puede escribir en la selección actual.|La selección actual del usuario no es compatible con una operación de escritura. (Por ejemplo, cuando el usuario selecciona una imagen).|
|2002|Error de escritura de datos|El objeto de datos proporcionado no es compatible con la forma o las dimensiones de la selección actual.|Se seleccionaron varias celdas (y la forma de la selección no coincide con la forma de los datos).Se seleccionaron varias celdas (y las dimensiones de la selección no coinciden con las dimensiones de los datos).|
|2003|Error de escritura de datos|La operación establecida no se pudo realizar porque el objeto de datos proporcionado sobrescribirá los datos.|Se ha seleccionado una sola celda y el objeto de datos proporcionado sobrescribe datos en la hoja de cálculo.|
|2004|Error de escritura de datos|El objeto de datos proporcionado no coincide con el tamaño de la selección actual.|El usuario proporciona un objeto que supera el tamaño de la selección actual.|
|2005|Error de escritura de datos|Los valores de startRow o startColumn especificados no son válidos.|El usuario proporciona valores de startRow o startCol que no son válidos.|
|2006|Error de formato no válido|El formato del objeto de datos especificado no es válido.|El desarrollador de soluciones proporciona una cadena HTML u OOXML no válida, o una cadena HTML mal formada.|
|2007|Objeto de datos no válido|El tipo de objeto de datos especificado no es compatible con la selección actual.|El desarrollador de soluciones proporciona un objeto de datos que no es compatible con el tipo de coerción especificado.|
|2008|Error de escritura de datos|Por determinar|Por determinar|
|2009|Error de escritura de datos|El objeto de datos especificado es demasiado grande.|El usuario intenta establecer datos que sobrepasan los límites de datos que definen los complementos host.|
|2010|Error de escritura de datos|No se pueden utilizar parámetros de coordenadas con el tipo de coerción Tabla cuando la tabla contiene celdas combinadas.|El usuario intenta establecer datos parciales de una tabla no uniforme (es decir, una tabla con celdas combinadas).|
|3000|No se pudo crear el enlace|No se puede enlazar a la selección actual.|La selección del usuario no es compatible con el enlace. (Por ejemplo, el usuario elige una imagen u otro tipo de objeto no admitido).|
|3001|No se pudo crear el enlace|Por determinar|Por determinar|
|3002|Error de enlace no válido|El enlace especificado no existe.|El desarrollador intenta enlazar a un enlace que no existe o se quitó.|
|3003|No se pudo crear el enlace|No se admiten las selecciones discontinuas.|El usuario está realizando selecciones múltiples.|
|3004|No se pudo crear el enlace|Un enlace no se puede crear con la selección actual y el tipo de enlace especificado.|Existen varias condiciones bajo las cuales puede ocurrir esto. Consulte la sección "Condiciones para no poder crear enlaces" más adelante en este artículo.|
|3005|Operación de enlace no válido|La operación no es compatible con este tipo de enlace.|El desarrollador envía una operación de agregar fila o columna en un tipo de enlace que no es  _tabla_.|
|3006|No se pudo crear el enlace|El elemento con nombre no existe.|No se pudo encontrar el elemento con nombre. No existe ninguna tabla o control de contenido con ese nombre.|
|3007|No se pudo crear el enlace|Se han encontrado varios objetos con el mismo nombre.|Error de colisión: hay varios controles de contenido con el mismo nombre y el error de colisión está establecido en  **true**.|
|3008|No se pudo crear el enlace|El tipo de enlace especificado no es compatible con el elemento con nombre suministrado.|El elemento con nombre no se puede enlazar al tipo. Por ejemplo, un control de contenido contiene texto, pero el desarrollador intentó enlazarlo mediante el tipo de coerción  _tabla_.|
|3009|Operación de enlace no válido|El tipo de enlace no es compatible.|Se usa para la compatibilidad con versiones anteriores.|
|3010|Operación de enlace no admitida|El contenido seleccionado debe estar en formato de tabla. Aplique el formato de tabla a los datos y vuelva a intentarlo.|El desarrollador está intentando usar los métodos **addRowsAsynch** o **deleteAllDataValuesAsynch** del objeto **TableBinding** en datos del tipo de coerción _matriz_.|
|4000|Error de configuración de lectura|El nombre de configuración especificado no existe.|Se proporcionó un nombre de configuración que no existe.|
|4001|Error de configuración de guardado|No se pudo guardar la configuración.|No se pudo guardar la configuración.|
|4002|Error de configuración obsoleto|No se pudo guardar la configuración porque está obsoleta.|La configuración está obsoleta y el desarrollador indicó que no se reemplazara.|
|5000|Error de configuración obsoleto|No se admite esta operación.|No se admite la operación en el host actual. Por ejemplo, Outlook llama a  **document.getSelectionAsync**.|
|5001|Error interno|Se ha producido un error interno.|Hace referencia a una condición de error interno, que se puede producir por cualquiera de las siguientes razones:<br/><table><tr><td>Un complemento en uso por parte de otro usuario con el que se comparte el libro de trabajo ha creado un enlace aproximadamente al mismo tiempo, de modo que el complemento debe reintentar el enlace.</tr></td><tr><td>Se ha producido un error desconocido.</tr></td><tr><td>Error.</tr></td><tr><td>Se denegó el acceso porque el usuario no es miembro de un rol autorizado.</tr></td><tr><td>Se denegó el acceso porque se requiere una comunicación segura y cifrada.</tr></td><tr><td>Los datos están obsoletos y el usuario debe confirmar la activación de las consultas para actualizarlos.</tr></td><tr><td>Se superó la cuota de CPU de la colección de sitios.</tr></td><tr><td>Se superó la cuota de memoria de la colección de sitios.</tr></td><tr><td>Se superó la cuota de memoria de sesión.</tr></td><tr><td>El libro se encuentra en un estado no válido y la operación no se puede realizar.</tr></td><tr><td>Se ha agotado el tiempo de espera de la sesión debido a la inactividad y el usuario debe volver a cargar el libro.</tr></td><tr><td>Se superó el número máximo de sesiones permitidas por usuario.</tr></td><tr><td>El usuario canceló la operación.</tr></td><tr><td>La operación no se puede completar porque está tardando demasiado.</tr></td><tr><td>La solicitud no se puede completar y debe volver a intentarse.</tr></td><tr><td>El período de prueba del producto ha expirado.</tr></td><tr><td>Se ha agotado el tiempo de espera de la sesión debido a la inactividad.</tr></td><tr><td>El usuario no tiene permiso para realizar la operación en el rango especificado.</tr></td><tr><td>La configuración regional del usuario no coincide con la sesión de colaboración actual.</tr></td><tr><td>El usuario ya no está conectado y debe actualizar o volver a abrir el libro.</tr></td><tr><td>El rango solicitado no existe en la hoja.</tr></td><tr><td>El usuario no tiene permiso para modificar el libro.</tr></td><tr><td>El libro no puede modificarse porque está bloqueado.</tr></td><tr><td>La sesión no puede guardar el libro automáticamente.</tr></td><tr><td>La sesión no puede actualizar el bloqueo en el archivo de libro.</tr></td><tr><td>La solicitud no se puede procesar y se debe volver a intentar.</tr></td><tr><td>La información de inicio de sesión del usuario no se pudo comprobar y debe volver a especificarse.</tr></td><tr><td>Se ha denegado el acceso al usuario.</tr></td><tr><td>El libro compartido debe actualizarse.</tr></td></table>|
|5002|Permiso denegado|La operación solicitada no se permite en el modo de documento actual.|El desarrollador de soluciones envía una operación establecida, pero el documento está en un modo que no permite modificaciones, como "Restringir edición".|
|5003|Error de registro de eventos|El objeto actual no admite el tipo de evento especificado.|El desarrollador de soluciones intenta registrar o anular el registro de un controlador en un evento que no existe.|
|5004|Llamada de API no válida|La llamada a API no es válida en el contexto actual.|Se ha realizado una llamada no válida para el contexto, por ejemplo, al intentar usar un objeto  **CustomXMLPart** en Excel.|
|5005|Datos obsoletos|Se produjo un error en la operación porque los datos están obsoletos en el servidor.|Hay que actualizar los datos del servidor.|
|5006|Tiempo de espera de sesión|Se agotó el tiempo de espera de la sesión del documento. Vuelva a cargar el documento. |Se agotó el tiempo de espera de la sesión.|
|5007|Llamada de API no válida|La enumeración no se admite en el contexto actual.|La enumeración no se admite en el contexto actual.|
|5009|Permiso denegado|Acceso denegado.|El complemento no tiene permiso para llamar a la API específica.|
|6000|Nodo no válido|No se encontró el nodo especificado.|No se encontró el nodo **CustomXmlPart**.|
|6100|Error de XML personalizado|Error de XML personalizado.|Llamada de API no válida|
|7000|Id. no válido|El id. especificado no existe.|Id. no válido.|
|7001|Navegación no válida|El objeto se encuentra en un lugar donde no se admite la navegación.|El usuario puede encontrar el objeto, pero no puede navegar hasta él. (Por ejemplo, en Word, el enlace es al encabezado, pie de página o un comentario).|
|7002|Navegación no válida|El objeto está bloqueado o protegido.|El usuario está intentando navegar hasta un intervalo bloqueado o protegido.|
|7004|Navegación no válida|No se pudo realizar la operación porque el índice está fuera del intervalo.|El usuario está intentando navegar hasta un índice que está fuera del intervalo.|
|8000|Parámetro ausente|No pudimos dar formato a la celda de la tabla porque faltan algunos valores de parámetro. Compruebe de nuevo los parámetros y vuelva a intentarlo.|Faltan algunos parámetros del método cellFormat. Por ejemplo, faltan parámetros de tableOptions, celdas o formato.|
|8010|Valor no válido|Uno o varios parámetros de celdas tienen valores que no están permitidos. Compruebe de nuevo los valores y vuelva a intentarlo.|No se ha definido la enumeración de referencia de celdas comunes. Por ejemplo, Todas, Datos o Encabezados.|
|8011|Valor no válido|Uno o varios parámetros tableOptions tienen valores que no están permitidos. Compruebe de nuevo los valores y vuelva a intentarlo.|Uno de los valores en tableOptions no es válido.|
|8012|Valor no válido|Uno o varios parámetros de formato tienen valores que no están permitidos. Compruebe de nuevo los valores y vuelva a intentarlo.|Uno de los valores en el formato no es válido.|
|8020|Fuera del intervalo|El valor de índice de la fila está fuera del intervalo permitido. Utilice un valor (0 o superior) que sea menor que el número de filas.|El índice de fila es superior al mayor índice de fila de la tabla o menor que 0.|
|8021|Fuera del intervalo|El valor de índice de la columna está fuera del intervalo permitido. Utilice un valor (0 o superior) que sea menor que el número de columnas.|El índice de la columna es mayor que el mayor índice de columna de la tabla o menor que 0.|
|8022|Fuera del intervalo|El valor está fuera del intervalo permitido.|Algunos de los valores en el formato están fuera de los intervalos admitidos.|
|9016|Permiso denegado|Permiso denegado|Acceso denegado.|
|12002|||Uno de los siguientes:<br> - No existe ninguna p?gina en la direcci?n URL pasada a `displayDialogAsync`.<br> - La p?gina pasada a `displayDialogAsync` se ha cargado, pero el cuadro de di?logo se ha dirigido a una p?gina que no se puede encontrar o cargar, o que se ha dirigido a una direcci?n URL con sintaxis no v?lida. Se produce en el cuadro de diálogo y trigges una `DialogEventReceived` (evento) en la página host.|
|12003|||El cuadro de diálogo se dirige a una dirección URL con el protocolo HTTP. Se necesita HTTPS. Se produce en el cuadro de diálogo y trigges una `DialogEventReceived` (evento) en la página host.|
|12004|||El dominio de la dirección URL se pasa a `displayDialogAsync` no es de confianza. El dominio debe ser el mismo dominio que la página host (incluido el número de puerto y protocolo). Iniciada por la llamada de `displayDialogAsync`.|
|12005|||La dirección URL pasa a `displayDialogAsync` utiliza el protocolo HTTP. Se necesita HTTPS. Iniciada por la llamada de `displayDialogAsync`. (En algunas versiones de Office, el mensaje de error devuelto con 12005 es el mismo devuelve uno para 12004.)|
|12006|||Se ha cerrado el cuadro de di?logo, normalmente porque el usuario elige el bot?n **X**. Se produce en el cuadro de diálogo y trigges una `DialogEventReceived` (evento) en la página host.|
|12007|||Ya se abre un cuadro de diálogo desde esta ventana host. Una ventana de host, como un panel de tareas, sólo puede tener un cuadro de diálogo Abrir a la vez. Iniciada por la llamada de `displayDialogAsync`.|
|13000 - 13010|||Vea [las causas y control de errores de getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/troubleshoot-sso-in-office-add-ins#causes-and-handling-of-errors-from-getaccesstokenasync).|

## <a name="binding-creation-error-conditions"></a>Condiciones para no poder crear enlaces

Cuando se crea un enlace en la API, debe indicar el tipo de enlace que quiere usar. En las tablas siguientes se enumeran los tipos de enlace y los comportamientos de enlace resultantes que se esperan.

### <a name="behavior-in-excel"></a>Comportamiento en Excel

La tabla siguiente resume el comportamiento del enlace en Excel.

|**Tipo de enlace especificado**|**Selección real**|**Comportamiento**|
|:-----|:-----|:-----|
|Matriz|Intervalo de celdas (incluye dentro de una tabla y una sola celda).|Un enlace del tipo _matriz_ se crea en las celdas seleccionadas. No se espera ninguna modificación en el documento.|
|Matriz|Texto seleccionado en la celda.|Un enlace del tipo _matriz_ se crea en toda la celda. No se espera ninguna modificación en el documento.|
|Matriz|Selección múltiple o no válida (por ejemplo, el usuario selecciona una imagen, un objeto o WordArt).|No se puede crear el enlace.|
|Tabla|Intervalo de celdas (incluye una sola celda).|No se puede crear el enlace.|
|Tabla|Intervalo de celdas en una tabla (incluye una sola celda de una tabla, toda la tabla o texto de una celda en una tabla).|Se crea un enlace en toda la tabla.|
|Tabla|Media selección en una tabla y media selección fuera de la tabla.|No se puede crear el enlace.|
|Tabla|Texto seleccionado en la celda (no en la tabla).|No se puede crear el enlace.|
|Tabla|Selección múltiple o no válida (por ejemplo, el usuario selecciona una imagen, objeto, WordArt, etc.).|No se puede crear el enlace.|
|Texto|Intervalo de celdas.|No se puede crear el enlace.|
|Texto|Intervalo de celdas en una tabla.|No se puede crear el enlace.|
|Texto|Una sola celda.|Se crea un enlace de tipo  _texto_.|
|Texto|Una sola celda en una tabla.|Se crea un enlace de tipo  _texto_.|
|Texto|Texto seleccionado en la celda.|Se crea un enlace de tipo  _texto_ en toda la celda.|

### <a name="behavior-in-word"></a>Comportamiento en Word

La tabla siguiente resume el comportamiento del enlace en Word.

|**Tipo de enlace especificado**|**Selección real**|**Comportamiento**|
|:-----|:-----|:-----|
|Matriz|Texto|No se puede crear el enlace.|
|Matriz|Toda la tabla.|Se crea un enlace de tipo  _matriz_.Se cambia el documento y un control de contenido debe ajustar la tabla. |
|Matriz|Intervalo en una tabla.|No se puede crear el enlace.|
|Matriz|Selección no válida (por ejemplo, objetos no válidos, múltiples, etc.).|No se puede crear el enlace.|
|Tabla|Texto|No se puede crear el enlace.|
|Tabla|Toda la tabla.|Se crea un enlace de tipo  _texto_.|
|Tabla|Intervalo en una tabla.|No se puede crear el enlace.|
|Tabla|Selección no válida (por ejemplo, objetos no válidos, múltiples, etc.).|No se puede crear el enlace.|
|Texto|Toda la tabla|Se crea un enlace de tipo  _texto_.|
|Texto|Intervalo en una tabla.|No se puede crear el enlace.|
|Texto|Selección múltiple|La última selección se ajustará con un control de contenido y un enlace a ese control. Se crea un control de contenido de tipo  _texto_.|
|Texto|Selección no válida (por ejemplo, objetos no válidos, múltiples, etc.).|No se puede crear el enlace.|

## <a name="see-also"></a>Vea también
   
- [Ciclo de vida del desarrollo de complementos de Office](https://docs.microsoft.com/office/dev/add-ins/concepts/add-in-development-lifecycle)
    
