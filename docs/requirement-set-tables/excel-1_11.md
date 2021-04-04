| Clase | Campos | Descripción |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Proporciona información basada en la configuración de referencia cultural del sistema actual.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Obtiene la cadena usada como separador decimal para valores numéricos.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Obtiene la cadena usada para separar grupos de dígitos a la izquierda del decimal para los valores numéricos.|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Especifica si los separadores del sistema de Excel están habilitados.|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Obtiene las entidades (por ejemplo, personas) que se mencionan en los comentarios.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Obtiene el contenido de comentario enriquecido (por ejemplo, menciones en comentarios).|
||[resuelto](/javascript/api/excel/excel.comment#resolved)|Estado del hilo de comentario.|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Actualiza el contenido del comentario con una cadena con formato especial y una lista de menciones.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Crea un nuevo comentario con el contenido específico de la celda.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|La dirección de correo electrónico de la entidad que se menciona en un comentario.|
||[id](/javascript/api/excel/excel.commentmention#id)|El identificador de la entidad.|
||[name](/javascript/api/excel/excel.commentmention#name)|Nombre de la entidad que se menciona en un comentario.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|Las entidades (por ejemplo, personas) que se mencionan en los comentarios.|
||[resuelto](/javascript/api/excel/excel.commentreply#resolved)|Estado de respuesta del comentario.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|Contenido de comentario enriquecido (por ejemplo, menciones en comentarios).|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Actualiza el contenido del comentario con una cadena con formato especial y una lista de menciones.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Crea una respuesta de comentario para un comentario.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|Una matriz que contiene todas las entidades (por ejemplo, personas) mencionadas en el comentario.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)|Especifica el contenido enriquecido del comentario (por ejemplo, contenido de comentario con menciones, la primera entidad mencionada tiene un atributo ID de 0 y la segunda entidad mencionada tiene un atributo ID de 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Obtiene el nombre de la referencia cultural con el formato languagecode2-country/regioncode2 (por ejemplo, "zh-cn" o "en-us").|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|Define el formato culturalmente adecuado para mostrar números.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Obtiene la cadena usada como separador decimal para valores numéricos.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Obtiene la cadena usada para separar grupos de dígitos a la izquierda del decimal para los valores numéricos.|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Range \| string)](/javascript/api/excel/excel.range#moveto-destinationrange-)|Mueve los valores de celda, el formato y las fórmulas del rango actual al intervalo de destino, reemplazando la información antigua de dichas celdas.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Ajusta la sangría del formato del intervalo.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Cierra el libro actual.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Guarda el libro actual.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Se produce cuando el estado oculto de una o más filas ha cambiado en una hoja de cálculo específica.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|Dirección del intervalo que completó el cálculo.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Se produce cuando el estado oculto de una o más filas ha cambiado en una hoja de cálculo específica.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtiene la dirección del intervalo que representa el área que ha cambiado en una hoja de cálculo específica.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Obtiene el tipo de cambio que representa cómo se desencadenó el evento.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtiene el origen del evento.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtiene el tipo del evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Obtiene el identificador de la hoja de cálculo en la que cambiaron los datos.|
