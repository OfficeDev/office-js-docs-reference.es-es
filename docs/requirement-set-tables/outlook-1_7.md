| Class | Campos | Descripción |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[addHandlerAsync (eventType: cadena de Office. EventType \| , handler: any, Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. asyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Agrega un controlador de eventos para un evento admitido.|
||[organizer](/javascript/api/outlook/outlook.appointmentcompose#organizer)|Obtiene el organizador de la reunión especificada.|
||[periodicidad](/javascript/api/outlook/outlook.appointmentcompose#recurrence)|Obtiene o establece el patrón de periodicidad de una cita.|
||[removeHandlerAsync (eventType: cadena de Office. EventType \| , opciones?: Office. AsyncContextOptions, callback?: (asyncResult: Office. asyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#removehandlerasync-eventtype--options--callback--asyncresult-)|Elimina el controlador de eventos de un tpo de evento admitido.|
||[seriesId](/javascript/api/outlook/outlook.appointmentcompose#seriesid)|Obtiene el identificador de la serie a la que pertenece una instancia.|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[addHandlerAsync (eventType: cadena de Office. EventType \| , handler: any, Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. asyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentread#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Agrega un controlador de eventos para un evento admitido.|
||[periodicidad](/javascript/api/outlook/outlook.appointmentread#recurrence)|Obtiene el patrón de periodicidad de una cita.|
||[removeHandlerAsync (eventType: cadena de Office. EventType \| , opciones?: Office. AsyncContextOptions, callback?: (asyncResult: Office. asyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentread#removehandlerasync-eventtype--options--callback--asyncresult-)|Elimina el controlador de eventos de un tpo de evento admitido.|
||[seriesId](/javascript/api/outlook/outlook.appointmentread#seriesid)|Obtiene el identificador de la serie a la que pertenece una instancia.|
|[AppointmentTimeChangedEventArgs](/javascript/api/outlook/outlook.appointmenttimechangedeventargs)|[end](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#end)||
||[start](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#start)||
||[type](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#type)||
|[From](/javascript/api/outlook/outlook.from)|[getAsync (Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <EmailAddressDetails> ) => void)](/javascript/api/outlook/outlook.from#getasync-options--callback--asyncresult-)|Obtiene el valor de from de un mensaje.|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[addHandlerAsync (eventType: cadena de Office. EventType \| , handler: any, Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. asyncResult <void> ) => void)](/javascript/api/outlook/outlook.messagecompose#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Agrega un controlador de eventos para un evento admitido.|
||[from](/javascript/api/outlook/outlook.messagecompose#from)|Obtiene la dirección de correo electrónico del remitente de un mensaje.|
||[removeHandlerAsync (eventType: cadena de Office. EventType \| , opciones?: Office. AsyncContextOptions, callback?: (asyncResult: Office. asyncResult <void> ) => void)](/javascript/api/outlook/outlook.messagecompose#removehandlerasync-eventtype--options--callback--asyncresult-)|Elimina el controlador de eventos de un tpo de evento admitido.|
||[seriesId](/javascript/api/outlook/outlook.messagecompose#seriesid)|Obtiene el identificador de la serie a la que pertenece una instancia.|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[addHandlerAsync (eventType: cadena de Office. EventType \| , handler: any, Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. asyncResult <void> ) => void)](/javascript/api/outlook/outlook.messageread#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Agrega un controlador de eventos para un evento admitido.|
||[periodicidad](/javascript/api/outlook/outlook.messageread#recurrence)|Obtiene el patrón de periodicidad de una cita.|
||[removeHandlerAsync (eventType: cadena de Office. EventType \| , opciones?: Office. AsyncContextOptions, callback?: (asyncResult: Office. asyncResult <void> ) => void)](/javascript/api/outlook/outlook.messageread#removehandlerasync-eventtype--options--callback--asyncresult-)|Elimina el controlador de eventos de un tpo de evento admitido.|
||[seriesId](/javascript/api/outlook/outlook.messageread#seriesid)|Obtiene el identificador de la serie a la que pertenece una instancia.|
|[Organizador](/javascript/api/outlook/outlook.organizer)|[getAsync (Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <EmailAddressDetails> ) => void)](/javascript/api/outlook/outlook.organizer#getasync-options--callback--asyncresult-)|Obtiene el valor de organizador de una cita como {@link Office. EmailAddressDetails | EmailAddressDetails} (objeto)|
|[RecipientsChangedEventArgs](/javascript/api/outlook/outlook.recipientschangedeventargs)|[changedRecipientFields](/javascript/api/outlook/outlook.recipientschangedeventargs#changedrecipientfields)||
||[type](/javascript/api/outlook/outlook.recipientschangedeventargs#type)||
|[RecipientsChangedFields](/javascript/api/outlook/outlook.recipientschangedfields)|[bcc](/javascript/api/outlook/outlook.recipientschangedfields#bcc)|Obtiene si los destinatarios del campo **CCO** se cambiaron.|
||[cc](/javascript/api/outlook/outlook.recipientschangedfields#cc)|Obtiene si se cambiaron los destinatarios del campo **CC** .|
||[optionalAttendees](/javascript/api/outlook/outlook.recipientschangedfields#optionalattendees)|Obtiene si se cambiaron los asistentes opcionales.|
||[requiredAttendees](/javascript/api/outlook/outlook.recipientschangedfields#requiredattendees)|Obtiene si se cambiaron los asistentes necesarios.|
||[resources](/javascript/api/outlook/outlook.recipientschangedfields#resources)|Obtiene si se cambiaron los recursos.|
||[to](/javascript/api/outlook/outlook.recipientschangedfields#to)|Obtiene si los destinatarios del campo **para** se cambiaron.|
|[Periodicidad](/javascript/api/outlook/outlook.recurrence)|[getAsync (Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <Recurrence> ) => void)](/javascript/api/outlook/outlook.recurrence#getasync-options--callback--asyncresult-)|Devuelve el objeto de periodicidad actual de una serie de citas.|
||[recurrenceProperties](/javascript/api/outlook/outlook.recurrence#recurrenceproperties)|Obtiene o establece las propiedades de la serie de citas periódicas.|
||[recurrenceTimeZone](/javascript/api/outlook/outlook.recurrence#recurrencetimezone)|Obtiene o establece las propiedades de la serie de citas periódicas.|
||[recurrenceType](/javascript/api/outlook/outlook.recurrence#recurrencetype)|Obtiene o establece el tipo de la serie de citas periódicas.|
||[seriesTime](/javascript/api/outlook/outlook.recurrence#seriestime)|{@Link Office. SeriesTime | SeriesTime} objeto permite administrar las fechas de inicio y finalización de la serie de citas periódicas y|
||[setAsync (recurrencePattern: recurrence, Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.recurrence#setasync-recurrencepattern--options--callback--asyncresult-)|Establece el patrón de periodicidad de una serie de citas.|
|[RecurrenceChangedEventArgs](/javascript/api/outlook/outlook.recurrencechangedeventargs)|[periodicidad](/javascript/api/outlook/outlook.recurrencechangedeventargs#recurrence)||
||[type](/javascript/api/outlook/outlook.recurrencechangedeventargs#type)||
|[RecurrenceProperties](/javascript/api/outlook/outlook.recurrenceproperties)|[dayOfMonth](/javascript/api/outlook/outlook.recurrenceproperties#dayofmonth)|Representa el día del mes.|
||[dayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#dayofweek)|Representa el día de la semana o el tipo del día, por ejemplo, día del fin de semana vs Weekday.|
||[próximos](/javascript/api/outlook/outlook.recurrenceproperties#days)|Representa el conjunto de días de esta periodicidad.|
||[firstDayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#firstdayofweek)|Representa el primer día de la semana elegido; de lo contrario, el valor predeterminado es el valor de la configuración del usuario actual.|
||[interval](/javascript/api/outlook/outlook.recurrenceproperties#interval)|Representa el período entre instancias de la misma serie periódica.|
||[month](/javascript/api/outlook/outlook.recurrenceproperties#month)|Representa el mes.|
||[weekNumber](/javascript/api/outlook/outlook.recurrenceproperties#weeknumber)|Representa el número de la semana del mes seleccionado, por ejemplo, ' primer ' para la primera semana del mes.|
|[RecurrenceTimeZone](/javascript/api/outlook/outlook.recurrencetimezone)|[name](/javascript/api/outlook/outlook.recurrencetimezone#name)|Representa el nombre de la zona horaria de periodicidad.|
||[desplaza](/javascript/api/outlook/outlook.recurrencetimezone#offset)|Valor entero que representa la diferencia en minutos entre la zona horaria local y la hora UTC en la fecha de inicio de la serie de reuniones.|
|[SeriesTime](/javascript/api/outlook/outlook.seriestime)|[getDuration()](/javascript/api/outlook/outlook.seriestime#getduration--)|Obtiene la duración en minutos de una instancia normal en una serie de citas periódicas.|
||[getEndDate()](/javascript/api/outlook/outlook.seriestime#getenddate--)|Obtiene la fecha de finalización de un patrón de periodicidad en el siguiente|
||[getEndTime()](/javascript/api/outlook/outlook.seriestime#getendtime--)|Obtiene la hora de finalización de una cita habitual o una instancia de convocatoria de reunión de un patrón de periodicidad en cualquier zona horaria que el usuario o|
||[getStartDate()](/javascript/api/outlook/outlook.seriestime#getstartdate--)|Obtiene la fecha de inicio de un patrón de periodicidad en el siguiente|
||[getStartTime()](/javascript/api/outlook/outlook.seriestime#getstarttime--)|Obtiene la hora de inicio de una instancia de cita normal de un patrón de periodicidad en la zona horaria que el usuario o complemento estableció el|
||[setDuration (minutos: número)](/javascript/api/outlook/outlook.seriestime#setduration-minutes-)|Establece la duración de todas las citas en un patrón de periodicidad.|
||[setEndDate (Date: String)](/javascript/api/outlook/outlook.seriestime#setenddate-date-)|Establece la fecha de finalización de una serie de citas periódicas.|
||[setEndDate (Year: Number, month: Number, Day: Number)](/javascript/api/outlook/outlook.seriestime#setenddate-year--month--day-)|Establece la fecha de finalización de una serie de citas periódicas.|
||[setStartDate (Date: String)](/javascript/api/outlook/outlook.seriestime#setstartdate-date-)|Establece la fecha de inicio de una serie de citas periódicas.|
||[setStartDate (Year: Number, month: Number, Day: Number)](/javascript/api/outlook/outlook.seriestime#setstartdate-year--month--day-)|Establece la fecha de inicio de una serie de citas periódicas.|
||[setStartTime (hours: Number, minutes: Number)](/javascript/api/outlook/outlook.seriestime#setstarttime-hours--minutes-)|Establece la hora de inicio de todas las instancias de una serie de citas periódicas en la zona horaria en la que se establece el patrón de periodicidad.|
||[setStartTime (Time: String)](/javascript/api/outlook/outlook.seriestime#setstarttime-time-)|Establece la hora de inicio de todas las instancias de una serie de citas periódicas en la zona horaria en la que se establece el patrón de periodicidad.|
