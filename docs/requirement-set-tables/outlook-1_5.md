| Class | Campos | Descripción |
|:---|:---|:---|
|[Buzón](/javascript/api/outlook/outlook.mailbox)|[addHandlerAsync (eventType: cadena de Office. EventType \| , handler: (tipo: Office. EventType) => void, Options?: Office. AsyncContextOptions, callback?: (asyncResult: Office. asyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#addhandlerasync-eventtype--handler--type-)|Agrega un controlador de eventos para un evento admitido.|
||[getCallbackTokenAsync (opciones: Office. AsyncContextOptions & {isRest?: Boolean}, callback: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.mailbox#getcallbacktokenasync-options--isrest--callback--asyncresult-)|Obtiene una cadena que contiene un token usado para llamar a las API de REST o a los servicios web Exchange (EWS).|
||[isRest](/javascript/api/outlook/outlook.mailbox#isrest)||
||[removeHandlerAsync (eventType: cadena de Office. EventType \| , opciones?: Office. AsyncContextOptions, callback?: (asyncResult: Office. asyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#removehandlerasync-eventtype--options--callback--asyncresult-)|Elimina el controlador de eventos de un tpo de evento admitido.|
||[restUrl](/javascript/api/outlook/outlook.mailbox#resturl)|Obtiene la URL del punto de conexión REST para esta cuenta de correo electrónico.|
