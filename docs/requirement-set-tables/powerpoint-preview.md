| Class | Campos | Descripción |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[formato](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Especifica el formato que se va a usar durante la inserción de diapositivas.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|Especifica las diapositivas de la presentación de origen que se insertarán en la presentación actual.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|Especifica en qué lugar de la presentación se insertarán las nuevas diapositivas.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64 (base64File: String, Options?: PowerPoint. InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|Inserta las diapositivas especificadas de una presentación en la presentación actual.|
||[quedar](/javascript/api/powerpoint/powerpoint.presentation#slides)|Devuelve una colección ordenada de diapositivas de la presentación.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|Elimina la diapositiva de la presentación.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Obtiene el identificador único de la diapositiva.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|Obtiene el número de diapositivas de la colección.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|Obtiene una diapositiva mediante su identificador único.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|Obtiene una diapositiva mediante su índice de base cero en la colección.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|Obtiene una diapositiva mediante su identificador único.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
