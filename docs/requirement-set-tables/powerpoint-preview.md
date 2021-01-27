| Clase | Campos | Descripción |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutid)|Especifica el identificador de un diseño de diapositiva que se usará para la nueva diapositiva.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slidemasterid)|Especifica el identificador de un patrón de diapositivas que se usará para la nueva diapositiva.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slidemasters)|Devuelve la colección de `SlideMaster` objetos que están en la presentación.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[id](/javascript/api/powerpoint/powerpoint.shape#id)|Obtiene el identificador único de la forma.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getcount--)|Obtiene el número de formas de la colección.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitem-key-)|Obtiene una forma mediante su identificador único.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemat-index-)|Obtiene una forma mediante su índice de base cero en la colección.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemornullobject-id-)|Obtiene una forma mediante su identificador único.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[layout](/javascript/api/powerpoint/powerpoint.slide#layout)|Obtiene el diseño de la diapositiva.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Devuelve una colección de formas de la diapositiva.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slidemaster)|Obtiene el `SlideMaster` objeto que representa el contenido predeterminado de la diapositiva.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add-options-)|Agrega una nueva diapositiva al final de la colección.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Obtiene el identificador único del diseño de diapositiva.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Obtiene el nombre del diseño de diapositiva.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getcount--)|Obtiene el número de diseños de la colección.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitem-key-)|Obtiene un diseño mediante su identificador único.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemat-index-)|Obtiene un diseño mediante su índice de base cero en la colección.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemornullobject-id-)|Obtiene un diseño mediante su identificador único.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Obtiene el identificador único del patrón de diapositivas.|
||[layouts](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Obtiene la colección de diseños proporcionados por el Patrón de diapositivas para diapositivas.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Obtiene el nombre único del patrón de diapositivas.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getcount--)|Obtiene el número de patrones de diapositivas de la colección.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitem-key-)|Obtiene un patrón de diapositivas mediante su identificador único.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemat-index-)|Obtiene un patrón de diapositivas mediante su índice de base cero en la colección.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemornullobject-id-)|Obtiene un patrón de diapositivas mediante su identificador único.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Obtiene los elementos secundarios cargados en esta colección.|
