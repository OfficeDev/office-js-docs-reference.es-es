# <a name="javascript-api-for-office"></a>API de JavaScript para Office

La API de JavaScript para Office permite crear aplicaciones web que interactúan con los modelos de objetos de aplicaciones host de Office. La aplicación hará referencia a la biblioteca de office.js, que es un cargador de secuencia de comandos. La biblioteca de office.js carga los modelos de objetos que son aplicables a la aplicación de Office que está ejecutando el complemento. Puede usar los siguientes modelos de objetos de JavaScript:

- **API común** - API que se introdujeron con **Office 2013**. Esto se carga para **todas las aplicaciones host de Office** y conecta a la aplicación de complemento con la aplicación de cliente de Office. El modelo de objetos contiene las API que son específicas de los clientes de Office y las API que son aplicables a varias aplicaciones de host de cliente de Office. Todo este contenido está bajo **Shared API**. 

  **Outlook** también usa la sintaxis de la API comunes. Todo el contenido en el alias Office contiene objetos que puede usar para escribir secuencias de comandos que interactúan con el contenido de documentos, hojas de cálculo, presentaciones, elementos de correo y proyectos desde los complementos de Office de Office. Debe usar estas API común si el complemento dirigirá a Office 2013 y versiones posterior. Este modelo de objetos utiliza las devoluciones de llamada.

- **Las API específicas de host** - API que se introdujeron con **Office 2016**. Este modelo de objetos proporciona objetos establecimiento inflexible de tipos específicos de host que corresponden a los objetos familiares que se muestra al usar a los clientes de Office y representa el futuro de APIs de JavaScript de Office. Las API de host específicos incluyen actualmente la API de JavaScript de Word y la API de JavaScript de Excel.

## <a name="supported-host-applications"></a>Aplicaciones host compatibles

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [Shared API](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint y Project](requirement-sets/powerpoint-and-project-note.md) admiten complementos realizados con la API de JavaScript. Sin embargo, actualmente no tienen API específica del host. Interactuar con estos hosts a través de la API de Shared.

Obtenga más información acerca de los [hosts compatibles y otros requisitos](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins).

## <a name="open-api-specifications"></a>Especificaciones de la API pública

Cuando diseñemos y desarrollemos nuevas interfaces de programación de aplicaciones (API) para complementos de Office, estarán disponibles para que pueda enviar sus comentarios en la página [Especificaciones de la API abierta](openspec.md). Descubra qué nuevas características están en proceso y envíe sus comentarios acerca de nuestras especificaciones de diseño.

## <a name="see-also"></a>Vea también

- [Referencia de la API de JavaScript de Office](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)