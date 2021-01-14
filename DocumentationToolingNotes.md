# <a name="how-the-office-javascript-api-documentation-is-generated"></a>Cómo se genera la documentación de la API de JavaScript de Office

Las páginas de documentación de referencia de JavaScript de Office se generan a partir de archivos de definición de tipo y fragmentos de código de ejemplo. Este proceso usa una combinación de herramientas de código abierto y scripts específicos del repositorio. El objetivo de este documento es hacer transparentes los procesos de este repositorio para que la comunidad pueda beneficiarse mejor de este contenido y contribuir a este contenido.

## <a name="content-sources"></a>Orígenes de contenido

Se combinan dos tipos de contenido para crear la documentación de referencia de Office-JS: definiciones de tipo y fragmentos de código. Estos garantizan la cobertura completa de la API y dan ejemplos de código pequeños y en línea.

### <a name="type-definition-files"></a>Escribir archivos de definición

Los archivos de definición de tipo [que se escriben](https://github.com/DefinitelyTyped/DefinitelyTyped) definitivamente son la única fuente de verdad de la documentación. Cualquier complemento de Office que use TypeScript compila mediante estos archivos de definición de tipo. También ofrece a los desarrolladores de JavaScript y TypeScript IntelliSense funcionalidades. Al crear la documentación de referencia a partir de estas definiciones, proporcionamos información más precisa.

Hay cuatro archivos d.ts relevantes que proporcionan contenido de origen para diferentes subsecciones de los documentos.

- [office-js/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts) (las definiciones de la versión).
  - [Excel (versión)](https://docs.microsoft.com/javascript/api/excel_release)
  - [OneNote](https://docs.microsoft.com/javascript/api/onenote)
  - [PowerPoint](https://docs.microsoft.com/javascript/api/powerpoint)
  - [Visio](https://docs.microsoft.com/javascript/api/visio)
  - [Word (versión)](https://docs.microsoft.com/javascript/api/word_release)
  - [Subsección OfficeExtensions de la API común](https://docs.microsoft.com/javascript/api/office)
- [office-js-preview/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) (las definiciones de vista previa).
  - [Excel (versión preliminar)](https://docs.microsoft.com/javascript/api/excel)
  - [Outlook (versión preliminar)](https://docs.microsoft.com/javascript/api/outlook)
  - [Word (versión preliminar)](https://docs.microsoft.com/javascript/api/word)
  - [API común](https://docs.microsoft.com/javascript/api/office)
- [custom-functions-runtime/index.d.ts](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts) (definiciones de tiempo de ejecución de funciones personalizadas de Excel).
  - [Funciones personalizadas](https://docs.microsoft.com/javascript/api/custom-functions-runtime)
- [office-runtime/index.d.ts](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts) (las definiciones de tiempo de ejecución de Office para la plataforma de funciones personalizadas).
  - [Office Runtime](https://docs.microsoft.com/javascript/api/office-runtime)

Las versiones anteriores de las API tienen sus propios archivos d.ts. Se conservan cuando se lanza un nuevo conjunto de requisitos de la API. También se pueden generar con la herramienta [Version Remover](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts). Estos archivos d.ts antiguos se mantienen para que, en caso de que las API se revisión o se altere, el comportamiento original aún se documente. Esto es útil si tiene que dirigirse a una versión anterior de la API.

#### <a name="testing-type-definition-file-changes"></a>Prueba de cambios en el archivo de definición de tipo

Cualquier cambio en la documentación de la API de JavaScript de Office se realiza mediante la edición de los cuatro archivos d.ts mencionados anteriormente. Sin embargo, puede probar un cambio antes de enviar un PR a DefinitelyTyped (si necesita, por ejemplo, probar cómo se traducirá el formato a markdown) editando el archivo correspondiente en [generate-docs/script-inputs](https://github.com/OfficeDev/office-js-docs-reference/tree/master/generate-docs/script-inputs) y ejecutando [GenerateDocs.cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd). Cuando se le solicite, seleccione la opción "Archivos locales".

Insertar cambios en una rama remota de este repositorio hace que la docs.microsoft.com de prueba cree una rama de prueba. Esta rama se representa en review.docs.microsoft.com, a la que solo puede acceder el personal interno de Microsoft. Cualquier persona que revise su PR comprobará la precisión del sitio de revisión.

### <a name="code-snippets"></a>Fragmentos de código

Los fragmentos de código de ejemplo se agregan a las páginas de referencia desde dos orígenes:

- [Ejemplos de laboratorio de scripts](https://github.com/OfficeDev/office-js-snippets)
- [Fragmentos de código local](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/code-snippets)

Los fragmentos de código locales están en archivos yaml específicos del host. Su contenido se organiza por clase y campo, por lo que se puede asignar al lugar adecuado en una página de referencia. El lenguaje del fragmento de código (JavaScript o TypeScript) se deduce mediante el uso de instrucciones await.

Los fragmentos de código de Script Lab se extrayó de ejemplos de trabajo. Actualmente, los ejemplos de Excel, Outlook, PowerPoint y Word se asignan a secciones de documento de referencia a través de [archivos de asignación.](https://github.com/OfficeDev/office-js-snippets/tree/prod/snippet-extractor-metadata) Estos coinciden con métodos de ejemplo individuales con propiedades o métodos de la API. Cuando se ejecuta el repositorio office-js-snippets, se crea un archivo yaml que contiene todos los `yarn start` fragmentos de código asignados. [](https://github.com/OfficeDev/office-js-snippets/blob/prod/snippet-extractor-output/snippets.yaml) Este archivo yaml es la entrada en las herramientas de documentación de referencia.

## <a name="tooling-pipeline"></a>Canalización de herramientas

![Imagen que muestra el flujo de control de Definitivamente escrito, al preprocesador, extractor de API, midprocessor, API Documenter y hasta el postprocesador.](ToolingPipeline.png)

Entre los orígenes de contenido y las páginas finales, el contenido de la documentación pasa por cinco pasos de herramientas:

1. [Script de preprocesador](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/preprocessor.ts)
1. [API Extractor](https://api-extractor.com/)
1. [Script de midprocessor](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/midprocessor.ts)
1. [Documentador de API](https://github.com/microsoft/rushstack/blob/master/apps/api-documenter/README.md)
1. [Script de posprocesador](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/postprocessor.ts)

El preprocesador toma los archivos d.ts y los divide en secciones específicas del host. Realiza cualquier limpieza necesaria para que las herramientas posteriores procese correctamente los datos.

El Extractor de API convierte los archivos d.ts en datos JSON. Esto tokeniza todos los datos de tipo, lo que permite un análisis más sencillo.

El midprocessor recupera los fragmentos de código y los empareja con los hosts adecuados, y limpia el enlace cruzado entre los objetos de Outlook y la API común.

Api Documenter convierte los datos JSON en archivos .yml. Los archivos .yml se convierten en markdown mediante el sistema de publicación abierto que publica nuestros documentos en docs.microsoft.com. Api Documenter también contiene una extensión específica de Office que inserta nuestros fragmentos de código.

El postprocesador limpia la tabla de contenido y mueve los archivos .yml a la [carpeta de publicación.](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen)

Estos cinco pasos se realizan cuando [se ejecuta GenerateDocs.cmd.](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd) Ese script también controla la instalación de módulos de nodo, limpia los conjuntos de archivos antiguos y los archivos de definición de tipo de versiones para cada conjunto de requisitos.
