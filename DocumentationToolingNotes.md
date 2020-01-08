# <a name="how-the-office-javascript-api-documentation-is-generated"></a>Cómo se genera la documentación de la API de JavaScript para Office

Las páginas de documentación de referencia de JavaScript de Office se generan a partir de archivos de definición de tipo y fragmentos de código de ejemplo. Este proceso utiliza una combinación de herramientas de código abierto y scripts específicos del repositorio. Este documento tiene como objetivo hacer que los procesos de este repositorio sean transparentes, de modo que la comunidad pueda beneficiarse mejor de este contenido y contribuir a él.

## <a name="content-sources"></a>Orígenes de contenido

Se combinan dos tipos de contenido para crear la documentación de referencia de Office-JS: definiciones de tipo y fragmentos de código. Esto garantiza una cobertura completa de la API y proporciona ejemplos pequeños de código en línea.

### <a name="type-definition-files"></a>Archivos de definición de tipo

Los archivos de definición de tipo en [tipos indefinidos](https://github.com/DefinitelyTyped/DefinitelyTyped) son el único origen de verdad para la documentación. Cualquier complemento de Office que usa TypeScript compila con estos archivos de definición de tipo. También ofrece capacidades de IntelliSense de JavaScript y TypeScript para desarrolladores. Al crear la documentación de referencia a partir de estas definiciones, proporcionamos información más precisa.

Hay cuatro archivos d. ts relevantes que proporcionan contenido de origen para diferentes subsecciones de los documentos.

- [Office-js/index. d. ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts) (las definiciones de la versión).
  - [Excel (versión)](https://docs.microsoft.com/javascript/api/excel_release)
  - [OneNote](https://docs.microsoft.com/javascript/api/onenote)
  - [PowerPoint](https://docs.microsoft.com/javascript/api/powerpoint)
  - [Visio](https://docs.microsoft.com/javascript/api/visio)
  - [Word (versión)](https://docs.microsoft.com/javascript/api/word_release)
  - [Subsección OfficeExtensions de la API común](https://docs.microsoft.com/javascript/api/office)
- [Office-js-Preview/index. d. ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) (las definiciones de vista previa).
  - [Excel (versión preliminar)](https://docs.microsoft.com/javascript/api/excel)
  - [Outlook (versión preliminar)](https://docs.microsoft.com/javascript/api/outlook)
  - [Word (vista previa)](https://docs.microsoft.com/javascript/api/word)
  - [API común](https://docs.microsoft.com/javascript/api/office)
- [Custom-Functions-Runtime](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts)(las definiciones de tiempo de ejecución de las funciones personalizadas de Excel).
  - [Funciones personalizadas](https://docs.microsoft.com/javascript/api/custom-functions-runtime)
- [Office-Runtime](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts)(las definiciones de tiempo de ejecución de Office para la plataforma de funciones personalizadas).
  - [Tiempo de ejecución de Office](https://docs.microsoft.com/javascript/api/office-runtime)

Las versiones anteriores de las API tienen sus propios archivos d. TS. Se conservan cuando se publica un nuevo conjunto de requisitos de la API. También pueden generarse con la [herramienta de descambio de versión](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts). Estos archivos d. ts antiguos se mantienen para que, en las API de evento se revisen o modifiquen, el comportamiento original sigue estando documentado. Esto es útil si tiene que dirigirse a una versión anterior de la API.

### <a name="code-snippets"></a>Fragmentos de código

Se agregan fragmentos de código de ejemplo a las páginas de referencia de dos orígenes:

- [Ejemplos del laboratorio de scripts](https://github.com/OfficeDev/office-js-snippets)
- [Fragmentos de código locales](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/code-snippets)

Los fragmentos de código locales se encuentran en archivos YAML específicos del host. El contenido se organiza por clase y campo, por lo que se puede asignar al espacio adecuado en una página de referencia. El idioma del fragmento de código (JavaScript o TypeScript) se deduce mediante el uso de instrucciones Await.

Los fragmentos de código del laboratorio de scripts se extraen de las muestras de trabajo. Actualmente, los ejemplos de Excel y Word se asignan para hacer referencia a secciones de documento a través [de un par de archivos de asignación](https://github.com/OfficeDev/office-js-snippets/tree/master/snippet-extractor-metadata). Estos métodos de ejemplo individuales coinciden con las propiedades o los métodos de la API. Cuando se `yarn start` ejecuta el repositorio Office-js-Snippets, se crea [un archivo YAML](https://github.com/OfficeDev/office-js-snippets/blob/master/snippet-extractor-output/snippets.yaml) que contiene todos los fragmentos de código asignados. Este archivo YAML es la entrada en la herramienta de documentación de referencia.

## <a name="tooling-pipeline"></a>Canalización de herramientas

![Una imagen que muestra el flujo de control de tipo definido de forma indefinida, al preprocesador, extractor de API, procesador, documentador de API y a a través del postprocesador.](ToolingPipeline.png)

Entre los orígenes de contenido y las páginas finales, el contenido de la documentación pasa por cinco pasos de herramientas:

1. [Script de preprocesador](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/preprocessor.ts)
1. [Extractor de API](https://api-extractor.com/)
1. [Script de procesador](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/midprocessor.ts)
1. [Documentador de la API](https://github.com/microsoft/rushstack/blob/master/apps/api-documenter/README.md)
1. [Script de postprocesador](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/postprocessor.ts)

El preprocesador toma los archivos d. TS y los divide en secciones específicas de host. Lleva a cabo las operaciones de limpieza necesarias para las herramientas siguientes para procesar los datos correctamente.

El extractor de API convierte los archivos d. TS en datos JSON. Esto tokenizes todos los datos de tipo, lo que permite un análisis más sencillo.

El procesador es el que recupera los fragmentos de código y los empareja con los hosts adecuados.

El Documentador de la API convierte los datos JSON en archivos. yml. Los archivos. yml se convierten en Markdown por el sistema de publicación abierto que publica nuestros documentos en docs.microsoft.com. El Documentador de la API también contiene una extensión específica de Office que inserta los fragmentos de código.

El postprocesador limpia la tabla de contenido y mueve los archivos. yml a la carpeta de [publicación](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen).

Estos cinco pasos se llevan a cabo cuando se ejecuta [GenerateDocs. cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd) . Ese script también controla la instalación del módulo de nodo, limpia los conjuntos de archivos antiguos y los archivos de definición de tipo de versiones para cada conjunto de requisitos.
