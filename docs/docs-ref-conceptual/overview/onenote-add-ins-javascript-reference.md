# <a name="onenote-javascript-api-overview"></a>Introducción a la API de JavaScript de OneNote

Válido para: OneNote Online

Los vínculos siguientes muestran los objetos de OneNote de nivel superior disponibles en la API. Cada vínculo de objeto de página contiene una descripción de las propiedades, los eventos y los métodos disponibles en el objeto. Explore los vínculos para obtener más información. 
    
- [Application](/javascript/api/onenote/onenote.application): El objeto de nivel superior que se usa para acceder a todos los objetos de OneNote a los que se puede hacer referencia globalmente, el bloc de notas activo y la sección activa.

- [Notebook](/javascript/api/onenote/onenote.notebook): un bloc de notas. Los blocs de notas contienen grupos de secciones y secciones.
    - [NotebookCollection](/javascript/api/onenote/onenote.notebookcollection): una colección de blocs de notas.

- [SectionGroup](/javascript/api/onenote/onenote.sectiongroup): un grupo de secciones. Los grupos de secciones contienen grupos de secciones y secciones.
    - [SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection): una colección de grupos de secciones.

- [Section](/javascript/api/onenote/onenote.section): una sección. Las secciones contienen páginas.
    - [SectionCollection](/javascript/api/onenote/onenote.sectioncollection): una colección de secciones.

- [Page](/javascript/api/onenote/onenote.page): una página. Las páginas contienen objetos PageContent.
    - [PageCollection](/javascript/api/onenote/onenote.pagecollection): una colección de páginas.

- [PageContent](/javascript/api/onenote/onenote.pagecontent): una región de nivel superior en una página que contiene tipos de contenido, como Outline o Image. Un objeto PageContent se puede asignar a una posición en la página.
    - [PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection): una colección de objetos PageContent que representa el contenido de una página.

- [Outline](/javascript/api/onenote/onenote.outline): un contenedor para objetos Paragraph. Un Outline es un elemento secundario directo de un objeto PageContent.

- [Image](/javascript/api/onenote/onenote.image): un objeto Image. Image puede ser un elemento secundario directo de un objeto PageContent o Paragraph.

- [Paragraph](/javascript/api/onenote/onenote.paragraph): un contenedor para el contenido visible en una página. Un objeto Paragraph es un elemento secundario directo de un Outline.
    - [ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection): Una colección de objetos Paragraph es un Outline.

- [RichText](/javascript/api/onenote/onenote.richtext): un objeto RichText.

- [Table](/javascript/api/onenote/onenote.table): un contenedor de objetos TableRow.

- [TableRow](/javascript/api/onenote/onenote.tablerow): un contenedor de objetos TableCell.
    - [TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection): una colección de objetos TableRow de una tabla.
 
- [TableCell](/javascript/api/onenote/onenote.tablecell): un contenedor para objetos Paragraph.
    - [TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection): una colección de objetos TableCell de un objeto TableRow.

## <a name="onenote-javascript-api-reference"></a>Referencia de la API de JavaScript de complementos de OneNote

Para obtener información detallada acerca de la API de JavaScript de OneNote, consulte la [documentación de referencia de la API de JavaScript de OneNote](/javascript/api/onenote).

## <a name="see-also"></a>Ver también

- [Introducción a la programación de API de JavaScript para OneNote](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [Crear el primer complemento de OneNote](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-getting-started)
- [Rubric Grader sample (Ejemplo de Rubric Grader)](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office Add-ins platform overview (Informaci?n general sobre la plataforma de complementos para Office)](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
