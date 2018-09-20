# <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos comunes de la API de Office

Los conjuntos de requisitos son grupos de miembros de la API con nombre. Complementos de Office utilizan conjuntos de requisitos especificados en el manifiesto o una comprobación en tiempo de ejecución para determinar si un host de Office admite las API que un complemento necesita. Para obtener más información, vea [conjuntos de versiones de Office y los requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

¿Necesita información sobre dónde se admiten complementos con un host de Office? Vea [Office Add-in host y plataforma de disponibilidad](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability).

¿Busca los conjuntos de requisitos de la APi *específica del host*? Vea los siguientes conjuntos de requisitos de API:
 
- [Conjuntos de requisitos de la API de JavaScript de Excel](excel-api-requirement-sets.md) (ExcelApi)
- [Conjuntos de requisitos de la API de JavaScript de Word](word-api-requirement-sets.md) (WordApi)
- [Conjuntos de requisitos de la API de JavaScript de OneNote](onenote-api-requirement-sets.md) (OneNoteApi)
- [Entender los conjuntos de requisitos de la API de Outlook](outlook-api-requirement-sets.md) (MailBox)

> [!IMPORTANT]
> Ya no se recomienda crear y usar aplicaciones web de Access y las bases de datos de SharePoint. Como alternativa, se recomienda que use [Microsoft PowerApps](https://powerapps.microsoft.com/) para crear soluciones de negocio sin código para web y dispositivos móviles.

## <a name="common-api-requirement-sets"></a>Conjuntos de requisitos comunes de la API

En la siguiente tabla se enumeran los conjuntos de requisitos comunes de la API, los métodos de cada conjunto, y las aplicaciones host de Office compatibles con el conjunto de requisitos. Todos estos conjuntos de requisitos de la API son de la versión 1.1.

|**Conjunto de requisitos**|**Host de Office**|**Métodos del conjunto**|
|:-----|:-----|:-----|
| ActiveView | PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac|Document.getActiveViewAsync|
| AddInCommands | Vea [conjuntos de requisitos de comando Add-in](add-in-commands-requirement-sets.md). | |
| BindingEvents  | Access Web Apps<br>Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Admite salida al formato Office Open XML (OOXML) como una matriz de bytes<br>(Office.FileType.Compressed) cuando se usa el método Document.getFileAsync.|
| CustomXmlParts    | Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| Dialog | Vea [conjuntos de API de cuadro de diálogo requisitos](dialog-api-requirement-sets.md). | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |
| DocumentEvents    | Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| Archivo  | Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | OneNote Online<br>Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Admite la coerción a HTML (Office.CoercionType.Html) al leer y escribir datos mediante los métodos Document.getSelectedDataAsync,<br>Document.setSelectedDataAsync, Binding.getDataAsync o Binding.setDataAsync.|
| IdentityAPI | Vea [conjuntos de requisitos de API de identidad](identity-api-requirement-sets.md). | Auth.getAccessTokenAsync |
| ImageCoercion | Excel<br>Excel para iPad<br>Excel para Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Admite la conversión a una imagen (Office.CoercionType.Image) al escribir datos mediante el método Document.setSelectedDataAsync.|
| Buzón   |Outlook para Windows<br>Outlook para web<br>Outlook para Android<br>Outlook para Mac<br>Outlook Web App |Consulte [Información sobre los conjuntos de requisitos de la API de Outlook](outlook-api-requirement-sets.md).|
| MatrixBindings    | Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word<br>Word Online<br>Word para iPad<br>Word para Mac|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Admite la coerción a la estructura de datos de "matrix" (matriz de matrices) (Office.CoercionType.Matrix) cuando se leen y escriben datos con los métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync o Binding.setDataAsync.|
| OoxmlCoercion | Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Admite la coerción al formato Open Office XML (OOXML) (Office.CoercionType.Ooxml) cuando se leen y escriben datos con los métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync o Binding.setDataAsync.|
| PartialTableBindings  | Access Web Apps||
| PdfFile   | Excel para Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Admite salida a formato PDF (Office.FileType.Pdf)<br>al usar el método Document.getFileAsync.|
| Selección | Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Project<br>Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Configuración  | Access Web Apps<br>Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | Access Web Apps<br>Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | Access Web Apps<br>Excel<br>Excel Online<br>Excel para iPad<br>Excel para Mac<br>Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Admite la coerción a la estructura de datos de "table" (Office.CoercionType.Table) cuando se leen y escriben datos con los métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync o Binding.setDataAsync.|
| TextBindings  | Excel<br>Excel Online<br>Excel para iPad<br>Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | Excel<br>Excel Online<br>Excel para iPad<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint para iPad<br>PowerPoint para Mac<br>Project<br>Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Admite la coerción al formato de texto (Office.CoercionType.Text) cuando se leen y escriben datos con los métodos Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync o Binding.setDataAsync.|
| TextFile  | Word 2013 y posterior<br>2016 de Word para Mac y versiones posteriores<br>Word Online<br>Word para iPad|Admite salida en formato de texto (Office.FileType.Text) cuando se usa el método Document.getFileAsync.|

## <a name="methods-that-arent-part-of-a-requirement-set"></a>Métodos no incluidos en un conjunto de requisitos

Los siguientes métodos de la API de JavaScript para Office no forman parte de un conjunto de requisitos. Si el complemento requiere cualquiera de estos métodos, utilice los elementos de **métodos** y el **método** en manifiesto del complemento para declarar que son necesarios, o realizar la verificación de tiempo de ejecución mediante un `if` instrucción. Para obtener más información, consulte [Especificar hosts de Office y requisitos de API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).

|**Nombre del método**|**Compatibilidad con host de Office**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Aplicaciones web de Access, Excel, Excel Online y Excel para iPad|
|Document.getFilePropertiesAsync|Excel, Excel Online, Excel para iPad, Excel para Mac, PowerPoint, PowerPoint en línea, PowerPoint para iPad, PowerPoint para Mac, Word, Word Online, Word para iPad y Word para Mac|
|Document.getProjectFieldAsync|Project Standard 2013 y Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013 y Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 y Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 y Project Professional 2013|
|Document.getSelectedViewAsync|Project Standard 2013 y Project Professional 2013|
|Document.getTaskAsync|Project Standard 2013 y Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013 y Project Professional 2013|
|Document.goToByIdAsync|Excel, Excel Online, Excel para iPad, Excel para Mac, PowerPoint, PowerPoint en línea, PowerPoint para iPad, PowerPoint para Mac, Word, Word Online, Word para iPad y Word para Mac|
|Settings.addHandlerAsync|Aplicaciones web de Access, Excel, en línea de Excel, PowerPoint, en línea de PowerPoint, Word y Word Online|
|Settings.refreshAsync|Aplicaciones web de Access, Excel, en línea de Excel, PowerPoint, en línea de PowerPoint, Word y Word Online|
|Settings.removeHandlerAsync|Aplicaciones web de Access, Excel, en línea de Excel, PowerPoint, en línea de PowerPoint, Word y Word Online|
|TableBinding.clearFormatsAsync|Excel para Mac, Excel y Excel en línea|
|TableBinding.setFormatsAsync|Excel, Excel Online, Excel para iPad y Excel para Mac|
|TableBinding.setTableOptionsAsync|Excel, Excel Online, Excel para iPad y Excel para Mac|

## <a name="see-also"></a>Vea también

- [Versiones de Office y conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar los hosts de Office y los requisitos de la API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifiesto XML de complementos para Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
