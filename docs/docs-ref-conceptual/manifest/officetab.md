# <a name="officetab-element"></a>Elemento OfficeTab

Define la ficha de la cinta en la que aparece el comando del complemento. Puede ser la pestaña predeterminada (**Inicio**, **Mensaje** o **Reunión**), o una pestaña personalizada definida por el complemento. Se requiere este elemento.

## <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  Grupo      | Sí |  Define un grupo de comandos. Solo se puede agregar un grupo por cada complemento a la ficha predeterminada.  |

Los siguientes valores son `id` valores de ficha válidos por host: Los valores en **negrita** son compatibles con escritorio y en línea (por ejemplo, Word 2016 o posterior para Windows y Word en línea).

### <a name="outlook"></a>Outlook

- **TabDefault**

### <a name="word"></a>Word

- **TabHome**
- **TabInsert**
- TabWordDesign
- **TabPageLayoutWord**
- TabReferences
- TabMailings
- TabReviewWord
- **TabView**
- TabDeveloper
- TabAddIns
- TabBlogPost
- TabBlogInsert
- TabPrintPreview
- TabOutlining
- TabConflicts
- TabBackgroundRemoval
- TabBroadcastPresentation

### <a name="excel"></a>Excel

- **TabHome**
- **TabInsert**
- TabPageLayoutExcel
- TabFormulas
- **TabData**
- **TabReview**
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabBackgroundRemoval 

### <a name="powerpoint"></a>PowerPoint

- **TabHome**
- **TabInsert**
- **TabDesign**
- **TabTransitions**
- **TabAnimations**
- TabSlideShow
- TabReview
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabMerge
- TabGrayscale
- TabBlackAndWhite
- TabBroadcastPresentation
- TabSlideMaster
- TabHandoutMaster
- TabNotesMaster
- TabBackgroundRemoval
- TabSlideMasterHome

### <a name="onenote"></a>OneNote

- **TabHome**
- **TabInsert**
- **TabView**
- TabDeveloper
- TabAddIns

## <a name="group"></a>Group

Un grupo de puntos de extensión de UI en una ficha. Un grupo puede tener hasta seis controles. El atributo **id** es obligatorio y cada **id** debe ser único en el manifiesto. El **id** es una cadena de 125 caracteres como máximo. Consulte el [elemento Group](group.md).

## <a name="officetab-example"></a>Ejemplo de OfficeTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
