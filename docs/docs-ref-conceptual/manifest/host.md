# <a name="host-element"></a>Elemento Host

Especifica un tipo individual de aplicación de Office en el que se debe activar el complemento.

> [!IMPORTANT] 
> La sintaxis de elemento de **Host** varía en función de si el elemento se ha definido dentro del [manifiesto básica](#basic-manifest) o dentro del nodo [VersionOverrides](#versionoverrides-node) . Sin embargo, las funciones son las mismas.  

## <a name="basic-manifest"></a>Manifiesto básico

Cuando se define en el manifiesto básico (bajo [OfficeApp](officeapp.md)), el tipo de host está determinado en el atributo `Name`.   

### <a name="attributes"></a>Atributos

| Atributo     | Tipo   | Necesario | Descripción                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [Name](#name) | string | necesario | El nombre del tipo de aplicación host de Office. |

### <a name="name"></a>Nombre
Especifica el tipo de host al que se dirige este complemento. El valor debe ser uno de los siguientes:

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

### <a name="example"></a>Ejemplo
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a>Nodo VersionOverrides
Cuando se define en [VersionOverrides](versionoverrides.md), el tipo de host está determinado en el atributo `xsi:type`. 

### <a name="attributes"></a>Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Sí  | Describe el host de Office donde se aplica esta configuración.|

### <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [DesktopFormFactor](desktopformfactor.md)    |  Sí   |  Define la configuración del factor de forma de escritorio. |
|  [MobileFormFactor](mobileformfactor.md)    |  No   |  Define la configuración del factor de forma móvil. **Nota:** Este elemento solo se admite en Outlook para iOS. |
|  [AllFormFactors](allformfactors.md)    |  No   |  Define la configuración de todos los factores de forma. Solo lo usan las funciones personalizadas en Excel. |

### <a name="xsitype"></a>xsi:type

Controla a qué host de Office (Word, Excel, PowerPoint, Outlook, OneNote) se aplica la configuración contenida. El valor debe ser uno de los siguientes:

- `Document` (Word)
- `MailHost` (Outlook)    
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## <a name="host-example"></a>Ejemplo de host 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
