# <a name="extensionpoint-element"></a>Elemento ExtensionPoint

 Define dónde expone su función un complemento en la interfaz de usuario de Office. El elemento **ExtensionPoint** es un elemento secundario de [AllFormFactors](allformfactors.md), de [DesktopFormFactor](desktopformfactor.md) o de [MobileFormFactor](mobileformfactor.md). 

## <a name="attributes"></a>Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Sí  | El tipo de punto de extensión que se está definiendo.|

## <a name="extension-points-for-excel-only"></a>Puntos de extensión solo para Excel

- **CustomFunctions** - una función personalizada que se escriben en JavaScript para Excel.

[XML de este ejemplo de código](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml) se muestra cómo usar el elemento **ExtensionPoint** con el valor del atributo **CustomFunctions** y los elementos secundarios que se usará.

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a>Puntos de extensión para comandos de complemento de Word, Excel, PowerPoint y OneNote

- **PrimaryCommandSurface**: la cinta de opciones en Office.
- **ContextMenu**: el menú contextual que aparece cuando se hace clic con el botón derecho en la interfaz de usuario de Office.

Los ejemplos siguientes muestran cómo usar el elemento  **ExtensionPoint** con los valores de atributo **PrimaryCommandSurface** y **ContextMenu**, así como los elementos secundarios que hay que usar con cada uno de ellos.

> [!IMPORTANT] 
> Para los elementos que contienen un atributo ID, asegúrese de que proporcionar un identificador único. Le recomendamos que use el nombre de su compañía con su id. Por ejemplo, use el formato siguiente. <CustomTab id="mycompanyname.mygroupname">

```XML
<ExtensionPoint xsi:type="PrimaryCommandSurface">
          <CustomTab id="Contoso Tab">
          <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
            <!-- <OfficeTab id="TabData"> -->
            <Label resid="residLabel4" />
            <Group id="Group1Id12">
              <Label resid="residLabel4" />
              <Icon>
                <bt:Image size="16" resid="icon1_32x32" />
                <bt:Image size="32" resid="icon1_32x32" />
                <bt:Image size="80" resid="icon1_32x32" />
              </Icon>
              <Tooltip resid="residToolTip" />
              <Control xsi:type="Button" id="Button1Id1">

                  <!-- information about the control -->
              </Control>
              <!-- other controls, as needed -->
            </Group>
          </CustomTab>
        </ExtensionPoint>

      <ExtensionPoint xsi:type="ContextMenu">
        <OfficeMenu id="ContextMenuCell">
          <Control xsi:type="Menu" id="ContextMenu2">
                  <!-- information about the control -->
          </Control>
          <!-- other controls, as needed -->
        </OfficeMenu>
        </ExtensionPoint>
```

#### <a name="child-elements"></a>Elementos secundarios
 
|**Elemento**|**Descripción**|
|:-----|:-----|
|**CustomTab**|Es obligatorio si quiere agregar una pestaña personalizada a la cinta de opciones (con  **PrimaryCommandSurface**). Si usa el elemento  **CustomTab**, no puede usar el elemento  **OfficeTab**. El atributo  **id** es obligatorio.|
|**OfficeTab**|Es obligatorio si quiere extender una pestaña de cinta de Office predeterminada (mediante **PrimaryCommandSurface**). Si usa el elemento **OfficeTab**, no puede usar el elemento **CustomTab**. Para obtener información detallada, consulte [OfficeTab](officetab.md).|
|**OfficeMenu**|Es obligatorio si agrega comandos de complemento a un menú contextual predeterminado (con **ContextMenu**). El atributo **id** debe establecerse en: <br/> - **ContextMenuText** para Excel o Word. Muestra el elemento en el menú contextual cuando se selecciona un texto y luego el usuario hace clic en él con el botón derecho. <br/> - **ContextMenuCell** para Excel. Muestra el elemento en el menú contextual cuando el usuario hace clic con el botón derecho en una celda de la hoja de cálculo.|
|**Group**|Un grupo de puntos de extensión de interfaz de usuario en una pestaña. Un grupo puede tener hasta seis controles. El atributo  **id** es obligatorio. Es una cadena con un máximo de 125 caracteres.|
|**Label**|Obligatorio. La etiqueta del grupo. El atributo  **resid** debe establecerse en el valor del atributo **id** de un elemento **String**. El elemento  **String** es un elemento secundario del elemento **ShortStrings**, que a su vez lo es de  **Resources**.|
|**Icon**|Obligatorio. Especifica el icono del grupo que se usará en dispositivos de factor de forma pequeños o cuando se muestren demasiados botones. El atributo  **resid** debe establecerse en el valor del atributo **id** de un elemento **Image**. El elemento  **Image** es un elemento secundario de **Images**, que a su vez lo es de  **Resources**. El atributo **size** determina el tamaño de la imagen en píxeles. Se necesitan tres tamaños de imagen: 16, 32 y 80. También se admiten cinco tamaños opcionales: 20, 24, 40, 48 y 64.|
|**Tooltip**|Opcional. La información sobre herramientas del grupo. El atributo  **resid** debe establecerse en el valor del atributo **id** de un elemento **String**. El elemento  **String** es un elemento secundario del elemento **LongStrings**, que a su vez lo es de  **Resources**.|
|**Control**|Cada grupo necesita como mínimo un control. Un elemento **Control** puede ser un **botón** o un **menú**. Use **Menu** para especificar una lista desplegable de controles de botón. Actualmente, solo se admiten botones y menús. Vea las secciones [Controles de botón](control.md#button-control) y [Controles de menú (de botones desplegables)](control.md#menu-dropdown-button-controls) para obtener más información.<br/>**Nota:**  Para hacer que la solución de problemas, se recomienda que un elemento de **Control** y los elementos secundarios de **recursos** relacionados se agregará uno a la vez.|
|**Script**|Vincula al archivo JavaScript con el código de registro y la definición de función personalizada. Este elemento no se usa en la vista previa para desarrolladores. En su lugar, la página HTML es responsable de cargar todos los archivos de JavaScript.|
|**Page**|Vínculos a la página HTML para sus funciones personalizadas.|

## <a name="extension-points-for-outlook"></a>Puntos de extensión para Outlook

- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (solo puede usarse en el elemento [DesktopFormFactor](desktopformfactor.md)).
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [Events](#events)
- [DetectedEntity](#detectedentity)

### <a name="messagereadcommandsurface"></a>MessageReadCommandSurface
Este punto de extensión coloca botones en la superficie del comando de la vista de lectura de correo. En el escritorio de Outlook aparece en la cinta.

#### <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Descripción  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Agrega los comandos a la pestaña de la cinta predeterminada.  |
|  [CustomTab](customtab.md) |  Agrega los comandos a la pestaña de la cinta personalizada.  |

#### <a name="officetab-example"></a>Ejemplo de OfficeTab
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Ejemplo de CustomTab
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a>MessageComposeCommandSurface
Este punto de extensión coloca botones en la cinta para complementos mediante el formulario de redacción de correo. 

#### <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Descripción  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Agrega los comandos a la pestaña de la cinta predeterminada.  |
|  [CustomTab](customtab.md) |  Agrega los comandos a la pestaña de la cinta personalizada.  |

#### <a name="officetab-example"></a>Ejemplo de OfficeTab
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Ejemplo de CustomTab

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a>AppointmentOrganizerCommandSurface

Este punto de extensión coloca botones en la cinta para el formulario que se muestra al organizador de la reunión. 

#### <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Descripción  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Agrega los comandos a la pestaña de la cinta predeterminada.  |
|  [CustomTab](customtab.md) |  Agrega los comandos a la pestaña de la cinta personalizada.  |

#### <a name="officetab-example"></a>Ejemplo de OfficeTab
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Ejemplo de CustomTab
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a>AppointmentAttendeeCommandSurface

Este punto de extensión coloca botones en la cinta para el formulario que se muestra al asistente de la reunión. 

#### <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Descripción  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Agrega los comandos a la pestaña de la cinta predeterminada.  |
|  [CustomTab](customtab.md) |  Agrega los comandos a la pestaña de la cinta personalizada.  |

#### <a name="officetab-example"></a>Ejemplo de OfficeTab
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Ejemplo de CustomTab
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a>Módulo

Este punto de extensión coloca botones en la cinta para la extensión de módulo. 

#### <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Descripción  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Agrega los comandos a la pestaña de la cinta predeterminada.  |
|  [CustomTab](customtab.md) |  Agrega los comandos a la pestaña de la cinta personalizada.  |

### <a name="mobilemessagereadcommandsurface"></a>MobileMessageReadCommandSurface
Este punto de extensión coloca botones en la superficie del comando de la vista de lectura de correo en el factor de forma móvil.

#### <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Descripción  |
|:-----|:-----|
|  [Group](group.md) |  Agrega un grupo de botones a la superficie del comando.  |

Los elementos **ExtensionPoint** de este tipo solo pueden tener un elemento secundario: un elemento **Group**.

Los elementos **Control** que se incluyen en este punto de extensión deben tener el atributo **xsi:type** establecido en `MobileButton`.

#### <a name="example"></a>Ejemplo
```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="mobileGroupID">
    <Label resid="residAppName"/>
      <Control id="mobileButton1" xsi:type="MobileButton">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="events"></a>Eventos

Este punto de extensión agrega un controlador de eventos para un evento especificado.

> [!NOTE]
> Este tipo de elemento sólo es compatible con Outlook en la web en Office 365.

| Elemento | Descripción  |
|:-----|:-----|
|  [Event](event.md) |  Especifica el evento y la función del controlador de eventos.  |

#### <a name="itemsend-event-example"></a>Ejemplo del evento ItemSend

```xml
<ExtensionPoint xsi:type="Events"> 
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
</ExtensionPoint> 
```

### <a name="detectedentity"></a>DetectedEntity

Este punto de extensión agrega una activación de complemento contextual en un tipo de entidad especificado.

El elemento que contenga [VersionOverrides](versionoverrides.md) debe tener un valor de atributo `xsi:type` de `VersionOverridesV1_1`.

> [!NOTE]
> Este tipo de elemento sólo es compatible con Outlook en la web en Office 365.

|  Elemento |  Descripción  |
|:-----|:-----|
|  [Label](#label) |  Especifica la etiqueta del complemento en la ventana contextual.  |
|  [SourceLocation](sourcelocation.md) |  Especifica la URL para la ventana contextual.  |
|  [Rule](rule.md) |  Especifica la regla o reglas que determinan cuándo se activa un complemento.  |

#### <a name="label"></a>Label

Obligatorio. La etiqueta del grupo. El atributo  **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento **ShortStrings** del elemento [Resources](resources.md).

#### <a name="highlight-requirements"></a>Requisitos de resaltado

La única manera que tiene un usuario de activar un complemento contextual es para interactuar con una entidad resaltada. Los desarrolladores pueden controlar qué entidades están resaltadas con el atributo `Highlight` del elemento `Rule` para los tipos de regla `ItemHasKnownEntity` y `ItemHasRegularExpressionMatch`.

En cambio, hay algunas limitaciones que tener en cuenta. Estas limitaciones se usan para garantizar que siempre habrá una entidad resaltada en citas o mensajes aplicables para proporcionar al usuario una manera de activar el complemento.

- Los tipos de entidad `EmailAddress` y `Url` no pueden resaltarse, y por lo tanto, no pueden usarse para activar un complemento.
- Si se usa una única regla, `Highlight` DEBE establecerse en `all`.
- Si se usa un tipo de regla `RuleCollection` con `Mode="AND"` para combinar varias reglas, al menos una de las reglas DEBE tener `Highlight` establecido en `all`.
- Si se usa un tipo de regla `RuleCollection` con `Mode="OR"` para combinar varias reglas, todas las reglas DEBEN tener `Highlight` establecido en `all`.

#### <a name="detectedentity-event-example"></a>Ejemplo de evento DetectedEntity

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint> 
```