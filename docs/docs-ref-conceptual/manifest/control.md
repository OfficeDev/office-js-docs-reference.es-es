# <a name="control-element"></a>Elemento Control

Define una función de JavaScript que ejecuta una acción o inicia un panel de tareas. Un elemento **Control** puede ser un botón o una opción de menú. Al menos un **Control** debe estar incluido en un elemento [Group](group.md).

## <a name="attributes"></a>Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|**xsi:type**|Sí|El tipo de control que se está definiendo. Puede ser `Button`, `Menu` o `MobileButton`. |
|**id**|No|El identificador del elemento de control. Puede tener 125 caracteres como máximo.|

> [!NOTE]
> El `MobileButton` el valor de **xsi: Type** se define en el esquema de VersionOverrides 1.1. Sólo se aplica a los elementos de **Control de** contenido dentro de un elemento [MobileFormFactor](mobileformfactor.md) .

## <a name="button-control"></a>Control de botón

Un botón realiza una sola acción cuando el usuario lo selecciona. Puede ejecutar una función o mostrar un panel de tareas. Cada control de botón ha que tener un `id` único para el manifiesto. 

### <a name="child-elements"></a>Elementos secundarios
|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  **Label**     | Sí |  El texto del botón. El atributo **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento **ShortStrings** del elemento [Resources](resources.md).        |
|  **ToolTip**  |No|La información sobre herramientas del botón. El atributo **resid** tiene que establecerse en el valor del atributo **id** de un elemento **String**. El elemento **String** es un elemento secundario del elemento **LongStrings**, que es un elemento secundario del elemento [Resources](resources.md).|        
|  [Supertip](supertip.md)  | Sí |  La sugerencia del botón.    |
|  [Icon](icon.md)      | Sí |  Una imagen del botón.         |
|  [Action](action.md)    | Sí |  Especifica la acción que se realizará.  |

### <a name="executefunction-button-example"></a>Ejemplo de botón ExecuteFunction

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-button-example"></a>Ejemplo de botón ShowTaskpane

```xml
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```

## <a name="menu-dropdown-button-controls"></a>Controles de menú (botón desplegable)

Un menú define una lista estática de opciones. Cada elemento de menú ejecuta una función o muestra un panel de tareas. No se admiten submenús. 

Cuando se usa con un [punto de extensión](extensionpoint.md) **PrimaryCommandSurface** o **ContextMenu**, el control de menú define:

- Un elemento de menú de nivel de raíz.

- Una lista de elementos de submenú.

Cuando se usa con  **PrimaryCommandSurface**, el elemento de menú raíz se muestra como un botón en la cinta de opciones. Cuando el botón está seleccionado, el submenú se muestra como una lista desplegable. Cuando se usa con  **ContextMenu**, se inserta un elemento de menú con un submenú en el menú contextual. En ambos casos, los elementos individuales del submenú pueden ejecutar una función de JavaScript o mostrar un panel de tareas. Por ahora solo se admite un nivel de submenús.

En el ejemplo siguiente se muestra cómo definir un elemento de menú con dos elementos de submenú. El primer elemento de submenú muestra un panel de tareas y el segundo elemento de submenú ejecuta una función JavaScript.

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

### <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  **Label**     | Sí |  El texto del botón. El atributo **resid** debe establecerse en el valor del atributo **id** de un elemento **String** en el elemento **ShortStrings** del elemento [Resources](resources.md).      |
|  **ToolTip**  |No|La información sobre herramientas del botón. El atributo **resid** tiene que establecerse en el valor del atributo **id** de un elemento **String**. El elemento **String** es un elemento secundario del elemento **LongStrings**, que es un elemento secundario del elemento [Resources](resources.md).|        
|  [Supertip](supertip.md)  | Sí |  La sugerencia de este botón.    |
|  [Icon](icon.md)      | Sí |  Una imagen del botón.         |
|  **Items**     | Sí |  Una colección de botones que se mostrarán en el menú Contiene los elementos **Item** para cada elemento de submenú. Cada elemento **Item** contiene los elementos secundarios del [control Button](#button-control).|

### <a name="menu-control-examples"></a>Ejemplos de control de menú

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

```xml
<Control xsi:type="Menu" id="msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="msgReadMenuItem1">
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

## <a name="mobilebutton-control"></a>Control MobileButton

Un botón móvil realiza una sola acción cuando el usuario lo selecciona. Puede ejecutar una función o mostrar un panel de tareas. Cada control de botón móvil debe tener un `id` único para el manifiesto.

El valor `MobileButton` de **xsi:type** se define en el esquema VersionOverrides 1.1. El elemento que contenga [VersionOverrides](versionoverrides.md) debe tener un valor de atributo `xsi:type` de `VersionOverridesV1_1`.

### <a name="child-elements"></a>Elementos secundarios
|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  **Label**     | Sí |  El texto del botón. El atributo **resid** debe estar establecido en el valor del atributo **id** de un elemento **String** en el elemento **ShortStrings** del elemento [Resources](resources.md).        |
|  [Icon](icon.md)      | Sí |  Una imagen del botón.         |
|  [Action](action.md)    | Sí |  Especifica la acción que se realizará.  |

### <a name="executefunction-mobile-button-example"></a>Ejemplo de botón móvil ExecuteFunction

```xml
<Control xsi:type="MobileButton" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-mobile-button-example"></a>Ejemplo de botón móvil ShowTaskpane

```xml
<Control xsi:type="MobileButton" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```