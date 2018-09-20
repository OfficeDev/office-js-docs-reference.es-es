# <a name="officemenu-element"></a>Elemento OfficeMenu

Define una colección de controles que se agregarán al menú contextual de Office. Se aplica a los complementos de Word, Excel, PowerPoint y OneNote.

## <a name="attributes"></a>Atributos

| Atributo            | Obligatorio | Descripción                          |
|:---------------------|:--------:|:-------------------------------------|
| [xsi:type](#xsitype) | Sí      | Tipo de OfficeMenu que va a definir.|

## <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [Control](#control)    | Sí |  Una colección de uno o varios objetos Control.  |

## <a name="xsitype"></a>xsi:type

Especifica un menú integrado de la aplicación de cliente de Office en el que se va a agregar este complemento de Office.

- `ContextMenuText` - Muestra el elemento en el menú contextual cuando se selecciona un texto y el usuario abre el menú contextual (con el botón derecho) sobre el texto seleccionado. Se aplica a Word, Excel, PowerPoint y OneNote.
- `ContextMenuCell` - Muestra el elemento en el menú contextual cuando el usuario abre el menú contextual (hace clic con el botón derecho) en una celda de la hoja de cálculo. Se aplica a Excel. 

## <a name="control"></a>Control

Cada elemento **OfficeMenu** requiere por lo menos uno o varios controles de [menú](control.md#menu-dropdown-button-controls). 

## <a name="example"></a>Ejemplo

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>   
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>    
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>    
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />    
          </Action>    
        </Item>
      </Items>
    </Control>   
</OfficeMenu>
```
