# <a name="resources-element"></a>Elemento Resources

Contiene iconos, cadenas y las direcciones URL para el nodo [VersionOverrides](versionoverrides.md). Un elemento del manifiesto especifica un recurso por medio del **identificador** del recurso. Así se contribuye a mantener el tamaño del manifiesto manejable, en especial cuando los recursos tienen diferentes versiones para diferentes configuraciones regionales. El **identificador** debe ser único dentro del manifiesto y puede tener un máximo de 32 caracteres.

Cada recurso puede tener uno o varios elementos secundarios **Override** para definir un recurso diferente en una configuración regional determinada.

## <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Tipo  |  Descripción  |
|:-----|:-----|:-----|
|  [Imágenes](#images)            |  image   |  Proporciona la dirección URL HTTPS a una imagen para un icono. |
|  **Urls**                |  url     |  Proporciona una ubicación URL HTTPS. La dirección URL puede tener 2048 caracteres como máximo. |
|  **ShortStrings** |  string  |  Texto de los elementos **Label** y **Title**. Cada **cadena** contiene un máximo de 125 caracteres. |
|  **LongStrings**  |  string  | Texto de los atributos **Description**. Cada **cadena** contiene un máximo de 250 caracteres.|

> [!NOTE]
> Debe usar capa de Sockets seguros (SSL) para todas las direcciones URL en los elementos de **imagen** y **dirección Url** .

### <a name="images"></a>Imágenes
Cada icono debe tener tres elementos **Image**, uno por cada uno de los tres tamaños obligatorios:

- 16x16
- 32x32
- 80x80

También se admiten los siguientes tamaños adicionales, pero no son obligatorios:

- 20x20
- 24x24
- 40x40
- 48x48
- 64x64

> [!IMPORTANT] 
> Outlook requiere la capacidad de recursos de imagen de la memoria caché para aumentar el rendimiento. Por este motivo, el servidor donde se hospede un recurso de imagen no tiene que agregar directivas Cache-Control al encabezado de respuesta. Si lo hace, Outlook lo sustituirá automáticamente por una imagen genérica o una imagen predeterminada.    

## <a name="resources-examples"></a>Ejemplos de recursos 

```XML
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp32-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp80-icon_default.png" />
        </bt:Image>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
        </bt:Url>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residLabel" DefaultValue="GetData">
          <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residToolTip" DefaultValue="Get data for your document.">
          <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
        </bt:String>
      </bt:LongStrings>
    </Resources>
```

```xml
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER//blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/blue-80.png"/>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
    <!-- other URLs -->
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="groupLabel" DefaultValue="Add-in Demo">
      <bt:Override Locale="ar-sa" Value="<Localized text>" />
    </bt:String>
    <!-- Other short strings -->
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment.">
      <bt:Override Locale="ar-sa" Value="<Localized text>." />
    </bt:String>
    <!-- Other long strings -->
  </bt:LongStrings>
</Resources>
```
