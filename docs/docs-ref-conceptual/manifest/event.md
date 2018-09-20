# <a name="event-element"></a>Elemento Event

Define un controlador de eventos en un complemento.

> [!NOTE] 
> El `Event` elemento actualmente es solo compatible con Outlook en la web en Office 365.

## <a name="attributes"></a>Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [Tipo](#type-attribute)  |  Sí  | Especifica el evento que se va a controlar. |
|  [FunctionExecution](#functionexecution-attribute)  |  Sí  | Especifica el estilo de ejecución para el controlador de eventos, asincrónico o sincrónico. Actualmente solo se admiten los controladores de eventos sincrónicos. |
|  [FunctionName](#functionname-attribute)  |  Sí  | Especifica el nombre de la función para el controlador de eventos. |

### <a name="type-attribute"></a>Atributo Type

Necesario. Especifica qué evento invocará el controlador de eventos. Los posibles valores para este atributo se especifican en la tabla siguiente.

|  Tipo de evento  |  Descripción  |
|:-----|:-----|
|  `ItemSend`  |  El controlador de eventos se invocará cuando el usuario envíe un mensaje o una invitación de reunión.  |

### <a name="functionexecution-attribute"></a>Atributo FunctionExecution

Necesario. DEBE establecerse en `synchronous`.

### <a name="functionname-attribute"></a>Atributo FunctionName

Necesario. Especifica el nombre de la función del controlador de eventos. Este valor debe coincidir con un nombre de función en el [archivo de función](functionfile.md) del complemento.

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```