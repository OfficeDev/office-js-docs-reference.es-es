# <a name="action-element"></a>Elemento Action

Especifica la acción que se realiza cuando el usuario selecciona los controles de [Botón](control.md#button-control) o [Menú](control.md#menu-dropdown-button-controls).
 
## <a name="attributes"></a>Atributos

|  Atributo  |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Sí  | Tipo de acción que se va a realizar|

## <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Descripción  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Especifica el nombre de la función que se va a ejecutar. |
|  [SourceLocation](#sourcelocation) |    Especifica la ubicación del archivo de origen para esta acción. |
|  [TaskpaneId](#taskpaneid) | Especifica el id. del contenedor del panel de tareas.|
|  [Title](#title) | Especifica el título personalizado del panel de tareas.|
|  [SupportsPinning](#supportspinning) | Especifica que un panel de tareas admite el anclado, lo que provoca que el panel de tareas siga abierto aunque el usuario cambie la selección.|
  

## <a name="xsitype"></a>xsi:type

Este atributo especifica el tipo de acción que se realiza cuando el usuario selecciona el botón. Puede ser uno de las siguientes:

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a>FunctionName

Elemento obligatorio cuando **xsi:type** es "ExecuteFunction". Especifica el nombre de la función que se va a ejecutar. La función está incluida en el archivo especificado en el elemento [FunctionFile](functionfile.md).

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation

Elemento obligatorio cuando **xsi:type** es "ShowTaskpane". Especifica la ubicación del archivo de origen para esta acción. El atributo **resid** debe establecerse en el valor del atributo **id** de un elemento **Url** en el elemento **Urls** del elemento [Resources](resources.md).

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId

Elemento opcional cuando **xsi: Type** es "ShowTaskpane". Especifica el identificador del contenedor de panel de tareas. Cuando haya varias acciones de "ShowTaskpane", utilice una **TaskpaneId** diferente si desea un panel independiente para cada uno. Utilice el mismo **TaskpaneId** para distintas acciones que comparten el mismo panel. Cuando los usuarios eligen comandos que comparten la misma **TaskpaneId**, el contenedor del panel permanecerá abierto pero el contenido del panel se reemplazará por la correspondiente acción "SourceLocation". 

> [!NOTE]
> Este elemento no se admite en Outlook.

En el ejemplo siguiente se muestran dos acciones que comparten el mismo **TaskpaneId**. 

```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="aTaskPaneUrl" />
</Action>

<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="anotherTaskPaneUrl" />
</Action>
```  

En los ejemplos siguientes se muestran dos acciones que usan un **TaskpaneId** distinto. Puede ver estos ejemplos en contexto en [Ejemplo sencillo de comandos de complemento](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).

```xml
<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID1</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane1.Url" />
</Action>

<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID2</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane2.Url" />
</Action>
```  

```xml
<bt:Urls>
   <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
   <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
</bt:Urls>
```  

## <a name="title"></a>Title
Elemento opcional cuando **xsi: Type** es "ShowTaskpane". Especifica el título personalizado del panel de tareas de esta acción. 

Los ejemplos siguientes muestran dos acciones diferentes que usan el elemento **Title**.

```xml
<Action xsi:type="ShowTaskpane">
<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
<SourceLocation resid="PG.Code.Url" />
<Title resid="PG.CodeCommand.Title" />
</Action>
``` 

```xml
<Action xsi:type="ShowTaskpane">
<SourceLocation resid="PG.Run.Url" />
<Title resid="PG.RunCommand.Title" />
</Action>
``` 

```xml
<bt:Urls>
<bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
<bt:Url id="PG.Run.Url" DefaultValue="https://localhost:3000/run.html" />
</bt:Urls>
``` 

```xml
<bt:ShortStrings>
<bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
<bt:String id="PG.RunCommand.Title" DefaultValue="Run" />
</bt:ShortStrings>
``` 

## <a name="supportspinning"></a>SupportsPinning

Elemento opcional cuando **xsi:type** es "ShowTaskpane". Los elementos que contengan [VersionOverrides](versionoverrides.md) deben tener un valor de atributo `xsi:type` de `VersionOverridesV1_1`. Incluya este elemento con el valor `true` para admitir el anclado de paneles de tareas. El usuario podrá "anclar" el panel de tareas, lo que hará que permanezca abierto cuando se cambie la selección. Para obtener más información, consulte [Implementar un panel de tareas anclable en Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).

> [!NOTE]
> SupportsPinning sólo admite actualmente 2016 de Outlook para Windows (compilación 7628.1000 o posterior).

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```


