# <a name="action-element"></a><span data-ttu-id="88dd4-101">Elemento Action</span><span class="sxs-lookup"><span data-stu-id="88dd4-101">Action element</span></span>

<span data-ttu-id="88dd4-102">Especifica la acción que se realiza cuando el usuario selecciona los controles de [Botón](control.md#button-control) o [Menú](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="88dd4-102">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>
 
## <a name="attributes"></a><span data-ttu-id="88dd4-103">Atributos</span><span class="sxs-lookup"><span data-stu-id="88dd4-103">Attributes</span></span>

|  <span data-ttu-id="88dd4-104">Atributo</span><span class="sxs-lookup"><span data-stu-id="88dd4-104">Attribute</span></span>  |  <span data-ttu-id="88dd4-105">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="88dd4-105">Required</span></span>  |  <span data-ttu-id="88dd4-106">Descripción</span><span class="sxs-lookup"><span data-stu-id="88dd4-106">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="88dd4-107">xsi:type</span><span class="sxs-lookup"><span data-stu-id="88dd4-107">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="88dd4-108">Sí</span><span class="sxs-lookup"><span data-stu-id="88dd4-108">Yes</span></span>  | <span data-ttu-id="88dd4-109">Tipo de acción que se va a realizar</span><span class="sxs-lookup"><span data-stu-id="88dd4-109">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="88dd4-110">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="88dd4-110">Child elements</span></span>

|  <span data-ttu-id="88dd4-111">Elemento</span><span class="sxs-lookup"><span data-stu-id="88dd4-111">Element</span></span> |  <span data-ttu-id="88dd4-112">Descripción</span><span class="sxs-lookup"><span data-stu-id="88dd4-112">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="88dd4-113">FunctionName</span><span class="sxs-lookup"><span data-stu-id="88dd4-113">FunctionName</span></span>](#functionname) |    <span data-ttu-id="88dd4-114">Especifica el nombre de la función que se va a ejecutar.</span><span class="sxs-lookup"><span data-stu-id="88dd4-114">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="88dd4-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="88dd4-115">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="88dd4-116">Especifica la ubicación del archivo de origen para esta acción.</span><span class="sxs-lookup"><span data-stu-id="88dd4-116">Specifies the source file location for this action.</span></span> |
|  [<span data-ttu-id="88dd4-117">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="88dd4-117">TaskpaneId</span></span>](#taskpaneid) | <span data-ttu-id="88dd4-118">Especifica el id. del contenedor del panel de tareas.</span><span class="sxs-lookup"><span data-stu-id="88dd4-118">Specifies the ID of the task pane container.</span></span>|
|  [<span data-ttu-id="88dd4-119">Title</span><span class="sxs-lookup"><span data-stu-id="88dd4-119">Title</span></span>](#title) | <span data-ttu-id="88dd4-120">Especifica el título personalizado del panel de tareas.</span><span class="sxs-lookup"><span data-stu-id="88dd4-120">Specifies the custom title for the task pane.</span></span>|
|  [<span data-ttu-id="88dd4-121">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="88dd4-121">SupportsPinning</span></span>](#supportspinning) | <span data-ttu-id="88dd4-122">Especifica que un panel de tareas admite el anclado, lo que provoca que el panel de tareas siga abierto aunque el usuario cambie la selección.</span><span class="sxs-lookup"><span data-stu-id="88dd4-122">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="88dd4-123">xsi:type</span><span class="sxs-lookup"><span data-stu-id="88dd4-123">xsi:type</span></span>

<span data-ttu-id="88dd4-p101">Este atributo especifica el tipo de acción que se realiza cuando el usuario selecciona el botón. Puede ser uno de las siguientes:</span><span class="sxs-lookup"><span data-stu-id="88dd4-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="88dd4-126">FunctionName</span><span class="sxs-lookup"><span data-stu-id="88dd4-126">FunctionName</span></span>

<span data-ttu-id="88dd4-p102">Elemento obligatorio cuando **xsi:type** es "ExecuteFunction". Especifica el nombre de la función que se va a ejecutar. La función está incluida en el archivo especificado en el elemento [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="88dd4-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="88dd4-130">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="88dd4-130">SourceLocation</span></span>

<span data-ttu-id="88dd4-p103">Elemento obligatorio cuando **xsi:type** es "ShowTaskpane". Especifica la ubicación del archivo de origen para esta acción. El atributo **resid** debe establecerse en el valor del atributo **id** de un elemento **Url** en el elemento **Urls** del elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="88dd4-p103">Required element when  **xsi:type** is "ShowTaskpane". Specifies the source file location for this action. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="88dd4-134">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="88dd4-134">TaskpaneId</span></span>

<span data-ttu-id="88dd4-p104">Elemento opcional cuando **xsi: Type** es "ShowTaskpane". Especifica el identificador del contenedor de panel de tareas. Cuando haya varias acciones de "ShowTaskpane", utilice una **TaskpaneId** diferente si desea un panel independiente para cada uno. Utilice el mismo **TaskpaneId** para distintas acciones que comparten el mismo panel. Cuando los usuarios eligen comandos que comparten la misma **TaskpaneId**, el contenedor del panel permanecerá abierto pero el contenido del panel se reemplazará por la correspondiente acción "SourceLocation".</span><span class="sxs-lookup"><span data-stu-id="88dd4-p104">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span> 

> [!NOTE]
> <span data-ttu-id="88dd4-140">Este elemento no se admite en Outlook.</span><span class="sxs-lookup"><span data-stu-id="88dd4-140">This element is not supported in Outlook.</span></span>

<span data-ttu-id="88dd4-141">En el ejemplo siguiente se muestran dos acciones que comparten el mismo **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="88dd4-141">The following example shows two actions that share the same **TaskpaneId**.</span></span> 

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

<span data-ttu-id="88dd4-p105">En los ejemplos siguientes se muestran dos acciones que usan un **TaskpaneId** distinto. Puede ver estos ejemplos en contexto en [Ejemplo sencillo de comandos de complemento](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="88dd4-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="88dd4-144">Title</span><span class="sxs-lookup"><span data-stu-id="88dd4-144">Title</span></span>
<span data-ttu-id="88dd4-p106">Elemento opcional cuando **xsi: Type** es "ShowTaskpane". Especifica el título personalizado del panel de tareas de esta acción.</span><span class="sxs-lookup"><span data-stu-id="88dd4-p106">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the custom title for the task pane for this action.</span></span> 

<span data-ttu-id="88dd4-147">Los ejemplos siguientes muestran dos acciones diferentes que usan el elemento **Title**.</span><span class="sxs-lookup"><span data-stu-id="88dd4-147">The following examples show two different actions that use the **Title** element.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="88dd4-148">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="88dd4-148">SupportsPinning</span></span>

<span data-ttu-id="88dd4-p107">Elemento opcional cuando **xsi:type** es "ShowTaskpane". Los elementos que contengan [VersionOverrides](versionoverrides.md) deben tener un valor de atributo `xsi:type` de `VersionOverridesV1_1`. Incluya este elemento con el valor `true` para admitir el anclado de paneles de tareas. El usuario podrá "anclar" el panel de tareas, lo que hará que permanezca abierto cuando se cambie la selección. Para obtener más información, consulte [Implementar un panel de tareas anclable en Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span><span class="sxs-lookup"><span data-stu-id="88dd4-p107">Optional element when **xsi:type** is "ShowTaskpane". The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`. Include this element with a value of `true` to support taskpane pinning. The user will be able to "pin" the taskpane, causing it to stay open when changing the selection. For more information, see [Implement a pinnable taskpane in Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span></span>

> [!NOTE]
> <span data-ttu-id="88dd4-154">SupportsPinning sólo admite actualmente 2016 de Outlook para Windows (compilación 7628.1000 o posterior).</span><span class="sxs-lookup"><span data-stu-id="88dd4-154">SupportsPinning currently only supported by Outlook 2016 for Windows (build 7628.1000 or later).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```


