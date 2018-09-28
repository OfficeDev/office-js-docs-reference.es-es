# <a name="permissions-element"></a><span data-ttu-id="31a04-101">Elemento Permissions</span><span class="sxs-lookup"><span data-stu-id="31a04-101">Permissions element</span></span>

<span data-ttu-id="31a04-102">Especifica el nivel de acceso a la API para su complemento de Office; debe solicitar permisos según el principio de privilegios mínimos.</span><span class="sxs-lookup"><span data-stu-id="31a04-102">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="31a04-103">**Tipo de complemento:** Contenido, panel de tareas, correo</span><span class="sxs-lookup"><span data-stu-id="31a04-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="31a04-104">Sintaxis</span><span class="sxs-lookup"><span data-stu-id="31a04-104">Syntax</span></span>

<span data-ttu-id="31a04-105">Para complementos de paneles de tareas:</span><span class="sxs-lookup"><span data-stu-id="31a04-105">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="31a04-106">Para los complementos de correo</span><span class="sxs-lookup"><span data-stu-id="31a04-106">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="31a04-107">Contenidos en</span><span class="sxs-lookup"><span data-stu-id="31a04-107">Contained in</span></span>

[<span data-ttu-id="31a04-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="31a04-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="31a04-109">Comentarios</span><span class="sxs-lookup"><span data-stu-id="31a04-109">Remarks</span></span>

<span data-ttu-id="31a04-110">Para obtener más información, consulte [Solicitar permisos para el uso de API en complementos de contenido y de panel de tareas](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) e [Información sobre los permisos de los complementos de Outlook](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="31a04-110">For more detail, see [Requesting permissions for API use in content and task pane add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
