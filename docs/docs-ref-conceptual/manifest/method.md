# <a name="method-element"></a><span data-ttu-id="7457f-101">Elemento Method</span><span class="sxs-lookup"><span data-stu-id="7457f-101">Method element</span></span>

<span data-ttu-id="7457f-102">Especifica un método individual de la API de JavaScript para Office que su complemento de Office necesita para activarse.</span><span class="sxs-lookup"><span data-stu-id="7457f-102">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="7457f-103">**Tipo de complemento:** Panel de tareas, contenido</span><span class="sxs-lookup"><span data-stu-id="7457f-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="7457f-104">Sintaxis</span><span class="sxs-lookup"><span data-stu-id="7457f-104">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="7457f-105">Contenidos en</span><span class="sxs-lookup"><span data-stu-id="7457f-105">Contained in</span></span>

[<span data-ttu-id="7457f-106">Métodos</span><span class="sxs-lookup"><span data-stu-id="7457f-106">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="7457f-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="7457f-107">Attributes</span></span>

|<span data-ttu-id="7457f-108">**Attribute**</span><span class="sxs-lookup"><span data-stu-id="7457f-108">**Attribute**</span></span>|<span data-ttu-id="7457f-109">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="7457f-109">**Type**</span></span>|<span data-ttu-id="7457f-110">**Obligatorio**</span><span class="sxs-lookup"><span data-stu-id="7457f-110">**Required**</span></span>|<span data-ttu-id="7457f-111">**Descripción**</span><span class="sxs-lookup"><span data-stu-id="7457f-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="7457f-112">Nombre</span><span class="sxs-lookup"><span data-stu-id="7457f-112">Name</span></span>|<span data-ttu-id="7457f-113">string</span><span class="sxs-lookup"><span data-stu-id="7457f-113">string</span></span>|<span data-ttu-id="7457f-114">necesario</span><span class="sxs-lookup"><span data-stu-id="7457f-114">required</span></span>|<span data-ttu-id="7457f-p101">Especifica el nombre del método necesario calificado con su objeto principal. Por ejemplo, para especificar el método **getSelectedDataAsync**, debe especificar `"Document.getSelectedDataAsync"`.</span><span class="sxs-lookup"><span data-stu-id="7457f-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="7457f-117">Comentarios</span><span class="sxs-lookup"><span data-stu-id="7457f-117">Remarks</span></span>

<span data-ttu-id="7457f-118">Los elementos de **métodos** y el **método** no son compatibles con los complementos de correo. Para obtener más información acerca de los conjuntos de requisitos, vea [conjuntos de versiones de Office y los requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="7457f-118">The  **Methods** and **Method** elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="7457f-119">Porque no hay ninguna forma de especificar los requisitos mínimos de versión para métodos individuales, para asegurarse de que un método está disponible en tiempo de ejecución, debe también utilizar una instrucción **if** cuando llama a ese método en la secuencia de comandos de complementos.</span><span class="sxs-lookup"><span data-stu-id="7457f-119">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="7457f-120">Para obtener más información acerca de cómo hacer esto, vea la [Descripción de la API de JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="7457f-120">For more information about how to do this, see [Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

