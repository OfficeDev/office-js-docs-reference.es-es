# <a name="defaultsettings-element"></a><span data-ttu-id="7c841-101">Elemento DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="7c841-101">DefaultSettings element</span></span>

<span data-ttu-id="7c841-102">Especifica la ubicación de origen de forma predeterminada y otras configuraciones predeterminadas de su contenido o complemento en el panel de tareas.</span><span class="sxs-lookup"><span data-stu-id="7c841-102">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="7c841-103">**Tipo de complemento:** Panel de tareas, contenido</span><span class="sxs-lookup"><span data-stu-id="7c841-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="7c841-104">Sintaxis</span><span class="sxs-lookup"><span data-stu-id="7c841-104">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="7c841-105">Contenidos en</span><span class="sxs-lookup"><span data-stu-id="7c841-105">Contained in</span></span>

[<span data-ttu-id="7c841-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="7c841-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="7c841-107">Puede contener</span><span class="sxs-lookup"><span data-stu-id="7c841-107">Can contain</span></span>

|<span data-ttu-id="7c841-108">**Element**</span><span class="sxs-lookup"><span data-stu-id="7c841-108">**Element**</span></span>|<span data-ttu-id="7c841-109">**Contenido**</span><span class="sxs-lookup"><span data-stu-id="7c841-109">**Content**</span></span>|<span data-ttu-id="7c841-110">**Correo**</span><span class="sxs-lookup"><span data-stu-id="7c841-110">**Mail**</span></span>|<span data-ttu-id="7c841-111">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="7c841-111">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="7c841-112">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="7c841-112">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="7c841-113">x</span><span class="sxs-lookup"><span data-stu-id="7c841-113">x</span></span>||<span data-ttu-id="7c841-114">x</span><span class="sxs-lookup"><span data-stu-id="7c841-114">x</span></span>|
|[<span data-ttu-id="7c841-115">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="7c841-115">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="7c841-116">x</span><span class="sxs-lookup"><span data-stu-id="7c841-116">x</span></span>|||
|[<span data-ttu-id="7c841-117">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="7c841-117">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="7c841-118">x</span><span class="sxs-lookup"><span data-stu-id="7c841-118">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="7c841-119">Comentarios</span><span class="sxs-lookup"><span data-stu-id="7c841-119">Remarks</span></span>

<span data-ttu-id="7c841-120">La ubicación del origen y otros parámetros del elemento **DefaultSettings** se aplican solo a complementos de panel de tares y contenido. Para complementos de correo, puede especificar las ubicaciones predeterminadas de los archivos de origen y otros parámetros predeterminados en el elemento [FormSettings](formsettings.md).</span><span class="sxs-lookup"><span data-stu-id="7c841-120">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

