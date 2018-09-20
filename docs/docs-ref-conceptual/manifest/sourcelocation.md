# <a name="sourcelocation-element"></a><span data-ttu-id="67143-101">Elemento SourceLocation</span><span class="sxs-lookup"><span data-stu-id="67143-101">SourceLocation element</span></span>

<span data-ttu-id="67143-p101">Especifica las ubicaciones del código fuente para su complemento de Office como una dirección URL de entre 1 y 2018 caracteres. La ubicación de origen debe ser una dirección HTTPS, no una ruta de acceso de archivo.</span><span class="sxs-lookup"><span data-stu-id="67143-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="67143-104">**Tipo de complemento:** Contenido, panel de tareas, correo</span><span class="sxs-lookup"><span data-stu-id="67143-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="67143-105">Sintaxis</span><span class="sxs-lookup"><span data-stu-id="67143-105">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="67143-106">Contenidos en</span><span class="sxs-lookup"><span data-stu-id="67143-106">Contained in</span></span>

- <span data-ttu-id="67143-107">[DefaultSettings](defaultsettings.md) (complementos de contenido y de panel de tareas)</span><span class="sxs-lookup"><span data-stu-id="67143-107">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="67143-108">[FormSettings](formsettings.md) (complementos de correo)</span><span class="sxs-lookup"><span data-stu-id="67143-108">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="67143-109">[ExtensionPoint](extensionpoint.md) (complementos de correo contextuales)</span><span class="sxs-lookup"><span data-stu-id="67143-109">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="67143-110">Puede contener</span><span class="sxs-lookup"><span data-stu-id="67143-110">Can contain</span></span>

[<span data-ttu-id="67143-111">Override</span><span class="sxs-lookup"><span data-stu-id="67143-111">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="67143-112">Atributos</span><span class="sxs-lookup"><span data-stu-id="67143-112">Attributes</span></span>

|<span data-ttu-id="67143-113">**Attribute**</span><span class="sxs-lookup"><span data-stu-id="67143-113">**Attribute**</span></span>|<span data-ttu-id="67143-114">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="67143-114">**Type**</span></span>|<span data-ttu-id="67143-115">**Obligatorio**</span><span class="sxs-lookup"><span data-stu-id="67143-115">**Required**</span></span>|<span data-ttu-id="67143-116">**Descripción**</span><span class="sxs-lookup"><span data-stu-id="67143-116">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="67143-117">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="67143-117">DefaultValue</span></span>|<span data-ttu-id="67143-118">Dirección URL</span><span class="sxs-lookup"><span data-stu-id="67143-118">URL</span></span>|<span data-ttu-id="67143-119">necesario</span><span class="sxs-lookup"><span data-stu-id="67143-119">required</span></span>|<span data-ttu-id="67143-120">Especifica el valor predeterminado de esta opción para la configuración regional especificada en el elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="67143-120">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
