# <a name="override-element"></a><span data-ttu-id="ebc2e-101">Elemento Override</span><span class="sxs-lookup"><span data-stu-id="ebc2e-101">Override element</span></span>

<span data-ttu-id="ebc2e-102">Permite especificar el valor de una opción para una configuración regional adicional.</span><span class="sxs-lookup"><span data-stu-id="ebc2e-102">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="ebc2e-103">**Tipo de complemento:** Contenido, panel de tareas, correo</span><span class="sxs-lookup"><span data-stu-id="ebc2e-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="ebc2e-104">Sintaxis</span><span class="sxs-lookup"><span data-stu-id="ebc2e-104">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="ebc2e-105">Contenidos en</span><span class="sxs-lookup"><span data-stu-id="ebc2e-105">Contained in</span></span>

|<span data-ttu-id="ebc2e-106">**Element**</span><span class="sxs-lookup"><span data-stu-id="ebc2e-106">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="ebc2e-107">CitationText</span><span class="sxs-lookup"><span data-stu-id="ebc2e-107">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="ebc2e-108">Descripción</span><span class="sxs-lookup"><span data-stu-id="ebc2e-108">Description</span></span>](description.md)|
|[<span data-ttu-id="ebc2e-109">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="ebc2e-109">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="ebc2e-110">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="ebc2e-110">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="ebc2e-111">DisplayName</span><span class="sxs-lookup"><span data-stu-id="ebc2e-111">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="ebc2e-112">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="ebc2e-112">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="ebc2e-113">IconUrl</span><span class="sxs-lookup"><span data-stu-id="ebc2e-113">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="ebc2e-114">QueryUri</span><span class="sxs-lookup"><span data-stu-id="ebc2e-114">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="ebc2e-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="ebc2e-115">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="ebc2e-116">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="ebc2e-116">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="ebc2e-117">Atributos</span><span class="sxs-lookup"><span data-stu-id="ebc2e-117">Attributes</span></span>

|<span data-ttu-id="ebc2e-118">**Attribute**</span><span class="sxs-lookup"><span data-stu-id="ebc2e-118">**Attribute**</span></span>|<span data-ttu-id="ebc2e-119">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="ebc2e-119">**Type**</span></span>|<span data-ttu-id="ebc2e-120">**Obligatorio**</span><span class="sxs-lookup"><span data-stu-id="ebc2e-120">**Required**</span></span>|<span data-ttu-id="ebc2e-121">**Descripción**</span><span class="sxs-lookup"><span data-stu-id="ebc2e-121">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ebc2e-122">Locale</span><span class="sxs-lookup"><span data-stu-id="ebc2e-122">Locale</span></span>|<span data-ttu-id="ebc2e-123">string</span><span class="sxs-lookup"><span data-stu-id="ebc2e-123">string</span></span>|<span data-ttu-id="ebc2e-124">necesario</span><span class="sxs-lookup"><span data-stu-id="ebc2e-124">required</span></span>|<span data-ttu-id="ebc2e-125">Especifica el nombre de referencia cultural de la configuración regional para esta invalidación en el formato de etiqueta de lenguaje BCP 47, como en `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="ebc2e-125">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="ebc2e-126">Valor</span><span class="sxs-lookup"><span data-stu-id="ebc2e-126">Value</span></span>|<span data-ttu-id="ebc2e-127">string</span><span class="sxs-lookup"><span data-stu-id="ebc2e-127">string</span></span>|<span data-ttu-id="ebc2e-128">necesario</span><span class="sxs-lookup"><span data-stu-id="ebc2e-128">required</span></span>|<span data-ttu-id="ebc2e-129">Especifica el valor de la opción de configuración expresado para la configuración regional especificada.</span><span class="sxs-lookup"><span data-stu-id="ebc2e-129">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="ebc2e-130">Recursos adicionales</span><span class="sxs-lookup"><span data-stu-id="ebc2e-130">See also</span></span>

- [<span data-ttu-id="ebc2e-131">Localización de complementos de Office</span><span class="sxs-lookup"><span data-stu-id="ebc2e-131">Localization for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
