# <a name="supporturl-element"></a><span data-ttu-id="f6f20-101">Elemento SupportUrl</span><span class="sxs-lookup"><span data-stu-id="f6f20-101">SupportUrl element</span></span>

<span data-ttu-id="f6f20-102">Especifica la dirección URL de una página que proporciona información de soporte para su complemento.</span><span class="sxs-lookup"><span data-stu-id="f6f20-102">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="f6f20-103">Sintaxis</span><span class="sxs-lookup"><span data-stu-id="f6f20-103">Syntax</span></span>

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="f6f20-104">Contenidos en</span><span class="sxs-lookup"><span data-stu-id="f6f20-104">Contained in</span></span>

[<span data-ttu-id="f6f20-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="f6f20-105">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="f6f20-106">Puede contener</span><span class="sxs-lookup"><span data-stu-id="f6f20-106">Can contain</span></span>

|  <span data-ttu-id="f6f20-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="f6f20-107">Element</span></span> | <span data-ttu-id="f6f20-108">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="f6f20-108">Required</span></span> | <span data-ttu-id="f6f20-109">Descripción</span><span class="sxs-lookup"><span data-stu-id="f6f20-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f6f20-110">Override</span><span class="sxs-lookup"><span data-stu-id="f6f20-110">Override</span></span>](override.md)   | <span data-ttu-id="f6f20-111">No</span><span class="sxs-lookup"><span data-stu-id="f6f20-111">No</span></span> | <span data-ttu-id="f6f20-112">Especifica la configuración de direcciones URL adicionales de configuración regional</span><span class="sxs-lookup"><span data-stu-id="f6f20-112">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="f6f20-113">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6f20-113">Attributes</span></span>

|<span data-ttu-id="f6f20-114">**Attribute**</span><span class="sxs-lookup"><span data-stu-id="f6f20-114">**Attribute**</span></span>|<span data-ttu-id="f6f20-115">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="f6f20-115">**Type**</span></span>|<span data-ttu-id="f6f20-116">**Obligatorio**</span><span class="sxs-lookup"><span data-stu-id="f6f20-116">**Required**</span></span>|<span data-ttu-id="f6f20-117">**Descripción**</span><span class="sxs-lookup"><span data-stu-id="f6f20-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="f6f20-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="f6f20-118">DefaultValue</span></span>|<span data-ttu-id="f6f20-119">Dirección URL</span><span class="sxs-lookup"><span data-stu-id="f6f20-119">URL</span></span>|<span data-ttu-id="f6f20-120">necesario</span><span class="sxs-lookup"><span data-stu-id="f6f20-120">required</span></span>|<span data-ttu-id="f6f20-121">Especifica el valor predeterminado de esta opción, expresado para la configuración regional especificada en el elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="f6f20-121">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
