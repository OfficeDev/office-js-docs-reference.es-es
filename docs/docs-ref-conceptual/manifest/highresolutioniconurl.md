# <a name="highresolutioniconurl-element"></a><span data-ttu-id="c2cbc-101">Elemento HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="c2cbc-101">HighResolutionIconUrl element</span></span>

<span data-ttu-id="c2cbc-102">Especifica la dirección URL de la imagen que se usa para representar su complemento de Office en la UX de inserción y la Tienda Office en pantallas con valores altos de PPP. </span><span class="sxs-lookup"><span data-stu-id="c2cbc-102">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="c2cbc-103">**Tipo de complemento:** Contenido, panel de tareas, correo</span><span class="sxs-lookup"><span data-stu-id="c2cbc-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c2cbc-104">Sintaxis</span><span class="sxs-lookup"><span data-stu-id="c2cbc-104">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="c2cbc-105">Puede contener</span><span class="sxs-lookup"><span data-stu-id="c2cbc-105">Can contain</span></span>

[<span data-ttu-id="c2cbc-106">Override</span><span class="sxs-lookup"><span data-stu-id="c2cbc-106">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="c2cbc-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="c2cbc-107">Attributes</span></span>

|<span data-ttu-id="c2cbc-108">**Attribute**</span><span class="sxs-lookup"><span data-stu-id="c2cbc-108">**Attribute**</span></span>|<span data-ttu-id="c2cbc-109">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="c2cbc-109">**Type**</span></span>|<span data-ttu-id="c2cbc-110">**Obligatorio**</span><span class="sxs-lookup"><span data-stu-id="c2cbc-110">**Required**</span></span>|<span data-ttu-id="c2cbc-111">**Descripción**</span><span class="sxs-lookup"><span data-stu-id="c2cbc-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="c2cbc-112">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="c2cbc-112">DefaultValue</span></span>|<span data-ttu-id="c2cbc-113">cadena (URL)</span><span class="sxs-lookup"><span data-stu-id="c2cbc-113">string (URL)</span></span>|<span data-ttu-id="c2cbc-114">necesario</span><span class="sxs-lookup"><span data-stu-id="c2cbc-114">required</span></span>|<span data-ttu-id="c2cbc-115">Especifica el valor predeterminado de esta opción, expresado para la configuración regional especificada en el elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="c2cbc-115">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="c2cbc-116">Comentarios</span><span class="sxs-lookup"><span data-stu-id="c2cbc-116">Remarks</span></span>

<span data-ttu-id="c2cbc-p101">Para un complemento de correo, se muestra el icono en la interfaz de usuario **Archivo**  >  **Administrar complementos**. Para un complemento de contenido o panel de tareas, se muestra el icono en la interfaz de usuario **Insertar**  >  **Complementos**.</span><span class="sxs-lookup"><span data-stu-id="c2cbc-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI . For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="c2cbc-119">La imagen debe estar en uno de los siguientes formatos de archivo con una resolución recomendada de 64 x 64 píxeles: GIF, JPG, PNG, EXIF, BMP o TIFF.</span><span class="sxs-lookup"><span data-stu-id="c2cbc-119">The image must be in one of the following file formats at a recommended resolution of 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP or TIFF.</span></span> <span data-ttu-id="c2cbc-120">Para obtener más información, vea la sección _crear una identidad visual coherente para su aplicación_ en [crear eficaces listados en AppSource y dentro de Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="c2cbc-120">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).</span></span>
