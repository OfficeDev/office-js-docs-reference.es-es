# <a name="allowsnapshot-element"></a><span data-ttu-id="fe327-101">Elemento AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="fe327-101">AllowSnapshot element</span></span>

<span data-ttu-id="fe327-102">Especifica si una imagen de instantánea de su complemento de contenido debe guardarse con el documento host.</span><span class="sxs-lookup"><span data-stu-id="fe327-102">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="fe327-103">**Tipo de complemento:** Contenido</span><span class="sxs-lookup"><span data-stu-id="fe327-103">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="fe327-104">Sintaxis</span><span class="sxs-lookup"><span data-stu-id="fe327-104">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="fe327-105">Contenidos en</span><span class="sxs-lookup"><span data-stu-id="fe327-105">Contained in</span></span>

[<span data-ttu-id="fe327-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="fe327-106">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="fe327-107">Comentarios</span><span class="sxs-lookup"><span data-stu-id="fe327-107">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="fe327-108">**AllowSnapshot** es `true` de forma predeterminada.</span><span class="sxs-lookup"><span data-stu-id="fe327-108">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="fe327-109">Esto hace que una imagen del complemento sea visible para los usuarios que abren el documento en una versión de la aplicación host que no admite complementos de Office, o proporciona una imagen estática del complemento si la aplicación host no puede conectarse al servidor que aloja el complemento.</span><span class="sxs-lookup"><span data-stu-id="fe327-109">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="fe327-110">Sin embargo, esto también significa que se puede tener acceso a información confidencial que se muestra en el complemento directamente desde el documento que lo hospeda.</span><span class="sxs-lookup"><span data-stu-id="fe327-110">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

