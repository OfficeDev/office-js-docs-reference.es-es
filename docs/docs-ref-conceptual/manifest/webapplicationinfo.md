# <a name="webapplicationinfo-element"></a><span data-ttu-id="77f71-101">Elemento WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="77f71-101">WebApplicationInfo element</span></span>

<span data-ttu-id="77f71-102">Admite inicio de sesión único (SSO) en Complementos de Office. Este elemento contiene información sobre el complemento como:</span><span class="sxs-lookup"><span data-stu-id="77f71-102">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="77f71-103">Un *recurso*OAuth 2.0 para el que la aplicación host de Office podría necesitar permisos.</span><span class="sxs-lookup"><span data-stu-id="77f71-103">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="77f71-104">Un *cliente* OAuth 2.0 que podría necesitar permisos para Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="77f71-104">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

<span data-ttu-id="77f71-105">**WebApplicationInfo** es un elemento secundario del elemento [VersionOverrides](versionoverrides.md) del manifiesto.</span><span class="sxs-lookup"><span data-stu-id="77f71-105">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="77f71-106">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="77f71-106">Child elements</span></span>

|  <span data-ttu-id="77f71-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="77f71-107">Element</span></span> |  <span data-ttu-id="77f71-108">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="77f71-108">Required</span></span>  |  <span data-ttu-id="77f71-109">Descripción</span><span class="sxs-lookup"><span data-stu-id="77f71-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="77f71-110">**Id**</span><span class="sxs-lookup"><span data-stu-id="77f71-110">**Id**</span></span>    |  <span data-ttu-id="77f71-111">Sí</span><span class="sxs-lookup"><span data-stu-id="77f71-111">Yes</span></span>   |  <span data-ttu-id="77f71-112">**Id. de aplicación** del servicio asociado al complemento según está registrado en el punto de conexión de Azure Active Directory v2.0.</span><span class="sxs-lookup"><span data-stu-id="77f71-112">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="77f71-113">**Resource**</span><span class="sxs-lookup"><span data-stu-id="77f71-113">**Resource**</span></span>  |  <span data-ttu-id="77f71-114">Sí</span><span class="sxs-lookup"><span data-stu-id="77f71-114">Yes</span></span>   |  <span data-ttu-id="77f71-115">Especifica el **URI de id. de aplicación** del complemento según está registrado en el punto de conexión de Azure Active Directory v2.0.</span><span class="sxs-lookup"><span data-stu-id="77f71-115">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="77f71-116">Scopes</span><span class="sxs-lookup"><span data-stu-id="77f71-116">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="77f71-117">No</span><span class="sxs-lookup"><span data-stu-id="77f71-117">No</span></span>  |  <span data-ttu-id="77f71-118">Especifica los permisos que el complemento necesita en Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="77f71-118">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="77f71-119">En la actualidad, es necesario que recurso del complemento coincide con su Host.</span><span class="sxs-lookup"><span data-stu-id="77f71-119">Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="77f71-120">Office no solicitará un token para un complemento, salvo que pueda demostrar la propiedad, lo que actualmente se hace hospedando el complemento en el nombre de dominio completo del Resource.</span><span class="sxs-lookup"><span data-stu-id="77f71-120">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="77f71-121">Ejemplo de WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="77f71-121">WebApplicationInfo example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>        
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
