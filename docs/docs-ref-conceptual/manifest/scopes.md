# <a name="scopes-element"></a><span data-ttu-id="427d0-101">Elemento de ámbitos</span><span class="sxs-lookup"><span data-stu-id="427d0-101">Scopes element</span></span>

<span data-ttu-id="427d0-102">Contiene los permisos que necesita el complemento en Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="427d0-102">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="427d0-103">La Tienda Office usa el elemento Scopes para crear un cuadro de diálogo de consentimiento.</span><span class="sxs-lookup"><span data-stu-id="427d0-103">The Office Store uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="427d0-104">Cuando los usuarios instalan el complemento de la Tienda, se les pide que concedan al complemento los permisos especificados para los datos de Microsoft Graph del usuario.</span><span class="sxs-lookup"><span data-stu-id="427d0-104">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="427d0-105">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="427d0-105">Child elements</span></span>

|  <span data-ttu-id="427d0-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="427d0-106">Element</span></span> |  <span data-ttu-id="427d0-107">Tipo</span><span class="sxs-lookup"><span data-stu-id="427d0-107">Type</span></span>  |  <span data-ttu-id="427d0-108">Descripción</span><span class="sxs-lookup"><span data-stu-id="427d0-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="427d0-109">**Ámbito**</span><span class="sxs-lookup"><span data-stu-id="427d0-109">**Scope**</span></span>                |  <span data-ttu-id="427d0-110">string</span><span class="sxs-lookup"><span data-stu-id="427d0-110">string</span></span>     |   <span data-ttu-id="427d0-111">El nombre de un permiso para Microsoft Graph; por ejemplo, Files.Read.All.</span><span class="sxs-lookup"><span data-stu-id="427d0-111">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="427d0-112">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="427d0-112">Example</span></span>

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
