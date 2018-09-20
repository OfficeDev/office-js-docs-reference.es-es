# <a name="webapplicationinfo-element"></a>Elemento WebApplicationInfo

Admite inicio de sesión único (SSO) en Complementos de Office. Este elemento contiene información sobre el complemento como:

- Un *recurso*OAuth 2.0 para el que la aplicación host de Office podría necesitar permisos.
- Un *cliente* OAuth 2.0 que podría necesitar permisos para Microsoft Graph.

**WebApplicationInfo** es un elemento secundario del elemento [VersionOverrides](versionoverrides.md) del manifiesto.  

## <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Obligatorio  |  Descripción  |
|:-----|:-----|:-----|
|  **Id**    |  Sí   |  **Id. de aplicación** del servicio asociado al complemento según está registrado en el punto de conexión de Azure Active Directory v2.0.|
|  **Resource**  |  Sí   |  Especifica el **URI de id. de aplicación** del complemento según está registrado en el punto de conexión de Azure Active Directory v2.0.|
|  [Scopes](scopes.md)                |  No  |  Especifica los permisos que el complemento necesita en Microsoft Graph.  |

> [!NOTE] 
> En la actualidad, es necesario que recurso del complemento coincide con su Host. Office no solicitará un token para un complemento, salvo que pueda demostrar la propiedad, lo que actualmente se hace hospedando el complemento en el nombre de dominio completo del Resource.

## <a name="webapplicationinfo-example"></a>Ejemplo de WebApplicationInfo

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
