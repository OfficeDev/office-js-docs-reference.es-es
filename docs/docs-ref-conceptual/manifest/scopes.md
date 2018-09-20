# <a name="scopes-element"></a>Elemento de ámbitos

Contiene los permisos que necesita el complemento en Microsoft Graph. La Tienda Office usa el elemento Scopes para crear un cuadro de diálogo de consentimiento. Cuando los usuarios instalan el complemento de la Tienda, se les pide que concedan al complemento los permisos especificados para los datos de Microsoft Graph del usuario.

## <a name="child-elements"></a>Elementos secundarios

|  Elemento |  Tipo  |  Descripción  |
|:-----|:-----|:-----|
|  **Ámbito**                |  string     |   El nombre de un permiso para Microsoft Graph; por ejemplo, Files.Read.All. |

## <a name="example"></a>Ejemplo

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
