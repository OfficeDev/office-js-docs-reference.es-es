# <a name="supporturl-element"></a>Elemento SupportUrl

Especifica la dirección URL de una página que proporciona información de soporte para su complemento.

## <a name="syntax"></a>Sintaxis

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

## <a name="contained-in"></a>Contenidos en

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Puede contener

|  Elemento | Obligatorio | Descripción  |
|:-----|:-----|:-----|
|  [Override](override.md)   | No | Especifica la configuración de direcciones URL adicionales de configuración regional |

## <a name="attributes"></a>Atributos

|**Attribute**|**Tipo**|**Obligatorio**|**Descripción**|
|:-----|:-----|:-----|:-----|
|DefaultValue|Dirección URL|necesario|Especifica el valor predeterminado de esta opción, expresado para la configuración regional especificada en el elemento [DefaultLocale](defaultlocale.md).|
