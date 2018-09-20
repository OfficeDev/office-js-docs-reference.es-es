# <a name="officeapp-element"></a>Elemento OfficeApp

El elemento raíz del manifiesto de un complemento de Office.

**Tipo de complemento:** Contenido, panel de tareas, correo

## <a name="syntax"></a>Sintaxis

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a>Contenidos en

 _none_

## <a name="must-contain"></a>Debe contener

|**Element**|**Contenido**|**Correo**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[id
](id.md)|x|x|x|
|[Version](version.md)|x|x|x|
|[ProviderName](providername.md)|x|x|x|
|[DefaultLocale](defaultlocale.md)|x|x|x|
|[DefaultSettings](defaultsettings.md)|x||x|
|[DisplayName](displayname.md)|x|x|x|
|[Descripción](description.md)|x|x|x|
|[FormSettings](formsettings.md)||x||
|[Permisos](permissions.md)|x||x|
|[Rule](rule.md)||x||

## <a name="can-contain"></a>Puede contener

|**Element**|**Contenido**|**Correo**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[AlternateId](alternateid.md)|x|x|x|
|[IconUrl](iconurl.md)|x|x|x|
|[HighResolutionIconUrl](highresolutioniconurl.md)|x|x|x|
|[SupportUrl](supporturl.md)|x|x|x|
|[AppDomains](appdomains.md)|x|x|x|
|[Hosts](hosts.md)|x|x|x|
|[Requisitos](requirements.md)|x|x|x|
|[AllowSnapshot](allowsnapshot.md)|x|||
|[Permisos](permissions.md)||x||
|[DisableEntityHighlighting](disableentityhighlighting.md)||x||
|[Diccionario](dictionary.md)|||x|
|[VersionOverrides](versionoverrides.md)|X|X|X|

## <a name="attributes"></a>Atributos

|||
|:-----|:-----|
|xmlns|Define el esquema de versión y el espacio de nombres del manifiesto del complemento de Office. Este atributo debe establecerse siempre en `"http://schemas.microsoft.com/office/appforoffice/1.1"`|
|xmlns:xsi|Define la instancia de esquema XML. Este atributo debe establecerse siempre en `"http://www.w3.org/2001/XMLSchema-instance"`|
|xsi:type|Define el tipo de complemento de Office. Este atributo debe establecerse en uno de los siguientes valores: `"ContentApp"`, `"MailApp"` o `"TaskPaneApp"`|
