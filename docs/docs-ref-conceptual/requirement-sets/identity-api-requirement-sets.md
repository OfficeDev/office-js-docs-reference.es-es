# <a name="identity-api-requirement-sets"></a>Conjuntos de requisitos de la API de identidad

Los conjuntos de requisitos son grupos de miembros de la API con nombre. Complementos de Office utilizan conjuntos de requisitos especificados en el manifiesto o una comprobación en tiempo de ejecución para determinar si un host de Office admite las API que un complemento necesita. Para obtener más información, vea [conjuntos de versiones de Office y los requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Complementos de Office que se ejecuten a través de varias versiones de Office. En la siguiente tabla se enumera los conjuntos de requisitos de API de identidad, las aplicaciones host de Office que admiten que los números de conjunto y la compilación o la versión de requisito para la aplicación de Office.

|  Conjunto de requisitos  | Office 2013 para Windows | Office 365 para profesionales de Windows   |  Office 365 para iPad  |  Office 365 para Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com y Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | N/D | Vista previa **& #42;** | Próximamente | Vista previa **& #42;**| Disponible | Disponible| Próximamente | Próximamente |

> **& #42;** Durante la fase de vista previa, la API de identidad es compatible con 2016 de Windows y Mac sólo para los usuarios en el programa de personal mediante la opción Fast. Para unirse al programa de personal, vea [ser un especialista en Office](https://products.office.com/office-insider?tab=tab-1). Para cambiar a la Fast track, vea [Fast información confidencial](https://answers.microsoft.com/en-us/msoffice/forum/msoffice_officeinsider-mso_win10-msoinsider_reg/its-here-office-insider-fast-for-office-2016-on/dbe8e7bb-9523-44a4-948b-9436fedfd961).

Para obtener más información sobre las versiones, los números de compilación y Office Online Server, vea:

- [Números de versión y compilación de las versiones del canal de actualización para los clientes de Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [¿Qué versión de Office estoy usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Dónde puede encontrar el número de versión y de compilación de una aplicación de cliente de Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- 
  [Información general de Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos comunes de la API de Office

Para obtener información sobre los conjuntos de requisitos comunes de la API, vea [Conjuntos de requisitos comunes de la API de Office](office-add-in-requirement-sets.md).

## <a name="identityapi-11"></a>IdentityAPI 1.1 

IdentityAPI 1.1 de inicio de sesión único es la primera versión de la API. Para obtener información detallada acerca de esta API, vea la sección de [referencia de la API de SSO](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) de [Habilitar SSO en un complemento](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins).

## <a name="see-also"></a>Vea también

- [Versiones de Office y conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar los hosts de Office y los requisitos de la API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifiesto XML de complementos para Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
