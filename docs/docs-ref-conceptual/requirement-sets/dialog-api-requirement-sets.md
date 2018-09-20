# <a name="dialog-api-requirement-sets"></a>Conjuntos de requisitos de la API de cuadros de diálogo

Los conjuntos de requisitos son grupos de miembros de la API con nombre. Complementos de Office utilizan conjuntos de requisitos especificados en el manifiesto o una comprobación en tiempo de ejecución para determinar si un host de Office admite las API que un complemento necesita. Para obtener más información, vea [conjuntos de versiones de Office y los requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Los complementos de Office se ejecutan en varias versiones de Office. En la siguiente tabla se enumeran los conjuntos de requisitos de la API de cuadros de diálogo, las aplicaciones de host de Office que admiten ese conjunto de requisitos y la compilación o números de versión de la aplicación de Office.

|  Conjunto de requisitos  | Office 2013 para Windows | Office 2016 para Windows (instalaciones MSI)   | Office 365 para profesionales de Windows (se instala C2R)   |  Office 365 para iPad  |  Office 365 para Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | Compilación 15.0.4855.1000 o posterior | Compilación 16.0.4390.1000 o posterior | Versión 1602 (compilación 6741.0000) o posterior | 1.22 o posterior | 15.20 o posterior| Enero de 2017 | Versión 1608 (compilación 7601.6800) o posterior|

Para obtener más información sobre las versiones, los números de compilación y Office Online Server, vea:

- [Números de versión y compilación de las versiones del canal de actualización para los clientes de Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [¿Qué versión de Office estoy usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Dónde puede encontrar el número de versión y de compilación de una aplicación de cliente de Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Información general de Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos comunes de la API de Office

Para obtener información sobre los conjuntos de requisitos comunes de la API, consulte [Office common API requirement sets (Conjuntos de requisitos comunes de la API de Office)](office-add-in-requirement-sets.md).

## <a name="dialog-api-11"></a>API de cuadros de diálogo 1.1 

La 1.1 de API de cuadro de diálogo es la primera versión de la API. Para obtener información detallada acerca de la API, vea el tema de referencia de [API de cuadro de diálogo](/javascript/api/office/office.ui) .

## <a name="see-also"></a>Vea también

- [Versiones de Office y conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar los hosts de Office y los requisitos de la API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifiesto XML de complementos para Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
