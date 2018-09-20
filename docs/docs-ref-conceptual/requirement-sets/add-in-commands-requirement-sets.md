# <a name="add-in-commands-requirement-sets"></a>Conjuntos de requisitos de comandos de complemento

Los conjuntos de requisitos son grupos de miembros de la API con nombre. Complementos de Office utilizan conjuntos de requisitos especificados en el manifiesto o una comprobación en tiempo de ejecución para determinar si un host de Office admite las API que un complemento necesita. Para obtener más información, vea [conjuntos de versiones de Office y los requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Los comandos de complemento son elementos de la interfaz de usuario que amplían la interfaz de usuario de Office e inician acciones en el complemento. Puede usar comandos de complemento para agregar un botón a la cinta de opciones o un elemento a un menú contextual. Para obtener más información, vea [Comandos de complemento para Excel, Word y PowerPoint](https://docs.microsoft.com/office/dev/add-ins/design/add-in-commands) y [Comandos de complemento para Outlook](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook).

La publicación inicial de los comandos de complemento no tiene un conjunto de requisitos correspondiente (es decir, no hay ningún conjunto de requisitos de AddinCommands 1.0). En la tabla siguiente se muestran las aplicaciones host de Office que admiten la versión de la publicación inicial, así como los números de compilación o versión de esas aplicaciones.  

| Versión   |  Office 2013 para Windows | 2016 de Office para Windows (sin suscripción) | Office 365 para profesionales de Windows   |  Office 365 para iPad  |  Office 365 para Mac  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Comandos de complemento (publicación inicial, sin conjunto de requisitos) | N/D | 16.0.4678.1000 *compatibles con solo en Outlook* |Versión 1603 (compilación 6769.0000) o posteriores | N/D | 15.33 o posteriores| Enero de 2016 | |

En el conjunto de requisitos de comando de complemento 1.1 se introduce la capacidad de [abrir automáticamente un panel de tareas con documentos](https://docs.microsoft.com/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document).

En la tabla siguiente se muestra una lista de los conjuntos de requisitos de los comandos de complemento 1.1, las aplicaciones host de Office que admiten ese conjunto de requisitos y los números de compilación o versión de la aplicación de Office. 

|  Conjunto de requisitos  |  Office 2013 para Windows | 2016 de Office para Windows (sin suscripción) | Office 365 para profesionales de Windows   |  Office 365 para iPad  |  Office 365 para Mac  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1.1  | N/D | 16.0.4678.1000 *compatibles con solo en Outlook*  | Versión 1705 (compilación 8121.1000) o posteriores | N/D | 15.34 o posteriores| Mayo de 2017 | |

Para obtener más información sobre las versiones, los números de compilación y Office Online Server, vea:

- [Números de versión y compilación de las versiones del canal de actualización para los clientes de Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [¿Qué versión de Office estoy usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Dónde puede encontrar el número de versión y de compilación de una aplicación de cliente de Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- 
  [Información general de Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos comunes de la API de Office

Para obtener información sobre los conjuntos de requisitos comunes de la API, vea [Conjuntos de requisitos comunes de la API de Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Vea también

- [Versiones de Office y conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar los hosts de Office y los requisitos de la API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifiesto XML de complementos para Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
