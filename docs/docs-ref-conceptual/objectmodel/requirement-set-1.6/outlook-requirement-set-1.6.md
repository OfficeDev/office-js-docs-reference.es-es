# <a name="outlook-add-in-api-requirement-set-16"></a>Requisito de Outlook en API establece 1.6

El Outlook complemento subconjunto de API de la API de JavaScript para Office incluye objetos, métodos, propiedades y eventos que puede usar en un Outlook complemento.

## <a name="whats-new-in-16"></a>¿Qué es una novedad de 1.6?

Conjunto de requisitos 1.6 incluye todas las características del [conjunto de requisito de 1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md). Agregan las siguientes características.

- Se agregaron nuevas API para complementos contextuales para obtener la entidad o RegEx coincide con la que el usuario seleccionado para activar el complemento.
- Se agregó una nueva API para abrir un nuevo formulario de mensaje.
- Se agregó la capacidad para el complemento determinar el tipo de cuenta del buzón del usuario.

### <a name="change-log"></a>Registro de cambios

- Agregado [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities): agrega una nueva función que obtiene las entidades que se encuentran en una coincidencia resaltada un usuario ha seleccionado. Las coincidencias resaltadas se aplican a complementos contextuales.
- Agregado [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object): agrega una nueva función que devuelve valores de cadena en una coincidencia resaltada que coinciden con las expresiones regulares definidas en el archivo XML del manifiesto. Las coincidencias resaltadas se aplican a complementos contextuales.
- Agregado [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters): agrega una nueva función que se abre un nuevo formulario de mensaje.
- Agregado [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string): agrega un nuevo miembro al perfil de usuario que indica el tipo de la cuenta de usuario.

## <a name="see-also"></a>Vea también

- [Complementos de Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Ejemplos de código de complementos de Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introducción](https://docs.microsoft.com/outlook/add-ins/quick-start)