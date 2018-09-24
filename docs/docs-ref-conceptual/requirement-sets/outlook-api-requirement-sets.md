# <a name="outlook-javascript-api-requirement-sets"></a>Conjuntos de requisitos de API de JavaScript de Outlook

Complementos de Outlook declaran qué versiones de API que requieren mediante el elemento de [los requisitos](/javascript/office/manifest/requirements) en su [manifiesto](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests). Los complementos de Outlook siempre incluyen un elemento [Set](/javascript/office/manifest/set) con un atributo `Name` establecido en `Mailbox` y un atributo `MinVersion` establecido en el conjunto de requisitos mínimos de la API que admite los escenarios de los complementos.

Por ejemplo, el siguiente fragmento de manifiesto indica un conjunto de requisitos mínimos de la versión 1.1:

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

Todas las API de Outlook pertenecen al [conjunto de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements) `Mailbox`. El conjunto de requisitos `Mailbox` tiene versiones, y cada nuevo conjunto de API que se publica pertenece a una versión posterior del conjunto. No todos los clientes de Outlook admiten el conjunto más reciente de API, pero si un cliente de Outlook declara la compatibilidad con un conjunto de requisitos, es compatible con todas las API de ese conjunto de requisitos.

Al establecer una versión de conjunto de requisitos mínimos en el manifiesto se controla en qué cliente de Outlook aparecerá el complemento. Si un cliente no admite el conjunto de requisitos mínimos, no carga el complemento. Por ejemplo, si se especifica la versión 1.3 del conjunto de requisitos, esto significa que el complemento no se mostrará en ningún cliente de Outlook que no admita al menos la versión 1.3.

## <a name="using-apis-from-later-requirement-sets"></a>Usar las API desde conjuntos de requisitos posteriores

Configuración de un conjunto de requisito no limitan las API disponibles que puede usar el complemento. Por ejemplo, si el complemento especifica requisito establece 1.1, pero se está ejecutando en un cliente de Outlook que admiten 1.3, el complemento puede utilizar las API de conjunto de requisitos 1.3.

Para usar las API más reciente, los programadores pueden comprobar sólo para su existencia mediante el uso de la técnica de JavaScript estándar:

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

Dichos controles no son necesarios para ninguna API que esté presente en la versión del conjunto de requisitos especificada en el manifiesto.

## <a name="choosing-a-minimum-requirement-set"></a>Elegir un conjunto de requisitos mínimos

Los desarrolladores deben usar el conjunto de requisitos más antiguo que contenga el conjunto fundamental de las API para su escenario, sin el que no funcionará el complemento.

## <a name="clients"></a>Clientes

Los siguientes clientes admiten complementos de Outlook.

| Client | Conjuntos admitidos de requisitos de la API |
| --- | --- |
| Outlook 2019 para Windows | 1.1, 1.2, 1.3, 1.4, 1.5, 1.6 |
| 2019 Outlook para Mac | 1.1, 1.2, 1.3, 1.4, 1.5, 1.6 |
| Outlook 2016 (Hacer clic y ejecutar) para Windows | 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7 |
| Outlook 2016 (MSI) para Windows | 1.1, 1.2, 1.3, 1.4 |
| Outlook 2016 para Mac | 1.1, 1.2, 1.3, 1.4, 1.5, 1.6 |
| Outlook 2013 para Windows | 1.1, 1.2, 1.3, 1.4 |
| Outlook para iPhone | 1.1, 1.2, 1.3, 1.4, 1.5 |
| Outlook para Android | 1.1, 1.2, 1.3, 1.4, 1.5 |
| Outlook en la Web (Office 365 y Outlook.com) | 1.1, 1.2, 1.3, 1.4, 1.5, 1.6 |
| Outlook Web App (Exchange 2013 local) | 1.1 |
| Outlook Web App (Exchange 2016 local) | 1.1, 1.2. 1.3 |

> [!NOTE]
> Se agregó compatibilidad para 1.3 en Outlook 2013 como parte de la [8 de diciembre 2015, actualización para Outlook 2013 (KB3114349)](https://support.microsoft.com/kb/3114349). Se agregó compatibilidad para 1.4 en Outlook 2013 como parte de la [13 de septiembre, 2016, actualización para Outlook 2013 (KB3118280)](https://support.microsoft.com/help/3118280).
