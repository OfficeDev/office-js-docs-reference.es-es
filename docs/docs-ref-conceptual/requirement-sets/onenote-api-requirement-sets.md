# <a name="onenote-javascript-api-requirement-sets"></a>Conjuntos de requisitos de la API de JavaScript de OneNote

Los conjuntos de requisitos son grupos de miembros de la API con nombre. Complementos de Office utilizan conjuntos de requisitos especificados en el manifiesto o una comprobación en tiempo de ejecución para determinar si un host de Office admite las API que un complemento necesita. Para obtener más información, vea [conjuntos de versiones de Office y los requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

En la siguiente tabla se enumeran los conjuntos de requisitos de OneNote, las aplicaciones de host de Office que admiten esos conjuntos de requisitos y las versiones de la compilación o fecha de disponibilidad.

|  Conjunto de requisitos  |  Office Online | 
|:-----|:-----|
| OneNoteApi 1.1  | Septiembre de 2016 |  

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos comunes de la API de Office

Para obtener información sobre los conjuntos de requisitos comunes de la API, consulte [Office common API requirement sets (Conjuntos de requisitos comunes de la API de Office)](office-add-in-requirement-sets.md).

## <a name="onenote-javascript-api-11"></a>API de JavaScript de OneNote 1.1 

API de JavaScript de OneNote 1.1 es la primera versión de la API. Para obtener información detallada acerca de la API, vea la [información general sobre programación de API de JavaScript de OneNote](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).

## <a name="runtime-requirement-support-check"></a>Comprobación de compatibilidad de los requisitos en tiempo de ejecución

Durante el tiempo de ejecución, los complementos pueden comprobar si un host determinado admite un conjunto de requisitos de la API establecido realizando la siguiente comprobación: 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a>Comprobación de compatibilidad de los requisitos basada en un manifiesto

Use el elemento Requirements en el manifiesto del complemento para especificar conjuntos de requisitos críticos o miembros de la API que debe usar el complemento. Si la plataforma o el host de Office no son compatibles con los conjuntos de requisitos o los miembros de la API especificados en el elemento Requirements, el complemento no se ejecutará en la plataforma o el host ni se mostrará en Mis complementos.

En el siguiente ejemplo de código se muestra un complemento que se carga en todas las aplicaciones host de Office que admiten el conjunto de requisitos OneNoteApi, versión 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a>Vea también

- [Versiones de Office y conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar los hosts de Office y los requisitos de la API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifiesto XML de complementos para Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
