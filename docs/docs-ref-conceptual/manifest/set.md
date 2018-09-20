# <a name="set-element"></a>Elemento Set

Especifica un conjunto de requisitos de la API de JavaScript para Office que su complemento de Office necesita para activarse.

**Tipo de complemento:** Contenido, panel de tareas, correo

## <a name="syntax"></a>Sintaxis

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>Contenidos en

[Conjuntos](sets.md)

## <a name="attributes"></a>Atributos

|**Attribute**|**Tipo**|**Obligatorio**|**Descripción**|
|:-----|:-----|:-----|:-----|
|Nombre|string|necesario|El nombre de un [conjunto de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).|
|MinVersion|string|opcional|Especifica la versión mínima del conjunto de API que necesita el complemento. Reemplaza el valor **DefaultMinVersion** si se especifica en el elemento primario [Sets](sets.md).|

## <a name="remarks"></a>Comentarios

Para obtener más información acerca de los conjuntos de requisitos, vea [conjuntos de versiones de Office y los requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Para obtener más información sobre el atributo **MinVersion** del elemento **Set**  y del atributo **DefaultMinVersion** del elemento **Sets**, consulte [Definir el elemento Requirements en el manifiesto](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

> [!IMPORTANT] 
> Para complementos de correo, sólo hay una `"Mailbox"` conjuntos de requisito. En este conjunto de requisitos contiene el subconjunto completo de API admitidas en los complementos de correo para Outlook, y debe especificar el `"Mailbox"` requisito establecido en manifiesto del su correo complemento (no es opcional como es el caso de tareas y de contenido complementos del panel). Además, no se puede declarar compatibilidad para métodos específicos en los complementos de correo.
