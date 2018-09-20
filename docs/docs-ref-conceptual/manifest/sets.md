# <a name="sets-element"></a>Elemento Sets

Especifica el subconjunto mínimo de la API de JavaScript para Office que su complemento de Office necesita para activarse.

**Tipo de complemento:** Contenido, panel de tareas, correo

## <a name="syntax"></a>Sintaxis

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>Contenidos en

[Requisitos](requirements.md)

## <a name="can-contain"></a>Puede contener

[Set](set.md)

## <a name="attributes"></a>Atributos

|**Attribute**|**Tipo**|**Obligatorio**|**Descripción**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|string|opcional|Especifica el valor de atributo predeterminado de **MinVersion** elementos [Set](set.md) secundarios. El valor predeterminado es "1.1".|

## <a name="remarks"></a>Comentarios

Para obtener más información acerca de los conjuntos de requisitos, vea [conjuntos de versiones de Office y los requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Para obtener más información sobre el atributo **MinVersion** del elemento **Set**  y del atributo **DefaultMinVersion** del elemento **Sets**, consulte [Definir el elemento Requirements en el manifiesto](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

