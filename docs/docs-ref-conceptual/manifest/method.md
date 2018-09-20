# <a name="method-element"></a>Elemento Method

Especifica un método individual de la API de JavaScript para Office que su complemento de Office necesita para activarse.

**Tipo de complemento:** Panel de tareas, contenido

## <a name="syntax"></a>Sintaxis

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>Contenidos en

[Métodos](methods.md)

## <a name="attributes"></a>Atributos

|**Attribute**|**Tipo**|**Obligatorio**|**Descripción**|
|:-----|:-----|:-----|:-----|
|Nombre|string|necesario|Especifica el nombre del método necesario calificado con su objeto principal. Por ejemplo, para especificar el método **getSelectedDataAsync**, debe especificar `"Document.getSelectedDataAsync"`.|

## <a name="remarks"></a>Comentarios

Los elementos de **métodos** y el **método** no son compatibles con los complementos de correo. Para obtener más información acerca de los conjuntos de requisitos, vea [conjuntos de versiones de Office y los requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

> [!IMPORTANT] 
> Porque no hay ninguna forma de especificar los requisitos mínimos de versión para métodos individuales, para asegurarse de que un método está disponible en tiempo de ejecución, debe también utilizar una instrucción **if** cuando llama a ese método en la secuencia de comandos de complementos. Para obtener más información acerca de cómo hacer esto, vea la [Descripción de la API de JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

