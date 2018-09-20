# <a name="permissions-element"></a>Elemento Permissions

Especifica el nivel de acceso a la API para su complemento de Office; debe solicitar permisos según el principio de privilegios mínimos.

**Tipo de complemento:** Contenido, panel de tareas, correo

## <a name="syntax"></a>Sintaxis

Para complementos de paneles de tareas:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Para los complementos de correo

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>Contenidos en

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Comentarios

Para obtener más información, consulte [Solicitar permisos para el uso de API en complementos de contenido y de panel de tareas](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) e [Información sobre los permisos de los complementos de Outlook](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).
