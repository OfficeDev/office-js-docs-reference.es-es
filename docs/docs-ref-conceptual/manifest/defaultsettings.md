# <a name="defaultsettings-element"></a>Elemento DefaultSettings

Especifica la ubicación de origen de forma predeterminada y otras configuraciones predeterminadas de su contenido o complemento en el panel de tareas.

**Tipo de complemento:** Panel de tareas, contenido

## <a name="syntax"></a>Sintaxis

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a>Contenidos en

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Puede contener

|**Element**|**Contenido**|**Correo**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>Comentarios

La ubicación del origen y otros parámetros del elemento **DefaultSettings** se aplican solo a complementos de panel de tares y contenido. Para complementos de correo, puede especificar las ubicaciones predeterminadas de los archivos de origen y otros parámetros predeterminados en el elemento [FormSettings](formsettings.md).

