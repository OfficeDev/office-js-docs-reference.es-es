# <a name="highresolutioniconurl-element"></a>Elemento HighResolutionIconUrl

Especifica la dirección URL de la imagen que se usa para representar su complemento de Office en la UX de inserción y la Tienda Office en pantallas con valores altos de PPP. 

**Tipo de complemento:** Contenido, panel de tareas, correo

## <a name="syntax"></a>Sintaxis

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>Puede contener

[Override](override.md)

## <a name="attributes"></a>Atributos

|**Attribute**|**Tipo**|**Obligatorio**|**Descripción**|
|:-----|:-----|:-----|:-----|
|DefaultValue|cadena (URL)|necesario|Especifica el valor predeterminado de esta opción, expresado para la configuración regional especificada en el elemento [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Comentarios

Para un complemento de correo, se muestra el icono en la interfaz de usuario **Archivo**  >  **Administrar complementos**. Para un complemento de contenido o panel de tareas, se muestra el icono en la interfaz de usuario **Insertar**  >  **Complementos**.

La imagen debe estar en uno de los siguientes formatos de archivo con una resolución recomendada de 64 x 64 píxeles: GIF, JPG, PNG, EXIF, BMP o TIFF. Para obtener más información, vea la sección _crear una identidad visual coherente para su aplicación_ en [crear eficaces listados en AppSource y dentro de Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).
