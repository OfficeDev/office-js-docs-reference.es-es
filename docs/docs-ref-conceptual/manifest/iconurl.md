# <a name="iconurl-element"></a>Elemento IconUrl

Especifica la dirección URL de la imagen que se usa para representar su complemento de Office en la UX de inserción y la Tienda Office.

**Tipo de complemento:** Contenido, panel de tareas, correo

## <a name="syntax"></a>Sintaxis

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>Puede contener

[Override](override.md)

## <a name="attributes"></a>Atributos

|**Attribute**|**Tipo**|**Obligatorio**|**Descripción**|
|:-----|:-----|:-----|:-----|
|DefaultValue|string|necesario|Especifica el valor predeterminado de esta opción, expresado para la configuración regional especificada en el elemento [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Comentarios

Para un complemento de correo, se muestra el icono en la interfaz de usuario **Archivo**  >  **Administrar complementos** (Outlook) o **Configuración**  >  **Administrar complementos** (Outlook Web App). Para un complemento de contenido o panel de tareas, se muestra el icono en la interfaz de usuario **Insertar**  >  **Complementos**. Para todos los tipos de complemento, también se usa el icono en el sitio de la Tienda Office si publica el complemento en la Tienda Office.

La imagen debe estar en uno de los siguientes formatos de archivo: GIF, JPG, PNG, EXIF, BMP o TIFF. Para aplicaciones de panel de tareas y de contenido, la imagen especificada debe ser de 32 x 32 píxeles. Para las aplicaciones de correo, la imagen debe ser 64 x 64 píxeles. También debe especificar un icono para su uso con aplicaciones de host de Office que se ejecuten en pantallas de PPP alta mediante el elemento [HighResolutionIconUrl](highresolutioniconurl.md) . Para obtener más información, vea la sección _crear una identidad visual coherente para su aplicación_ en [crear eficaces listados en AppSource y dentro de Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).
