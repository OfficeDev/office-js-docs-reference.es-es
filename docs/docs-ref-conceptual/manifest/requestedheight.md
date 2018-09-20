# <a name="requestedheight-element"></a>Elemento RequestedHeight.

Especifica el alto inicial (en píxeles) de un contenido complemento o un complemento en el correo. 

**Tipo de complemento:** Contenido, correo

## <a name="syntax"></a>Sintaxis

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a>Contenidos en

- [DefaultSettings](defaultsettings.md) (Contenido de complementos) con un valor que puede estar entre 32 y 1.000
- [DesktopSettings](desktopsettings.md) y [TabletSettings](tabletsettings.md) (complementos de correo) con un valor que puede estar comprendido entre 32 y 450
- [ExtensionPoint](extensionpoint.md) (Complementos de correo contextual) con un valor que puede ser entre 140 y 450 para el punto de extensión **DetectedEntity** y entre 32 y 450 para el punto de extensión **CustomPane**