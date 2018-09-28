# <a name="versionoverrides-element"></a><span data-ttu-id="c312e-101">Elemento VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="c312e-101">VersionOverrides element</span></span>

<span data-ttu-id="c312e-p101">Elemento raíz que contiene la información de los comandos del complemento implementados por el complemento. **VersionOverrides** es un elemento secundario del elemento [OfficeApp](./officeapp.md) del manifiesto. Este elemento se admite en la versión 1.1 del esquema del manifiesto y en las posteriores, pero se define en el esquema de la versión 1.0 o 1.1 de VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="c312e-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="c312e-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="c312e-105">Attributes</span></span>

|  <span data-ttu-id="c312e-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="c312e-106">Attribute</span></span>  |  <span data-ttu-id="c312e-107">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="c312e-107">Required</span></span>  |  <span data-ttu-id="c312e-108">Descripción</span><span class="sxs-lookup"><span data-stu-id="c312e-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c312e-109">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="c312e-109">**xmlns**</span></span>       |  <span data-ttu-id="c312e-110">Sí</span><span class="sxs-lookup"><span data-stu-id="c312e-110">Yes</span></span>  |  <span data-ttu-id="c312e-111">Ubicación del esquema, que tiene que ser `http://schemas.microsoft.com/office/mailappversionoverrides` cuando `xsi:type` sea `VersionOverridesV1_0` y `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` cuando `xsi:type` sea `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="c312e-111">The schema location, which must be `http://schemas.microsoft.com/office/mailappversionoverrides` when `xsi:type` is `VersionOverridesV1_0`, and `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` when `xsi:type` is `VersionOverridesV1_1`.</span></span>|
|  <span data-ttu-id="c312e-112">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="c312e-112">**xsi:type**</span></span>  |  <span data-ttu-id="c312e-113">Sí</span><span class="sxs-lookup"><span data-stu-id="c312e-113">Yes</span></span>  | <span data-ttu-id="c312e-p102">Versión del esquema. En este momento, los únicos valores válidos son `VersionOverridesV1_0` y `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="c312e-p102">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

> [!NOTE]
> <span data-ttu-id="c312e-116">Actualmente sólo 2016 de Outlook es compatible con el esquema v1.1 de VersionOverrides y la `VersionOverridesV1_1` tipo.</span><span class="sxs-lookup"><span data-stu-id="c312e-116">Currently only Outlook 2016 supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c312e-117">Elementos secundarios</span><span class="sxs-lookup"><span data-stu-id="c312e-117">Child elements</span></span>

|  <span data-ttu-id="c312e-118">Elemento</span><span class="sxs-lookup"><span data-stu-id="c312e-118">Element</span></span> |  <span data-ttu-id="c312e-119">Obligatorio</span><span class="sxs-lookup"><span data-stu-id="c312e-119">Required</span></span>  |  <span data-ttu-id="c312e-120">Descripción</span><span class="sxs-lookup"><span data-stu-id="c312e-120">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c312e-121">**Descripción**</span><span class="sxs-lookup"><span data-stu-id="c312e-121">**Description**</span></span>    |  <span data-ttu-id="c312e-122">No</span><span class="sxs-lookup"><span data-stu-id="c312e-122">No</span></span>   |  <span data-ttu-id="c312e-p103">Describe el complemento. Esto reemplaza el elemento `Description` en cualquier parte principal del manifiesto. El texto de la descripción está contenido en un elemento secundario del elemento **LongString**, contenido en el elemento [Resources](./resources.md). El atributo `resid` del elemento **Description** está establecido en el valor del atributo `id` del elemento `String` que contiene el texto.</span><span class="sxs-lookup"><span data-stu-id="c312e-p103">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="c312e-127">**Requisitos**</span><span class="sxs-lookup"><span data-stu-id="c312e-127">**Requirements**</span></span>  |  <span data-ttu-id="c312e-128">No</span><span class="sxs-lookup"><span data-stu-id="c312e-128">No</span></span>   |  <span data-ttu-id="c312e-p104">Especifica el conjunto de requisitos mínimos y la versión de Office.js que necesita el complemento. Esto reemplaza el elemento `Requirements` en la parte principal del manifiesto.</span><span class="sxs-lookup"><span data-stu-id="c312e-p104">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="c312e-131">Hosts</span><span class="sxs-lookup"><span data-stu-id="c312e-131">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="c312e-132">Sí</span><span class="sxs-lookup"><span data-stu-id="c312e-132">Yes</span></span>  |  <span data-ttu-id="c312e-p105">Especifica una colección de hosts de Office. El elemento Hosts secundario reemplaza el elemento Hosts en cualquier parte principal del manifiesto.</span><span class="sxs-lookup"><span data-stu-id="c312e-p105">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="c312e-135">Recursos</span><span class="sxs-lookup"><span data-stu-id="c312e-135">Resources</span></span>](./resources.md)    |  <span data-ttu-id="c312e-136">Sí</span><span class="sxs-lookup"><span data-stu-id="c312e-136">Yes</span></span>  | <span data-ttu-id="c312e-137">Define una colección de recursos (cadenas, direcciones URL e imágenes) a las que hacen referencia otros elementos del manifiesto.</span><span class="sxs-lookup"><span data-stu-id="c312e-137">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  <span data-ttu-id="c312e-138">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="c312e-138">**VersionOverrides**</span></span>    |  <span data-ttu-id="c312e-139">No</span><span class="sxs-lookup"><span data-stu-id="c312e-139">No</span></span>  | <span data-ttu-id="c312e-p106">Define comandos de complemento en una versión más reciente del esquema. Consulte [Implementar varias versiones](#implementing-multiple-versions) para obtener detalles.</span><span class="sxs-lookup"><span data-stu-id="c312e-p106">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  <span data-ttu-id="c312e-142">**WebApplicationInfo**</span><span class="sxs-lookup"><span data-stu-id="c312e-142">**WebApplicationInfo**</span></span>    |  <span data-ttu-id="c312e-143">No</span><span class="sxs-lookup"><span data-stu-id="c312e-143">No</span></span>  | <span data-ttu-id="c312e-144">Especifica detalles sobre la aplicación web asociada al complemento.</span><span class="sxs-lookup"><span data-stu-id="c312e-144">Specifies details about the add-in's associated Web application.</span></span> |



### <a name="versionoverrides-example"></a><span data-ttu-id="c312e-145">Ejemplo de VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="c312e-145">VersionOverrides example</span></span>
```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a><span data-ttu-id="c312e-146">Implementar varias versiones</span><span class="sxs-lookup"><span data-stu-id="c312e-146">Implementing multiple versions</span></span>

<span data-ttu-id="c312e-p107">Un manifiesto puede implementar varias versiones del elemento `VersionOverrides` que admiten diferentes versiones del esquema de VersionOverrides. Esta acción puede realizarse para tener la opción de admitir nuevas características en un esquema más reciente, pero, a la vez, admitir clientes anteriores que no sean compatibles con las nuevas características.</span><span class="sxs-lookup"><span data-stu-id="c312e-p107">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="c312e-149">Para poder implementar varias versiones, el elemento `VersionOverrides` de la versión más reciente deberá ser un elemento secundario del elemento `VersionOverrides` de la versión anterior.</span><span class="sxs-lookup"><span data-stu-id="c312e-149">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="c312e-150">El elemento `VersionOverrides` secundario no hereda ningún valor del elemento primario.</span><span class="sxs-lookup"><span data-stu-id="c312e-150">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="c312e-151">Para implementar tanto el esquema de la versión 1.0 como el de la versión 1.1 de VersionOverrides, el manifiesto debería ser similar al del siguiente ejemplo:</span><span class="sxs-lookup"><span data-stu-id="c312e-151">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
...
</OfficeApp>
```
