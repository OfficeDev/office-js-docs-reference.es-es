# <a name="appdomains-element"></a><span data-ttu-id="b4979-101">Elemento AppDomains</span><span class="sxs-lookup"><span data-stu-id="b4979-101">AppDomains element</span></span>

<span data-ttu-id="b4979-p101">Enumera los dominios además del dominio especificado en el elemento SourceLocation que el complemento de Office utilizará para cargar páginas. Para cada dominio adicional, especifique un elemento AppDomain.</span><span class="sxs-lookup"><span data-stu-id="b4979-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="b4979-104">**Tipo de complemento:** Contenido, panel de tareas, correo</span><span class="sxs-lookup"><span data-stu-id="b4979-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b4979-105">Sintaxis</span><span class="sxs-lookup"><span data-stu-id="b4979-105">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

## <a name="contained-in"></a><span data-ttu-id="b4979-106">Contenidos en</span><span class="sxs-lookup"><span data-stu-id="b4979-106">Contained in</span></span>

[<span data-ttu-id="b4979-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b4979-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="b4979-108">Puede contener</span><span class="sxs-lookup"><span data-stu-id="b4979-108">Can contain</span></span>

[<span data-ttu-id="b4979-109">AppDomain</span><span class="sxs-lookup"><span data-stu-id="b4979-109">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="b4979-110">Comentarios</span><span class="sxs-lookup"><span data-stu-id="b4979-110">Remarks</span></span>

<span data-ttu-id="b4979-p102">De forma predeterminada, el complemento puede cargar cualquier página que esté en el mismo dominio que la ubicación especificada en el elemento **SourceLocation**. Para cargar páginas que no estén en el mismo dominio que el complemento, especifique los dominios usando los elementos **AppDomains** y **AppDomain**. Este elemento no puede estar vacío.</span><span class="sxs-lookup"><span data-stu-id="b4979-p102">By default, your add-in can load any page that is in the same domain as the location specified in the **SourceLocation** element. To load pages that are not in the same domain as the add-in, specify the domains by using the **AppDomains** and **AppDomain** elements. This element can't be empty.</span></span> 
