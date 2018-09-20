# <a name="contribute-to-this-documentation"></a>Contribuir a este contenido

Gracias por su inter?s en nuestra documentaci?n.

* [Formas de contribuir](#ways-to-contribute)
* [Colaborar con GitHub](#contribute-using-github)
* [Colaborar con Git](#contribute-using-git)
* [Uso de Markdown para dar formato al tema](#how-to-use-markdown-to-format-your-topic)
* [Preguntas m?s frecuentes](#faq)
* [Más recursos](#more-resources)

## <a name="ways-to-contribute"></a>Formas de colaborar

Estas son algunas formas en las que puede colaborar con esta documentaci?n:

* Para realizar peque?os cambios en un art?culo, [colabore con GitHub](#contribute-using-github).
* Para realizar cambios importantes o cambios relacionados con el c?digo, [colabore con Git](#contribute-using-git).
* Informar sobre errores de documentación a través de los problemas de depósito.
* Solicite nueva documentaci?n en el sitio de [UserVoice de la plataforma para desarrolladores de Office](http://officespdev.uservoice.com).

## <a name="contribute-using-github"></a>Colaborar con GitHub

Use GitHub para colaborar en esta documentaci?n sin tener que clonar el repositorio en el escritorio. Esta es la forma m?s sencilla de crear una solicitud de incorporaci?n de cambios en este repositorio. Use este m?todo para realizar un cambio secundario que no implique cambios en el c?digo. 

**Nota** Con este m?todo podr? colaborar en un art?culo cada vez.

### <a name="to-contribute-using-github"></a>Para colaborar con GitHub

1. Busque el art?culo en el que quiera colaborar en GitHub.
2. Cuando esté en el artículo de GitHub, inicie sesión en GitHub ([únase a GitHub](https://github.com/join) para obtener una cuenta gratuita).
3. Seleccione el **icono de lápiz** (edite el archivo en la bifurcación de este proyecto) y realice los cambios en la ventana **<> Editar archivo**. 
4. Despl?cese hasta la parte inferior y escriba una descripci?n.
5. Seleccione **Proponer cambio de archivo**>**Crear solicitud de incorporaci?n de cambios**.

Envi? correctamente una solicitud de incorporaci?n de cambios. Las solicitudes de incorporaci?n de cambios se suelen revisar en un plazo de 10 d?as h?biles. 


## <a name="contribute-using-git"></a>Colaborar con Git

Use Git para colaborar con cambios importantes, como:

* Colaboraciones de c?digo.
* Colaboraciones de cambios que afectan al significado.
* Colaboraciones de cambios importantes en el texto.
* Para agregar nuevos temas.

### <a name="to-contribute-using-git"></a>Para colaborar con Git

1. Si no tiene una cuenta de GitHub, cree una en [GitHub](https://github.com/join). 
2. Cuando tenga una cuenta, instale Git en el equipo. Siga los pasos del tutorial [Configurar Git] .
3. Para enviar una solicitud de incorporación de cambios con Git, siga los pasos que se indican en las instrucciones para [usar GitHub, Git y este repositorio](#use-github-git-and-this-repository).
4. Se le pedir? que firme el Contrato de licencia de colaborador si es:

    * Miembro del grupo Microsoft Open Technologies.
    * Un colaborador que no trabaja para Microsoft.

Como miembro de la comunidad, necesita firmar el contrato de licencia de colaborador (CLA) antes de realizar envíos significativos a un proyecto. Solo necesita completar y enviar la documentación una vez. Revise detenidamente el documento. Es posible que un responsable de su empresa tenga que firmar el documento.

El analizador de firma no le concede derechos para confirmar en el repositorio principal, pero significa que los equipos de desarrolladores de Office y Office Developer Content Publishing podrán revisar y aprobar sus contribuciones. Se abona para los envíos.

Las solicitudes de incorporaci?n de cambios se suelen revisar en un plazo de 10 d?as h?biles.

## <a name="use-github-git-and-this-repository"></a>Usar GitHub, Git y este repositorio

**Nota:** La mayoría de la información de esta sección se puede encontrar en artículos de la [Ayuda de GitHub].  Si está familiarizado con Git y GitHub, vaya a la sección **Colaborar y editar contenido** para consultar los detalles del flujo de código y contenido de este repositorio.

### <a name="to-set-up-your-fork-of-the-repository"></a>Para configurar una bifurcaci?n del repositorio

1.  Configure una cuenta de GitHub para colaborar con este proyecto. Si aún no lo ha hecho, vaya a [GitHub](https://github.com/join) ahora y realice este procedimiento.
2.  Instale Git en el equipo. Siga los pasos del tutorial [Configurar Git] .
3.  Cree su propia bifurcación de este repositorio. Para ello, en la parte superior de la página, elija el botón de **horquilla** .
4.  Copie la bifurcación en el equipo. Para hacerlo, abra Git Bash. En el símbolo del sistema, escriba:

        git clone https://github.com/<your user name>/<repo name>.git

    Después, use estos comandos para crear una referencia al repositorio raíz:

        cd <repo name>
        git remote add upstream https://github.com/OfficeDev/<repo name>.git
        git fetch upstream

Enhorabuena. Ya ha configurado el repositorio. No tendrá que repetir estos pasos de nuevo.

### <a name="contribute-and-edit-content"></a>Colaborar y editar contenido

Para que el proceso de colaboraci?n sea lo m?s fluido posible, siga estos pasos.

#### <a name="to-contribute-and-edit-content"></a>Para colaborar y editar contenido

1. Cree una rama.
2. Agregue contenido nuevo o edite contenido existente.
3. Envíe una solicitud de incorporación de cambios al repositorio principal.
4. Elimine la rama.

**Importante** Limite cada rama a un solo artículo o concepto para simplificar el flujo de trabajo y reducir la posibilidad de conflictos de combinación. El contenido adecuado para una nueva rama es:

* Un nuevo art?culo.
* Ediciones de ortograf?a y gram?tica.
* Aplicar un ?nico cambio de formato en un conjunto amplio de art?culos (por ejemplo, aplicar un nuevo pie de p?gina de copyright).

#### <a name="to-create-a-new-branch"></a>Para crear una rama

1.  Abra Git Bash.
2.  En el símbolo del sistema de Git Bash, escriba `git pull upstream master:<new branch name>`. Se creará una rama de forma local que se copiará de la última rama principal de OfficeDev.
3.  En el símbolo del sistema de Git Bash, escriba `git push origin <new branch name>`. De esta forma, se enviará una alerta a GitHub sobre la nueva rama. Después, verá la nueva rama en la bifurcación del repositorio en GitHub.
4.  En el s?mbolo del sistema de Git Bash, escriba `git checkout <new branch name>` para cambiar a la nueva rama.

#### <a name="add-new-content-or-edit-existing-content"></a>Agregar contenido nuevo o editar contenido existente

Abra el repositorio en el equipo local con el Explorador de archivos. Los archivos del repositorio están en `C:\Users\<yourusername>\<repo name>`.

Para editar archivos, ?bralos en un editor de su elecci?n y modif?quelos. Para crear un archivo, use el editor que prefiera y guarde el nuevo archivo en la ubicaci?n adecuada de la copia local del repositorio. Mientras trabaje, guarde su trabajo con frecuencia.

Los archivos de `C:\Users\<yourusername>\<repo name>` son una copia de trabajo de la rama que creó en el repositorio local. Los cambios en esta carpeta no afectan al repositorio local hasta que se confirme un cambio. Para confirmar un cambio en el repositorio local, escriba los siguientes comandos en GitBash:

    git add .
    git commit -v -a -m "<Describe the changes made in this commit>"

El comando `add` agrega los cambios a un área de ensayo como preparación para confirmarlos en el repositorio. El punto después del comando `add` especifica que quiere organizar todos los archivos que agregó o modificó, comprobando las subcarpetas de forma recursiva. (Si no quiere confirmar todos los cambios, puede agregar archivos específicos. También puede deshacer una confirmación. Para obtener ayuda, escriba `git add -help` o `git status`).

El comando `commit` aplica los cambios provisionales en el repositorio. El modificador `-m` quiere decir que proporciona el comentario de la confirmación en la línea de comandos. Los modificadores -v y -a se pueden omitir. El modificador -v muestra resultados detallados del comando, y -a hace lo mismo que ya hizo con el comando add.

Puede confirmar varias veces mientras trabaja, o bien puede esperar y confirmar solo una vez cuando termine.

#### <a name="submit-a-pull-request-to-the-main-repository"></a>Enviar una solicitud de incorporación de cambios al repositorio principal

Cuando termine el trabajo y est? listo para combinarlo con el repositorio principal, siga estos pasos.

#### <a name="to-submit-a-pull-request-to-the-main-repository"></a>Para enviar una solicitud de incorporaci?n de cambios al repositorio principal

1.  En el símbolo del sistema de Git Bash, escriba `git push origin <new branch name>`. En el repositorio local, `origin` hace referencia a su repositorio de GitHub desde el que se clonó el repositorio local. Este comando inserta el estado actual de la nueva rama, incluyendo todas las confirmaciones realizadas en los pasos anteriores, en la bifurcación de GitHub.
2.  En el sitio de GitHub, vaya a la bifurcaci?n de la rama nueva.
3.  Haga clic en el bot?n **Pull Request** (solicitud de incorporaci?n de cambios) en la parte superior de la p?gina.
4.  Compruebe que la rama base sea `OfficeDev/<repo name>@master` y que la rama principal sea `<your username>/<repo name>@<branch name>`.
5.  Haga clic en el bot?n **Update Commit Range** (Actualizar intervalo de confirmaci?n).
6.  Agregue un t?tulo a la solicitud de incorporaci?n de cambios y describa todos los cambios que realiz?.
7.  Envíe la solicitud de incorporación de cambios.

Uno de los administradores del sitio procesará su solicitud de incorporación de cambios. La solicitud de incorporación de cambios se mostrará en el sitio OfficeDev/<repo name>, en la sección de problemas. Cuando se acepte la solicitud de incorporación de cambios, el problema se solucionará.

#### <a name="create-a-new-branch-after-merge"></a>Crear una rama después de una combinación

Despu?s de combinar correctamente una rama (es decir, despu?s de que se acepte la solicitud de incorporaci?n de cambios), no siga trabajando en la rama local. Esto puede provocar conflictos de combinaci?n si env?a otra solicitud de incorporaci?n de cambios. Para realizar otra actualizaci?n, cree otra rama local desde la rama precedente que se combin? correctamente y, despu?s, elimine la rama local inicial.

Por ejemplo, si la rama local X se combin? correctamente en la rama principal OfficeDev/microsoft-graph-docs y quiere realizar actualizaciones adicionales en el contenido que se combin?. Cree una rama local (X2) desde la rama principal OfficeDev/microsoft-graph-docs. Para hacerlo, abra Git Bash y ejecute los comandos siguientes:

    cd microsoft-graph-docs
    git pull upstream master:X2
    git push origin X2

Ahora tiene copias locales (en una nueva rama local) del trabajo que envi? en la rama X. La rama X2 tambi?n contiene todo el trabajo que otros autores combinaron y, por lo tanto, si su trabajo depende del de otros (por ejemplo, im?genes compartidas), estar? disponible en la nueva rama. Puede comprobar que el trabajo anterior (y el de otros) est? en la rama si extrae del repositorio la nueva rama?

    git checkout X2

... y comprueba el contenido. (El comando `checkout` actualiza los archivos de `C:\Users\<yourusername>\microsoft-graph-docs` al estado actual de la rama X2). Una vez que compruebe la rama nueva, puede realizar actualizaciones en el contenido y confirmarlos como de costumbre. Pero para evitar trabajar en la rama combinada (X) por error, es mejor eliminarla (vea la siguiente sección **Eliminar una rama**).

#### <a name="delete-a-branch"></a>Eliminar una rama

Despu?s de combinar correctamente los cambios en el repositorio principal, elimine la rama que us? (ya no la necesita).  Para realizar otras modificaciones, necesita hacerlas en una nueva rama.  

#### <a name="to-delete-a-branch"></a>Para eliminar una rama

1.  En el símbolo del sistema de Git Bash, escriba `git checkout master`. Esto garantiza que no esté en la rama que se eliminará (lo que no está permitido).
2.  Después, en el símbolo del sistema, escriba `git branch -d <branch name>`. Esto elimina la rama en el equipo local solo si se combinó correctamente en el repositorio precedente. (Puede invalidar este comportamiento con la marca `–D`, pero primero asegúrese de que quiere hacerlo).
3.  Por último, escriba `git push origin :<branch name>` en el símbolo del sistema (un espacio antes de los dos puntos y sin espacios después).  Esto eliminará la rama en la bifurcación de github.  

Colabor? correctamente con el proyecto.

## <a name="how-to-use-markdown-to-format-your-topic"></a>Uso de Markdown para dar formato al tema

### <a name="markdown"></a>Markdown

Todos los artículos de este repositorio usan Markdown. Una introducción completa (y lista de todos los la sintaxis) pueden encontrarse en [Bola audacia de fuego - descuento].
 
## <a name="faq"></a>Preguntas m?s frecuentes

### <a name="how-do-i-get-a-github-account"></a>?C?mo se puede obtener una cuenta de GitHub?

Rellene el formulario en [Join GitHub](https://github.com/join) (Unirse a GitHub) para abrir una cuenta gratuita de GitHub. 

### <a name="where-do-i-get-a-contributors-license-agreement"></a>?D?nde se puede conseguir el contrato de licencia de colaborador? 

Se enviar? autom?ticamente un aviso de que necesita firmar el contrato de licencia de colaborador (CLA) si su solicitud de incorporaci?n de cambios requiere uno. 

Como miembro de la comunidad, **debe firmar el contrato de licencia de colaborador (CLA) antes de poder realizar grandes envíos a este proyecto**. Solo necesita completar y enviar la documentación una vez. Revise detenidamente el documento. Es posible que un responsable de su empresa tenga que firmar el documento.

### <a name="what-happens-with-my-contributions"></a>?Qu? ocurre con mis contribuciones?

Al enviar los cambios con una solicitud de incorporaci?n de cambios, nuestro equipo recibir? una notificaci?n y revisar? la solicitud de incorporaci?n de cambios. Recibir? notificaciones sobre la solicitud de incorporaci?n de cambios de GitHub; tambi?n es posible que un miembro de nuestro equipo le env?e una notificaci?n si necesitamos m?s informaci?n. Si se aprueba su solicitud de incorporaci?n de cambios, actualizaremos la documentaci?n. Nos reservamos el derecho de editar el env?o por motivos legales, de estilo o claridad, o por otros problemas.

### <a name="can-i-become-an-approver-for-this-repositorys-github-pull-requests"></a>?Me puedo convertir en un aprobador de solicitudes de incorporaci?n de cambios de GitHub de este repositorio?

Actualmente no permitimos que colaboradores externos aprueben solicitudes de incorporaci?n de cambios en este repositorio.

### <a name="how-soon-will-i-get-a-response-about-my-change-request"></a>?Cu?ndo recibir? una respuesta sobre mi solicitud de cambio?

Las solicitudes de incorporaci?n de cambios se suelen revisar en un plazo de 10 d?as h?biles.


## <a name="more-resources"></a>Más recursos

* Para obtener más información acerca de descuento, vaya al sitio del creador de descuento [Audacia bola de fuego].
* Para obtener más información acerca del uso de Git y depósito, desproteger primero la [Ayuda de depósito].

[GitHub Home]: http://github.com
[Ayuda de GitHub]: http://help.github.com/
[Configurar Git]: https://help.github.com/articles/set-up-git/
[Audacia bola de fuego - descuento]: http://daringfireball.net/projects/markdown/
[Daring Fireball]: http://daringfireball.net/
