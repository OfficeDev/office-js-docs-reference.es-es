# <a name="contribute-to-this-documentation"></a>Contribuir a este contenido

Gracias por su interés en nuestro contenido.

* [Formas de colaborar](#ways-to-contribute)
* [Colaborar con GitHub](#contribute-using-github)
* [Colaborar con Git](#contribute-using-git)
* [Uso de Markdown para dar formato al tema](#how-to-use-markdown-to-format-your-topic)
* [Preguntas más frecuentes](#faq)
* [Más recursos](#more-resources)

## <a name="ways-to-contribute"></a>Formas de contribuir

Estas son algunas formas en las que puede colaborar con esta documentación:

* Para realizar pequeños cambios en un artículo, [colabore con GitHub](#contribute-using-github).
* Para realizar cambios importantes o cambios relacionados con el código, [colabore con Git](#contribute-using-git).
* Informar de errores de documentación a través de problemas de GitHub.
* Solicite nueva documentación en el sitio de [UserVoice de la plataforma para desarrolladores de Office](http://officespdev.uservoice.com).

## <a name="contribute-using-github"></a>Contribuir con GitHub

Use GitHub para colaborar en esta documentación sin tener que clonar el repositorio en el escritorio. Esta es la forma más sencilla de crear una solicitud de incorporación de cambios en este repositorio. Use este método para realizar un cambio secundario que no implique cambios en el código. 

**Nota** Con este método podrá colaborar en un artículo cada vez.

### <a name="to-contribute-using-github"></a>Para colaborar con GitHub

1. Busque el artículo en el que quiera colaborar en GitHub.
2. Cuando esté en el artículo de GitHub, inicie sesión en GitHub ([únase a GitHub](https://github.com/join) para obtener una cuenta gratuita).
3. Seleccione el **icono de lápiz** (edite el archivo en la bifurcación de este proyecto) y realice los cambios en la ventana **<> Editar archivo**. 
4. Desplácese hasta la parte inferior y escriba una descripción.
5. Seleccione **Proponer cambio de archivo**>**Crear solicitud de incorporación de cambios**.

Envió correctamente una solicitud de incorporación de cambios. Las solicitudes de incorporación de cambios se suelen revisar en un plazo de 10 días hábiles. 


## <a name="contribute-using-git"></a>Contribuir con Git

Use Git para colaborar con cambios importantes, como:

* Colaboraciones de código.
* Colaboraciones de cambios que afectan al significado.
* Colaboraciones de cambios importantes en el texto.
* Para agregar nuevos temas.

### <a name="to-contribute-using-git"></a>Para colaborar con Git

1. Si no tiene una cuenta de GitHub, cree una en [GitHub](https://github.com/join). 
2. Cuando tenga una cuenta, instale Git en el equipo. Siga los pasos descritos en el tutorial de [configuración de Git] .
3. Para enviar una solicitud de incorporación de cambios con Git, siga los pasos que se indican en las instrucciones para [usar GitHub, Git y este repositorio](#use-github-git-and-this-repository).
4. Se le pedirá que firme el Contrato de licencia de colaborador si es:

    * Miembro del grupo Microsoft Open Technologies.
    * Un colaborador que no trabaja para Microsoft.

Como miembro de la comunidad, necesita firmar el contrato de licencia de colaborador (CLA) antes de realizar envíos significativos a un proyecto. Solo necesita completar y enviar la documentación una vez. Revise detenidamente el documento. Es posible que un responsable de su empresa tenga que firmar el documento.

Firmar el CLA no le concede los derechos para realizar confirmaciones en el repositorio principal, pero el equipo de desarrolladores de Office y el equipo de publicación de contenido de desarrolladores de Office podrán revisar y aprobar sus contribuciones. Se le abonarán los envíos.

Las solicitudes de incorporación de cambios se suelen revisar en un plazo de 10 días hábiles.

## <a name="use-github-git-and-this-repository"></a>Usar GitHub, Git y este repositorio

**Nota:** La mayoría de la información de esta sección se puede encontrar en artículos de la [Ayuda de GitHub].  Si está familiarizado con Git y GitHub, vaya a la sección sobre cómo **colaborar y editar contenido** para consultar los detalles del flujo de código y contenido de este repositorio.

### <a name="to-set-up-your-fork-of-the-repository"></a>Para configurar una bifurcación del repositorio

1.  Configure una cuenta de GitHub para colaborar con este proyecto. Si aún no lo ha hecho, vaya a [GitHub](https://github.com/join) ahora y realice este procedimiento.
2.  Instale Git en el equipo. Siga los pasos descritos en el tutorial de [configuración de Git] .
3.  Cree su propia bifurcación de este repositorio. Para ello, en la parte superior de la página, elija el **** botón bifurcar.
4.  Copie la bifurcación en el equipo. Para hacerlo, abra Git Bash. En el símbolo del sistema, escriba:

        git clone https://github.com/<your user name>/<repo name>.git

    Después, use estos comandos para crear una referencia al repositorio raíz:

        cd <repo name>
        git remote add upstream https://github.com/OfficeDev/<repo name>.git
        git fetch upstream

Enhorabuena. Ya ha configurado el repositorio. No tendrá que repetir estos pasos de nuevo.

### <a name="contribute-and-edit-content"></a>Colaborar y editar contenido

Para que el proceso de colaboración sea lo más fluido posible, siga estos pasos.

#### <a name="to-contribute-and-edit-content"></a>Para colaborar y editar contenido

1. Cree una rama.
2. Agregue contenido nuevo o edite contenido existente.
3. Envíe una solicitud de incorporación de cambios al repositorio principal.
4. Elimine la rama.

**Importante** Limite cada rama a un solo artículo o concepto para simplificar el flujo de trabajo y reducir la posibilidad de conflictos de combinación. El contenido adecuado para una nueva rama es:

* Un nuevo artículo.
* Ediciones de ortografía y gramática.
* Aplicar un único cambio de formato en un conjunto amplio de artículos (por ejemplo, aplicar un nuevo pie de página de copyright).

#### <a name="to-create-a-new-branch"></a>Para crear una rama

1.  Abra Git Bash.
2.  En el símbolo del sistema de Git Bash, escriba `git pull upstream master:<new branch name>`. Se creará una rama de forma local que se copiará de la última rama principal de OfficeDev.
3.  En el símbolo del sistema de Git Bash, escriba `git push origin <new branch name>`. De esta forma, se enviará una alerta a GitHub sobre la nueva rama. Después, verá la nueva rama en la bifurcación del repositorio en GitHub.
4.  En el símbolo del sistema de Git Bash, escriba `git checkout <new branch name>` para cambiar a la nueva rama.

#### <a name="add-new-content-or-edit-existing-content"></a>Agregar contenido nuevo o editar contenido existente

Abra el repositorio en el equipo local con el Explorador de archivos. Los archivos del repositorio están en `C:\Users\<yourusername>\<repo name>`.

Para editar archivos, ábralos en un editor de su elección y modifíquelos. Para crear un archivo, use el editor que prefiera y guarde el nuevo archivo en la ubicación adecuada de la copia local del repositorio. Mientras trabaje, guarde su trabajo con frecuencia.

Los archivos de `C:\Users\<yourusername>\<repo name>` son una copia de trabajo de la rama que creó en el repositorio local. Los cambios en esta carpeta no afectan al repositorio local hasta que se confirme un cambio. Para confirmar un cambio en el repositorio local, escriba los siguientes comandos en GitBash:

    git add .
    git commit -v -a -m "<Describe the changes made in this commit>"

El comando `add` agrega los cambios a un área de ensayo como preparación para confirmarlos en el repositorio. El punto después del comando `add` especifica que quiere organizar todos los archivos que agregó o modificó, comprobando las subcarpetas de forma recursiva. (Si no quiere confirmar todos los cambios, puede agregar archivos específicos. También puede deshacer una confirmación. Para obtener ayuda, escriba `git add -help` o `git status`).

El comando `commit` aplica los cambios provisionales en el repositorio. El modificador `-m` quiere decir que proporciona el comentario de la confirmación en la línea de comandos. Los modificadores -v y -a se pueden omitir. El modificador -v muestra resultados detallados del comando, y -a hace lo mismo que ya hizo con el comando add.

Puede confirmar varias veces mientras trabaja, o bien puede esperar y confirmar solo una vez cuando termine.

#### <a name="submit-a-pull-request-to-the-main-repository"></a>Enviar una solicitud de incorporación de cambios al repositorio principal

Cuando termine el trabajo y esté listo para combinarlo con el repositorio principal, siga estos pasos.

#### <a name="to-submit-a-pull-request-to-the-main-repository"></a>Para enviar una solicitud de incorporación de cambios al repositorio principal

1.  En el símbolo del sistema de Git Bash, escriba `git push origin <new branch name>`. En el repositorio local, `origin` hace referencia a su repositorio de GitHub desde el que se clonó el repositorio local. Este comando inserta el estado actual de la nueva rama, incluyendo todas las confirmaciones realizadas en los pasos anteriores, en la bifurcación de GitHub.
2.  En el sitio de GitHub, vaya a la bifurcación de la rama nueva.
3.  Haga clic en el botón **Pull Request** (solicitud de incorporación de cambios) en la parte superior de la página.
4.  Compruebe que la rama base sea `OfficeDev/<repo name>@master` y que la rama principal sea `<your username>/<repo name>@<branch name>`.
5.  Haga clic en el botón **Update Commit Range** (Actualizar intervalo de confirmación).
6.  Agregue un título a la solicitud de incorporación de cambios y describa todos los cambios que realizó.
7.  Envíe la solicitud de incorporación de cambios.

Uno de los administradores del sitio procesará su solicitud de incorporación de cambios. La solicitud de incorporación de cambios se mostrará en el sitio OfficeDev/<repo name>, en la sección de problemas. Cuando se acepte la solicitud de incorporación de cambios, el problema se solucionará.

#### <a name="create-a-new-branch-after-merge"></a>Crear una rama después de una combinación

Después de combinar correctamente una rama (es decir, después de que se acepte la solicitud de incorporación de cambios), no siga trabajando en la rama local. Esto puede provocar conflictos de combinación si envía otra solicitud de incorporación de cambios. Para realizar otra actualización, cree otra rama local desde la rama precedente que se combinó correctamente y, después, elimine la rama local inicial.

Por ejemplo, si la rama local X se combinó correctamente en la rama principal OfficeDev/microsoft-graph-docs y quiere realizar actualizaciones adicionales en el contenido que se combinó. Cree una rama local (X2) desde la rama principal OfficeDev/microsoft-graph-docs. Para hacerlo, abra Git Bash y ejecute los comandos siguientes:

    cd microsoft-graph-docs
    git pull upstream master:X2
    git push origin X2

Ahora tiene copias locales (en una nueva rama local) del trabajo que envió en la rama X. La rama X2 también contiene todo el trabajo que otros autores combinaron y, por lo tanto, si su trabajo depende del de otros (por ejemplo, imágenes compartidas), estará disponible en la nueva rama. Puede comprobar que el trabajo anterior (y el de otros) está en la rama si extrae del repositorio la nueva rama…

    git checkout X2

... y comprueba el contenido. (El comando `checkout` actualiza los archivos de `C:\Users\<yourusername>\microsoft-graph-docs` al estado actual de la rama X2). Una vez que compruebe la rama nueva, puede realizar actualizaciones en el contenido y confirmarlos como de costumbre. Pero para evitar trabajar en la rama combinada (X) por error, es mejor eliminarla (vea la siguiente sección **Eliminar una rama**).

#### <a name="delete-a-branch"></a>Eliminar una rama

Después de combinar correctamente los cambios en el repositorio principal, elimine la rama que usó (ya no la necesita).  Para realizar otras modificaciones, necesita hacerlas en una nueva rama.  

#### <a name="to-delete-a-branch"></a>Para eliminar una rama

1.  En el símbolo del sistema de Git Bash, escriba `git checkout master`. Esto garantiza que no esté en la rama que se eliminará (lo que no está permitido).
2.  Después, en el símbolo del sistema, escriba `git branch -d <branch name>`. Esto elimina la rama en el equipo local solo si se combinó correctamente en el repositorio precedente. (Puede invalidar este comportamiento con la marca `–D`, pero primero asegúrese de que quiere hacerlo).
3.  Por último, escriba `git push origin :<branch name>` en el símbolo del sistema (un espacio antes de los dos puntos y sin espacios después).  Esto eliminará la rama en la bifurcación de GitHub.  

Colaboró correctamente con el proyecto.

## <a name="how-to-use-markdown-to-format-your-topic"></a>Cómo usar Markdown para dar formato al tema

### <a name="markdown"></a>Markdown

En todos los artículos de este repositorio se usa Markdown. Puede encontrar una introducción completa (y una lista de toda la sintaxis) en [Daring Fireball-Markdown].
 
## <a name="faq"></a>Preguntas más frecuentes

### <a name="how-do-i-get-a-github-account"></a>¿Cómo se puede obtener una cuenta de GitHub?

Rellene el formulario en [Join GitHub](https://github.com/join) (Unirse a GitHub) para abrir una cuenta gratuita de GitHub. 

### <a name="where-do-i-get-a-contributors-license-agreement"></a>¿Dónde se puede conseguir el contrato de licencia de colaborador? 

Se enviará automáticamente un aviso de que necesita firmar el contrato de licencia de colaborador (CLA) si su solicitud de incorporación de cambios requiere uno. 

Como miembro de la comunidad, **debe firmar el contrato de licencia de colaborador (CLA) antes de poder realizar grandes envíos a este proyecto**. Solo necesita completar y enviar la documentación una vez. Revise detenidamente el documento. Es posible que un responsable de su empresa tenga que firmar el documento.

### <a name="what-happens-with-my-contributions"></a>¿Qué ocurre con mis contribuciones?

Al enviar los cambios con una solicitud de incorporación de cambios, nuestro equipo recibirá una notificación y revisará la solicitud de incorporación de cambios. Recibirá notificaciones sobre la solicitud de incorporación de cambios de GitHub; también es posible que un miembro de nuestro equipo le envíe una notificación si necesitamos más información. Si se aprueba la solicitud de incorporación de documentos, actualizaremos la documentación. Nos reservamos el derecho de editar el envío por motivos legales, de estilo o claridad, o por otros problemas.

### <a name="can-i-become-an-approver-for-this-repositorys-github-pull-requests"></a>¿Me puedo convertir en un aprobador de solicitudes de incorporación de cambios de GitHub de este repositorio?

Actualmente no permitimos que colaboradores externos aprueben solicitudes de incorporación de cambios en este repositorio.

### <a name="how-soon-will-i-get-a-response-about-my-change-request"></a>¿Cuándo recibirá una respuesta sobre mi solicitud de cambio?

Las solicitudes de incorporación de cambios se suelen revisar en un plazo de 10 días hábiles.


## <a name="more-resources"></a>Más recursos

* Para obtener más información sobre Markdown, vaya al sitio de Daring de la [Fireball]del creador de Markdown.
* Para obtener más información acerca del uso de Git y GitHub, consulte primero la [ayuda de github].

[GitHub Home]: http://github.com
[Ayuda de GitHub]: http://help.github.com/
[Configurar git]: https://help.github.com/articles/set-up-git/
[Daring Fireball-Markdown]: http://daringfireball.net/projects/markdown/
[Daring Fireball]: http://daringfireball.net/
