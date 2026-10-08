# Adjuntos del hilo — complemento para el Outlook nuevo

Panel lateral que, con un correo abierto, lista **todos los archivos adjuntos de la conversación**
(de todos los mensajes del hilo, incluidos los que enviaste tú), con botón de descarga y
acceso directo al correo que los lleva. Equivale al botón «Explorar archivos» de la app móvil.

Funciona en el Outlook nuevo de Windows, Outlook web y Outlook para Mac. No necesita servidor:
son archivos estáticos (HTML/JS) que leen el buzón a través de Microsoft Graph con tu propia sesión.

## Qué hay en la carpeta

| Archivo | Para qué |
|---|---|
| `manifest.template.xml` | Plantilla del manifiesto del complemento |
| `config.template.js` | Plantilla con el Id de aplicación y de inquilino |
| `configurar.py` | Rellena las plantillas y genera `manifest.xml` y `config.js` |
| `taskpane.html / .css / .js` | El panel |
| `visor.html` | Ventana de vista previa ampliada |
| `msal-browser.min.js` | Biblioteca de autenticación de Microsoft (MSAL) |
| `assets/` | Iconos |

## Puesta en marcha (una sola vez)

### 1. Registrar la aplicación en Entra ID

1. Entra en <https://entra.microsoft.com> con tu cuenta de empresa → **Identidad → Aplicaciones → Registros de aplicaciones → Nuevo registro**.
2. Nombre: `Adjuntos del hilo`. Tipos de cuenta: **solo las cuentas de este directorio**. Sin URI de redirección por ahora. **Registrar**.
3. En **Información general**, copia el **Id. de aplicación (cliente)** y el **Id. de directorio (inquilino)**.
4. **Permisos de API → Agregar un permiso → Microsoft Graph → Permisos delegados** → marca `Mail.Read` (y `User.Read`, que suele venir ya) → **Agregar permisos**.
   - Si ves el botón **Conceder consentimiento de administrador** y puedes pulsarlo, hazlo. Si no, al primer uso Outlook te pedirá el consentimiento; si tu organización exige aprobación del administrador, aparecerá una pantalla para solicitarla.
5. Deja la pestaña **Autenticación** para después del paso 2, porque necesitas saber la URL donde alojas el complemento.

> Si no te deja crear registros de aplicaciones ("no tienes permisos"), tu administrador tiene restringido ese paso y tendrá que crear el registro por ti con estos mismos datos.

### 2. Alojar los archivos en HTTPS

Vale cualquier sitio estático con HTTPS. El más sencillo es **GitHub Pages**:

1. Crea un repositorio (puede ser público; aquí no hay nada secreto, el Id de cliente no es una contraseña).
2. Sube el contenido de esta carpeta a la raíz del repositorio.
3. **Settings → Pages → Deploy from a branch → `main` / root**. Al cabo de un minuto tendrás la URL, del estilo `https://TUUSUARIO.github.io/NOMBRE-REPO`.

### 3. Generar el manifiesto con tus datos

```bash
python configurar.py --url https://TUUSUARIO.github.io/NOMBRE-REPO \
                     --client-id <Id. de aplicación> \
                     --tenant-id <Id. de directorio>
```

Genera `manifest.xml` y `config.js`. Sube ambos al hosting (vuelve a hacer push).

### 4. URIs de redirección en Entra ID

En el registro de la aplicación → **Autenticación → Agregar una plataforma → Aplicación de página única** y añade estos dos URI (el script te los imprime):

- `brk-multihub://TUUSUARIO.github.io` (solo el dominio, sin ruta — es el que usa el Outlook nuevo)
- `https://TUUSUARIO.github.io/NOMBRE-REPO/taskpane.html` (reserva para hosts antiguos)

Guarda.

### 5. Cargar el complemento en Outlook

1. Abre <https://aka.ms/olksideload> con tu sesión de empresa (ventana InPrivate si el navegador tiene otras cuentas).
2. **Mis complementos → Complementos personalizados → Agregar un complemento personalizado → Agregar desde archivo** → elige `manifest.xml` → **Instalar**.
3. En el Outlook nuevo, abre cualquier correo. El botón **Adjuntos del hilo** aparece en la cinta (o dentro de **Aplicaciones** en la barra del mensaje).
4. La primera vez pulsa **Iniciar sesión** y acepta el permiso de lectura del correo.

El panel se puede **anclar** (icono de chincheta): así se queda abierto y se actualiza solo al cambiar de correo.

## Versión 2: redactar, arrastrar, vista previa

- **Al responder o reenviar**, el botón «Adjuntos del hilo» también aparece en la ventana de redacción. Cada archivo tiene un clip **Adjuntar** (y «Adjuntar todo»): lo baja del correo original y lo mete en tu respuesta sin pasar por el disco. Límite de Outlook: 25 MB por archivo.
- **Arrastrar** desde el panel al correo no es posible: Outlook no acepta archivos soltados desde un complemento. Usa el clip desde la ventana de redacción.
- **Vista previa** (icono del ojo) para PDF e imágenes, dentro del panel, con «Ampliar» para abrirla en una ventana grande. Excel, Word, DWG… no tienen previsualizador en el navegador: «Descargar» o «Abrir correo» (el correo original, donde Outlook sí los previsualiza).
- «Abrir correo» está ahora una vez por mensaje, en la cabecera.

**Actualizar de la v1 a la v2**: el manifiesto cambia (añade el modo redacción y el permiso `ReadWriteItem`), así que hay que regenerarlo con `configurar.py`, subir **todos** los archivos a GitHub (incluido el nuevo `visor.html`) y volver a cargar `manifest.xml` en Outlook (Mis complementos → el complemento anterior se sustituye; si da problemas, quítalo primero y vuelve a añadirlo).

## Cómo funciona

1. Toma el correo abierto y pide a Graph su `conversationId`.
2. Pide todos los mensajes del buzón con ese `conversationId` (todas las carpetas).
3. Para cada mensaje con adjuntos, lista sus archivos (sin descargar contenido).
4. «Descargar» baja el archivo por Graph; «Abrir correo» abre el mensaje original en Outlook.

Los enlaces a OneDrive/SharePoint y los correos adjuntos (.eml) se listan, pero para abrirlos se va al correo original.

## Problemas habituales

| Síntoma | Causa / solución |
|---|---|
| `AADSTS50011` / «redirect URI no coincide» | Falta el URI `brk-multihub://<dominio>` o la URL `…/taskpane.html` en Autenticación → SPA. |
| `AADSTS65001` / «consentimiento necesario» | La organización exige que un administrador apruebe `Mail.Read`. Pídeselo o usa el botón de solicitud que aparece. |
| `AADSTS700016` / «aplicación no encontrada» | `client-id` o `tenant-id` mal copiados. Vuelve a ejecutar `configurar.py` y resube `config.js`. |
| El botón no aparece en Outlook | Reinicia el Outlook nuevo tras instalar, o comprueba que `manifest.xml` apunta a la URL correcta y que los iconos cargan en el navegador. |
| «Falta configurar clientId» en el panel | Subiste `config.template.js` pero no `config.js` generado. |
| Descarga no arranca | Abre el correo con «Abrir correo» y descárgalo desde ahí; algunos entornos bloquean descargas desde el panel. |

## Cambiar el color

En `taskpane.css`, variable `--acento` (por defecto verde). Pon ahí el verde corporativo.
