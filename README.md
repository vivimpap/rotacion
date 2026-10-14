# VolApp

Versión web de un solo archivo (`index.html`) para usar VolApp desde el navegador
del móvil mientras se construye la app nativa. Funciona sin conexión una vez abierta
(service worker) y guarda los datos en el propio móvil (`localStorage`).

Publicada con GitHub Pages en `https://vivimpap.github.io/rotacion/`.

## Instalar en iPhone
1. Abre el enlace en **Safari**.
2. Botón Compartir → **Añadir a pantalla de inicio**.
3. Ábrela siempre desde ese icono: sus datos son distintos de los de Safari.

## Primer uso
La app empieza vacía: primero pide el nombre del equipo y después añades
la plantilla en Equipo.

## Copias de seguridad
Inicio → **Copia de seguridad** → Enviar copia (WhatsApp, correo, Notas). Para
recuperarla, pega el texto en la misma pantalla y pulsa Restaurar.

## Código
El proyecto completo (incluida la app nativa en Expo, en construcción) vive en el
repo privado `vivimpap/experiments`, carpeta `voley/`. Este repo solo publica la web app.
