# Guía de Cumplimiento por Departamento

Página única (`index.html`, sin dependencias locales) conectada a Google Sheets vía Apps Script.
Es la versión v3 de trabajo: departamentos con submenú por materia, auditoría interna (con selección de obligaciones,
filtro por materia y agrupación), auditoría financiera y fiscal, descarga en CSV/PDF.

## Archivos
- `index.html` — toda la aplicación (CSS, JS y logos incrustados).
- `apps-script/Code.gs` — pegar en Google Sheets → Extensiones → Apps Script; publicar "Nueva versión".
- `excel/SincronizarGoogle.bas` — macro para el libro .xlsm (URL y token en la hoja SINCRONIZACION B4/B5).
- `INSTRUCCIONES.md` — puesta en marcha.
- `sw.js` — limpia el service worker que dejó la versión anterior; se puede borrar cuando ya nadie lo tenga instalado.

## Publicar
GitHub Pages: Settings → Pages → rama `main`, carpeta `/ (root)`.
Al abrir: ⚙ → Conectar con Google Sheets → pegar URL `/exec` y clave. No guardes la clave en el repositorio.
