# Guía de cumplimiento v3: puesta en marcha y uso

## Qué cambió respecto a la versión anterior
- La página ahora **guarda lo que capturas** en tu hoja de Google (pestañas ocultas RESPUESTAS y BITACORA_APP).
- Hay dos claves: **TOKEN** (edición) y **TOKEN_LECTURA** (solo consulta, opcional).
- Nuevos módulos: **Auditoría interna** y **Auditoría financiera y fiscal**. Ambos se pueden descargar en **CSV** o **PDF**.
- Ocho departamentos (REPSE, Sostenibilidad NIS e ISO son departamentos nuevos) con submenú por materia y rol Responsable/Apoyo.

## 1. Google Sheets y Apps Script (actualizar lo que ya tienes)
1. Abre tu hoja de Google → Extensiones → Apps Script. **Reemplaza todo el código** por el de `Code.gs`.
2. Configuración del proyecto → Propiedades del script: conserva `TOKEN`. Si quieres dar acceso de consulta a otras personas, agrega `TOKEN_LECTURA` con otra clave larga.
3. Implementar → **Administrar implementaciones** → lápiz → Versión: **Nueva versión** → Implementar. (No uses "Nueva implementación": cambiaría la URL.)
4. Si es la primera vez, sigue el mismo procedimiento del paso inicial: Aplicación web, ejecutar como **Yo**, acceso **Cualquier persona**.

## 2. Excel (este libro será el definitivo)
1. Abre `GUIA_CUMPLIMIENTO_v3.xlsx` y guárdalo como **Libro habilitado para macros (.xlsm)**. Este será tu archivo base para todas las empresas.
2. **Llevar la macro de tu Excel anterior:** en ese archivo, Alt+F11 → abre el módulo de sincronización → copia la URL `/exec` y el token que ya tenías en las dos primeras líneas. En el libro nuevo, Alt+F11 → Archivo → Importar archivo → `SincronizarGoogle.bas`. Ya no hay constantes dentro del código: pega la URL y el token en la hoja **SINCRONIZACION** (celdas B4 y B5).
3. En CONFIG responde el perfil de la empresa (Sí/No). ISO, NIS y "presta servicios REPSE" vienen en No.
4. Flujo: **antes de editar CONTROL o AUD_FISCAL en Excel, ejecuta `TraerCambiosDeLaApp`** (Alt+F8). Después edita y ejecuta `SincronizarConGoogleSheets`.
5. Para otra empresa: "Guardar como" una copia del .xlsm y cambia solo la URL y el token en SINCRONIZACION (cada empresa con su propia hoja de Google). **Borra el token antes de compartir el libro con alguien.**

## 3. Netlify
Entra a tu sitio → Deploys → arrastra la carpeta `netlify` (conserva la misma URL). Abre la página → ⚙ → Conectar con Google Sheets → pega la URL y la clave.

## Uso diario
- Departamentos (panel izquierdo): al elegir uno se despliegan sus materias y "Apoyo a otras áreas".
- ↻ actualiza; ⚙ abre tema, conexión, carga manual, exportar, tu nombre y borrar datos.
- Cada vista tiene **Descargar CSV** y **Descargar PDF** (el PDF usa el diálogo de impresión: elige "Guardar como PDF").

## Límites a tener presentes
- Si falla la conexión al guardar, la página lo avisa pero **no reintenta sola**: repite el cambio.
- Si dos personas editan lo mismo, gana el último guardado.
- Las auditorías internas viven en tu hoja de Google y en la página; no se pasan al Excel.
- Las evidencias marcadas se guardan en Google Sheets, no en el Excel.
- Las filas "Sugerida" del catálogo y la lista AUD_FISCAL son propuestas: confirma la vigencia antes de usarlas como base de una auditoría.
