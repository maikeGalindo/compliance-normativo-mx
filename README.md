# Guía de Cumplimiento por Departamento (Compliance Normativo MX)

Aplicación web de una sola página (SPA) para gestionar y dar seguimiento al **cumplimiento normativo por departamento** en organizaciones mexicanas.

## Características principales

- **Panel por departamento**: Visualiza obligaciones aplicables a cada área.
- **Estados de cumplimiento**: Pendiente, Parcial, Cumple, No cumple, N/A.
- **KPIs en tiempo real**: Índice de cumplimiento, pendientes y obligaciones en riesgo.
- **Modo auditoría**: Preparación y seguimiento de auditorías internas.
- **Sincronización opcional**: Con Google Apps Script / Google Sheets.
- **Importación Excel**: Carga masiva desde plantilla (hojas CATALOGO, CONTROL, CONFIG, AUD_FISCAL).
- **Tema claro/oscuro**: Soporte nativo y manual.
- **Almacenamiento local**: Persistencia en el navegador (localStorage).
- **Responsive**: Funciona en escritorio y móvil.
- **Notificaciones de vencimientos**: Avisa de obligaciones próximas a vencer (próximos 7 días).
- **Registro de auditoría**: Log de cambios (quién, cuándo, qué campo, valor anterior/nuevo).
- **PWA + Offline**: Instalable y funciona sin conexión (Service Worker).
- **Código modular**: CSS y JS separados para mejor mantenibilidad.

## Cómo usarlo

1. Abre `index.html` en el navegador o publícalo con GitHub Pages.
2. Carga tu archivo Excel de la guía (versión 3) o conecta con tu Apps Script.
3. Selecciona departamentos en la barra lateral.
4. Actualiza estados, responsables, fechas y evidencias.
5. Exporta o sincroniza cuando lo necesites.

### Publicar en GitHub Pages

1. Ve a **Settings → Pages** del repositorio.
2. En **Source** selecciona la rama `main` y carpeta `/ (root)`.
3. Guarda. En 1-2 minutos estará disponible en:
   `https://maikeGalindo.github.io/compliance-normativo-mx/`

Para dominio personalizado: crea un archivo `CNAME` en la raíz con tu dominio y configura el DNS.

## Estructura del Excel esperado

| Hoja       | Contenido                          |
|------------|------------------------------------|
| CATALOGO   | Catálogo de obligaciones           |
| CONTROL    | Seguimiento y estados              |
| CONFIG     | Configuración de departamentos     |
| AUD_FISCAL | Auditorías fiscales (opcional)     |

## Mejoras aplicadas

- Documentación completa.
- Soporte PWA (`manifest.json` + Service Worker).
- CSS y JS separados (`styles.css`, `app.js`).
- Notificaciones de vencimientos.
- Registro de auditoría de cambios.
- Metadatos y theme-color.

## Próximas mejoras sugeridas

- Plantilla Excel de ejemplo descargable.
- Exportación a PDF/Excel del estado actual.
- Roles y permisos más avanzados (con autenticación).
- Tests automatizados.

## Autor

[maikeGalindo](https://github.com/maikeGalindo)

---

*Herramienta interna de seguimiento de cumplimiento normativo mexicano.*
