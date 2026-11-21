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

## Cómo usarlo

1. Abre `index.html` en el navegador o publícalo con GitHub Pages.
2. Carga tu archivo Excel de la guía (versión 3) o conecta con tu Apps Script.
3. Selecciona departamentos en la barra lateral.
4. Actualiza estados, responsables, fechas y evidencias.
5. Exporta o sincroniza cuando lo necesites.

## Estructura del Excel esperado

| Hoja       | Contenido                          |
|------------|------------------------------------|
| CATALOGO   | Catálogo de obligaciones           |
| CONTROL    | Seguimiento y estados              |
| CONFIG     | Configuración de departamentos     |
| AUD_FISCAL | Auditorías fiscales (opcional)     |

## Mejoras aplicadas

- Documentación completa (este README).
- Soporte PWA básico (`manifest.json`) para instalación en dispositivos.
- Metadatos mejorados.

## Próximas mejoras sugeridas

- Plantilla Excel de ejemplo descargable.
- Exportación a PDF/Excel del estado actual.
- Notificaciones de vencimientos.
- Roles y permisos más granulares.
- Separación de CSS/JS para mayor mantenibilidad.
- Tests automatizados.

## Autor

[maikeGalindo](https://github.com/maikeGalindo)

---

*Herramienta interna de seguimiento de cumplimiento normativo mexicano.*
