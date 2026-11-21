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
2. Carga tu archivo Excel de la guía (**plantilla_compliance_mx_v5.xlsx** recomendada) o conecta con tu Apps Script.
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

## Plantilla actualizada (v5)

Descarga `plantilla_compliance_mx_v5.xlsx` del repositorio. Incluye obligaciones **al día** a noviembre 2026 para:

- **PLD / LFPIORPI + Acuerdo 115/2026** (EBR, Manual, Beneficiario Controlador 25%, PEP, capacitación, monitoreo automatizado, auditoría).
- **Fiscal** (reformas CFF 2025-2026: CSD, art. 30-B plataformas digitales).
- **Laboral** (NOM-035 actualizada por violencia de género, NOM-037, registro electrónico de jornada desde 2027).

## Auditoría normativa

Consulta `AUDITORIA_NORMATIVA_2026.md` para el detalle de normas, reformas y estado de cumplimiento por módulo.

## Mejoras aplicadas

- Documentación completa.
- Soporte PWA (`manifest.json` + Service Worker).
- CSS y JS separados (`styles.css`, `app.js`).
- Notificaciones de vencimientos.
- Registro de auditoría de cambios.
- Metadatos y theme-color.
- Plantilla Excel v5 con catálogo actualizado a normas 2026.
- Auditoría normativa detallada.

## Próximas mejoras sugeridas

- Exportación a PDF/Excel del estado actual.
- Roles y permisos más avanzados (con autenticación).
- Tests automatizados.

## Autor

[maikeGalindo](https://github.com/maikeGalindo)

---

*Herramienta interna de seguimiento de cumplimiento normativo mexicano.*
