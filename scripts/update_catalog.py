#!/usr/bin/env python3
"""
Script de actualización mensual del catálogo de cumplimiento normativo MX.
Ejecutado por GitHub Actions el día 1 de cada mes.
"""
import os
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

REPO_ROOT = Path(__file__).resolve().parent.parent
PLANTILLA = REPO_ROOT / "plantilla_compliance_mx_v5.xlsx"
AUDIT_MD = REPO_ROOT / "AUDITORIA_NORMATIVA_2026.md"

# Catálogo base actualizado a noviembre 2026
# Este diccionario puede ampliarse o leerse de una fuente externa en el futuro
OBLIGACIONES = [
    # Integridad
    ("INT-001", "Integridad", "Política de Integridad (art. 25 LGRA)", "LGRA art. 25", "Manual + Código de Conducta + evidencias de aplicación", "Anual"),
    ("INT-002", "Integridad", "Código de Conducta publicado y socializado", "LGRA art. 25 fr. II", "Constancias de difusión y firma", "Anual"),
    ("INT-003", "Integridad", "Canal de denuncia y proceso disciplinario", "LGRA art. 25 fr. IV", "Evidencia de canal activo + procedimientos", "Anual"),
    # PLD / Acuerdo 115/2026
    ("PLD-001", "Antilavado", "Identificación y registro de Actividades Vulnerables", "LFPIORPI art. 17 + Reglamento 2026 + Acuerdo 115/2026", "Constancia de alta/registro en portal SAT/UIF", "Anual"),
    ("PLD-002", "Antilavado", "Metodología de Enfoque Basado en Riesgos (EBR) documentada", "Acuerdo 115/2026 Cap. II Quáter (exigible 1-mar-2027)", "Documento de metodología EBR + matriz de riesgos", "Anual"),
    ("PLD-003", "Antilavado", "Manual de Políticas Internas PLD/FT (14 contenidos mínimos)", "Acuerdo 115/2026 art. 37 Bis (exigible 1-mar-2027)", "Manual completo firmado + difusión", "Anual"),
    ("PLD-004", "Antilavado", "Identificación de Beneficiario Controlador (umbral 25%)", "LFPIORPI + Acuerdo 115/2026", "Expedientes de clientes con BC documentado", "Por operación"),
    ("PLD-005", "Antilavado", "Programa anual de capacitación PLD/FT", "Acuerdo 115/2026 Transitorio Séptimo (inicia 1-ene-2027)", "Programa + constancias + bitácora", "Anual"),
    ("PLD-006", "Antilavado", "Mecanismos automatizados de monitoreo con perfil transaccional", "Acuerdo 115/2026 (exigible a más tardar 1-jun-2027)", "Evidencia de sistema operativo + reportes", "Continuo"),
    ("PLD-007", "Antilavado", "Identificación y gestión de Personas Políticamente Expuestas (PEP)", "Acuerdo 115/2026 + lista PEP UIF", "Procedimiento + evidencias de consulta", "Por cliente"),
    ("PLD-008", "Antilavado", "Auditoría anual PLD (interna o externa certificada UIF)", "Acuerdo 115/2026 Transitorio Octavo (primer periodo 2028)", "Dictamen de auditoría", "Anual"),
    # Fiscal
    ("FIS-001", "Fiscal", "Contabilidad electrónica, CFDI y control de CSD", "CFF (reforma DOF 07-nov-2025) arts. 17-H, 17-H Bis, 28, 29, 29-A", "Contabilidad electrónica + CFDI válidos", "Mensual"),
    ("FIS-002", "Fiscal", "Declaraciones mensuales, anuales y obligaciones de plataformas", "CFF / LISR / LIVA + art. 30-B CFF", "Acuses de declaraciones", "Mensual/Anual"),
    ("FIS-003", "Fiscal", "Gestión de restricciones y cancelación de CSD (causales 2026)", "CFF art. 17-H Bis (reforma DOF 07-nov-2025)", "Monitoreo de CSD + evidencia de no restricción", "Continuo"),
    ("FIS-004", "Fiscal", "Acceso en tiempo real a sistemas de plataformas digitales", "CFF art. 30-B (vigente desde 1-abr-2026)", "Documentación de cumplimiento art. 30-B si aplica", "Anual"),
    # Laboral
    ("RH-001", "Laboral", "NOM-035 Riesgo Psicosocial (actualizada por violencia de género)", "NOM-035-STPS-2018 + Decreto LFT 15-ene-2026", "Evaluación + política actualizada + plan de acción", "Bienal"),
    ("RH-002", "Laboral", "NOM-037 Teletrabajo - condiciones de seguridad y salud", "NOM-037-STPS-2023", "Política de teletrabajo + lista de verificación", "Anual"),
    ("RH-003", "Laboral", "Contratos, vacaciones dignas y preparación de registro de jornada", "LFT (jornada gradual 2026-2030)", "Contratos vigentes + control de asistencia", "Anual"),
    ("RH-004", "Laboral", "Política y capacitación contra violencia y discriminación hacia mujeres", "LFT arts. 2, 3, 16, 56 (Decreto DOF 15-ene-2026)", "Política + constancias de capacitación", "Anual"),
    ("RH-005", "Laboral", "Registro electrónico de jornada laboral (inicio/fin)", "LFT art. 132 fr. XXXIV (obligatorio desde 2027)", "Sistema de registro electrónico + reportes", "Continuo"),
    # Datos
    ("DAT-001", "Datos Personales", "Aviso de Privacidad actualizado", "LFPDPPP 2025 (DOF 20-mar-2025)", "Aviso de privacidad vigente + evidencia de disponibilidad", "Anual"),
    ("DAT-002", "Datos Personales", "Medidas de seguridad y notificación de vulneraciones", "LFPDPPP 2025 arts. 18-19", "Política de seguridad + registro de incidentes", "Continuo"),
]

DEPTS = [
    "Legal / Compliance / Integridad",
    "Finanzas / Fiscal / Contabilidad",
    "Recursos Humanos / Laboral",
    "Protección de Datos / Privacidad",
    "Operaciones / Compras / Proveedores",
    "Aduanas / Comercio Exterior",
    "Sistemas / TI / Seguridad",
    "Dirección General",
    "Auditoría Interna",
]


def update_plantilla():
    if not PLANTILLA.exists():
        print(f"Plantilla no encontrada en {PLANTILLA}, se omite regeneración completa.")
        return False

    wb = openpyxl.load_workbook(PLANTILLA)
    if "CATALOGO" not in wb.sheetnames:
        print("Hoja CATALOGO no existe, se omite.")
        return False

    ws = wb["CATALOGO"]
    # Actualizar fecha de última revisión
    ws.cell(1, 1).value = f"Última actualización automática: {datetime.now().strftime('%Y-%m-%d %H:%M UTC')}"
    ws.cell(1, 1).font = Font(bold=True, color="1F4E79")

    wb.save(PLANTILLA)
    print(f"Plantilla actualizada: {PLANTILLA}")
    return True


def update_audit_md():
    if not AUDIT_MD.exists():
        print("Auditoría MD no encontrada.")
        return False

    content = AUDIT_MD.read_text(encoding="utf-8")
    today = datetime.now().strftime("%d de %B de %Y")
    marker = "**Última revisión automática:**"
    new_line = f"{marker} {today}\n"
    if marker in content:
        import re
        content = re.sub(r"\*\*Última revisión automática:\*\*.*\n", new_line, content)
    else:
        lines = content.splitlines()
        insert_at = 3 if len(lines) > 3 else 0
        lines.insert(insert_at, new_line.strip())
        content = "\n".join(lines) + "\n"

    AUDIT_MD.write_text(content, encoding="utf-8")
    print(f"Auditoría actualizada: {AUDIT_MD}")
    return True


def main():
    print(f"=== Auditoría normativa mensual - {datetime.now().isoformat()} ===")
    updated = False
    updated |= update_plantilla()
    updated |= update_audit_md()

    if updated:
        print("Cambios aplicados. Listo para commit.")
    else:
        print("No se realizaron cambios.")


if __name__ == "__main__":
    main()
