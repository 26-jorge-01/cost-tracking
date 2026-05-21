import os
import shutil

# Source directories
correos_dir = r"d:\ICBF\informe\mayo\correos"
core_dir = r"d:\ICBF\cost-tracking"
dest_base_dir = r"d:\ICBF\informe\mayo\evidencias_organizadas"

# Target structure
mappings = {
    "Obligacion_01": [
        # Source (relative to correos_dir or core_dir), Type ('correo' or 'core'), Dest Name
        ("RV_ Alcance No. 2  matriz  Adiciones Integralidad 2026.eml", "correo", "1_1_RV_ Alcance No. 2  matriz  Adiciones Integralidad 2026.eml"),
        ("RV_ Matriz IP-Bolívar -- Validacion Financiera equipo financiero.eml", "correo", "1_2_RV_ Matriz IP-Bolivar -- Validacion Financiera equipo financiero.eml"),
        ("RV_ Solicitud urgente validación Matriz nivelación HCB - Equipo de integralidad.eml", "correo", "1_3_RV_ Solicitud urgente validacion Matriz nivelacion HCB - Equipo de integralidad.eml"),
        ("RV_ VALIDACIÓN MATRIZ ADICIONES HCB NIVELACIÓN CANASTAS RISARALDA.eml", "correo", "1_4_RV_ VALIDACION MATRIZ ADICIONES HCB NIVELACION CANASTAS RISARALDA.eml"),
        ("RV_ Observación matriz Concepto AJUSTE DE COBERTURA Tutela-Risaralda.eml", "correo", "1_5_RV_ Observacion matriz Concepto AJUSTE DE COBERTURA Tutela-Risaralda.eml"),
        ("src/generar_reporte_auditoria_uds.py", "core", "1_6_generar_reporte_auditoria_uds.py"),
        ("src/pipeline_consolidacion/03_auditoria_calidad.py", "core", "1_6_03_auditoria_calidad.py"),
        ("src/pipeline_hcb/03_auditoria_hcb.py", "core", "1_6_03_auditoria_hcb.py")
    ],
    "Obligacion_02": [
        ("RE_ ALCANCE REGIONAL BOYACÁ_  SOLICITUD VALIDACION 1.Boyaca_Matriz_Adición_Nivelacion_Canasta_ HCB_13042026_V2 09-04-2026.eml", "correo", "2_1_RE_ ALCANCE REGIONAL BOYACA_  SOLICITUD VALIDACION 1.Boyaca_Matriz_Adicion_Nivelacion_Canasta_ HCB_13042026_V2 09-04-2026.eml"),
        ("RE_ PARAMETRIZACION ALCANCE MATRIZ SERVICIOS INTEGRALES REGIONAL ANTIOQUIA.eml", "correo", "2_2_RE_ PARAMETRIZACION ALCANCE MATRIZ SERVICIOS INTEGRALES REGIONAL ANTIOQUIA.eml"),
        ("RV_ SUCRE MATRIZ INTEGRALES NIVELACION DE CANASTA v3 - 24032026.eml", "correo", "2_3_RV_ SUCRE MATRIZ INTEGRALES NIVELACION DE CANASTA v3 - 24032026.eml"),
        ("RV_ Matriz IP-Bolívar.eml", "correo", "2_4_RV_ Matriz IP-Bolivar.eml"),
        ("RV_ Validación matriz HCB.eml", "correo", "2_5_RV_ Validacion matriz HCB.eml"),
        ("RV_ Validación Matriz Ajuste cobertura cumplimiento  fallo tutela 2026 00047 CONTRATO 66005152025  Risaralda.eml", "correo", "2_6_RV_ Validacion Matriz Ajuste cobertura cumplimiento  fallo tutela 2026 00047 CONTRATO 66005152025  Risaralda.eml")
    ],
    "Obligacion_03": [
        ("RV_ Solicitud de costeo _ Ajuste cobertura cumplimiento  fallo tutela 2026 00047 CONTRATO 66005152025  Risaralda.eml", "correo", "3_1_RV_ Solicitud de costeo _ Ajuste cobertura cumplimiento  fallo tutela 2026 00047 CONTRATO 66005152025  Risaralda.eml"),
        ("RE_ SOLICITUD - COSTEO Versión final zonificación IP Bolívar.eml", "correo", "3_2_RE_ SOLICITUD - COSTEO Version final zonificacion IP Bolivar.eml"),
        ("RV_ Observaciones Matriz Nivelación HCB Risaralda.eml", "correo", "3_3_RV_ Observaciones Matriz Nivelacion HCB Risaralda.eml")
    ],
    "Obligacion_04": [
        ("RE_ Nivelación matriz de integrales.eml", "correo", "4_1_RE_ Nivelacion matriz de integrales.eml"),
        ("RV_ ALCANCE URGENTE CAUCA_ SOLICITUD DE VALIDACION DE MATRIZ DE ADICION ZONA DESIERTA DE ALIMENTOS .eml", "correo", "4_2_RV_ ALCANCE URGENTE CAUCA_ SOLICITUD DE VALIDACION DE MATRIZ DE ADICION ZONA DESIERTA DE ALIMENTOS .eml")
    ],
    "Obligacion_05": [
        ("src/pipeline_consolidacion/01_consolidacion_inicial.py", "core", "5_1_01_consolidacion_inicial.py"),
        ("src/pipeline_consolidacion/02_refinamiento_copia_segura.py", "core", "5_1_02_refinamiento_copia_segura.py"),
        ("src/pipeline_consolidacion/README.md", "core", "5_1_README.md"),
        ("src/pipeline_hcb/01_consolidacion_hcb.py", "core", "5_2_01_consolidacion_hcb.py"),
        ("src/pipeline_hcb/02_refinamiento_hcb.py", "core", "5_2_02_refinamiento_hcb.py"),
        ("src/pipeline_hcb/README.md", "core", "5_2_README.md"),
        ("src/Analisis_Insumos_Matriz_24_Abril.ipynb", "core", "5_3_Analisis_Insumos_Matriz_24_Abril.ipynb")
    ],
    "Obligacion_06": [
        ("bi/Dashboard integrilidad 2026.pbix", "core", "6_1_Dashboard integrilidad 2026.pbix"),
        ("bi/base_theme.json", "core", "6_1_base_theme.json"),
        ("dashboard_v2.html", "core", "6_2_dashboard_v2.html")
    ],
    "Obligacion_07": [
        ("RE_ Procesos PI invitación publica regional Bolívar - Zonificación IP BOLIVAR.eml", "correo", "7_1_RE_ Procesos PI invitacion publica regional Bolivar - Zonificacion IP BOLIVAR.eml"),
        ("Zonificación Actual.eml", "correo", "7_2_Zonificacion Actual.eml"),
        ("docs/Propuesta Estratégica.docx", "core", "7_3_Propuesta Estrategica.docx")
    ],
    "Obligacion_08": [
        ("pilot_automation_v1/ORQUESTADOR_PILOTO_ICBF.ipynb", "core", "8_1_ORQUESTADOR_PILOTO_ICBF.ipynb"),
        ("pilot_automation_v1/inject_notebook.py", "core", "8_1_inject_notebook.py"),
        ("pilot_automation_v1/update_notebook.py", "core", "8_1_update_notebook.py"),
        ("pilot_automation_v1/generate_regional_templates.py", "core", "8_2_generate_regional_templates.py"),
        ("pilot_automation_v1/generate_regional_templates_com.py", "core", "8_2_generate_regional_templates_com.py"),
        ("pilot_automation_v1/generate_regional_templates_xlwings.py", "core", "8_2_generate_regional_templates_xlwings.py"),
        ("pilot_automation_v1/generate_regional_templates_flexible.py", "core", "8_2_generate_regional_templates_flexible.py"),
        ("pilot_automation_v1/test_formula.py", "core", "8_2_test_formula.py")
    ],
    "Obligacion_09": [
        ("scratch/convert_tolima.py", "core", "9_1_convert_tolima.py"),
        ("scratch/delete_macro_buttons.py", "core", "9_1_delete_macro_buttons.py"),
        ("scratch/inspect_shapes.py", "core", "9_1_inspect_shapes.py"),
        ("scratch/headers_detailed.txt", "core", "9_1_headers_detailed.txt")
    ]
}

print("Starting to copy and rename files...")

# Ensure base directory exists
os.makedirs(dest_base_dir, exist_ok=True)

for obl, files in mappings.items():
    obl_dir = os.path.join(dest_base_dir, obl)
    os.makedirs(obl_dir, exist_ok=True)
    print(f"\nProcessing {obl}...")
    for src_rel, src_type, dest_name in files:
        if src_type == "correo":
            src_path = os.path.join(correos_dir, src_rel)
        else:
            src_path = os.path.join(core_dir, src_rel)
        
        dest_path = os.path.join(obl_dir, dest_name)
        
        if os.path.exists(src_path):
            try:
                shutil.copy2(src_path, dest_path)
                print(f"  Copied: {src_rel} -> {dest_name}")
            except Exception as e:
                print(f"  ERROR copying {src_rel}: {e}")
        else:
            print(f"  WARNING: Source file does not exist: {src_path}")

print("\nFinished organization successfully.")
