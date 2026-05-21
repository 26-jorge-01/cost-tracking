# Pilar 2: Automatización y Distribución de Insumos (Plantillas Regionales)
### 📅 Última Modificación: 20 de Mayo de 2026 (Foco Activo de Desarrollo)

Este pilar automatiza la creación de las plantillas regionales personalizadas e inyecta medidas de seguridad digital para proteger los datos contra manipulaciones accidentales del usuario final.

---

## 1. Contexto de Negocio
El ICBF distribuye mensualmente a las 33 regionales una hoja de cálculo maestra para que planifiquen adiciones de cupos y nivelación de tarifas de canastas.
*   **Problema de Integridad**: Si se les envía el archivo nacional completo, las regionales pueden ver datos confidenciales de otros departamentos o alterar accidentalmente fórmulas críticas del censo, rompiendo la macro nacional.
*   **Solución del Pilar**: Se programó un motor que toma la sábana maestra nacional, extrae solo las filas correspondientes a cada regional, protege las celdas con contraseñas seguras y bloquea las hojas antes de guardarlas por separado.

---

## 2. Catálogo de Desarrollos y Scripts

### A. Partidor de Archivos en Windows (PowerShell)
*   **Archivo**: [Generar_Reportes_ICBF.ps1](file:///d:/ICBF/cost-tracking/Generar_Reportes_ICBF.ps1)
*   **Función**: Script de infraestructura para sistemas Windows. Inicia una instancia invisible de Microsoft Excel en segundo plano, abre la plantilla nacional `.xlsm` y genera los 33 archivos regionales eliminando eficientemente las filas que no pertenecen a la regional destino.
*   **Flujo de Datos**:
    *   *Entrada*: `20260204_Plantilla_Matriz_de_Contratacion - copia.xlsx`
    *   *Salida*: 33 Archivos individuales en `salidas_definitivas/` (Ej. `Reporte_ANTIOQUIA.xlsm`).

### B. Orquestador Basado en Objetos COM (Excel Nativo)
*   **Archivo**: [ORQUESTADOR_PILOTO_ICBF.ipynb](file:///d:/ICBF/cost-tracking/pilot_automation_v1/ORQUESTADOR_PILOTO_ICBF.ipynb)
*   **Función**: Cuaderno Jupyter que gestiona el proceso piloto de distribución. Estructura el directorio del censo creando
<truncated 1532 bytes>