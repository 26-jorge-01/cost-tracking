# Pilar 4: Reconstrucción e Integridad del Censo Nacional (Metaverso)
### 📅 Última Modificación: 23 de Abril de 2026

El **Metaverso** es la base de datos geográfica y operativa unificada de todas las Unidades de Servicio (UDS) del ICBF. Este pilar gestiona la actualización mensual del censo e identifica de forma visual el estado de cada sede de cara a la contratación del nuevo año.

---

## 1. Contexto de Negocio
El inventario físico del ICBF muta constantemente: locales se abren, se trasladan o se cierran por motivos sanitarios o de infraestructura.
*   **Problema de Control**: Al iniciar un nuevo año fiscal, es vital asegurar que no se sigan pagando costos fijos o nóminas de UDS que fueron cerradas físicamente en el censo (Metaverso) del año anterior.
*   **Solución del Pilar**: Se programó un motor de fusión de bases de datos que compara el censo del Metaverso histórico contra la nueva planeación de canastas del año (Integrales, HCB y Alimentos Campesinos). El sistema clasifica el estado de cada UDS y aplica estilos de celda de colores en Excel para agilizar el análisis del auditor.

---

## 2. Catálogo de Desarrollos y Scripts

### A. Fusionador y Formateador Nacional del Censo
*   **Archivo**: [CONSOLIDACION_NACIONAL_METAVERSO.py](file:///d:/ICBF/cost-tracking/src/metaverso/CONSOLIDACION_NACIONAL_METAVERSO.py)
*   **Función**: Script en Python que carga la base histórica, acopla los nuevos insumos de planeación de cupos 2026, y evalúa el estatus de cada registro utilizando la librería `openpyxl`.
*   **Semaforización Visual Automatizada**:
    *   💚 **Verde (C6EFCE)** - Estado `ACTUALIZADO`: La UDS existía en el censo anterior y tiene contrato y cupos asignados para la vigencia actual.
    *   💙 **Azul (DDEBF7)** - Estado `NUEVO`: UDS que se abren por primera vez en la planeación actual (Requiere visita técnica de habilitación).
    *   ❤️ **Rojo (FFC7CE)** - Estado `HISTORICO`: UDS que existían en el censo pero no reportan cupos ni presupue
<truncated 981 bytes>