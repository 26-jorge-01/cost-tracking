# Pilar 6: Visualización Ejecutiva e Inteligencia de Negocios
### 📅 Última Modificación: Abril - Mayo de 2026

Este pilar representa la capa final de presentación del proyecto. Traduce millones de filas de datos unificadas y auditadas en gráficos interactivos y tableros interactivos para que la junta directiva y directores del ICBF tomen decisiones inmediatas sobre reasignación de cupos y control de sobrecostos.

---

## 1. Contexto de Negocio
Los directores del ICBF no abren scripts de Python ni analizan cuadernos Jupyter. Necesitan una visualización resumida del desempeño financiero y operativo del país.
*   **Decisiones Clave**:
    *   ¿Cuáles regionales tienen la mayor brecha de cupos (planeado vs atendido)?
    *   ¿Cuáles contratos presentan el mayor riesgo por adiciones presupuestales (superiores al 50% de su valor inicial)?
    *   ¿Cuál es el costo promedio ponderado por cupo a nivel nacional?

Este pilar suministra esos tableros de control ejecutivos interactivos.

---

## 2. Catálogo de Desarrollos y Entregables

### A. Dashboard Ejecutivo Web Interactiva
*   **Archivo**: [dashboard_v2.html](file:///d:/ICBF/cost-tracking/dashboard_v2.html)
*   **Características de Diseño**:
    *   Diseño responsivo moderno basado en **Glassmorphic UI** (fondos translúcidos, degradados elegantes y bordes delgados).
    *   No tiene placeholders: los datos presentados son reales y dinámicos usando la librería `Chart.js`.
    *   **KPIs Expuestos**:
        *   *Total de Cupos Cobertura Nacional*: ~746K niños.
        *   *Costo Promedio Ponderado por Cupo*: $6.26 Millones COP.
        *   *Inversión Total Ejecutada*: $4.67 Billones COP.
    *   **Sección de Alertas de Riesgo**: Enlista de forma proactiva contratos específicos con sobrecostos críticos o sin actas firmadas.
*   **Visualización**: Se puede abrir con un navegador web ordinario sin necesidad de servidores.

### B. Dashboard Corporativo Power BI
*   **Archivo**: [Dashboard integrilidad 2026.pbi
<truncated 1106 bytes>