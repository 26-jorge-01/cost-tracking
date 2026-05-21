# Pilar 1: Consolidación y Control de Calidad Regional (Integrales y HCB)
### 📅 Última Modificación: 06 de Mayo de 2026

Este pilar se encarga del procesamiento inicial de los archivos de Excel entregados mensualmente por las 33 regionales del ICBF. El objetivo principal es depurar y unificar estos insumos en una base nacional estructurada y libre de duplicidades.

---

## 1. Contexto de Negocio
Las regionales del ICBF reportan periódicamente sus datos de planeación presupuestal y de cobertura (cupos) en formatos de hoja de cálculo. Sin embargo, estas entregas suelen contener:
*   Filas de totales acumulados a mano que duplican las sumas si se consolidan a ciegas.
*   Retrasos en los envíos de algunas regionales que pueden dejar incompleto el censo nacional.
*   Campos numéricos con formatos inconsistentes o caracteres extraños.

Este pilar automatiza la limpieza y garantiza la integridad de los datos consolidados.

---

## 2. Catálogo de Desarrollos y Scripts

### A. Consolidación de Canastas Integrales
*   **Archivo**: [01_consolidacion_inicial.py](file:///d:/ICBF/cost-tracking/src/pipeline_consolidacion/01_consolidacion_inicial.py)
*   **Función**: Escanea recursivamente las carpetas asignadas a las regionales, identifica las hojas que corresponden a la modalidad "Integrales", filtra celdas vacías y genera el primer borrador consolidado.
*   **Flujo de Datos**:
    *   *Entrada*: Carpetas con plantillas de las 33 regionales.
    *   *Salida*: `consolidacion_matriz_integrales_28042026.xlsx`

### B. Ajuste de Entregas Tardías
*   **Archivo**: [02_refinamiento_copia_segura.py](file:///d:/ICBF/cost-tracking/src/pipeline_consolidacion/02_refinamiento_copia_segura.py)
*   **Función**: Compara las regionales incluidas en la base maestra contra el listado oficial de las 33 regionales del país. Si alguna falta por envío tardío, la recupera de una carpeta de contingencia y genera una versión final unificada con seguridad.
*   **Flujo de Datos**:
    *   *Entrada*: B
<truncated 1753 bytes>