# Pilar 3: Digitalización y Extracción de Actas Físicas (Cuentas)
### 📅 Última Modificación: 06 de Mayo de 2026

Este pilar funciona como un sistema de extracción automática de datos (Data Scraper) a nivel de archivos. Permite procesar de forma paralela miles de actas de legalización de cuentas mensuales en Excel y contrastarlas con la existencia de su soporte legal físico firmado en formato PDF.

---

## 1. Contexto de Negocio
Mensualmente, los operadores (EAS) envían sus actas de cobro en archivos de Excel con múltiples pestañas de cálculo. Al mismo tiempo, deben subir la versión escaneada y firmada en PDF.
*   **Problema de Auditoría**: Verificar manualmente que la información cargada en el sistema financiero coincide con la del Excel firmado y que existe el soporte PDF para cada mes del año toma semanas para un equipo de 20 analistas.
*   **Solución del Pilar**: Se construyó un extractor automático de alto rendimiento. En un piloto real, el sistema procesó **1,551 carpetas y archivos en solo 13 minutos**, identificando de inmediato qué operadores cobraron dinero pero no adjuntaron su acta digital firmada.

---

## 2. Catálogo de Desarrollos y Scripts

### A. Panel de Mapeo de Celdas (Configuración)
*   **Archivo**: [config.py](file:///d:/ICBF/cost-tracking/src/orm/config.py)
*   **Función**: Diccionario estructurado en Python que almacena las coordenadas exactas de las celdas a extraer en las actas (ej: celda `W29` para N° de Contrato, `W32` para NIT, `W18` para Regional, `U46` para Valor de Cobro, etc.). Si el formato del acta cambia en el futuro, solo se edita este archivo sin alterar el código del motor extractor.

### B. Extractor Multihilo Paralelo
*   **Archivo**: [obtener_datos_unificados_cf_2025.py](file:///d:/ICBF/cost-tracking/src/orm/obtener_datos_unificados_cf_2025.py)
*   **Función**: Motor extractor programado con `ProcessPoolExecutor` para aprovechar todos los núcleos del procesador. 
*   **Gestión de Fórmulas**: Si una celda clave conti
<truncated 1516 bytes>