# Pilar 5: Conciliación Avanzada y Cruce Sincrónico (NIT Bridge, SECOP y Fuzzy Matching)
### 📅 Última Modificación: 24 de Abril de 2026

Este pilar resuelve los problemas más complejos de cruce de datos y verificación externa del proyecto. Integra algoritmos de similitud lingüística, puentes de conversión de llaves primarias y conexiones con bases de datos del Estado en tiempo real.

---

## 1. El Concepto del "NIT Bridge" (Puente de Datos)
> 💡 **El Desafío de Negocio**: Las matrices de planeación regional utilizan números de contrato de la vigencia 2026, mientras que las bases de auditoría y censo histórico utilizan números de contrato de la vigencia 2025. Al no existir un identificador común, cruzarlos directamente por "Código de Contrato" genera una desconexión total del censo (0% de coincidencia).
>
> 🔧 **La Solución**: Implementamos el **NIT Bridge** utilizando la base *Cuéntame*. 
> 1. Mapeamos cada contrato 2026 hacia el **NIT de la EAS (Contratista)**.
> 2. Con el NIT como llave única, unimos la información del Metaverso histórico y la planeación presupuestal.
> 3. Agrupamos los contratos bajo el mismo contratista en la región, permitiendo una comparación consistente.

---

## 2. Búsqueda por Similitud Lingüística (Fuzzy Matching)
Los nombres de los servicios asignados a los niños varían por tildes, comas o abreviaciones según la regional (Ej. `"CDI - DIRECTO"` vs `"CDI DIRECTO"`). 
Para solucionar esto, el sistema limpia las cadenas de texto (eliminando conectores y diacritics con `nltk` y [CleanData.py](file:///d:/ICBF/cost-tracking/src/cruces_plantilla_matriz_contratacion/utils/CleanData.py)) y calcula la similitud porcentual con la librería `thefuzz`. Si la similitud supera el 85%, empareja los cupos del Metaverso con una de las 7 columnas horizontales de servicios de la base original.

---

## 3. Catálogo de Desarrollos y Scripts

### A. Motor Maestro de Integración (NIT Bridge)
*   **Archivo**: [Analisis_Insumos_Matriz_24
<truncated 2884 bytes>