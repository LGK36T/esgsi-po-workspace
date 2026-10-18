# Documentación Técnica Integral - E-SGSI PO Toolkit (FO-PO-002)

Manual de arquitectura, componentes, estructura de datos y procedimientos de automatización para la toma de requerimientos y trazabilidad de control de calidad (QA) de la plataforma **E-SGSI**.

---

## 1. Propósito y Arquitectura General

El **E-SGSI PO Toolkit** es un sistema modular en Python diseñado para estandarizar el ciclo de captura, formalización, codificación correlativa y generación documental de requerimientos de software según la norma documental institucional **FO-PO-002**.

### Principios de Diseño
1. **Excel como Única Fuente de Verdad (*Single Source of Truth*):** Ningún código de requerimiento (`BUG-XXX`, `ME-XXX`, `HU-XXX`, etc.) ni de prueba (`QA-XXX`) se asigna arbitrariamente. El estado correlativo se calcula escaneando el libro maestro de Excel antes de cada inserción.
2. **Desacoplamiento Total vía Configuración:** Rutas, prefijos, estilos tipográficos, paletas de color institucionales y catálogos de negocio residen en [`config.json`](config.json), sin valores cableados (*hardcoded*) en los scripts ejecutables.
3. **Consistencia Tipológica Bidireccional:** El tipo de requerimiento seleccionado rige tanto el prefijo del código asignado como la ficha técnica y la matriz de impacto.
4. **Trazabilidad QA Garantizada:** Todo requerimiento técnico registrado exige la existencia de al menos un escenario de prueba en formato BDD (*Behavior-Driven Development*: Dado / Cuando / Entonces), sincronizado en las hojas de requerimientos y trazabilidad.

---

## 2. Estructura del Proyecto

```text
esgsi-po-workspace/
├── config.json                                          # Variables globales, reglas, catálogos y estilos
├── Requerimientos_y_QA_Normalizados_E-SGSI_Completo.xlsx # Excel Maestro (Única Fuente de Verdad)
├── README.md                                            # Resumen de inicio rápido
├── DOCUMENTACION.md                                     # Manual técnico exhaustivo (este documento)
├── SKILL.md                                             # Directivas de automatización y comandos de skill
├── scripts/
│   ├── excel_manager.py                                 # Lógica de escaneo, asignación y escritura en Excel
│   ├── req_drafter.py                                   # Validador de esquemas, normalizador y acumulador
│   └── word_builder.py                                  # Compilador de documentos Word (.docx) formato FO-PO-002
├── templates/
│   └── formato_base_requerimiento.json                  # Esquema modelo para especificación de requerimientos
├── sprints/
│   └── sprint_8.json                                    # Acumulador de datos del sprint en formato JSON
└── sprint_8/
    └── sprint_8.json                                    # Directorio de trabajo y artefactos del Sprint actual
```

---

## 3. Detalle de Componentes y Módulos

### 3.1. Archivo Central de Configuración (`config.json`)
Centraliza los parámetros del entorno, divididos en cinco secciones principales:

* **`paths`:**
  * `default_excel`: Nombre y ruta del libro Excel maestro.
  * `default_sprints_dir`: Directorio por defecto para almacenar acumuladores de sprint.
  * `default_output_docx`: Nombre del archivo de salida Word compilado.
* **`excel_settings`:**
  * `sheet_requirements`: Nombre exacto de la hoja de requerimientos (`1_Requerimientos_Codificados`).
  * `sheet_qa_traceability`: Nombre de la hoja de pruebas (`2_Trazabilidad_QA_Completa`).
  * `excluded_sheet_keywords`: Hojas excluidas del escaneo correlativo (`Backlog_Sin_Codigo`, `Hallazgos`).
* **`prefixes`:** Mapeo formal de cadenas de tipo a su acrónimo oficial:
  * `BUG` $\rightarrow$ `BUG`
  * `MEJORA` / `ME` $\rightarrow$ `ME`
  * `HISTORIA DE USUARIO` / `HU` $\rightarrow$ `HU`
  * `NUEVA CARACTERÍSTICA` / `NC` $\rightarrow$ `NC`
  * `DEUDA TÉCNICA` / `TD` $\rightarrow$ `TD`
  * `TAREA` / `TASK` $\rightarrow$ `TASK`
  * `VULNERABILIDADES` / `VULN` $\rightarrow$ `VULN`
  * `ÉPICA` / `EPIC` $\rightarrow$ `EPIC`
  * `LEGACY-HOTFIX` $\rightarrow$ `LEGACY-HOTFIX`
* **`document_defaults`:** Metadatos institucionales (Folio `FO-PO-002`, Área `Desarrollo`, Clasificación `Uso Interno`, Autor, Responsable QA por defecto).
* **`catalogs`:** Listas de valores válidos para `tipos_permitidos`, `prioridades` (`Crítica`, `Alta`, `Media`, `Baja`), `solicitantes` y `modulos_comunes` (rutas oficiales de la aplicación y microservicios).
* **`word_styles`:** Definición tipográfica institucional (fuente Calibri, tamaños diferenciados) y paletas de color en código hexadecimal (para relleno XML de tablas) y tuplas RGB (para textos y encabezados).

---

### 3.2. Excel Maestro (`Requerimientos_y_QA_Normalizados_E-SGSI_Completo.xlsx`)
Es el registro central de gobernanza de requerimientos de la plataforma.

* **Hoja `1_Requerimientos_Codificados`:**
  * Almacena cada requerimiento formal con columnas: `Código Actual`, `Código Normalizado / Sugerido`, `Nombre Requerimiento`, `Códigos QA Asociados`, `Tipo`, `Sprint / Origen`, `Estado` y `Observaciones / Correcciones`.
* **Hoja `2_Trazabilidad_QA_Completa`:**
  * Mantiene la correspondencia 1:N entre requerimientos y pruebas con columnas: `Código QA`, `ID Req en Documento/Excel`, `ID Req Correcto`, `Sprint`, `Escenario / Requerimiento` (con desglose BDD Dado/Cuando/Entonces) y `Estado`.
* **Hojas `3_Backlog_Sin_Codigo` y `4_Hallazgos_Drive_OldBacklog`:**
  * Registros históricos mantenidos para fines de auditoría y cruce de información, excluidos automáticamente del cálculo del correlativo activo.

---

### 3.3. Gestor del Excel Maestro (`scripts/excel_manager.py`)
Módulo encargado de interactuar con el libro Excel mediante `openpyxl`.

* **`load_config(config_path)`:** Carga resiliente de configuración con búsqueda automática en directorio local, directorio del script o carpeta superior.
* **`scan_max_codes(excel_path, config)`:**
  * Abre el libro en modo solo lectura (`data_only=True`).
  * Itera por todas las celdas de las hojas activas aplicando expresiones regulares (`\b(PREFIJO)-(\d+)\b` y `\bQA-0*(\d+)\b`).
  * Omite celdas marcadas con observaciones de sugerencia o auditoría para evitar falsos positivos.
  * Retorna el número máximo detectado para cada prefijo y para QA.
* **`get_next_codes(excel_path, tipo_req, num_qa_scenarios, config)`:**
  * Determina el código secuencial inmediato formateado a tres dígitos (ej. `ME-018`) y la lista de correlativos QA requeridos (ej. `['QA-087', 'QA-088']`).
* **`append_requirement_to_excel(excel_path, req_data, config)`:**
  * Registra la fila del requerimiento en `1_Requerimientos_Codificados`.
  * Registra cada escenario BDD en `2_Trazabilidad_QA_Completa`.
  * Guarda y cierra el archivo preservando la integridad del libro.
* **Interfaz de Línea de Comandos (CLI):**
  * `--status`: Imprime en consola la tabla completa de últimos correlativos y siguientes disponibles.
  * `--next <TIPO> [--qa-count N]`: Emite una salida JSON estructurada con el siguiente código del tipo especificado y sus respectivos QA.

---

### 3.4. Validador y Redactor de Requerimientos (`scripts/req_drafter.py`)
Controlador que procesa los archivos de entrada individuales en formato JSON.

* **`auto_generate_smoke_test(req_data, config)`:**
  * Si el requerimiento no incluye escenarios de prueba explícitos (y no es de tipo ÉPICA), construye automáticamente un Smoke Test funcional estructurado a partir del rol, módulo principal, acción y especificación técnica.
* **`process_and_register_requirement(...)`:**
  * Lee el JSON de entrada y limpia claves auxiliares de comentarios (`_instrucciones`, `_opciones_*`).
  * Normaliza el tipo a mayúsculas y asigna fecha actual si no existe.
  * Si `codigo_requerimiento` o `id_prueba` están en `"AUTO"`, consulta a `excel_manager.py` y asigna los números correlativos reales.
  * Agrega el requerimiento validado al archivo acumulador del sprint (`sprint_X.json`).
  * Si el flag `update_excel` está activo (por defecto), persiste inmediatamente los cambios en el libro Excel maestro.
* **Interfaz CLI:**
  * `--input <ruta_json>`: Archivo con la ficha del requerimiento.
  * `--sprint-file <ruta_sprint_json>`: Archivo acumulador de destino.
  * `--no-excel`: Modo simulación para validar datos sin modificar el archivo Excel.

---

### 3.5. Compilador de Documentos Word (`scripts/word_builder.py`)
Motor de renderizado documental basado en `python-docx` que produce el archivo `.docx` oficial listo para entrega o firma.

* **Manipulación de Estilos y XML Directo:**
  * Implementa `set_cell_background` inyectando elementos `w:shd` en el árbol XML (`tcPr`) para garantizar sombreado exacto de celdas con independencia de plantillas externas de Microsoft Office.
  * Aplica tipografía institucional Calibri y códigos de color corporativos definidos en `config.json`.
* **Estructura Documental Generada:**
  1. **Encabezado Institucional (Tabla 2x4):** Metadatos formales (Área, Folio Minuta, Clasificación, Fecha).
  2. **Sección 1 - Título y Resumen Ejecutivo:** Tabla consolidada de requerimientos del Sprint (Sección, Código, Tipo, Título, Códigos QA). Salto de página posterior.
  3. **Sección 2 - Desarrollo Detallado:** Para cada requerimiento registrado:
     * Ficha Técnica (Tabla 5x2): Tipo, Código resaltado en rojo, Prioridad funcional, Componentes afectados, Solicitante.
     * Contexto del problema y necesidad.
     * Contenedores de evidencias gráficas (`add_image_placeholder_box`) delimitados con instrucciones de inserción.
     * **2.X.1 Definición:** Narrativa estructurada en formato historia (*Como [Rol], quiero [acción], para [beneficio]*).
     * **2.X.2 Matriz de Impacto:** Detalle de módulo principal, módulos secundarios, actores impactados y necesidad de diseño de proceso.
     * **2.X.3 Especificación Técnica:** Descripción general del funcionamiento y notas operativas.
     * **2.X.4 Escenarios de Prueba BDD:** Tabla formal de trazabilidad con ID Prueba, ID Técnico, escenario desglosado (Dado / Cuando / Entonces) y Responsable QA.
* **Interfaz CLI:**
  * `--sprint-file <ruta_sprint_json>`: Archivo de datos fuente.
  * `--output <ruta_salida_docx>`: Nombre del archivo `.docx` generado.

---

### 3.6. Plantilla Modelo (`templates/formato_base_requerimiento.json`)
Esqueleto estandarizado para la especificación de cualquier tipo de ítem de trabajo. Permite definir requerimientos completos o parciales, delegando en los scripts la asignación de identificadores y la generación de pruebas complementarias.

---

## 4. Flujo Operativo Estándar

El ciclo para incorporar nuevos requerimientos se ejecuta en cuatro fases:

```
[1. Consulta Estado]       -->  python3 scripts/excel_manager.py --status
          │
[2. Ficha JSON]            -->  Completar templates/formato_base_requerimiento.json con "codigo_requerimiento": "AUTO"
          │
[3. Registro y Sincronía]  -->  python3 scripts/req_drafter.py --input req.json --sprint-file sprint_8/sprint_8.json
          │
[4. Compilación Word]      -->  python3 scripts/word_builder.py --sprint-file sprint_8/sprint_8.json --output sprint_8/FO-PO-002_Sprint_8.docx
```

---

## 5. Directrices de Calidad Documental

1. **Prioridad Funcional:** No se registran Story Points ni fechas tentativas no acordadas en la ficha inicial. Se debe especificar siempre una prioridad válida (`Crítica`, `Alta`, `Media`, `Baja`) y el `Solicitante`.
2. **Normalización de Módulos:** Toda ruta indicada en la matriz de impacto debe corresponder a un módulo canónico del sistema (ej. `/esgsi/doc-flow/`, `/esgsi/governance/`, `/esgsi/risk-management/`, `api-sgsi`).
3. **Persistencia Transaccional:** Cada requerimiento procesado con `req_drafter.py` actualiza en una sola ejecución el archivo JSON del sprint y las dos hojas correspondientes del Excel Maestro.
