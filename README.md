# E-SGSI Product Owner Toolkit (FO-PO-002)

Herramientas y scripts de automatización para la toma de requerimientos de software y control de trazabilidad QA de la plataforma E-SGSI, bajo la norma documental institucional FO-PO-002.

---

## Estructura del Repositorio

- `config.json`: Configuración centralizada de variables, prefijos, rutas, catálogos y estilos Word.
- `Requerimientos_y_QA_Normalizados_E-SGSI_Completo.xlsx`: Excel Maestro (Única Fuente de Verdad).
- `DOCUMENTACION.md`: Documentación técnica detallada de cada componente y módulo.
- `SKILL.md`: Directivas operativas para el uso del toolkit.
- `scripts/`:
  - `excel_manager.py`: Consulta de correlativos y sincronización con el Excel Maestro.
  - `req_drafter.py`: Validación de esquemas, asignación de códigos y registro de requerimientos.
  - `word_builder.py`: Compilador de documentos Word (.docx) formato FO-PO-002.
- `templates/`:
  - `formato_base_requerimiento.json`: Plantilla JSON modelo para redactar requerimientos.
- `sprints/`: Directorio de sprints del proyecto.
  - `sprint_8/`: Directorio de trabajo y entregables del Sprint 8 actual (`sprint_8.json` y documentos compilados).

---

## Requisitos de Entorno

Python 3.8+ con las siguientes dependencias:

```bash
pip install openpyxl python-docx
```

---

## Comandos Frecuentes

### 1. Consultar estado y próximos correlativos disponibles
```bash
python3 scripts/excel_manager.py --status
```

### 2. Procesar y registrar un nuevo requerimiento
```bash
python3 scripts/req_drafter.py --input templates/formato_base_requerimiento.json --sprint-file sprints/sprint_8/sprint_8.json
```

### 3. Generar el documento Word (.docx) del Sprint
```bash
python3 scripts/word_builder.py --sprint-file sprints/sprint_8/sprint_8.json --output sprints/sprint_8/FO-PO-002_Sprint_8.docx
```

Para más detalles sobre la arquitectura y la lógica interna de los scripts, consultar [`DOCUMENTACION.md`](DOCUMENTACION.md).
