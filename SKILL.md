---
name: esgsi-req-po
description: Toolkit de Toma de Requerimientos y Matriz de Trazabilidad (FO-PO-002) para E-SGSI. Valida datos, sincroniza correlativos con el Excel Maestro, genera escenarios BDD y compila minutas en Word (.docx).
---

# Skill: Gestión y Toma de Requerimientos E-SGSI (FO-PO-002)

## Reglas de Oro
1. El Excel Maestro es la Única Fuente de Verdad: Consulta siempre `python3 scripts/excel_manager.py --status` antes de registrar.
2. Cero Contradicciones: El tipo se sincroniza automáticamente en la ficha técnica y la matriz de impacto.
3. Sin Story Points ni Fechas Indefinidas: Pide Prioridad (Crítica, Alta, Media, Baja) y Solicitante.
4. Cobertura QA Obligatoria: Todo requerimiento o corrección debe llevar al menos 1 escenario BDD con código QA-XXX.
5. Recuadros para Imágenes: Si el PO indica que pegará imágenes, incluye el placeholder para generar el recuadro en el Word.

## Comandos Principales
- Ver estado correlativos: `python3 scripts/excel_manager.py --status`
- Redactar y registrar requerimiento: `python3 scripts/req_drafter.py --input req.json --sprint-file sprints/sprint_8.json`
- Generar Word editable: `python3 scripts/word_builder.py --sprint-file sprints/sprint_8.json --output FO-PO-002_Sprint_8.docx`
