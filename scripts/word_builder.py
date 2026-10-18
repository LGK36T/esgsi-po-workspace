#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
word_builder.py
Genera el documento Word editable (.docx) de Toma de Requerimientos del Sprint (FO-PO-002).
Todas las tipografías, colores, fondos, márgenes y títulos institucionales se cargan desde config.json.
"""

import os
import json
import argparse
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def load_config(config_path: str = "config.json") -> dict:
    if not os.path.exists(config_path):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        alt_path = os.path.join(script_dir, config_path)
        if os.path.exists(alt_path):
            config_path = alt_path
        else:
            alt_parent = os.path.join(script_dir, "..", config_path)
            if os.path.exists(alt_parent):
                config_path = alt_parent

    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def set_cell_background(cell, fill_hex: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)


def set_cell_bold_text(cell, text: str, font_size: int = 10, color_rgb=None, font_name: str = "Calibri"):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(str(text))
    run.bold = True
    run.font.name = font_name
    run.font.size = Pt(font_size)
    if color_rgb:
        run.font.color.rgb = RGBColor(*color_rgb)


def set_cell_normal_text(cell, text: str, font_size: int = 10, font_name: str = "Calibri"):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(str(text))
    run.font.name = font_name
    run.font.size = Pt(font_size)


def add_image_placeholder_box(doc, caption_text: str, styles: dict):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, styles["colors_hex"]["placeholder_bg"])
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"\n[ ESPACIO PARA EVIDENCIA / DIAGRAMA ]\n{caption_text}\n(Haga clic aquí para insertar o pegar la imagen de referencia)\n")
    run.italic = True
    run.font.name = styles["font_family"]
    run.font.size = Pt(styles["font_size_placeholder"])
    run.font.color.rgb = RGBColor(*styles["colors_rgb"]["placeholder_text"])
    doc.add_paragraph()


def build_word_document(sprint_json_path: str, output_docx_path: str, config: dict):
    with open(sprint_json_path, "r", encoding="utf-8") as f:
        sprint_data = json.load(f)

    doc_defs = config["document_defaults"]
    styles = config["word_styles"]

    sprint_num = sprint_data.get("sprint", 8)
    folio = sprint_data.get("folio", doc_defs["folio"])
    area = sprint_data.get("area", doc_defs["area"])
    clasificacion = sprint_data.get("clasificacion", doc_defs["clasificacion"])
    fecha_solicitada = sprint_data.get("fecha_solicitada", "")
    reqs = sprint_data.get("requerimientos", [])

    doc = docx.Document()

    # Configurar estilo base
    style = doc.styles['Normal']
    style.font.name = styles["font_family"]
    style.font.size = Pt(styles["font_size_normal"])

    # 1. Encabezado Institucional
    header_tbl = doc.add_table(rows=2, cols=4)
    header_tbl.style = "Table Grid"
    header_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr_bg_light = styles["colors_hex"]["header_bg_light"]
    fn = styles["font_family"]
    fs_tbl = styles["font_size_table"]

    set_cell_bold_text(header_tbl.cell(0, 0), "Área", fs_tbl, font_name=fn)
    set_cell_background(header_tbl.cell(0, 0), hdr_bg_light)
    set_cell_normal_text(header_tbl.cell(0, 1), area, fs_tbl, font_name=fn)

    set_cell_bold_text(header_tbl.cell(0, 2), "Folio Minuta", fs_tbl, font_name=fn)
    set_cell_background(header_tbl.cell(0, 2), hdr_bg_light)
    set_cell_normal_text(header_tbl.cell(0, 3), folio, fs_tbl, font_name=fn)

    set_cell_bold_text(header_tbl.cell(1, 0), "Clasificación Documento", fs_tbl, font_name=fn)
    set_cell_background(header_tbl.cell(1, 0), hdr_bg_light)
    set_cell_normal_text(header_tbl.cell(1, 1), clasificacion, fs_tbl, font_name=fn)

    set_cell_bold_text(header_tbl.cell(1, 2), "Fecha Solicitada", fs_tbl, font_name=fn)
    set_cell_background(header_tbl.cell(1, 2), hdr_bg_light)
    set_cell_normal_text(header_tbl.cell(1, 3), fecha_solicitada, fs_tbl, font_name=fn)

    doc.add_paragraph()

    # Título Principal del Sprint
    h1 = doc.add_heading(f"1. Formato Toma de Requerimientos - Sprint {sprint_num}", level=1)
    h1.runs[0].font.name = fn
    h1.runs[0].font.color.rgb = RGBColor(*styles["colors_rgb"]["title_primary"])

    # Tabla Resumen del Sprint
    p_res = doc.add_paragraph()
    r_res = p_res.add_run("Resumen de Requerimientos incluidos en este Sprint:")
    r_res.bold = True
    r_res.font.name = fn

    idx_tbl = doc.add_table(rows=1, cols=5)
    idx_tbl.style = "Table Grid"
    hdr_titles = ["Sección", "Código", "Tipo", "Título del Requerimiento", "Códigos QA"]
    for i, t in enumerate(hdr_titles):
        set_cell_bold_text(idx_tbl.cell(0, i), t, fs_tbl, styles["colors_rgb"]["header_text_white"], font_name=fn)
        set_cell_background(idx_tbl.cell(0, i), styles["colors_hex"]["header_bg_dark"])

    for i, req in enumerate(reqs, start=1):
        row_cells = idx_tbl.add_row().cells
        qa_ids = ", ".join([s.get("id_prueba", "") for s in req.get("escenarios_qa", [])]) or "N/A"
        set_cell_normal_text(row_cells[0], f"2.{i}", fs_tbl, font_name=fn)
        set_cell_bold_text(row_cells[1], req.get("codigo_requerimiento", ""), fs_tbl, font_name=fn)
        set_cell_normal_text(row_cells[2], req.get("tipo", ""), fs_tbl, font_name=fn)
        set_cell_normal_text(row_cells[3], req.get("titulo", ""), fs_tbl, font_name=fn)
        set_cell_normal_text(row_cells[4], qa_ids, fs_tbl, font_name=fn)

    doc.add_page_break()

    # 2. Desarrollo de Requerimientos
    h2_main = doc.add_heading("2. Requerimientos", level=1)
    h2_main.runs[0].font.name = fn
    h2_main.runs[0].font.color.rgb = RGBColor(*styles["colors_rgb"]["title_primary"])

    for idx, req in enumerate(reqs, start=1):
        sec_prefix = f"2.{idx}"
        titulo = req.get("titulo", "Sin Título")
        codigo = req.get("codigo_requerimiento", "PENDIENTE")
        tipo = req.get("tipo", "MEJORA")
        prioridad = req.get("prioridad", doc_defs["prioridad_default"])
        solicitante = req.get("solicitante", "Product Owner")
        matriz = req.get("matriz_impacto", {})
        mod_principal = matriz.get("modulo_principal", "N/D")
        mod_secundarios = matriz.get("modulos_secundarios", [])
        todos_componentes = "\n".join([mod_principal] + [m for m in mod_secundarios if m])

        h_req = doc.add_heading(f"{sec_prefix} {titulo}", level=2)
        h_req.runs[0].font.name = fn
        h_req.runs[0].font.color.rgb = RGBColor(*styles["colors_rgb"]["title_secondary"])

        # Ficha Técnica
        ficha_tbl = doc.add_table(rows=5, cols=2)
        ficha_tbl.style = "Table Grid"
        ficha_rows = [
            ("Tipo Requerimiento", tipo),
            ("Código Requerimiento", codigo),
            ("Prioridad", prioridad),
            ("Componente Afectado", todos_componentes),
            ("Solicitante", solicitante)
        ]
        for r_i, (label, val) in enumerate(ficha_rows):
            set_cell_bold_text(ficha_tbl.cell(r_i, 0), label, fs_tbl, font_name=fn)
            set_cell_background(ficha_tbl.cell(r_i, 0), styles["colors_hex"]["cell_bg_neutral"])
            if label == "Código Requerimiento":
                set_cell_bold_text(ficha_tbl.cell(r_i, 1), val, fs_tbl, styles["colors_rgb"]["code_highlight"], font_name=fn)
            else:
                set_cell_normal_text(ficha_tbl.cell(r_i, 1), val, fs_tbl, font_name=fn)

        doc.add_paragraph()

        if req.get("contexto_problema"):
            p_ctx = doc.add_paragraph(req["contexto_problema"])
            p_ctx.paragraph_format.space_after = Pt(6)

        for img_caption in req.get("placeholders_imagenes", []):
            if img_caption and img_caption.strip():
                add_image_placeholder_box(doc, img_caption, styles)

        # 2.X.1 Definición
        doc.add_heading(f"{sec_prefix}.1 Definición del Requerimiento:", level=3)
        definicion = req.get("definicion", {})
        if isinstance(definicion, dict) and definicion.get("rol"):
            texto_def = f"Como {definicion.get('rol')}, quiero {definicion.get('quiero')}, para {definicion.get('para')}."
        else:
            texto_def = str(definicion)
        doc.add_paragraph(texto_def)

        # 2.X.2 Matriz de Impacto
        doc.add_heading(f"{sec_prefix}.2 Matriz de Impacto", level=3)
        p_tipo = doc.add_paragraph()
        p_tipo.add_run("Tipo: ").bold = True
        p_tipo.add_run(f"{tipo}.")

        p_imp = doc.add_paragraph()
        p_imp.add_run("Modificación impacta a:").bold = True

        mod_list = [("1 [Principal]", mod_principal)]
        for s_idx, sec_mod in enumerate(mod_secundarios, start=2):
            if sec_mod and sec_mod.strip():
                mod_list.append((str(s_idx), sec_mod))

        imp_tbl = doc.add_table(rows=1, cols=2)
        imp_tbl.style = "Table Grid"
        set_cell_bold_text(imp_tbl.cell(0, 0), "Índice", fs_tbl, font_name=fn)
        set_cell_background(imp_tbl.cell(0, 0), hdr_bg_light)
        set_cell_bold_text(imp_tbl.cell(0, 1), "Módulo", fs_tbl, font_name=fn)
        set_cell_background(imp_tbl.cell(0, 1), hdr_bg_light)

        for idx_label, mod_name in mod_list:
            rc = imp_tbl.add_row().cells
            set_cell_normal_text(rc[0], idx_label, fs_tbl, font_name=fn)
            set_cell_normal_text(rc[1], mod_name, fs_tbl, font_name=fn)

        doc.add_paragraph()
        p_apl = doc.add_paragraph()
        p_apl.add_run("Modificación aplicada a: ").bold = True
        p_apl.add_run(f"{matriz.get('modificacion_aplicada_a', definicion.get('rol', 'Administrador'))}.")

        p_proc = doc.add_paragraph()
        p_proc.add_run("Aplica diseño de proceso: ").bold = True
        p_proc.add_run(f"{matriz.get('aplica_diseno_proceso', doc_defs['aplica_diseno_proceso_default'])}.")

        # 2.X.3 Requerimiento
        doc.add_heading(f"{sec_prefix}.3 Requerimiento", level=3)
        p_req_desc = doc.add_paragraph()
        p_req_desc.add_run("Descripción General: ").bold = True
        p_req_desc.add_run(req.get("especificacion_requerimiento", ""))

        if req.get("notas_adicionales"):
            p_nota = doc.add_paragraph()
            p_nota.add_run("NOTA: ").bold = True
            p_nota.add_run(req["notas_adicionales"])

        # 2.X.4 Escenarios de pruebas BDD
        escenarios = req.get("escenarios_qa", [])
        if escenarios:
            doc.add_heading(f"{sec_prefix}.4 Escenarios de pruebas (BDD - Dado / Cuando / Entonces)", level=3)
            qa_tbl = doc.add_table(rows=1, cols=4)
            qa_tbl.style = "Table Grid"
            qa_headers = ["ID Prueba", "ID Requerimiento Técnico", "Escenarios", "Responsable QA"]
            for col_i, h_text in enumerate(qa_headers):
                set_cell_bold_text(qa_tbl.cell(0, col_i), h_text, fs_tbl, font_name=fn)
                set_cell_background(qa_tbl.cell(0, col_i), hdr_bg_light)

            for sc in escenarios:
                q_row = qa_tbl.add_row().cells
                set_cell_bold_text(q_row[0], sc.get("id_prueba", "AUTO"), fs_tbl, font_name=fn)
                set_cell_bold_text(q_row[1], codigo, fs_tbl, font_name=fn)
                
                esc_cell = q_row[2]
                esc_cell.text = ""
                p_d = esc_cell.paragraphs[0]
                p_d.add_run("• Dado: ").bold = True
                p_d.add_run(f"{sc.get('dado', '')}\n")
                p_d.add_run("• Cuando: ").bold = True
                p_d.add_run(f"{sc.get('cuando', '')}\n")
                p_d.add_run("• Entonces: ").bold = True
                p_d.add_run(f"{sc.get('entonces', '')}")

                set_cell_normal_text(q_row[3], sc.get("responsable_qa", doc_defs["responsable_qa_default"]), fs_tbl, font_name=fn)

        doc.add_paragraph()
        doc.add_paragraph("─" * 55)

    os.makedirs(os.path.dirname(os.path.abspath(output_docx_path)), exist_ok=True)
    doc.save(output_docx_path)
    return output_docx_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generador de Documento Word (.docx) de Toma de Requerimientos E-SGSI")
    parser.add_argument("--sprint-file", required=True, help="Archivo JSON del Sprint (ej: sprints/sprint_8.json)")
    parser.add_argument("--config", default="config.json", help="Ruta al archivo config.json")
    parser.add_argument("--output", default=None, help="Ruta de salida del archivo .docx (sobreescribe config.json)")
    args = parser.parse_args()

    cfg = load_config(args.config)
    out_file = args.output if args.output else cfg["paths"]["default_output_docx"]

    out = build_word_document(args.sprint_file, out_file, cfg)
    print(f"[OK] Documento Word generado exitosamente en: {out}")
