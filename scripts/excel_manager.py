#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
excel_manager.py
Mantiene el archivo Excel Maestro de Requerimientos y QA como Única Fuente de Verdad.
Todas las variables de entorno, nombres de hojas, prefijos y rutas se cargan desde config.json.
"""

import os
import re
import json
import argparse
import openpyxl


def load_config(config_path: str = "config.json") -> dict:
    if not os.path.exists(config_path):
        # Intentar ruta relativa en directorio superior o del script
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


def get_prefix(tipo_str: str, prefix_map: dict) -> str:
    key = tipo_str.strip().upper()
    return prefix_map.get(key, "REQ")


def scan_max_codes(excel_path: str, config: dict):
    """
    Escanea las hojas principales del Excel para encontrar el número correlativo más alto
    de cada prefijo de requerimiento y de QA según las reglas de config.json.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"No se encontró el archivo Excel maestro: {excel_path}")

    prefix_map = config["prefixes"]
    excluded_keywords = config["excel_settings"]["excluded_sheet_keywords"]

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    max_req = {p: 0 for p in set(prefix_map.values())}
    max_qa = 0

    req_pattern = re.compile(r"\b(BUG|ME|HU|NC|TD|TASK|VULN|EPIC|LEGACY-HOTFIX)-(\d+)\b", re.IGNORECASE)
    qa_pattern = re.compile(r"\bQA-0*(\d+)\b", re.IGNORECASE)

    target_sheets = [s for s in wb.sheetnames if not any(k in s for k in excluded_keywords)]
    for sheet_name in target_sheets:
        ws = wb[sheet_name]
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if cell is None:
                    continue
                val = str(cell)
                if "Sugerido" in val or "Disponible" in val or "Salto" in val:
                    continue
                for m in req_pattern.finditer(val):
                    pref = m.group(1).upper()
                    num = int(m.group(2))
                    if pref in max_req and num > max_req[pref]:
                        max_req[pref] = num
                for m_qa in qa_pattern.finditer(val):
                    q_num = int(m_qa.group(1))
                    if q_num > max_qa:
                        max_qa = q_num

    wb.close()
    return max_req, max_qa


def get_next_codes(excel_path: str, tipo_req: str, num_qa_scenarios: int = 1, config: dict = None):
    if config is None:
        config = load_config()

    max_req, max_qa = scan_max_codes(excel_path, config)
    prefix = get_prefix(tipo_req, config["prefixes"])
    next_req_num = max_req.get(prefix, 0) + 1
    next_req_code = f"{prefix}-{next_req_num:03d}"

    next_qa_codes = [f"QA-{(max_qa + i):03d}" for i in range(1, num_qa_scenarios + 1)]
    return next_req_code, next_qa_codes


def append_requirement_to_excel(excel_path: str, req_data: dict, config: dict = None):
    if config is None:
        config = load_config()

    sheet_req_name = config["excel_settings"]["sheet_requirements"]
    sheet_qa_name = config["excel_settings"]["sheet_qa_traceability"]

    wb = openpyxl.load_workbook(excel_path)

    # 1. Registro en Hoja de Requerimientos
    ws_req = wb[sheet_req_name] if sheet_req_name in wb.sheetnames else wb.active
    qa_list = [sc.get("id_prueba", "") for sc in req_data.get("escenarios_qa", []) if sc.get("id_prueba")]
    qa_str = ", ".join(qa_list) if qa_list else "N/A"

    sprint_label = f"Sprint {req_data.get('sprint', 'N/D')}"
    ws_req.append([
        req_data["codigo_requerimiento"],
        req_data["codigo_requerimiento"],
        req_data["titulo"],
        qa_str,
        req_data["tipo"].upper(),
        sprint_label,
        req_data.get("estado", config["document_defaults"]["estado_default"]),
        f"Solicitante: {req_data.get('solicitante', 'PO')} | Prioridad: {req_data.get('prioridad', config['document_defaults']['prioridad_default'])} | Mod: {req_data.get('matriz_impacto', {}).get('modulo_principal', 'N/D')}"
    ])

    # 2. Registro en Hoja de QA
    if sheet_qa_name in wb.sheetnames:
        ws_qa = wb[sheet_qa_name]
        for sc in req_data.get("escenarios_qa", []):
            resumen_bdd = f"Dado: {sc.get('dado', '')} | Cuando: {sc.get('cuando', '')} | Entonces: {sc.get('entonces', '')}"
            ws_qa.append([
                sc.get("id_prueba", "N/A"),
                req_data["codigo_requerimiento"],
                req_data["codigo_requerimiento"],
                sprint_label,
                f"{req_data['titulo']} ({resumen_bdd})",
                "Registrado"
            ])

    wb.save(excel_path)
    wb.close()
    return req_data["codigo_requerimiento"], qa_list


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Gestor del Excel Maestro de Requerimientos y QA")
    parser.add_argument("--config", default="config.json", help="Ruta al archivo config.json")
    parser.add_argument("--excel", default=None, help="Ruta al archivo Excel maestro (opcional, sobreescribe config.json)")
    parser.add_argument("--status", action="store_true", help="Muestra los últimos códigos registrados en el Excel")
    parser.add_argument("--next", type=str, help="Consulta el siguiente código para un Tipo (ej: BUG, ME, HU)")
    parser.add_argument("--qa-count", type=int, default=1, help="Cantidad de códigos QA a reservar")
    args = parser.parse_args()

    cfg = load_config(args.config)
    excel_file = args.excel if args.excel else cfg["paths"]["default_excel"]

    if args.status:
        max_r, max_q = scan_max_codes(excel_file, cfg)
        print("=== ESTADO ACTUAL LEYENDO EL EXCEL MAESTRO ===")
        for k, v in sorted(max_r.items()):
            print(f"  Último {k:14s}: {k}-{v:03d}  -> Siguiente: {k}-{(v+1):03d}")
        print(f"  Último QA            : QA-{max_q:03d}  -> Siguiente: QA-{(max_q+1):03d}")
    elif args.next:
        r_code, q_codes = get_next_codes(excel_file, args.next, args.qa_count, cfg)
        print(json.dumps({"next_req_code": r_code, "next_qa_codes": q_codes}, ensure_ascii=False))
