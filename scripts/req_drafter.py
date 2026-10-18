#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
req_drafter.py
Valida y normaliza un requerimiento a partir de un JSON base:
1. Elimina contradicciones entre Tipo, Matriz de Impacto y BDD.
2. Autogenera 1 Smoke Test BDD si no se proporcionan escenarios.
3. Asigna códigos correlativos mediante excel_manager.py.
4. Guarda en sprints/sprint_X.json y actualiza el Excel Maestro.
Todas las constantes, defaults y catálogos se cargan desde config.json.
"""

import os
import sys
import json
import argparse
from datetime import datetime

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from excel_manager import get_next_codes, append_requirement_to_excel, load_config


def auto_generate_smoke_test(req_data: dict, config: dict) -> list:
    tipo_upper = req_data.get("tipo", "MEJORA").strip().upper()
    if tipo_upper in ["ÉPICA", "EPICA", "EPIC"]:
        return []

    definicion = req_data.get("definicion", {})
    rol = definicion.get("rol", "Usuario / Administrador")
    modulo = req_data.get("matriz_impacto", {}).get("modulo_principal", "el módulo correspondiente")
    quiero = definicion.get("quiero", req_data.get("titulo", "ejecutar la acción"))
    especificacion = req_data.get("especificacion_requerimiento", "el sistema procesa la solicitud correctamente.")
    responsable_default = config["document_defaults"]["responsable_qa_default"]

    return [
        {
            "id_prueba": "AUTO",
            "dado": f"Que un {rol} se encuentra autenticado en {modulo}",
            "cuando": f"Realiza la acción asociada a: {quiero}",
            "entonces": f"El sistema valida la operación y cumple con: {especificacion}",
            "responsable_qa": responsable_default
        }
    ]


def process_and_register_requirement(input_json_path: str, excel_path: str, sprint_file_path: str, config: dict, update_excel: bool = True):
    with open(input_json_path, "r", encoding="utf-8") as f:
        req = json.load(f)

    req = {k: v for k, v in req.items() if not k.startswith("_")}
    req["tipo"] = req.get("tipo", "MEJORA").strip().upper()
    if not req.get("fecha_solicitada"):
        req["fecha_solicitada"] = datetime.now().strftime("%d-%m-%Y")

    escenarios = req.get("escenarios_qa", [])
    if not escenarios:
        escenarios = auto_generate_smoke_test(req, config)
        req["escenarios_qa"] = escenarios

    next_req_code, next_qa_codes = get_next_codes(excel_path, req["tipo"], len(escenarios), config)

    if req.get("codigo_requerimiento", "AUTO").upper() == "AUTO" or not req.get("codigo_requerimiento"):
        req["codigo_requerimiento"] = next_req_code

    default_qa_resp = config["document_defaults"]["responsable_qa_default"]
    for idx, sc in enumerate(req["escenarios_qa"]):
        if sc.get("id_prueba", "AUTO").upper() == "AUTO" or not sc.get("id_prueba"):
            sc["id_prueba"] = next_qa_codes[idx]
        sc["id_requerimiento_tecnico"] = req["codigo_requerimiento"]
        if not sc.get("responsable_qa"):
            sc["responsable_qa"] = default_qa_resp

    if os.path.exists(sprint_file_path):
        with open(sprint_file_path, "r", encoding="utf-8") as sf:
            sprint_doc = json.load(sf)
    else:
        doc_defs = config["document_defaults"]
        sprint_doc = {
            "folio": doc_defs["folio"],
            "area": doc_defs["area"],
            "clasificacion": doc_defs["clasificacion"],
            "sprint": req.get("sprint", 8),
            "fecha_solicitada": req.get("fecha_solicitada"),
            "autor": doc_defs["autor_default"],
            "requerimientos": []
        }

    sprint_doc["requerimientos"].append(req)
    os.makedirs(os.path.dirname(os.path.abspath(sprint_file_path)), exist_ok=True)
    with open(sprint_file_path, "w", encoding="utf-8") as sf:
        json.dump(sprint_doc, sf, indent=2, ensure_ascii=False)

    if update_excel:
        append_requirement_to_excel(excel_path, req, config)

    return req


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Redactor y Validador de Requerimientos E-SGSI")
    parser.add_argument("--input", required=True, help="Archivo JSON del requerimiento")
    parser.add_argument("--config", default="config.json", help="Ruta al archivo config.json")
    parser.add_argument("--excel", default=None, help="Ruta al archivo Excel Maestro (sobreescribe config.json)")
    parser.add_argument("--sprint-file", default=None, help="Archivo JSON acumulador del Sprint")
    parser.add_argument("--no-excel", action="store_true", help="Modo simulación: no escribe en el Excel")
    args = parser.parse_args()

    cfg = load_config(args.config)
    excel_file = args.excel if args.excel else cfg["paths"]["default_excel"]
    sprint_file = args.sprint_file if args.sprint_file else os.path.join(cfg["paths"]["default_sprints_dir"], "sprint_actual.json")

    drafted = process_and_register_requirement(args.input, excel_file, sprint_file, cfg, update_excel=(not args.no_excel))
    print(f"[OK] Requerimiento registrado como {drafted['codigo_requerimiento']} con QA: {[s['id_prueba'] for s in drafted['escenarios_qa']]}")
