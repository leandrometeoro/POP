#!/usr/bin/env python3
from __future__ import annotations
import argparse
from pathlib import Path

from POP.service import generate_pop_odt

def main():
    ap = argparse.ArgumentParser(description="Gera ODT do POP a partir de um BPMN do Camunda")
    ap.add_argument("--bpmn", required=True, help="Caminho para o arquivo .bpmn")
    ap.add_argument("--out-dir", required=False, help="Diretório de saída (opcional). Se ausente, usa o diretório do BPMN.")
    args = ap.parse_args()

    res = generate_pop_odt(bpmn_path=args.bpmn, out_dir=args.out_dir)
    print(f"OK: {res['output_path']}")
    print(f"contexto: {res['context_path']}")

if __name__ == "__main__":
    main()

