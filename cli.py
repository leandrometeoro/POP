#!/usr/bin/env python3
from __future__ import annotations
import argparse
from pathlib import Path
from POP.service import generate_pop_odt

def main():
    ap = argparse.ArgumentParser(description="Gera ODT do POP a partir de um BPMN.")
    ap.add_argument("--bpmn", required=True, help="Caminho para o .bpmn do Camunda")
    ap.add_argument("--out", required=False, help="Pasta de saída (opcional).")
    args = ap.parse_args()

    final = generate_pop_odt(args.bpmn, args.out)
    print(f"OK: {final}")

if __name__ == "__main__":
    main()

