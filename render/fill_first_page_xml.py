#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Preenche a 1ª página do ODT manipulando XML de forma segura (lxml).

- Substitui <text:user-field-get text:name="...">...</text:user-field-get>
- Injeta listas (com ;, "; e", ".") nos bookmarks BM_OE_LIST e BM_IE_LIST
- Calcula NVL_GERENCIAL / NVL_OPERACIONAL pelas regras combinadas
- Regrava ODT preservando 'mimetype' como primeiro entry, sem compressão

Uso:
  python fill_first_page_xml.py \
    --contexto output/primeira_pagina.contexto.json \
    --odt-in modelo_POP.odt \
    --odt-out output/primeira_pagina.odt
"""
import argparse, json, zipfile, io, re
from pathlib import Path
from lxml import etree as ET

TEXT_NS   = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
NS = {"text": TEXT_NS, "office": OFFICE_NS}

def _t(tag): return f"{{{TEXT_NS}}}{tag}"

# ---------- Regras de listas ----------
def format_lista_semicolas(itens):
    xs = [str(s).strip() for s in (itens or []) if s and str(s).strip()]
    n = len(xs)
    if n == 0:
        return []
    if n == 1:
        return [f"{xs[0]}."]
    if n == 2:
        return [f"{xs[0]}; e", f"{xs[1]}."]
    linhas = [f"{xs[i]};" for i in range(0, n-2)]
    linhas.append(f"{xs[-2]}; e")
    linhas.append(f"{xs[-1]}.")
    return linhas

# ---------- Regras NVL ----------
_IEAPM_CODE_RE = re.compile(r"\(IEAPM-(\d+(?:\.\d+)?)\)", re.IGNORECASE)

def _extrai_codigo_ieapm(rotulo_setor: str):
    if not rotulo_setor: return None
    m = _IEAPM_CODE_RE.search(rotulo_setor)
    return m.group(1) if m else None

def calcula_nvl_gerencial(pop_setor_superior: str) -> str:
    cod = _extrai_codigo_ieapm(pop_setor_superior)
    if cod and cod.split('.')[0] in {"10", "20", "30"}:
        return "Superintendência Responsável"
    return "Setor Responsável"

def calcula_nvl_operacional(pop_setor_executor: str) -> str:
    s = (pop_setor_executor or "").lower()
    if "depto" in s or "depart" in s:
        return "Departamento Responsável"
    if "divisão" in s or "divisao" in s or "div." in s:
        return "Divisão Responsável"
    if "coord." in s or "coordenação" in s or "coordenacao" in s:
        return "Coordenação Responsável"
    if "gerência" in s or "gerencia" in s or "ger." in s:
        return "Gerência Responsável"
    return "Unidade responsável"

# ---------- Operações XML ----------
def replace_userfield(root, name: str, value: str) -> int:
    """Troca <text:user-field-get name=name> por <text:span>valor</text:span>."""
    if value is None: value = ""
    path = f".//text:user-field-get[@text:name='{name}']"
    hits = root.xpath(path, namespaces=NS)
    for el in hits:
        parent = el.getparent()
        idx = parent.index(el)
        span = ET.Element(_t("span"))
        span.text = str(value)
        parent.insert(idx, span)
        parent.remove(el)
    return len(hits)

def _insert_lines(parent, idx, linhas):
    """Insere linhas (span + line-break) a partir do índice idx dentro de um <text:p>."""
    inserted = 0
    for i, ln in enumerate(linhas):
        span = ET.Element(_t("span"))
        span.text = ln
        parent.insert(idx, span)
        idx += 1; inserted += 1
        if i < len(linhas) - 1:
            lb = ET.Element(_t("line-break"))
            parent.insert(idx, lb)
            idx += 1; inserted += 1
    return inserted

def fill_bookmark_single(root, name: str, linhas) -> int:
    """Substitui <text:bookmark name=name/> por linhas com <text:line-break/>."""
    path = f".//text:bookmark[@text:name='{name}']"
    hits = root.xpath(path, namespaces=NS)
    n = 0
    for bm in hits:
        parent = bm.getparent()
        idx = parent.index(bm)
        _insert_lines(parent, idx, linhas)
        parent.remove(bm)
        n += 1
    return n

def fill_bookmark_range_same_parent(root, name: str, linhas) -> int:
    """
    Quando existir <text:bookmark-start name=name/> ... <text:bookmark-end name=name/>
    no MESMO parágrafo, apaga o conteúdo entre eles e injeta as linhas no lugar.
    """
    starts = root.xpath(f".//text:bookmark-start[@text:name='{name}']", namespaces=NS)
    changed = 0
    for st in starts:
        par = st.getparent()
        if par is None: continue
        i0 = par.index(st)
        end = None
        for j in range(i0 + 1, len(par)):
            node = par[j]
            if node.tag == _t("bookmark-end") and node.get(f"{{{TEXT_NS}}}name") == name:
                end = node
                i1 = j
                break
        if end is None:
            continue
        # remove conteúdo entre start e end
        for k in range(i1 - 1, i0, -1):
            par.remove(par[k])
        # injeta linhas na posição do start
        _insert_lines(par, i0, linhas)
        # remove start e end (agora end mudou de índice, mas continua após i0)
        par.remove(end)
        par.remove(st)
        changed += 1
    return changed

def write_odt_like_template(src_zip: zipfile.ZipFile, new_content: bytes, out_path: Path):
    """Regrava ODT preservando 'mimetype' como primeiro (STORED) e demais arquivos."""
    names = src_zip.namelist()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, "w") as zout:
        # 1) mimetype primeiro, STORED
        if "mimetype" in names:
            data = src_zip.read("mimetype")
            info = zipfile.ZipInfo("mimetype")
            info.compress_type = zipfile.ZIP_STORED
            zout.writestr(info, data)
        # 2) demais arquivos
        for name in names:
            if name == "mimetype":
                continue
            if name == "content.xml":
                zout.writestr(name, new_content, compress_type=zipfile.ZIP_DEFLATED)
            else:
                zout.writestr(name, src_zip.read(name), compress_type=zipfile.ZIP_DEFLATED)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--contexto", required=True)
    ap.add_argument("--odt-in", required=True)
    ap.add_argument("--odt-out", required=True)
    args = ap.parse_args()

    ctx = json.load(open(args.contexto, "r", encoding="utf-8"))

    # Dados base esperados (ajuste as chaves se necessário)
    nome_processo  = ctx.get("nome_processo", "")
    codigo         = ctx.get("codigo", "")
    versao         = str(ctx.get("versao", ""))
    setor_superior = ctx.get("setor_superior", "")
    setor_executor = ctx.get("setor_executor", "")
    objetivos      = ctx.get("objetivos_estrategicos", [])
    indicadores    = ctx.get("indicadores_estrategicos", [])

    nvl_g = calcula_nvl_gerencial(setor_superior)
    nvl_o = calcula_nvl_operacional(setor_executor)

    oe_lines = format_lista_semicolas(objetivos)
    ie_lines = format_lista_semicolas(indicadores)

    with zipfile.ZipFile(args.odt_in, "r") as zin:
        content = zin.read("content.xml")
        parser = ET.XMLParser(remove_blank_text=False, recover=False)
        root = ET.fromstring(content, parser)

        # Troca user fields
        fields = {
            "POP_NOME_PROCESSO": nome_processo,
            "POP_CODIGO": codigo,
            "POP_VERSAO": versao,
            "POP_SETOR_SUPERIOR": setor_superior,
            "POP_SETOR_EXECUTOR": setor_executor,
            "NVL_GERENCIAL": nvl_g,
            "NVL_OPERACIONAL": nvl_o,
        }
        n_user = 0
        for k, v in fields.items():
            n_user += replace_userfield(root, k, v)

        # Preenche bookmarks (single e range no mesmo parágrafo)
        n_bm_oe = fill_bookmark_single(root, "BM_OE_LIST", oe_lines)
        n_bm_ie = fill_bookmark_single(root, "BM_IE_LIST", ie_lines)
        # fallback: se existirem start/end no mesmo <text:p>
        if n_bm_oe == 0:
            n_bm_oe = fill_bookmark_range_same_parent(root, "BM_OE_LIST", oe_lines)
        if n_bm_ie == 0:
            n_bm_ie = fill_bookmark_range_same_parent(root, "BM_IE_LIST", ie_lines)

        new_xml = ET.tostring(root, encoding="UTF-8", xml_declaration=True)

        write_odt_like_template(zin, new_xml, Path(args.odt_out))

    print(f"OK: {args.odt_out} | user-fields: {n_user} | bookmark_OE: {n_bm_oe} | bookmark_IE: {n_bm_ie}")

if __name__ == "__main__":
    main()
