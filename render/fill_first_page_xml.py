#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import zipfile
from pathlib import Path
import re
from lxml import etree as ET

def _find_paragraph(el):
    """Sobe na árvore até achar o <text:p> que contém o elemento."""
    cur = el
    while cur is not None and cur.tag != _t("p"):
        cur = cur.getparent()
    return cur

def _remove_paragraph_with_userfield(root, name: str) -> int:
    """Remove o <text:p> que contém o user-field `name` (se existir)."""
    path = f".//text:user-field-get[@text:name='{name}']"
    hits = root.xpath(path, namespaces=NS)
    n = 0
    for el in hits:
        par = el
        while par is not None and par.tag != _t("p"):
            par = par.getparent()
        if par is not None and par.getparent() is not None:
            par.getparent().remove(par)
            n += 1
    return n

def _insert_lines_as_paragraphs(par: ET._Element, linhas) -> int:
    """
    Substitui o parágrafo `par` por N parágrafos (um por linha),
    preservando text:style-name do parágrafo original.
    """
    parent = par.getparent()
    if parent is None: 
        return 0
    # preserva o estilo do parágrafo original
    style = par.get(f"{{{TEXT_NS}}}style-name")
    idx = parent.index(par)
    parent.remove(par)

    n = 0
    for ln in linhas:
        p = ET.Element(_t("p"))
        if style:
            p.set(f"{{{TEXT_NS}}}style-name", style)
        span = ET.Element(_t("span"))
        span.text = ln
        p.append(span)
        parent.insert(idx + n, p)
        n += 1
    return n

_VERSAO_NUM_RE = re.compile(r"(\d+)")

def _versao_fem_ordinal(v) -> str:
    s = str(v or "").strip()
    if not s:
        return ""
    # se já vier com ordinal, mantém
    if "ª" in s or "º" in s:
        return s
    m = _VERSAO_NUM_RE.search(s)
    if not m:
        return s
    num = str(int(m.group(1)))  # remove zeros à esquerda (ex.: "01" -> "1")
    return f"{num}ª"

TEXT_NS   = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
NS = {"text": TEXT_NS, "office": OFFICE_NS}
def _t(tag): return f"{{{TEXT_NS}}}{tag}"

# ---------- listas com ; "; e" "." ----------
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

# ---------- XML helpers ----------
def replace_userfield(root, name: str, value: str) -> int:
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

def replace_userfield_cleanup(root, name: str, value: str, remove_prev_break_if_empty: bool = False) -> int:
    """
    Versão 'inteligente' do replace_userfield:
    - Se value != "", substitui normalmente.
    - Se value == "", remove o campo e:
        * remove um <text:line-break/> imediatamente anterior, se houver, e
        * se o parágrafo ficar vazio, remove o parágrafo.
    """
    path = f".//text:user-field-get[@text:name='{name}']"
    hits = root.xpath(path, namespaces=NS)
    changed = 0
    for el in hits:
        par = _find_paragraph(el)
        parent = el.getparent()
        idx = parent.index(el)

        if value:  # substitui normalmente
            span = ET.Element(_t("span"))
            span.text = str(value)
            parent.insert(idx, span)
            parent.remove(el)
            changed += 1
            continue

        # value vazio -> remover campo
        parent.remove(el)

        # se pedimos limpeza, remova quebra de linha imediatamente anterior
        if remove_prev_break_if_empty and idx - 1 >= 0:
            prev = parent[idx - 1]
            if prev.tag == _t("line-break"):
                parent.remove(prev)

        # se o parágrafo ficou vazio, remove parágrafo inteiro
        if par is not None:
            has_text = (par.text or "").strip()
            has_children = len(par) > 0
            if not has_text and not has_children:
                container = par.getparent()
                if container is not None:
                    container.remove(par)
        changed += 1
    return changed

def _insert_lines(parent, idx, linhas):
    inserted = 0
    for i, ln in enumerate(linhas):
        span = ET.Element(_t("span")); span.text = ln
        parent.insert(idx, span); idx += 1; inserted += 1
        if i < len(linhas) - 1:
            lb = ET.Element(_t("line-break"))
            parent.insert(idx, lb); idx += 1; inserted += 1
    return inserted

def fill_bookmark_single(root, name: str, linhas, as_paragraphs=False) -> int:
    path = f".//text:bookmark[@text:name='{name}']"
    hits = root.xpath(path, namespaces=NS)
    n = 0
    for bm in hits:
        par = _find_paragraph(bm)
        if par is None:
            continue
        if as_paragraphs:
            _insert_lines_as_paragraphs(par, linhas)
        else:
            idx = par.index(bm)
            _insert_lines(par, idx, linhas)
            par.remove(bm)
        n += 1
    return n

# def fill_bookmark_single(root, name: str, linhas, as_paragraphs=False) -> int:
#     """
#     Substitui <text:bookmark name=.../>.
#     - se as_paragraphs=True: troca o <text:p> inteiro por N parágrafos (ENTER real).
#     - caso contrário: injeta spans + <text:line-break/> dentro do mesmo parágrafo.
#     """
#     path = f".//text:bookmark[@text:name='{name}']"
#     hits = root.xpath(path, namespaces=NS)
#     n = 0
#     for bm in hits:
#         # acha o <text:p> que contém o bookmark
#         p = bm
#         while p is not None and p.tag != _t("p"):
#             p = p.getparent()
#         if p is None:
#             continue

#         if as_paragraphs:
#             container = p.getparent()
#             if container is None:
#                 continue
#             p_idx = container.index(p)
#             # preserva o estilo do parágrafo original, se houver
#             style_attr = f"{{{TEXT_NS}}}style-name"
#             p_style = p.get(style_attr)

#             # cria um parágrafo por linha DEPOIS do original
#             for i, ln in enumerate(linhas, start=1):
#                 new_p = ET.Element(_t("p"))
#                 if p_style:
#                     new_p.set(style_attr, p_style)
#                 span = ET.Element(_t("span"))
#                 span.text = ln
#                 new_p.append(span)
#                 container.insert(p_idx + i, new_p)

#             # remove o parágrafo original (que só tinha o bookmark)
#             container.remove(p)
#         else:
#             # modo antigo: injeta <text:line-break/> dentro do mesmo <text:p>
#             parent = bm.getparent()
#             idx = parent.index(bm)
#             _insert_lines(parent, idx, linhas)
#             parent.remove(bm)

#         n += 1
#     return n


def fill_bookmark_range_same_parent(root, name: str, linhas, as_paragraphs=False) -> int:
    starts = root.xpath(f".//text:bookmark-start[@text:name='{name}']", namespaces=NS)
    changed = 0
    for st in starts:
        par = _find_paragraph(st)
        if par is None:
            continue
        # se for range, limpamos conteúdo entre start/end
        end = par.xpath(f".//text:bookmark-end[@text:name='{name}']", namespaces=NS)
        if as_paragraphs:
            _insert_lines_as_paragraphs(par, linhas)
        else:
            # apaga tudo entre start e end (se existir) e injeta no lugar
            i0 = par.index(st)
            if end:
                i1 = par.index(end[0])
                for _ in range(i1 - i0 + 1):
                    del par[i0]
            else:
                par.remove(st)
            _insert_lines(par, i0, linhas)
        changed += 1
    return changed

# def fill_bookmark_range_same_parent(root, name: str, linhas, as_paragraphs=False) -> int:
#     """
#     Quando há <text:bookmark-start name=.../> ... <text:bookmark-end name=.../> no MESMO <text:p>.
#     - se as_paragraphs=True: substitui o <text:p> inteiro por N parágrafos (ENTER real).
#     - senão: apaga o conteúdo entre start/end e insere spans + <text:line-break/>.
#     """
#     starts = root.xpath(f".//text:bookmark-start[@text:name='{name}']", namespaces=NS)
#     changed = 0
#     for st in starts:
#         # garante que o end correspondente está no mesmo <text:p>
#         p = st
#         while p is not None and p.tag != _t("p"):
#             p = p.getparent()
#         if p is None:
#             continue
#         ends = [el for el in p.findall(f".//{{{TEXT_NS}}}bookmark-end") if el.get(f"{{{TEXT_NS}}}name") == name]
#         if not ends:
#             continue
#         en = ends[0]

#         if as_paragraphs:
#             container = p.getparent()
#             if container is None:
#                 continue
#             p_idx = container.index(p)
#             style_attr = f"{{{TEXT_NS}}}style-name"
#             p_style = p.get(style_attr)

#             for i, ln in enumerate(linhas, start=1):
#                 new_p = ET.Element(_t("p"))
#                 if p_style:
#                     new_p.set(style_attr, p_style)
#                 span = ET.Element(_t("span"))
#                 span.text = ln
#                 new_p.append(span)
#                 container.insert(p_idx + i, new_p)

#             container.remove(p)
#         else:
#             # remove conteúdo entre start e end dentro do mesmo <text:p> e insere line-breaks
#             i0 = p.index(st)
#             i1 = p.index(en)
#             for i in range(i1, i0 - 1, -1):
#                 p.remove(p[i])
#             _insert_lines(p, i0, linhas)

#         changed += 1
#     return changed

def _serialize(root) -> bytes:
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8")

def _write_odt_like_template(src_zip: zipfile.ZipFile, new_content: bytes, out_path: Path) -> bytes:
    # retorna bytes do ODT final
    mem = Path(out_path).with_suffix(".tmp.bin")  # apenas para nome; gravaremos em memória via writestr
    from io import BytesIO
    buff = BytesIO()
    with zipfile.ZipFile(buff, "w") as zout:
        # exige mimetype como primeira entrada, STORED
        mt = src_zip.read("mimetype") if "mimetype" in src_zip.namelist() else b"application/vnd.oasis.opendocument.text"
        zi = zipfile.ZipInfo("mimetype"); zi.compress_type = zipfile.ZIP_STORED
        zout.writestr(zi, mt)
        for name in src_zip.namelist():
            if name in {"mimetype", "content.xml"}:
                continue
            zout.writestr(name, src_zip.read(name))
        zout.writestr("content.xml", new_content)
    return buff.getvalue()



# def render_odt(template_path: str | Path, ctx: dict) -> bytes:
#     """Gera bytes do ODT a partir do template + contexto já normalizado."""
#     template_path = str(template_path)
#     with zipfile.ZipFile(template_path, "r") as zin:
#         xml = zin.read("content.xml")
#         root = ET.fromstring(xml)

#         # 1) Troca dos user fields (estes nomes devem existir no ODT)
#         fields = {
#             "POP_NOME_PROCESSO":  ctx.get("nome_processo",""),
#             "POP_CODIGO":         ctx.get("codigo",""),
#             "POP_VERSAO":         _versao_fem_ordinal(ctx.get("versao","")),
#             "POP_SETOR_SUPERIOR": ctx.get("POP_SETOR_SUPERIOR",""),
#             "POP_SETOR_EXECUTOR": ctx.get("POP_SETOR_EXECUTOR",""),
#             "NVL_GERENCIAL":      ctx.get("NVL_GERENCIAL",""),
#             "NVL_OPERACIONAL":    ctx.get("NVL_OPERACIONAL",""),
#             "EORG_SUP":           ctx.get("EORG_SUP",""),
#             "EORG_EXEC":          ctx.get("EORG_EXEC",""),
#         }
#         for k, v in fields.items():
#             replace_userfield(root, k, v)

#         # 2) Listas OE/IE (com ENTER real entre itens)
#         oe_lines = format_lista_semicolas(ctx.get("objetivos_estrategicos", []))

#         ie_raw   = ctx.get("indicadores_estrategicos", [])
#         ie_lines = format_lista_semicolas(ie_raw) if ie_raw else ["Não há indicador sensibilizado"]

#         _ = (fill_bookmark_single(root, "BM_OE_LIST", oe_lines, as_paragraphs=True)
#              or fill_bookmark_range_same_parent(root, "BM_OE_LIST", oe_lines, as_paragraphs=True))

#         _ = (fill_bookmark_single(root, "BM_IE_LIST", ie_lines, as_paragraphs=True)
#              or fill_bookmark_range_same_parent(root, "BM_IE_LIST", ie_lines, as_paragraphs=True))

#         new_content = _serialize(root)
#         return _write_odt_like_template(zin, new_content, Path(template_path))

def render_odt(template_path: str | Path, ctx: dict) -> bytes:
    template_path = str(template_path)
    with zipfile.ZipFile(template_path, "r") as zin:
        xml = zin.read("content.xml")
        root = ET.fromstring(xml)

        # 1) User fields "comuns"
        fields = {
            "POP_NOME_PROCESSO":  ctx.get("nome_processo",""),
            "POP_CODIGO":         ctx.get("codigo",""),
            "POP_VERSAO":         _versao_fem_ordinal(ctx.get("versao","")),
            "POP_SETOR_SUPERIOR": ctx.get("POP_SETOR_SUPERIOR",""),
            "POP_SETOR_EXECUTOR": ctx.get("POP_SETOR_EXECUTOR",""),
            "NVL_GERENCIAL":      ctx.get("NVL_GERENCIAL",""),
            "NVL_OPERACIONAL":    ctx.get("NVL_OPERACIONAL",""),
            # (EORG_* ficam fora daqui para tratamento especial)
        }
        for k, v in fields.items():
            replace_userfield(root, k, v)

        # 1.1) EORGs com limpeza de quebra quando vazios
        eorg_sup  = ctx.get("EORG_SUP", "")
        eorg_exec = ctx.get("EORG_EXEC", "")
        replace_userfield_cleanup(root, "EORG_SUP",  eorg_sup,  remove_prev_break_if_empty=True)
        replace_userfield_cleanup(root, "EORG_EXEC", eorg_exec, remove_prev_break_if_empty=True)

        # 2) Listas (ENTER real entre itens)
        oe_lines = format_lista_semicolas(ctx.get("objetivos_estrategicos", []))
        ie_raw   = ctx.get("indicadores_estrategicos", [])
        ie_lines = format_lista_semicolas(ie_raw) if ie_raw else ["Não há indicador sensibilizado"]

        _ = (fill_bookmark_single(root, "BM_OE_LIST", oe_lines, as_paragraphs=True)
             or fill_bookmark_range_same_parent(root, "BM_OE_LIST", oe_lines, as_paragraphs=True))

        _ = (fill_bookmark_single(root, "BM_IE_LIST", ie_lines, as_paragraphs=True)
             or fill_bookmark_range_same_parent(root, "BM_IE_LIST", ie_lines, as_paragraphs=True))

        new_content = _serialize(root)
        return _write_odt_like_template(zin, new_content, Path(template_path))