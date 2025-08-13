# POP/render/__init__.py
from __future__ import annotations
from pathlib import Path
from io import BytesIO
import zipfile

from lxml import etree as ET

from .fill_first_page_xml import (
    TEXT_NS, OFFICE_NS, NS, _t,
    format_lista_semicolas,
    replace_userfield,
    fill_bookmark_single,
    fill_bookmark_range_same_parent,
)

def _etree_tostring(root) -> bytes:
    return ET.tostring(
        root, encoding="UTF-8", xml_declaration=True, pretty_print=False
    )

def _write_odt_bytes_from_template(zin: zipfile.ZipFile, new_content: bytes) -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w") as zout:
        # 1) mimetype deve ser a primeira entry e sem compressão
        if "mimetype" in zin.namelist():
            info = zipfile.ZipInfo("mimetype")
            info.compress_type = zipfile.ZIP_STORED
            data = zin.read("mimetype")
            zout.writestr(info, data)
        # 2) Demais arquivos (trocando o content.xml)
        for name in zin.namelist():
            if name == "mimetype":
                continue
            if name == "content.xml":
                zout.writestr(name, new_content)
            else:
                zout.writestr(name, zin.read(name))
    return buf.getvalue()

def render_odt(template_path: str, ctx: dict) -> bytes:
    """
    Lê um .odt de template, injeta contexto (user-field-get e bookmarks) e
    retorna os bytes do ODT final.
    """
    tpl = Path(template_path)
    if not tpl.is_file():
        raise FileNotFoundError(f"Template não encontrado: {tpl}")

    with zipfile.ZipFile(tpl, "r") as zin:
        xml_bytes = zin.read("content.xml")
        root = ET.fromstring(xml_bytes)

        # --- Campos simples (user-field-get) ---
        fields = {
            "POP_NOME_PROCESSO": ctx.get("nome_processo", ""),
            "POP_CODIGO":        ctx.get("codigo", ""),
            "POP_VERSAO":        str(ctx.get("versao", "")),
            "POP_SETOR_SUPERIOR": ctx.get("POP_SETOR_SUPERIOR", ctx.get("setor_superior","")),
            "POP_SETOR_EXECUTOR": ctx.get("POP_SETOR_EXECUTOR", ctx.get("setor_executor","")),
            "NVL_GERENCIAL":     ctx.get("NVL_GERENCIAL", ""),
            "NVL_OPERACIONAL":   ctx.get("NVL_OPERACIONAL", ""),
            # Se você criar user-fields para EORG_* no template, já cobre aqui:
            "EORG_SUP":          ctx.get("EORG_SUP",""),
            "EORG_EXEC":         ctx.get("EORG_EXEC",""),
        }
        for k, v in fields.items():
            replace_userfield(root, k, v if v is not None else "")

        # --- Listas (bookmarks BM_OE_LIST / BM_IE_LIST) ---
        oe_lines = format_lista_semicolas(ctx.get("objetivos_estrategicos", []))
        ie_lines = format_lista_semicolas(ctx.get("indicadores_estrategicos", []))
        if oe_lines:
            # tenta marcador simples (<text:bookmark name="BM_OE_LIST"/>)
            n = fill_bookmark_single(root, "BM_OE_LIST", oe_lines)
            if n == 0:
                # tenta intervalo no mesmo parágrafo (<bookmark-start/end>)
                fill_bookmark_range_same_parent(root, "BM_OE_LIST", oe_lines)
        if ie_lines:
            n = fill_bookmark_single(root, "BM_IE_LIST", ie_lines)
            if n == 0:
                fill_bookmark_range_same_parent(root, "BM_IE_LIST", ie_lines)

        new_content = _etree_tostring(root)
        return _write_odt_bytes_from_template(zin, new_content)

