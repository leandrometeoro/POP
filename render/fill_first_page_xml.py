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



def _serialize(root) -> bytes:
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8")

def _write_odt_like_template(src_zip: zipfile.ZipFile, files_to_update: dict, out_path: Path) -> bytes:
    """
    Grava um novo arquivo ODT baseado em um template, atualizando os arquivos
    cujos conteúdos são passados no dicionário `files_to_update`.
    """
    from io import BytesIO
    buff = BytesIO()
    with zipfile.ZipFile(buff, "w") as zout:
        # Exige mimetype como primeira entrada, sem compressão
        mt = src_zip.read("mimetype") if "mimetype" in src_zip.namelist() else b"application/vnd.oasis.opendocument.text"
        zi = zipfile.ZipInfo("mimetype")
        zi.compress_type = zipfile.ZIP_STORED
        zout.writestr(zi, mt)

        # Copia todos os arquivos do original, exceto os que vamos atualizar
        files_to_ignore = {"mimetype"} | files_to_update.keys()
        for name in src_zip.namelist():
            if name in files_to_ignore:
                continue
            zout.writestr(name, src_zip.read(name))

        # Escreve os arquivos modificados
        for name, content in files_to_update.items():
            zout.writestr(name, content)
            
    return buff.getvalue()

import ast # Garanta que esta linha está no topo do seu arquivo

def processa_lista_aninhada(itens_brutos: list) -> list:
    """
    Processa uma lista do contexto onde apenas o último item pode
    ser uma sub-lista representada como string.
    """
    if not itens_brutos:
        return []
    
    # Pega todos os itens, exceto o último, que são sempre simples.
    resultados = [str(item).strip() for item in itens_brutos[:-1]]
    
    # Agora, trata apenas o último item, que pode ser especial.
    ultimo_item_str = str(itens_brutos[-1]).strip()
    
    # Tenta interpretar o último item como uma literal Python (ex: "['a', 'b']")
    if ultimo_item_str.startswith('[') and ultimo_item_str.endswith(']'):
        try:
            sub_lista = ast.literal_eval(ultimo_item_str)
            if isinstance(sub_lista, list):
                resultados.extend([str(sub).strip() for sub in sub_lista])
            else:
                # Não era uma lista, adiciona a string como está
                resultados.append(ultimo_item_str)
        except (ValueError, SyntaxError):
            # Se não for uma lista válida, adiciona a string como está
            resultados.append(ultimo_item_str)
    else:
        # Se não parece uma lista, é um item simples
        resultados.append(ultimo_item_str)
            
    return resultados

def insere_lista_como_bullets(root, bookmark_name: str, lista_de_itens: list):
    """
    Encontra um marcador de texto e o substitui por uma lista de bullets,
    gerando a estrutura XML correta <text:list> e <text:list-item>.
    """
    path = f".//*[self::text:bookmark or self::text:bookmark-start][@text:name='{bookmark_name}']"
    hits = root.xpath(path, namespaces=NS)
    for bm in hits:
        par_original = _find_paragraph(bm)
        if par_original is None: continue

        parent = par_original.getparent()
        if parent is None: continue

        # Guarda a posição e o estilo do parágrafo original
        idx = parent.index(par_original)
        p_style = par_original.get(f"{{{TEXT_NS}}}style-name")

        # Cria o elemento principal da lista <text:list>
        lista_xml = ET.Element(_t("list"))
        
        # Itera sobre os itens para criar a estrutura de cada bullet
        for item_text in lista_de_itens:
            if not item_text: continue # Pula itens vazios

            # Cria o item da lista <text:list-item>
            list_item_xml = ET.SubElement(lista_xml, _t("list-item"))
            
            # Cria o parágrafo <text:p> dentro do item da lista
            p = ET.SubElement(list_item_xml, _t("p"))
            if p_style:
                p.set(f"{{{TEXT_NS}}}style-name", p_style)
            
            # Adiciona o texto dentro do parágrafo
            span = ET.SubElement(p, _t("span"))
            span.text = item_text

        # Remove o parágrafo original do placeholder
        parent.remove(par_original)
        # Insere a lista completa no lugar
        parent.insert(idx, lista_xml)

    return len(hits)

def insere_lista_numerada_atividades(root, bookmark_name: str, lista_atividades: list):
    """
    Encontra um marcador e o substitui por uma lista numerada de atividades.
    Para cada atividade, cria um parágrafo para o título numerado e outro
    para a descrição.
    """
    path = f".//*[self::text:bookmark or self::text:bookmark-start][@text:name='{bookmark_name}']"
    hits = root.xpath(path, namespaces=NS)
    for bm in hits:
        par_original = _find_paragraph(bm)
        if par_original is None: continue

        parent = par_original.getparent()
        if parent is None: continue

        # Guarda a posição e o estilo do parágrafo numerado
        idx = parent.index(par_original)
        style_titulo_numerado = par_original.get(f"{{{TEXT_NS}}}style-name")
        
        # Define um estilo padrão para o parágrafo de descrição.
        # 'Text_20_body' é um estilo comum em muitos templates.
        # Se a formatação da descrição não ficar boa, este nome pode ser ajustado.
        style_descricao = 'Text_20_body'

        # Remove o parágrafo original que contém o marcador
        parent.remove(par_original)
        
        # Itera sobre os dados para criar os novos parágrafos
        for i, atividade in enumerate(lista_atividades, start=1):
            elemento_titulo = f"{i}. {atividade.get('elemento', '')}"
            texto_descricao = atividade.get('descricao', '')

            # 1. Cria o parágrafo do Título Numerado
            p_titulo = ET.Element(_t("p"))
            if style_titulo_numerado:
                p_titulo.set(f"{{{TEXT_NS}}}style-name", style_titulo_numerado)
            p_titulo.text = elemento_titulo
            parent.insert(idx, p_titulo)
            idx += 1

            # 2. Cria o parágrafo da Descrição
            p_desc = ET.Element(_t("p"))
            p_desc.set(f"{{{TEXT_NS}}}style-name", style_descricao)
            
            # Trata possíveis quebras de linha na descrição
            linhas_desc = texto_descricao.split('\n')
            for j, linha in enumerate(linhas_desc):
                span = ET.Element(_t("span"))
                span.text = linha
                p_desc.append(span)
                if j < len(linhas_desc) - 1:
                    p_desc.append(ET.Element(_t("line-break")))

            parent.insert(idx, p_desc)
            idx += 1
            
    return len(hits)

# def insere_lista_numerada_atividades(root, bookmark_name: str, lista_atividades: list):
#     """
#     Encontra um marcador e o substitui por uma lista numerada de atividades.
#     Aprende os estilos de título e descrição a partir de dois marcadores
#     no template ODT:
#     - bookmark_name: Marca o parágrafo do título (numerado, negrito).
#     - f"{bookmark_name}_DESC_STYLE": Marca o parágrafo da descrição (justificado).
#     """
#     # Marcador para o título (o principal)
#     path_titulo = f".//*[self::text:bookmark or self::text:bookmark-start][@text:name='{bookmark_name}']"
#     # Marcador para aprender o estilo da descrição
#     path_desc = f".//*[self::text:bookmark or self::text:bookmark-start][@text:name='{bookmark_name}_DESC_STYLE']"
    
#     hits_titulo = root.xpath(path_titulo, namespaces=NS)
#     hits_desc = root.xpath(path_desc, namespaces=NS)
    
#     # Se não encontrar os dois marcadores, não faz nada para evitar erros.
#     if not hits_titulo or not hits_desc:
#         return 0

#     # Captura as informações do parágrafo do TÍTULO
#     bm_titulo = hits_titulo[0]
#     par_titulo_original = _find_paragraph(bm_titulo)
#     parent = par_titulo_original.getparent()
#     idx = parent.index(par_titulo_original)
#     style_titulo_numerado = par_titulo_original.get(f"{{{TEXT_NS}}}style-name")
    
#     # Captura as informações do parágrafo da DESCRIÇÃO
#     bm_desc = hits_desc[0]
#     par_desc_original = _find_paragraph(bm_desc)
#     style_descricao = par_desc_original.get(f"{{{TEXT_NS}}}style-name")
    
#     # Remove os dois parágrafos de placeholder do documento
#     parent.remove(par_titulo_original)
#     # Verifica se o parágrafo de descrição ainda existe antes de remover
#     if par_desc_original in parent:
#         parent.remove(par_desc_original)

#     # Itera sobre os dados para criar os novos parágrafos
#     for i, atividade in enumerate(lista_atividades, start=1):
#         elemento_titulo = f"{i}. {atividade.get('elemento', '')}"
#         texto_descricao = atividade.get('descricao', '')

#         # 1. Cria o parágrafo do Título com o estilo capturado
#         p_titulo = ET.Element(_t("p"))
#         if style_titulo_numerado:
#             p_titulo.set(f"{{{TEXT_NS}}}style-name", style_titulo_numerado)
#         p_titulo.text = elemento_titulo
#         parent.insert(idx, p_titulo)
#         idx += 1

#         # 2. Cria o parágrafo da Descrição com o estilo capturado
#         p_desc = ET.Element(_t("p"))
#         if style_descricao:
#             p_desc.set(f"{{{TEXT_NS}}}style-name", style_descricao)
        
#         linhas_desc = texto_descricao.split('\n')
#         for j, linha in enumerate(linhas_desc):
#             span = ET.Element(_t("span"))
#             span.text = linha
#             p_desc.append(span)
#             if j < len(linhas_desc) - 1:
#                 p_desc.append(ET.Element(_t("line-break")))

#         parent.insert(idx, p_desc)
#         idx += 1
            
#     return len(hits_titulo)

def insert_toc_at_bookmark(root, name="BM_TOC", title="SUMÁRIO",
                           outline_levels=3, toc_name="TableOfContent1",
                           protect=True) -> int:
    hits = root.xpath(f".//text:bookmark[@text:name='{name}']", namespaces=NS)
    starts = root.xpath(f".//text:bookmark-start[@text:name='{name}']", namespaces=NS)
    changed = 0

    def _build_toc():
        toc = ET.Element(_t("table-of-content"))
        toc.set(f"{{{TEXT_NS}}}name", toc_name)
        if protect:
            toc.set(f"{{{TEXT_NS}}}protected", "true")
        source = ET.SubElement(toc, _t("table-of-content-source"))
        source.set(f"{{{TEXT_NS}}}outline-level", str(outline_levels))
        for level in range(1, outline_levels + 1):
            templ = ET.SubElement(source, _t("index-entry-template"))
            templ.set(f"{{{TEXT_NS}}}outline-level", str(level))
            ET.SubElement(templ, _t("index-entry-chapter"))
            ET.SubElement(templ, _t("index-entry-text"))
            ET.SubElement(templ, _t("index-entry-tab-stop"))
            ET.SubElement(templ, _t("index-entry-page-number"))
        index_title = ET.SubElement(toc, _t("index-title"))
        ptitle = ET.SubElement(index_title, _t("p"))
        ET.SubElement(ptitle, _t("span")).text = title
        ET.SubElement(toc, _t("index-body"))
        return toc

    # Caso 1: marcador de ponto (<text:bookmark/>)
    for bm in hits:
        par = _find_paragraph(bm)
        if par is None: 
            continue
        parent, idx = par.getparent(), par.getparent().index(par)
        parent.remove(par)
        parent.insert(idx, _build_toc())
        changed += 1

    # Caso 2: marcador de intervalo (<bookmark-start/end>) no MESMO parágrafo
    for st in starts:
        par = _find_paragraph(st)
        if par is None: 
            continue
        end = par.xpath(f".//text:bookmark-end[@text:name='{name}']", namespaces=NS)
        parent, idx = par.getparent(), par.getparent().index(par)
        # remove conteúdo do marcador e o parágrafo, substitui pelo TOC
        if end:
            i0, i1 = par.index(st), par.index(end[0])
            for _ in range(i1 - i0 + 1):
                del par[i0]
        parent.remove(par)
        parent.insert(idx, _build_toc())
        changed += 1

    return changed

def render_odt(template_path: str | Path, ctx: dict) -> bytes:
    template_path = str(template_path)
    with zipfile.ZipFile(template_path, "r") as zin:
        # Dicionários para guardar as árvores XML e os novos conteúdos
        roots = {}
        files_to_update = {}
        
        # Lê os arquivos XML relevantes (content.xml e styles.xml)
        if 'content.xml' in zin.namelist():
            xml_content = zin.read("content.xml")
            roots['content.xml'] = ET.fromstring(xml_content)
            
        if 'styles.xml' in zin.namelist():
            xml_styles = zin.read("styles.xml")
            roots['styles.xml'] = ET.fromstring(xml_styles)

        # --- Início da Lógica de Substituição ---

        # 1) User fields "comuns"
        fields = {
            "POP_NOME_PROCESSO":  ctx.get("nome_processo",""),
            "POP_CODIGO":         ctx.get("codigo",""),
            "POP_VERSAO":         _versao_fem_ordinal(ctx.get("versao","")),
            "POP_SETOR_SUPERIOR": ctx.get("POP_SETOR_SUPERIOR",""),
            "POP_SETOR_EXECUTOR": ctx.get("POP_SETOR_EXECUTOR",""),
            "NVL_GERENCIAL":      ctx.get("NVL_GERENCIAL",""),
            "NVL_OPERACIONAL":    ctx.get("NVL_OPERACIONAL",""),
            "POP_REVISOR":        ctx.get("rodape_elaborador", ""),
            "POP_APROVADOR":      ctx.get("aprovacao_responsavel", ""),
            "POP_DATA_APROVACAO": ctx.get("aprovacao_data", ""),
            # Remova a linha de teste se ainda estiver aqui
            # "TESTE_ABC": "SE ISTO APARECER, FUNCIONOU!" 
        }
        
        # Aplica a substituição em todos os XMLs carregados (content e styles)
        for root in roots.values():
            for k, v in fields.items():
                replace_userfield(root, k, v)

        # 1.1) EORGs com limpeza de quebra quando vazios
        eorg_sup  = ctx.get("EORG_SUP", "")
        eorg_exec = ctx.get("EORG_EXEC", "")
        for root in roots.values():
            replace_userfield_cleanup(root, "EORG_SUP",  eorg_sup,  remove_prev_break_if_empty=True)
            replace_userfield_cleanup(root, "EORG_EXEC", eorg_exec, remove_prev_break_if_empty=True)

        # 2) Listas (ENTER real entre itens) - Geralmente ficam só no content.xml
        if 'content.xml' in roots:
            content_root = roots['content.xml']
            oe_lines = format_lista_semicolas(ctx.get("objetivos_estrategicos", []))
            ie_raw   = ctx.get("indicadores_estrategicos", [])
            ie_lines = format_lista_semicolas(ie_raw) if ie_raw else ["Não há indicador sensibilizado"]

            _ = (fill_bookmark_single(content_root, "BM_OE_LIST", oe_lines, as_paragraphs=True)
                 or fill_bookmark_range_same_parent(content_root, "BM_OE_LIST", oe_lines, as_paragraphs=True))

            _ = (fill_bookmark_single(content_root, "BM_IE_LIST", ie_lines, as_paragraphs=True)
                 or fill_bookmark_range_same_parent(content_root, "BM_IE_LIST", ie_lines, as_paragraphs=True))
            # 1. Processa a lista "suja" do contexto para garantir que esteja limpa
            palavras_chave_processadas = processa_lista_aninhada(ctx.get("palavras_chave", []))
            # 2. Insere a lista limpa como bullets, encontrando qualquer tipo de marcador
            insere_lista_como_bullets(content_root, "BM_PALAVRAS_CHAVE", palavras_chave_processadas)
            
            atividades_list = ctx.get("descricao_processo_atividades", [])
            insere_lista_numerada_atividades(content_root, "BM_ATIVIDADES", atividades_list)

            _ = insert_toc_at_bookmark(content_root, name="BM_TOC", title="SUMÁRIO", outline_levels=3)
        
        # --- Fim da Lógica de Substituição ---

        # Serializa todos os arquivos XML que foram modificados
        for name, root in roots.items():
            files_to_update[name] = _serialize(root)

        # Grava o novo ODT com todas as alterações
        return _write_odt_like_template(zin, files_to_update, Path(template_path))