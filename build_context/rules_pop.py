import re
from html import unescape
from typing import Optional

_IEAPM_CODE_RE = re.compile(r"\(IEAPM-(\d+(?:\.\d+)?)\)", re.IGNORECASE)

def strip_html_preserve_breaks(html_text: str) -> str:
    if not isinstance(html_text, str):
        return ""
    s = unescape(html_text)
    s = re.sub(r"(?i)<\s*br\s*/?\s*>", "\n", s)
    s = re.sub(r"(?i)</\s*p\s*>", "\n\n", s)
    s = re.sub(r"<[^>]+>", "", s)
    s = s.replace("\xa0", " ").replace("\u200b", "")
    s = re.sub(r"[ \t]+", " ", s).strip()
    return re.sub(r"\n{3,}", "\n\n", s)

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

def _extrai_codigo_ieapm(rotulo_setor: str) -> Optional[str]:
    if not rotulo_setor:
        return None
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
