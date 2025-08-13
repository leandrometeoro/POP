# POP/service.py
from __future__ import annotations
import re, unicodedata
from pathlib import Path

from .build_context.pipeline_pop import hydrate_from_bpmn
from .build_context.rules_pop import calcula_nvl_gerencial, calcula_nvl_operacional
from .render import render_odt
from .workspace import new_job, stage_input, write_context, write_artifact, deliver

PKG_DIR = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = PKG_DIR / "templates" / "modelo_POP.odt"
DEFAULT_CAM_MAP  = PKG_DIR / "templates" / "pop-template.json"

_IEAPM = re.compile(r"\((IEAPM-[^)]+)\)")

def _split_eorg(label: str):
    if not label:
        return "", ""
    m = _IEAPM.search(label)
    code = m.group(1) if m else ""
    clean = _IEAPM.sub("", label).strip()
    return clean, code

def _paren(code: str) -> str:
    return f"({code})" if code else ""

def _is_na(text: str) -> bool:
    s = (text or "").strip().lower()
    return ("não aplicável" in s or "nao aplicavel" in s) or (s in {"—x—", "—x-", "-x—", "-x-", "—x—"})

def _slug(s: str) -> str:
    if s is None:
        return ""
    # remove acentos
    s = unicodedata.normalize("NFKD", s)
    s = s.encode("ascii", "ignore").decode("ascii")
    # troca não-alfanum por _
    s = re.sub(r"[^A-Za-z0-9]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s

def _normalize_oe_label(label: str) -> str:
    # "OE 02 - Texto" => "Texto (OE 02)"
    m = re.match(r"^\s*(OE\s*\d+(?:\.\d+)?)\s*-\s*(.+)$", label or "")
    if m:
        code, desc = m.group(1).strip(), m.group(2).strip()
        return f"{desc} ({code})"
    return label or ""

def _apply_business_rules(ctx: dict) -> dict:
    # 1) EORG + limpeza dos setores
    sup_raw = ctx.get("setor_superior", "")
    exe_raw = ctx.get("setor_executor", "")

    sup_clean, sup_code = _split_eorg(sup_raw)
    exe_clean, exe_code = _split_eorg(exe_raw)

    # Execução NA => apenas —X— e sem EORG_EXEC
    if _is_na(exe_clean):
        pop_exec = "—X—"
        eorg_exec = ""
    else:
        pop_exec = exe_clean
        eorg_exec = _paren(exe_code)

    # Parênteses no EORG_SUP sempre que houver
    eorg_sup = _paren(sup_code)

    # 2) NVL_* calculados a partir dos rótulos originais (que ainda têm o código)
    ctx["NVL_GERENCIAL"]   = calcula_nvl_gerencial(sup_raw)
    ctx["NVL_OPERACIONAL"] = calcula_nvl_operacional(exe_raw)

    # 3) Preenche os campos consumidos pelo ODT
    ctx["POP_SETOR_SUPERIOR"] = sup_clean
    ctx["POP_SETOR_EXECUTOR"] = pop_exec
    ctx["EORG_SUP"]  = eorg_sup
    ctx["EORG_EXEC"] = eorg_exec

    # 4) Objetivos estratégicos no formato "Texto (OE 0X)"
    oes = [ _normalize_oe_label(x) for x in ctx.get("objetivos_estrategicos", []) ]
    ctx["objetivos_estrategicos"] = [x for x in oes if x]

    # 5) Indicadores: fallback quando vazio
    ies = ctx.get("indicadores_estrategicos") or []
    if not ies:
        ctx["indicadores_estrategicos"] = ["Não há indicador sensibilizado."]

    return ctx

def generate_pop_odt(
    bpmn_path: str,
    out_dir: str | None = None,
    template_path: str | Path = DEFAULT_TEMPLATE,
    camunda_map_path: str | Path = DEFAULT_CAM_MAP,
):
    job_id, _ = new_job(prefix="pop")

    # isola insumos
    bpmn_in = stage_input(job_id, bpmn_path)
    tpl_in  = stage_input(job_id, template_path)
    cmap_in = stage_input(job_id, camunda_map_path)

    # contexto base (BPMN + maps)
    ctx = hydrate_from_bpmn(str(bpmn_in), str(cmap_in))

    # aplica regras de negócio locais
    ctx = _apply_business_rules(ctx)

    # salva contexto para auditoria/depuração
    ctx_path = write_context(job_id, ctx, "primeira_pagina.contexto.json")

    # renderiza ODT -> bytes
    odt_bytes = render_odt(str(tpl_in), ctx)
    odt_int   = write_artifact(job_id, odt_bytes, "odt", "primeira_pagina.odt")

    # nome de entrega: codigo_nomeprocesso.odt
    codigo = _slug(ctx.get("codigo", "") or "CODIGO")
    nome   = _slug(ctx.get("nome_processo", "") or "NOME_PROCESSO")
    final_name = f"{codigo}_{nome}.odt"

    # destino: mesmo diretório do BPMN, salvo se out_dir for passado
    out_dir = Path(out_dir) if out_dir else Path(bpmn_path).resolve().parent
    final = deliver(odt_int, out_dir / final_name)

    return {
        "job_id": job_id,
        "context_path": str(ctx_path),
        "output_path": str(final),
        "filename": final_name,
    }

