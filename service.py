# POP/service.py
from __future__ import annotations
from pathlib import Path
import re, unicodedata
from .build_context.pipeline_pop import hydrate_from_bpmn
from .build_context.rules_pop import calcula_nvl_gerencial, calcula_nvl_operacional
from .render import render_odt

BASE = Path(__file__).resolve().parent
CAM_MAP_PATH  = BASE / "templates" / "pop-template.json"
TEMPLATE_PATH = BASE / "templates" / "modelo_POP.odt"

_IEAPM = re.compile(r"\((IEAPM-[^)]+)\)")

def _split_eorg(label: str):
    if not label: return "", ""
    m = _IEAPM.search(label)
    code = m.group(1) if m else ""
    clean = _IEAPM.sub("", label).strip()
    return clean, code

def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))

def _clean_proc_name(s: str) -> str:
    s = _strip_accents(s or "")
    return re.sub(r"[^A-Za-z0-9]+", "", s) or "SemNome"

def _clean_code(s: str) -> str:
    s = _strip_accents(s or "")
    s = re.sub(r"[^A-Za-z0-9\-]+", "", s).strip("-")
    return s or "POP"

def _compute_filename(ctx: dict) -> str:
    codigo = _clean_code(ctx.get("codigo", "POP"))
    nome   = _clean_proc_name(ctx.get("nome_processo", "SemNome"))
    return f"{codigo}_{nome}.odt"

def _enrich_ctx(ctx: dict) -> dict:
    sup_clean, eorg_sup   = _split_eorg(ctx.get("setor_superior", ""))
    exec_clean, eorg_exec = _split_eorg(ctx.get("setor_executor", ""))
    ctx.update({
        "POP_SETOR_SUPERIOR": sup_clean,
        "POP_SETOR_EXECUTOR": exec_clean,
        "EORG_SUP": eorg_sup,
        "EORG_EXEC": eorg_exec,
        "NVL_GERENCIAL":   calcula_nvl_gerencial(ctx.get("setor_superior", "")),
        "NVL_OPERACIONAL": calcula_nvl_operacional(ctx.get("setor_executor", "")),
    })
    return ctx

def generate_pop_odt(bpmn_path: str | Path, out_dir: str | Path | None = None) -> Path:
    """Caminho → Caminho. Gera o ODT no disco e retorna o path final."""
    bpmn_path = Path(bpmn_path)
    if not TEMPLATE_PATH.is_file():
        raise FileNotFoundError(f"Template ODT não encontrado: {TEMPLATE_PATH}")
    if not CAM_MAP_PATH.is_file():
        raise FileNotFoundError(f"Mapa Camunda não encontrado: {CAM_MAP_PATH}")

    ctx = hydrate_from_bpmn(str(bpmn_path), str(CAM_MAP_PATH))
    ctx = _enrich_ctx(ctx)

    odt_bytes = render_odt(str(TEMPLATE_PATH), ctx)
    out_dir = Path(out_dir) if out_dir else bpmn_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    final = out_dir / _compute_filename(ctx)
    final.write_bytes(odt_bytes)
    return final

def generate_pop_odt_from_upload(bpmn_bytes: bytes, original_name: str = "processo.bpmn") -> tuple[str, bytes]:
    """Bytes → (nome_sugerido, odt_bytes). Útil para 'botão' com upload no navegador."""
    # salva temporário só para reusar o pipeline existente
    tmp = BASE / ".work" / "tmp_web"
    tmp.mkdir(parents=True, exist_ok=True)
    in_path = tmp / original_name
    in_path.write_bytes(bpmn_bytes)

    ctx = hydrate_from_bpmn(str(in_path), str(CAM_MAP_PATH))
    ctx = _enrich_ctx(ctx)
    odt_bytes = render_odt(str(TEMPLATE_PATH), ctx)
    return _compute_filename(ctx), odt_bytes

