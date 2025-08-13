
# pipeline_pop.py
# Orquestra: lê BPMN -> aplica maps -> gera contexto_clean.json (sem HTML) focado na primeira página

import json, os
from typing import Any, Dict

from .mapping_builder import build_maps_from_template_json
from .rules_pop import strip_html_preserve_breaks

def hydrate_from_bpmn(bpmn_path: str, template_json: str) -> dict:
    """Lê o .bpmn via seu parser e retorna um contexto 'bruto' + campos mapeados legíveis."""
    try:
        from .parser_bpmn import parse_bpmn_pop
        raw = parse_bpmn_pop(bpmn_path)  # espera um dict
    except Exception as e:
        raise RuntimeError(f"Falha ao ler BPMN: {e}")

    maps = build_maps_from_template_json(template_json)
    props = raw.get("propriedades_pop", raw)

    ctx: Dict[str, Any] = {}

    ctx["nome_processo"] = props.get("nomeProcesso") or raw.get("nome_processo") or ""
    ctx["codigo"]        = props.get("codigo") or props.get("pop:codigo") or raw.get("codigo") or ""
    ctx["versao"]        = props.get("versao") or raw.get("versao") or "1"

    def map_choice(field_name: str, value: str) -> str:
        if not value:
            return ""
        m = maps.get(field_name, {})
        return m.get(value, value)

    sup_val = props.get("superintendenciaResponsavel") or props.get("pop:superintendenciaResponsavel") or ""
    exe_val = props.get("departamentoResponsavel")     or props.get("pop:departamentoResponsavel")     or ""

    ctx["setor_superior"] = map_choice("pop:superintendenciaResponsavel", sup_val)
    ctx["setor_executor"] = map_choice("pop:departamentoResponsavel", exe_val)

    oe_codes = [
        props.get("objetivoEstrategico1") or props.get("pop:objetivoEstrategico1"),
        props.get("objetivoEstrategico2") or props.get("pop:objetivoEstrategico2"),
        props.get("objetivoEstrategico3") or props.get("pop:objetivoEstrategico3"),
    ]
    oe_map = maps.get("pop:objetivoEstrategico1", {})
    ctx["objetivos_estrategicos"] = [oe_map.get(c, c) for c in oe_codes if c]

    ie_codes = [
        props.get("indicadorEstrategico1") or props.get("pop:indicadorEstrategico1"),
        props.get("indicadorEstrategico2") or props.get("pop:indicadorEstrategico2"),
        props.get("indicadorEstrategico3") or props.get("pop:indicadorEstrategico3"),
    ]
    ie_map = maps.get("pop:indicadorEstrategico1", {})
    ctx["indicadores_estrategicos"] = [ie_map.get(c, c) for c in ie_codes if c]

    palavras = []
    for k in ("palavraChave1","palavraChave2","palavraChave3"):
        v = props.get(k) or props.get(f"pop:{k}")
        if v:
            palavras.append(v)
    add = props.get("palavrasChaveAdicionais") or props.get("pop:palavrasChaveAdicionais")
    if add:
        palavras.extend([s.strip() for s in str(add).split("//") if s.strip()])
    ctx["palavras_chave"] = palavras

    dicionario = []
    for i in (1,2,3):
        t = props.get(f"dicionario{i}_termo") or props.get(f"pop:dicionario{i}_termo")
        s = props.get(f"dicionario{i}_significado") or props.get(f"pop:dicionario{i}_significado")
        if t or s:
            dicionario.append({"termo": t or "", "significado": s or ""})
    termos_ad = props.get("dicionarioAdicionais_termos") or props.get("pop:dicionarioAdicionais_termos")
    sig_ad    = props.get("dicionarioAdicionais_significados") or props.get("pop:dicionarioAdicionais_significados")
    if termos_ad and sig_ad:
        ts = [s.strip() for s in str(termos_ad).split("//")]
        ss = [s.strip() for s in str(sig_ad).split("//")]
        for i in range(min(len(ts), len(ss))):
            if ts[i] or ss[i]:
                dicionario.append({"termo": ts[i], "significado": ss[i]})
    ctx["dicionario"] = dicionario

    desc = []
    for item in raw.get("descricao_processo_atividades", []):
        elemento = item.get("elemento","");
        texto = strip_html_preserve_breaks(item.get("descricao",""))
        if elemento or texto:
            desc.append({"elemento": elemento, "descricao": texto})
    if desc:
        ctx["descricao_processo_atividades"] = desc

    for k in ("rodape_elaborador","aprovacao_data","aprovacao_responsavel","aprovacao_setor"):
        if props.get(k) or props.get(f"pop:{k}"):
            ctx[k] = props.get(k) or props.get(f"pop:{k}")

    return ctx

def write_json(data: dict, path: str):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
