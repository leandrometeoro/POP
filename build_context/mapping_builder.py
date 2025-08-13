
# mapping_builder.py
# Constrói dicionários de tradução (code -> texto) a partir do pop-template.json

import json

def build_maps_from_template_json(template_json_path: str) -> dict:
    with open(template_json_path, "r", encoding="utf-8") as f:
        arr = json.load(f)
    tpl = arr[0]
    props = tpl.get("properties", [])
    maps = {}
    for p in props:
        binding = p.get("binding", {})
        name = binding.get("name")
        choices = p.get("choices")
        if name and choices:
            m = {}
            for ch in choices:
                val = ch.get("value")
                label = ch.get("name")
                if val is not None:
                    m[val] = label
            maps[name] = m
    return maps
