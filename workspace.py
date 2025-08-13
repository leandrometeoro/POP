from __future__ import annotations
from pathlib import Path
import os, time, uuid, shutil, json

_DEFAULT = Path(__file__).resolve().parent / ".work"
WORKDIR = Path(os.environ.get("POP_WORKDIR", _DEFAULT))
_SUBS = ["inbox", "contexts", "odt", "pdf", "tmp", "logs", "archive"]

def ensure_workdirs(): 
    for s in _SUBS: (WORKDIR / s).mkdir(parents=True, exist_ok=True)

def new_job(prefix: str = "job"):
    ensure_workdirs()
    ts = time.strftime("%Y%m%d-%H%M%S")
    jid = f"{prefix}-{ts}-{uuid.uuid4().hex[:6]}"
    jobdir = WORKDIR / "tmp" / jid
    jobdir.mkdir(parents=True, exist_ok=True)
    return jid, jobdir

def stage_input(job_id: str, src, name: str | None = None) -> Path:
    src = Path(src); name = name or src.name
    dst = WORKDIR / "inbox" / f"{job_id}-{name}"
    shutil.copy2(src, dst); return dst

def write_context(job_id: str, ctx: dict, filename="contexto.json") -> Path:
    out = WORKDIR / "contexts" / f"{job_id}-{filename}"
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "w", encoding="utf-8") as f: json.dump(ctx, f, ensure_ascii=False, indent=2)
    return out

def write_artifact(job_id: str, blob: bytes, kind: str, filename: str) -> Path:
    outdir = WORKDIR / kind; outdir.mkdir(parents=True, exist_ok=True)
    out = outdir / f"{job_id}-{filename}"
    with open(out, "wb") as f: f.write(blob)
    return out

def deliver(src, dst) -> Path:
    src, dst = Path(src), Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_suffix(dst.suffix + ".tmp")
    shutil.copy2(src, tmp); os.replace(tmp, dst)
    return dst

