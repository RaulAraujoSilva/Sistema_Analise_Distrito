# -*- coding: utf-8 -*-
"""FastAPI application for the gas audit pipeline web interface."""
import sys
from pathlib import Path

# Ensure src/ is importable
SRC_DIR = Path(__file__).resolve().parent.parent
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from fastapi import FastAPI, UploadFile, File, Request, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
import json
import os
import time

from config import (
    PROJECT_ROOT, DATA_DIR, GRAFICOS_DIR, DIAGRAMAS_DIR,
    CACHE_DIR, REPORTS_DIR, PRESENT_DIR, EXCEL_DEFAULT, IS_VERCEL,
)
from prompts_auditoria import CHAPTER_CONFIG

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------
WEB_DIR = Path(__file__).resolve().parent

app = FastAPI(title="Auditoria de Distrito de Gás", version="1.0")

# Static files
app.mount("/static", StaticFiles(directory=str(WEB_DIR / "static")), name="static")

# On Vercel, create writable output dirs in /tmp at startup
# Note: DIAGRAMAS_DIR is read-only (bundled in repo), not created here
if IS_VERCEL:
    for d in [DATA_DIR, GRAFICOS_DIR, REPORTS_DIR, PRESENT_DIR, CACHE_DIR]:
        d.mkdir(parents=True, exist_ok=True)

# Serve bundled read-only files via StaticFiles (diagramas are in repo, not /tmp)
if DIAGRAMAS_DIR.exists():
    app.mount("/files/diagramas", StaticFiles(directory=str(DIAGRAMAS_DIR)), name="diagramas")

# Templates
templates = Jinja2Templates(directory=str(WEB_DIR / "templates"))

# In-memory state
_state = {
    "api_key": None,
    "uploaded_file": None,
    "pipeline_running": False,
}

# ---------------------------------------------------------------------------
# Pipeline runner (lazy import to avoid circular deps)
# ---------------------------------------------------------------------------
_runner = None

def get_runner():
    global _runner
    if _runner is None:
        from web.pipeline_runner import PipelineRunner
        _runner = PipelineRunner()
    return _runner


# ---------------------------------------------------------------------------
# Page route
# ---------------------------------------------------------------------------
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """Main SPA page."""
    return templates.TemplateResponse("index.html", {
        "request": request,
        "chapter_config": CHAPTER_CONFIG,
    })


# ---------------------------------------------------------------------------
# Upload API
# ---------------------------------------------------------------------------
@app.post("/api/upload")
async def upload_excel(file: UploadFile = File(...)):
    """Upload Excel file to data/input/."""
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Arquivo deve ser .xlsx ou .xls")

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    dest = DATA_DIR / EXCEL_DEFAULT
    content = await file.read()
    dest.write_bytes(content)

    _state["uploaded_file"] = str(dest)
    return {
        "status": "ok",
        "filename": EXCEL_DEFAULT,
        "size_mb": round(len(content) / (1024 * 1024), 2),
    }


# ---------------------------------------------------------------------------
# Pipeline control API
# ---------------------------------------------------------------------------
class PipelineStartRequest(BaseModel):
    api_key: str
    mode: str = "gerar"  # "gerar", "resume", "montar"


@app.post("/api/pipeline/start")
async def pipeline_start(req: PipelineStartRequest):
    """Start the pipeline in background."""
    runner = get_runner()
    if runner.is_running:
        raise HTTPException(409, "Pipeline já está em execução")

    _state["api_key"] = req.api_key
    runner.start(
        api_key=req.api_key,
        resume=(req.mode == "resume"),
        montar=(req.mode == "montar"),
    )
    return {"status": "started", "mode": req.mode}


@app.get("/api/pipeline/events")
async def pipeline_events():
    """SSE stream of pipeline progress events."""
    runner = get_runner()

    async def generate():
        async for event in runner.events():
            data = json.dumps(event, ensure_ascii=False)
            yield f"event: {event.get('type', 'info')}\ndata: {data}\n\n"

    return StreamingResponse(
        generate(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


@app.get("/api/pipeline/status")
async def pipeline_status():
    """Current pipeline status (for reconnection)."""
    runner = get_runner()
    return runner.status()


@app.post("/api/pipeline/cancel")
async def pipeline_cancel():
    """Cancel running pipeline."""
    runner = get_runner()
    runner.cancel()
    return {"status": "cancelled"}


# ---------------------------------------------------------------------------
# Dynamic file serving (graphs, reports, presentations live in /tmp on Vercel)
# ---------------------------------------------------------------------------
@app.get("/files/graficos/{filename}")
async def serve_grafico(filename: str):
    """Serve a generated graph PNG."""
    fpath = GRAFICOS_DIR / filename
    if not fpath.exists():
        raise HTTPException(404, f"Gráfico {filename} não encontrado")
    return FileResponse(fpath, media_type="image/png")


@app.get("/files/reports/{filename}")
async def serve_report(filename: str):
    """Serve a generated DOCX report."""
    fpath = REPORTS_DIR / filename
    if not fpath.exists():
        raise HTTPException(404, f"Relatório {filename} não encontrado")
    return FileResponse(fpath)


@app.get("/files/presentations/{filename}")
async def serve_presentation(filename: str):
    """Serve a generated PPTX presentation."""
    fpath = PRESENT_DIR / filename
    if not fpath.exists():
        raise HTTPException(404, f"Apresentação {filename} não encontrado")
    return FileResponse(fpath)


# ---------------------------------------------------------------------------
# Output listing APIs
# ---------------------------------------------------------------------------
@app.get("/api/outputs/graficos")
async def list_graficos():
    """List all graph PNGs organized by chapter."""
    result = {}
    for cap_num, cfg in CHAPTER_CONFIG.items():
        graphs = []
        for fname in cfg.get("graph_files", []):
            fpath = GRAFICOS_DIR / fname
            graphs.append({
                "filename": fname,
                "caption": cfg.get("graph_captions", {}).get(fname, fname),
                "exists": fpath.exists(),
                "url": f"/files/graficos/{fname}",
            })
        if graphs:
            result[cap_num] = {
                "titulo": cfg["titulo"],
                "graphs": graphs,
            }
    return result


@app.get("/api/outputs/diagramas")
async def list_diagramas():
    """List all diagram PNGs."""
    diagrams = []
    if DIAGRAMAS_DIR.exists():
        for f in sorted(DIAGRAMAS_DIR.glob("*.png")):
            # Find caption from chapter config
            caption = f.stem.replace("_", " ").title()
            for cfg in CHAPTER_CONFIG.values():
                if f.name in cfg.get("diagram_captions", {}):
                    caption = cfg["diagram_captions"][f.name]
                    break
            diagrams.append({
                "filename": f.name,
                "caption": caption,
                "url": f"/files/diagramas/{f.name}",
            })
    return diagrams


@app.get("/api/outputs/cache")
async def list_cache():
    """List all cached markdown sections."""
    sections = []
    if CACHE_DIR.exists():
        for f in sorted(CACHE_DIR.glob("*.md")):
            sections.append({
                "filename": f.stem,
                "size_kb": round(f.stat().st_size / 1024, 1),
            })
    return sections


@app.get("/api/outputs/cache/{filename}")
async def get_cache_content(filename: str):
    """Get raw markdown content of a cached section."""
    fpath = CACHE_DIR / f"{filename}.md"
    if not fpath.exists():
        raise HTTPException(404, f"Cache {filename} não encontrado")
    return {"filename": filename, "content": fpath.read_text(encoding="utf-8")}


@app.get("/api/outputs/downloads")
async def list_downloads():
    """List available download files (DOCX, PPTX)."""
    files = []
    for directory, ftype in [(REPORTS_DIR, "docx"), (PRESENT_DIR, "pptx")]:
        if directory.exists():
            for f in directory.iterdir():
                if f.suffix in (".docx", ".pptx"):
                    files.append({
                        "filename": f.name,
                        "type": ftype,
                        "size_mb": round(f.stat().st_size / (1024 * 1024), 2),
                        "url": f"/files/{'reports' if ftype == 'docx' else 'presentations'}/{f.name}",
                    })
    return files


# ---------------------------------------------------------------------------
# Phase-by-phase execution API
# ---------------------------------------------------------------------------
@app.post("/api/phase/extract")
async def phase_extract():
    """Phase 0a: Extract data from Excel, save as JSON."""
    excel_path = DATA_DIR / EXCEL_DEFAULT
    if not excel_path.exists():
        raise HTTPException(400, "Excel não encontrado. Faça upload primeiro.")

    try:
        from extrator_dados import extrair_todos, salvar_json
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        dados = extrair_todos(str(excel_path))
        json_path = CACHE_DIR / "extracted_data.json"
        salvar_json(dados, str(json_path))

        # Build summary for response
        config = dados.get("config")
        vol = dados.get(2)
        pcs = dados.get(3)
        ene = dados.get(4)
        perf = dados.get(5)
        bal = dados.get(7)
        return {
            "status": "ok",
            "config": {
                "periodo": f"{config.periodo_inicio} a {config.periodo_fim}",
                "dias": config.dias,
                "n_clientes": config.n_clientes,
            },
            "resumo": {
                "vol_total_nm3": round(vol.vol_total_nm3),
                "vol_medio_nm3d": round(vol.vol_medio_nm3d),
                "pcs_media_kcal": round(pcs.media_kcal, 2),
                "energia_total_gcal": round(ene.total_gcal),
                "n_clientes": len(perf.clientes),
                "balanco_pct": bal.diferenca_pct,
                "balanco_resultado": bal.resultado,
            },
        }
    except Exception as e:
        raise HTTPException(500, f"Erro na extração: {e}")


@app.get("/api/phase/extract/preview")
async def phase_extract_preview():
    """Returns extracted data summary for validation UI."""
    json_path = CACHE_DIR / "extracted_data.json"
    if not json_path.exists():
        raise HTTPException(404, "Dados não extraídos. Execute /api/phase/extract primeiro.")

    try:
        from extrator_dados import carregar_json
        from dados_distrito import formatar_dados_secao
        dados = carregar_json(str(json_path))
        config = dados.get("config")
        sections = {}
        labels = {2: "Volumes", 3: "PCS", 4: "Energia", 5: "Clientes", 6: "Incertezas", 7: "Balanço"}
        for key in [2, 3, 4, 5, 6, 7]:
            if key in dados:
                sections[labels[key]] = formatar_dados_secao(dados[key], config=config)
        return {
            "config": {
                "periodo": f"{config.periodo_inicio} a {config.periodo_fim}",
                "dias": config.dias,
            } if config else None,
            "sections": sections,
        }
    except Exception as e:
        raise HTTPException(500, f"Erro ao carregar dados: {e}")


@app.post("/api/phase/graphs")
async def phase_graphs():
    """Phase 0b: Generate all graphs from Excel."""
    excel_path = DATA_DIR / EXCEL_DEFAULT
    if not excel_path.exists():
        raise HTTPException(400, "Excel não encontrado. Faça upload primeiro.")

    try:
        from graph_generator import gerar_todos_graficos
        GRAFICOS_DIR.mkdir(parents=True, exist_ok=True)
        gerados = gerar_todos_graficos(
            excel_path=str(excel_path),
            output_dir=str(GRAFICOS_DIR),
        )
        return {"status": "ok", "count": len(gerados), "files": gerados}
    except Exception as e:
        raise HTTPException(500, f"Erro ao gerar gráficos: {e}")


# ---------------------------------------------------------------------------
# Presentation generation
# ---------------------------------------------------------------------------
@app.post("/api/presentation/generate")
async def generate_presentation():
    """Generate PPTX presentation."""
    try:
        from gerar_apresentacao import gerar_apresentacao
        output_path = gerar_apresentacao()
        return {
            "status": "ok",
            "filename": Path(output_path).name,
            "size_mb": round(Path(output_path).stat().st_size / (1024 * 1024), 2),
        }
    except Exception as e:
        raise HTTPException(500, f"Erro ao gerar apresentação: {e}")
