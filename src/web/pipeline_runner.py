# -*- coding: utf-8 -*-
"""Runs the audit pipeline in a background thread with progress events via queue."""
import asyncio
import queue
import threading
import time
import traceback
from typing import AsyncGenerator


# Step map: all discrete steps with human-readable labels
STEP_MAP = [
    # Phase 0: Graph generation (6 steps)
    {"step": "graphs_volumes", "label": "Gráficos — Volumes de Entrada", "phase": 0, "chapter": None},
    {"step": "graphs_pcs", "label": "Gráficos — Poder Calorífico", "phase": 0, "chapter": None},
    {"step": "graphs_energia", "label": "Gráficos — Energia", "phase": 0, "chapter": None},
    {"step": "graphs_clientes", "label": "Gráficos — Perfis de Clientes", "phase": 0, "chapter": None},
    {"step": "graphs_incertezas", "label": "Gráficos — Incertezas", "phase": 0, "chapter": None},
    {"step": "graphs_balanco", "label": "Gráficos — Balanço de Massa", "phase": 0, "chapter": None},
    # Phase 1: Chapter generation (26 steps)
    {"step": "cap1_a_conteudo", "label": "Cap. 1 — Conteúdo", "phase": 1, "chapter": 1},
    {"step": "cap1_b_sintese", "label": "Cap. 1 — Síntese", "phase": 1, "chapter": 1},
    {"step": "cap2_a_metodologia", "label": "Cap. 2 — Metodologia", "phase": 1, "chapter": 2},
    {"step": "cap2_b_dados", "label": "Cap. 2 — Dados", "phase": 1, "chapter": 2},
    {"step": "cap2_c_graficos", "label": "Cap. 2 — Gráficos", "phase": 1, "chapter": 2},
    {"step": "cap2_d_sintese", "label": "Cap. 2 — Síntese", "phase": 1, "chapter": 2},
    {"step": "cap3_a_metodologia", "label": "Cap. 3 — Metodologia", "phase": 1, "chapter": 3},
    {"step": "cap3_b_dados", "label": "Cap. 3 — Dados", "phase": 1, "chapter": 3},
    {"step": "cap3_c_graficos", "label": "Cap. 3 — Gráficos", "phase": 1, "chapter": 3},
    {"step": "cap3_d_sintese", "label": "Cap. 3 — Síntese", "phase": 1, "chapter": 3},
    {"step": "cap4_a_metodologia", "label": "Cap. 4 — Metodologia", "phase": 1, "chapter": 4},
    {"step": "cap4_b_dados", "label": "Cap. 4 — Dados", "phase": 1, "chapter": 4},
    {"step": "cap4_c_graficos", "label": "Cap. 4 — Gráficos", "phase": 1, "chapter": 4},
    {"step": "cap4_d_sintese", "label": "Cap. 4 — Síntese", "phase": 1, "chapter": 4},
    {"step": "cap5_a_metodologia", "label": "Cap. 5 — Metodologia", "phase": 1, "chapter": 5},
    {"step": "cap5_b_dados", "label": "Cap. 5 — Dados", "phase": 1, "chapter": 5},
    {"step": "cap5_c_graficos", "label": "Cap. 5 — Gráficos", "phase": 1, "chapter": 5},
    {"step": "cap5_d_sintese", "label": "Cap. 5 — Síntese", "phase": 1, "chapter": 5},
    {"step": "cap6_a_metodologia", "label": "Cap. 6 — Metodologia", "phase": 1, "chapter": 6},
    {"step": "cap6_b_dados", "label": "Cap. 6 — Dados", "phase": 1, "chapter": 6},
    {"step": "cap6_c_graficos", "label": "Cap. 6 — Gráficos", "phase": 1, "chapter": 6},
    {"step": "cap6_d_sintese", "label": "Cap. 6 — Síntese", "phase": 1, "chapter": 6},
    {"step": "cap7_a_metodologia", "label": "Cap. 7 — Metodologia", "phase": 1, "chapter": 7},
    {"step": "cap7_b_dados", "label": "Cap. 7 — Dados", "phase": 1, "chapter": 7},
    {"step": "cap7_c_graficos", "label": "Cap. 7 — Gráficos", "phase": 1, "chapter": 7},
    {"step": "cap7_d_sintese", "label": "Cap. 7 — Síntese", "phase": 1, "chapter": 7},
    # Phase 2: Conclusions (1 step)
    {"step": "conclusoes", "label": "Conclusões e Recomendações", "phase": 2, "chapter": None},
    # Phase 3: Executive summary (1 step)
    {"step": "resumo_executivo", "label": "Resumo Executivo", "phase": 3, "chapter": None},
    # Phase 4: DOCX assembly (1 step)
    {"step": "docx_assembly", "label": "Montagem do Documento", "phase": 4, "chapter": None},
]

PHASE_NAMES = {
    0: "Geração dos Gráficos",
    1: "Geração dos Capítulos",
    2: "Conclusões e Recomendações",
    3: "Resumo Executivo",
    4: "Montagem do Documento",
}

TOTAL_STEPS = len(STEP_MAP)


class PipelineCancelled(Exception):
    pass


class PipelineRunner:
    """Manages pipeline execution in a background thread with event queuing."""

    def __init__(self):
        self._queue: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()
        self._running = False
        self._cancelled = False
        self._completed_steps: list[str] = []
        self._current_step: str | None = None
        self._current_phase: int = 0
        self._start_time: float = 0
        self._error: str | None = None

    @property
    def is_running(self) -> bool:
        return self._running

    def start(self, api_key: str, resume: bool = False, montar: bool = False):
        """Start the pipeline in a background thread."""
        self._running = True
        self._cancelled = False
        self._completed_steps = []
        self._current_step = None
        self._current_phase = 0
        self._start_time = time.time()
        self._error = None
        # Drain any old events
        while not self._queue.empty():
            try:
                self._queue.get_nowait()
            except queue.Empty:
                break

        self._thread = threading.Thread(
            target=self._run,
            args=(api_key, resume, montar),
            daemon=True,
        )
        self._thread.start()

    def cancel(self):
        """Request pipeline cancellation."""
        self._cancelled = True

    def _emit(self, event: dict):
        """Put an event on the queue (called from pipeline thread)."""
        event.setdefault("timestamp", time.time())
        event.setdefault("elapsed", time.time() - self._start_time)
        event.setdefault("progress", len(self._completed_steps) / TOTAL_STEPS)
        self._queue.put(event)

    def _on_progress(self, event: dict):
        """Callback injected into run_pipeline (thread-safe for parallel calls)."""
        if self._cancelled:
            raise PipelineCancelled("Pipeline cancelado pelo usuário")

        etype = event.get("type")
        step = event.get("step")

        with self._lock:
            if etype == "step_start":
                self._current_step = step
                # Find step info
                info = next((s for s in STEP_MAP if s["step"] == step), None)
                phase = info["phase"] if info else self._current_phase
                if phase != self._current_phase:
                    self._current_phase = phase
                    self._emit({
                        "type": "phase_start",
                        "phase": phase,
                        "phase_name": PHASE_NAMES.get(phase, ""),
                    })
                self._emit({
                    "type": "step_start",
                    "step": step,
                    "step_label": info["label"] if info else step,
                    "phase": phase,
                    "chapter": info["chapter"] if info else None,
                })

            elif etype == "step_complete":
                self._completed_steps.append(step)
                self._current_step = None
                info = next((s for s in STEP_MAP if s["step"] == step), None)
                self._emit({
                    "type": "step_complete",
                    "step": step,
                    "step_label": info["label"] if info else step,
                    "phase": info["phase"] if info else self._current_phase,
                    "chapter": info["chapter"] if info else None,
                    "completed": len(self._completed_steps),
                    "total": TOTAL_STEPS,
                })

            elif etype == "phase_complete":
                self._emit(event)

    def _run(self, api_key: str, resume: bool, montar: bool):
        """Execute pipeline in background thread."""
        try:
            from gerar_relatorio_auditoria import run_pipeline
            run_pipeline(
                api_key=api_key,
                resume=resume,
                montar=montar,
                on_progress=self._on_progress,
            )
            self._emit({"type": "done", "progress": 1.0})
        except PipelineCancelled:
            self._emit({"type": "cancelled", "detail": "Pipeline cancelado pelo usuário"})
        except Exception as e:
            self._error = str(e)
            self._emit({
                "type": "error",
                "detail": str(e),
                "traceback": traceback.format_exc(),
            })
        finally:
            self._running = False

    async def events(self) -> AsyncGenerator[dict, None]:
        """Async generator yielding events for the SSE endpoint."""
        while self._running or not self._queue.empty():
            try:
                event = self._queue.get_nowait()
                yield event
                if event.get("type") in ("done", "error", "cancelled"):
                    return
            except queue.Empty:
                await asyncio.sleep(0.3)

    def status(self) -> dict:
        """Current pipeline status snapshot."""
        return {
            "running": self._running,
            "cancelled": self._cancelled,
            "current_step": self._current_step,
            "current_phase": self._current_phase,
            "completed_steps": list(self._completed_steps),
            "completed_count": len(self._completed_steps),
            "total_steps": TOTAL_STEPS,
            "progress": len(self._completed_steps) / TOTAL_STEPS if TOTAL_STEPS else 0,
            "elapsed": time.time() - self._start_time if self._start_time else 0,
            "error": self._error,
        }
