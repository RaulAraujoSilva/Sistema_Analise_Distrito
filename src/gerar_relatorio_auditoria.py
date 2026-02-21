# -*- coding: utf-8 -*-
"""
Gerador automatizado de relatório de auditoria — Versão 3 (segmentada).

Pipeline: 4 chamadas por capítulo (Metodologia → Dados → Gráficos → Síntese),
depois Conclusões e Resumo Executivo.

Uso:
    python gerar_relatorio_auditoria.py --api-key SUA_CHAVE_API
    python gerar_relatorio_auditoria.py --api-key SUA_CHAVE_API --resume
    python gerar_relatorio_auditoria.py --api-key SUA_CHAVE_API --montar
"""
import argparse
import logging
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path

from dados_distrito import (
    VolumesEntrada, PCSData, EnergiaData,
    PerfisClientes, IncertezasData, BalancoMassa,
    formatar_dados_secao, gerar_tabelas_resumo,
)
from gemini_client import GeminiAuditClient
from prompts_auditoria import (
    CHAPTER_CONFIG,
    prompt_metodologia,
    prompt_dados,
    prompt_graficos,
    prompt_sintese,
    prompt_secao1_conteudo,
    prompt_secao1_sintese,
    prompt_conclusoes_recomendacoes,
    prompt_resumo_executivo,
)
from docx_builder import AuditReportBuilder

# =====================================================================
# CAMINHOS (centralizados em config.py)
# =====================================================================
from config import (
    GRAFICOS_DIR, CACHE_DIR, METODOLOGIA_DIR, DIAGRAMAS_DIR,
    REPORTS_DIR, NOTEBOOKS_DIR, NOTEBOOK_LIST, DATA_DIR, EXCEL_DEFAULT,
)
from graph_generator import gerar_todos_graficos

OUTPUT_DEFAULT = "Relatorio_Auditoria_Distrito.docx"

# =====================================================================
# LOGGING
# =====================================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# =====================================================================
# MAPA DE DADOS POR CAPÍTULO
# =====================================================================
DATA_CLASSES = {
    2: VolumesEntrada,
    3: PCSData,
    4: EnergiaData,
    5: PerfisClientes,
    6: IncertezasData,
    7: BalancoMassa,
}


# =====================================================================
# DATACLASSES
# =====================================================================

@dataclass
class ChapterResult:
    """Armazena sub-resultados de um capítulo."""
    chapter_num: int
    titulo: str
    introducao: str = ""
    metodologia: str = ""
    dados: str = ""
    graficos: str = ""
    parecer: str = ""
    conteudo_cap1: str = ""  # Apenas para Capítulo 1 (caso especial)

    @property
    def full_text(self) -> str:
        """Combina todas as partes para uso como contexto."""
        parts = []
        if self.introducao:
            parts.append(self.introducao)
        if self.conteudo_cap1:
            parts.append(self.conteudo_cap1)
        if self.metodologia:
            parts.append(f"### Fundamentação Teórica\n{self.metodologia}")
        if self.dados:
            parts.append(f"### Análise dos Dados\n{self.dados}")
        if self.graficos:
            parts.append(f"### Discussão dos Gráficos\n{self.graficos}")
        if self.parecer:
            parts.append(f"### Parecer Regulatório\n{self.parecer}")
        return "\n\n".join(parts)


# =====================================================================
# FUNÇÕES AUXILIARES
# =====================================================================

def load_metodologia(secao_num: int) -> str:
    """Carrega texto de metodologia pré-extraída para uma seção."""
    config = CHAPTER_CONFIG.get(secao_num, {})
    arquivo = config.get("metodologia_file", "")
    path = METODOLOGIA_DIR / arquivo
    if path.exists():
        text = path.read_text(encoding="utf-8")
        logger.info(f"      Metodologia: {arquivo} ({len(text)} chars)")
        return text
    logger.warning(f"      Metodologia não encontrada: {arquivo}")
    return "(Metodologia não disponível para esta seção)"


def load_cache(cache_key: str) -> str | None:
    """Carrega conteúdo do cache se existir."""
    cache_file = CACHE_DIR / f"{cache_key}.md"
    if cache_file.exists():
        text = cache_file.read_text(encoding="utf-8")
        if len(text) > 50:
            return text
    return None


def save_cache(cache_key: str, text: str):
    """Salva conteúdo no cache."""
    CACHE_DIR.mkdir(exist_ok=True)
    cache_file = CACHE_DIR / f"{cache_key}.md"
    cache_file.write_text(text, encoding="utf-8")


def parse_synthesis(text: str) -> tuple[str, str]:
    """
    Separa output da chamada de síntese em (introdução, parecer).
    O LLM é instruído a separar com '---SEPARADOR---'.
    """
    if "---SEPARADOR---" in text:
        parts = text.split("---SEPARADOR---", 1)
        intro = parts[0].strip()
        parecer = parts[1].strip()
    else:
        # Fallback: tenta encontrar "### Parecer Regulatório"
        match = re.search(r'(###\s*Parecer\s+Regulat)', text, re.IGNORECASE)
        if match:
            intro = text[:match.start()].strip()
            parecer = text[match.start():].strip()
        else:
            # Último recurso: tudo como introdução, parecer vazio
            intro = text.strip()
            parecer = ""
    return intro, parecer


def _emit(on_progress, event_type, **kwargs):
    """Emit a progress event if callback is provided."""
    if on_progress:
        on_progress({"type": event_type, **kwargs})


def generate_subcall(
    client: GeminiAuditClient,
    cache_key: str,
    prompt_fn,
    prompt_args: tuple,
    thinking_level: str = "high",
    resume: bool = False,
    montar: bool = False,
    on_progress=None,
) -> str:
    """Executa uma sub-chamada com cache."""
    _emit(on_progress, "step_start", step=cache_key)

    # Verificar cache
    if resume or montar:
        cached = load_cache(cache_key)
        if cached:
            logger.info(f"      Cache: {cache_key} ({len(cached)} chars)")
            _emit(on_progress, "step_complete", step=cache_key)
            return cached

    if montar:
        logger.warning(f"      Sem cache: {cache_key} (modo montar)")
        _emit(on_progress, "step_complete", step=cache_key)
        return f"[{cache_key} não gerado — modo montar sem cache]"

    # Gerar via API
    sys_prompt, sec_prompt, image_files = prompt_fn(*prompt_args)

    # Converter nomes de arquivos em caminhos completos
    image_paths = []
    for f in image_files:
        p = Path(f)
        if p.is_absolute() and p.exists():
            image_paths.append(str(p))
        else:
            full = GRAFICOS_DIR / f
            if full.exists():
                image_paths.append(str(full))

    text = client.analyze_section(
        sys_prompt, sec_prompt, image_paths,
        thinking_level=thinking_level,
    )
    save_cache(cache_key, text)
    logger.info(f"      Gerado: {cache_key} ({len(text)} chars)")
    _emit(on_progress, "step_complete", step=cache_key)
    return text


# =====================================================================
# GERAÇÃO DE CAPÍTULOS
# =====================================================================

def generate_chapter_1(client, resume=False, montar=False, on_progress=None) -> ChapterResult:
    """Capítulo 1: Visão Geral (caso especial — conteúdo + síntese)."""
    logger.info("  [Cap 1] Visão Geral do Distrito e Dados Disponíveis")
    ch = ChapterResult(chapter_num=1, titulo=CHAPTER_CONFIG[1]["titulo"])

    # Chamada A: Conteúdo com diagramas como imagens
    logger.info("    [A] Conteúdo (com diagramas)")
    metodologia = load_metodologia(1)
    ch.conteudo_cap1 = generate_subcall(
        client, "cap1_a_conteudo",
        prompt_secao1_conteudo, (metodologia,),
        thinking_level="high",
        resume=resume, montar=montar, on_progress=on_progress,
    )

    # Chamada B: Síntese (introdução + parecer)
    logger.info("    [B] Síntese")
    sintese_text = generate_subcall(
        client, "cap1_b_sintese",
        prompt_secao1_sintese, (ch.conteudo_cap1,),
        thinking_level="low",
        resume=resume, montar=montar, on_progress=on_progress,
    )
    ch.introducao, ch.parecer = parse_synthesis(sintese_text)

    return ch


def generate_chapter_standard(client, cap_num: int, resume=False, montar=False, on_progress=None) -> ChapterResult:
    """Capítulos 2-7: 4 sub-chamadas (Metodologia → Dados → Gráficos → Síntese)."""
    config = CHAPTER_CONFIG[cap_num]
    logger.info(f"  [Cap {cap_num}] {config['titulo']}")
    ch = ChapterResult(chapter_num=cap_num, titulo=config["titulo"])

    # Carregar metodologia e dados
    metodologia = load_metodologia(cap_num)
    data_cls = DATA_CLASSES.get(cap_num)
    dados_texto = formatar_dados_secao(data_cls()) if data_cls else ""

    # Chamada A: Metodologia
    logger.info(f"    [A] Metodologia")
    ch.metodologia = generate_subcall(
        client, f"cap{cap_num}_a_metodologia",
        prompt_metodologia, (cap_num, metodologia),
        thinking_level="low",
        resume=resume, montar=montar, on_progress=on_progress,
    )

    # Chamada B: Dados
    logger.info(f"    [B] Dados")
    ch.dados = generate_subcall(
        client, f"cap{cap_num}_b_dados",
        prompt_dados, (cap_num, dados_texto),
        thinking_level="low",
        resume=resume, montar=montar, on_progress=on_progress,
    )

    # Chamada C: Gráficos
    if config["graph_files"]:
        logger.info(f"    [C] Gráficos ({len(config['graph_files'])} imagens)")
        ch.graficos = generate_subcall(
            client, f"cap{cap_num}_c_graficos",
            prompt_graficos, (cap_num,),
            thinking_level="high",
            resume=resume, montar=montar, on_progress=on_progress,
        )

    # Chamada D: Síntese
    logger.info(f"    [D] Síntese")
    sintese_text = generate_subcall(
        client, f"cap{cap_num}_d_sintese",
        prompt_sintese, (cap_num, ch.metodologia, ch.dados, ch.graficos),
        thinking_level="low",
        resume=resume, montar=montar, on_progress=on_progress,
    )
    ch.introducao, ch.parecer = parse_synthesis(sintese_text)

    return ch


# =====================================================================
# MONTAGEM DOCX
# =====================================================================

def assemble_docx(chapters: dict, resumo: str, conclusoes: str, output: str = OUTPUT_DEFAULT):
    """Monta o documento DOCX final."""
    report = AuditReportBuilder(graficos_dir=str(GRAFICOS_DIR))

    # Capa e cabeçalhos
    report.add_cover_page()
    report.add_headers_footers()

    # Tabelas de dados
    tabelas = gerar_tabelas_resumo()

    # Montar sumário com títulos reais
    toc_titles = ["Resumo Executivo"]
    for num in sorted(chapters.keys()):
        toc_titles.append(CHAPTER_CONFIG[num]["titulo_docx"])
    toc_titles.append("8. Conclusões e Recomendações")
    report.add_table_of_contents(toc_titles)

    # ---- Resumo Executivo ----
    report.doc.add_page_break()
    report.doc.add_heading("Resumo Executivo", level=1)
    if resumo:
        report.add_section_from_markdown(resumo)

    # ---- Capítulos 1-7 ----
    for num in sorted(chapters.keys()):
        ch = chapters[num]
        config = CHAPTER_CONFIG[num]

        # Preparar itens de gráficos
        graph_items = [
            (fname, config["graph_captions"][fname])
            for fname in config["graph_files"]
            if fname in config["graph_captions"]
        ]

        # Preparar diagramas (Cap 1)
        diagram_items = None
        if config.get("diagram_files"):
            diagram_items = [
                (str(DIAGRAMAS_DIR / fname), config["diagram_captions"].get(fname, fname))
                for fname in config["diagram_files"]
            ]

        # Preparar tabela
        tabela_key = config.get("tabela_key")
        tabela = tabelas.get(tabela_key) if tabela_key else None

        # Para Cap 1 (caso especial): conteúdo vai no campo metodologia
        if config.get("special"):
            report.add_chapter_structured(
                title=config["titulo_docx"],
                introducao=ch.introducao,
                tabela=tabela,
                metodologia_text=ch.conteudo_cap1,
                dados_text="",
                graph_items=[],
                graficos_text="",
                parecer_text=ch.parecer,
                diagram_items=diagram_items,
            )
        else:
            report.add_chapter_structured(
                title=config["titulo_docx"],
                introducao=ch.introducao,
                tabela=tabela,
                metodologia_text=ch.metodologia,
                dados_text=ch.dados,
                graph_items=graph_items,
                graficos_text=ch.graficos,
                parecer_text=ch.parecer,
            )

    # ---- Conclusões e Recomendações ----
    report.doc.add_page_break()
    report.doc.add_heading("8. Conclusões e Recomendações", level=1)
    if conclusoes:
        report.add_section_from_markdown(conclusoes)

    # ---- Apêndice: Notebooks Jupyter ----
    NOTEBOOKS = [
        {"path": str(NOTEBOOKS_DIR / nb["file"]), "titulo": nb["titulo"]}
        for nb in NOTEBOOK_LIST
    ]
    logger.info("  Adicionando Apêndice A (Notebooks)...")
    report.add_notebook_appendix(NOTEBOOKS)
    logger.info("  Apêndice concluído.")

    # Salvar
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    output_path = str(REPORTS_DIR / output)
    report.save(output_path)
    return output_path


# =====================================================================
# MAIN
# =====================================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="Gerar relatório de auditoria com Gemini AI (v3 — segmentado)"
    )
    parser.add_argument("--api-key", required=True, help="Gemini API key")
    parser.add_argument("--output", default=OUTPUT_DEFAULT, help="Arquivo DOCX de saída")
    parser.add_argument("--resume", action="store_true",
                       help="Retomar geração usando cache de sub-chamadas anteriores")
    parser.add_argument("--montar", action="store_true",
                       help="Apenas montar DOCX a partir do cache (sem chamadas API)")
    return parser.parse_args()


def run_pipeline(
    api_key: str,
    output: str = OUTPUT_DEFAULT,
    resume: bool = False,
    montar: bool = False,
    on_progress=None,
):
    """
    Executa o pipeline completo de geração do relatório.

    Args:
        api_key: Chave da API Gemini
        output: Nome do arquivo DOCX de saída
        resume: Retomar usando cache existente
        montar: Apenas montar DOCX a partir do cache (sem chamadas API)
        on_progress: Callback opcional chamado a cada etapa com um dict de evento
    """
    start_time = time.time()

    # ================================================================
    # FASE 0: INICIALIZAÇÃO
    # ================================================================
    logger.info("=" * 60)
    logger.info("RELATÓRIO DE AUDITORIA — PIPELINE SEGMENTADO v3")
    logger.info("=" * 60)

    CACHE_DIR.mkdir(exist_ok=True)
    GRAFICOS_DIR.mkdir(parents=True, exist_ok=True)

    met_files = list(METODOLOGIA_DIR.glob("*.md")) if METODOLOGIA_DIR.exists() else []
    logger.info(f"  Metodologia: {len(met_files)} arquivos")
    logger.info(f"  Diagramas: {len(list(DIAGRAMAS_DIR.glob('*.png')))} PNGs")
    logger.info(f"  Cache: {CACHE_DIR}")
    logger.info(f"  Modo: {'montar' if montar else 'resume' if resume else 'gerar'}")

    # ================================================================
    # FASE 0: GERAÇÃO DOS GRÁFICOS (a partir do Excel)
    # ================================================================
    excel_path = DATA_DIR / EXCEL_DEFAULT
    if excel_path.exists():
        logger.info("")
        logger.info("=" * 60)
        logger.info("FASE 0: GERAÇÃO DOS GRÁFICOS")
        logger.info("=" * 60)
        _emit(on_progress, "phase_start", phase=0, phase_name="Geração dos Gráficos")

        def _graph_progress(info):
            step_id = f"graphs_{info['group']}"
            _emit(on_progress, "step_start", step=step_id)
            _emit(on_progress, "step_complete", step=step_id)

        gerados = gerar_todos_graficos(
            excel_path=str(excel_path),
            output_dir=str(GRAFICOS_DIR),
            on_progress=_graph_progress,
        )
        logger.info(f"  Gráficos gerados: {len(gerados)} PNGs")
        _emit(on_progress, "phase_complete", phase=0)
    else:
        logger.warning(f"  Excel não encontrado: {excel_path}")
        logger.info(f"  Gráficos pré-existentes: {len(list(GRAFICOS_DIR.glob('*.png')))} PNGs")

    logger.info(f"  Gráficos disponíveis: {len(list(GRAFICOS_DIR.glob('*.png')))} PNGs")

    client = None
    if not montar:
        client = GeminiAuditClient(api_key=api_key)
        logger.info(f"  Modelo: {GeminiAuditClient.MODEL}")

    # ================================================================
    # FASE 1: GERAÇÃO DOS CAPÍTULOS (26 chamadas — PARALELO)
    # ================================================================
    logger.info("")
    logger.info("=" * 60)
    logger.info("FASE 1: GERAÇÃO DOS CAPÍTULOS (paralelo)")
    logger.info("=" * 60)
    _emit(on_progress, "phase_start", phase=1, phase_name="Geração dos Capítulos")

    chapters = {}
    call_count = 0

    # Preparar dados de cada capítulo (antes de lançar threads)
    cap1_met = load_metodologia(1)
    cap_data = {}
    for n in range(2, 8):
        cap_data[n] = {
            "met": load_metodologia(n),
            "dados": formatar_dados_secao(DATA_CLASSES[n]()) if n in DATA_CLASSES else "",
        }

    # Inicializar ChapterResults
    for n in range(1, 8):
        chapters[n] = ChapterResult(chapter_num=n, titulo=CHAPTER_CONFIG[n]["titulo"])

    # Helper: executa subcall com cliente próprio (thread-safe)
    def _pcall(cache_key, prompt_fn, prompt_args, thinking="low"):
        thread_client = GeminiAuditClient(api_key=api_key) if not montar else None
        text = generate_subcall(
            thread_client or client, cache_key, prompt_fn, prompt_args,
            thinking_level=thinking,
            resume=resume, montar=montar, on_progress=on_progress,
        )
        return cache_key, text

    # --- Wave 1: Seções independentes (A, B, C) — até 19 chamadas em paralelo ---
    logger.info("  Wave 1: Seções independentes (19 chamadas em paralelo)")
    with ThreadPoolExecutor(max_workers=19) as pool:
        futs = []
        # Cap 1A
        futs.append(pool.submit(_pcall, "cap1_a_conteudo", prompt_secao1_conteudo, (cap1_met,), "high"))
        # Cap 2-7: A, B, C
        for n in range(2, 8):
            cd = cap_data[n]
            futs.append(pool.submit(_pcall, f"cap{n}_a_metodologia", prompt_metodologia, (n, cd["met"]), "low"))
            futs.append(pool.submit(_pcall, f"cap{n}_b_dados", prompt_dados, (n, cd["dados"]), "low"))
            if CHAPTER_CONFIG[n]["graph_files"]:
                futs.append(pool.submit(_pcall, f"cap{n}_c_graficos", prompt_graficos, (n,), "high"))
        w1 = {}
        for f in as_completed(futs):
            k, v = f.result()
            w1[k] = v
        call_count += len(futs)

    # Atribuir resultados Wave 1
    chapters[1].conteudo_cap1 = w1.get("cap1_a_conteudo", "")
    for n in range(2, 8):
        chapters[n].metodologia = w1.get(f"cap{n}_a_metodologia", "")
        chapters[n].dados = w1.get(f"cap{n}_b_dados", "")
        chapters[n].graficos = w1.get(f"cap{n}_c_graficos", "")

    # --- Wave 2: Sínteses (B/D) — 7 chamadas em paralelo ---
    logger.info("  Wave 2: Sínteses (7 chamadas em paralelo)")
    with ThreadPoolExecutor(max_workers=7) as pool:
        futs = []
        # Cap 1B (depende de Cap 1A)
        futs.append(pool.submit(_pcall, "cap1_b_sintese", prompt_secao1_sintese, (chapters[1].conteudo_cap1,), "low"))
        # Cap 2-7 D (depende de A, B, C do mesmo capítulo)
        for n in range(2, 8):
            ch = chapters[n]
            futs.append(pool.submit(_pcall, f"cap{n}_d_sintese", prompt_sintese, (n, ch.metodologia, ch.dados, ch.graficos), "low"))
        w2 = {}
        for f in as_completed(futs):
            k, v = f.result()
            w2[k] = v
        call_count += len(futs)

    # Atribuir resultados Wave 2
    chapters[1].introducao, chapters[1].parecer = parse_synthesis(w2.get("cap1_b_sintese", ""))
    for n in range(2, 8):
        chapters[n].introducao, chapters[n].parecer = parse_synthesis(w2.get(f"cap{n}_d_sintese", ""))

    logger.info(f"\n  Fase 1 concluída: {call_count} chamadas (2 waves paralelas)")
    _emit(on_progress, "phase_complete", phase=1)

    # ================================================================
    # FASE 2: CONCLUSÕES (1 chamada)
    # ================================================================
    logger.info("")
    logger.info("=" * 60)
    logger.info("FASE 2: CONCLUSÕES E RECOMENDAÇÕES")
    logger.info("=" * 60)
    _emit(on_progress, "phase_start", phase=2, phase_name="Conclusões e Recomendações")

    all_chapters_text = "\n\n---\n\n".join(
        f"## Capítulo {num}: {ch.titulo}\n{ch.full_text}"
        for num, ch in sorted(chapters.items())
    )

    conclusoes = generate_subcall(
        client, "conclusoes",
        prompt_conclusoes_recomendacoes, (all_chapters_text,),
        thinking_level="high",
        resume=resume, montar=montar, on_progress=on_progress,
    )
    call_count += 1
    _emit(on_progress, "phase_complete", phase=2)

    # ================================================================
    # FASE 3: RESUMO EXECUTIVO (1 chamada)
    # ================================================================
    logger.info("")
    logger.info("=" * 60)
    logger.info("FASE 3: RESUMO EXECUTIVO")
    logger.info("=" * 60)
    _emit(on_progress, "phase_start", phase=3, phase_name="Resumo Executivo")

    resumo = generate_subcall(
        client, "resumo_executivo",
        prompt_resumo_executivo, (all_chapters_text, conclusoes),
        thinking_level="low",
        resume=resume, montar=montar, on_progress=on_progress,
    )
    call_count += 1
    _emit(on_progress, "phase_complete", phase=3)

    # ================================================================
    # FASE 4: MONTAGEM DO DOCX
    # ================================================================
    logger.info("")
    logger.info("=" * 60)
    logger.info("FASE 4: MONTAGEM DO DOCUMENTO DOCX")
    logger.info("=" * 60)
    _emit(on_progress, "phase_start", phase=4, phase_name="Montagem do Documento")
    _emit(on_progress, "step_start", step="docx_assembly")

    output_path = assemble_docx(chapters, resumo, conclusoes, output)

    _emit(on_progress, "step_complete", step="docx_assembly")
    _emit(on_progress, "phase_complete", phase=4)

    elapsed = time.time() - start_time
    logger.info("")
    logger.info("=" * 60)
    logger.info("CONCLUÍDO!")
    logger.info(f"  Relatório: {output_path}")
    logger.info(f"  Chamadas API: {call_count}")
    logger.info(f"  Capítulos: {len(chapters)}")
    logger.info(f"  Tempo total: {elapsed:.0f}s ({elapsed / 60:.1f} min)")
    logger.info("=" * 60)

    return output_path


def main():
    args = parse_args()
    run_pipeline(
        api_key=args.api_key,
        output=args.output,
        resume=args.resume,
        montar=args.montar,
    )


if __name__ == "__main__":
    main()
