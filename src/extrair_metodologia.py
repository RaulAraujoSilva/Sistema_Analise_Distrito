# -*- coding: utf-8 -*-
"""
Extração única de metodologia da apostila ABAR.
Envia o PDF ao Gemini UMA VEZ e extrai textos teóricos para cada seção.
Salva em metodologia/*.md e metodologia/equacoes.json.

Uso:
    python extrair_metodologia.py --api-key SUA_CHAVE_API
"""
import argparse
import json
import logging
import os
import time
from pathlib import Path

from gemini_client import GeminiAuditClient

from config import APOSTILA_DIR, METODOLOGIA_DIR

PDF_PATH = APOSTILA_DIR / "APOSTILA COMPLETA_Curso ABAR_Medições Inteligentes e Gestão Integrada_rev0 (1).pdf"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# Seções e seus temas para extração
SECOES_EXTRACAO = [
    {
        "arquivo": "secao_1_visao_geral.md",
        "tema": "Visão Geral - Análise de Condições de Operação de Distritos",
        "prompt": """Extraia do documento anexo APENAS os conceitos teóricos, definições e procedimentos
relevantes para uma ANÁLISE DE CONDIÇÕES DE OPERAÇÃO DE UM DISTRITO DE DISTRIBUIÇÃO DE GÁS NATURAL.

Inclua:
- O que é uma análise de condições de operação e sua importância regulatória
- Quais elementos são obrigatórios em uma verificação de distrito
- Conceitos de transferência de custódia e medição fiscal
- Estrutura típica de um distrito (entrada, rede, pontos de saída)
- Requisitos de dados para auditoria (periodicidade, variáveis, completude)

Retorne como texto em português formal. NÃO referencie páginas, slides, blocos ou o documento.
Apresente como conhecimento técnico direto de um especialista.""",
    },
    {
        "arquivo": "secao_2_volumes.md",
        "tema": "Medição Volumétrica e Transferência de Custódia",
        "prompt": """Extraia do documento anexo APENAS os conceitos teóricos sobre MEDIÇÃO VOLUMÉTRICA
e TRANSFERÊNCIA DE CUSTÓDIA em gás natural.

Inclua:
- Princípios da medição fiscal de gás natural
- Conceito de transferência de custódia entre transportadora e distribuidora
- Unidades de medição (Nm³, m³/h, condições de referência)
- Critérios de concordância entre medições independentes
- Tolerâncias aceitáveis para diferenças entre medições fiscais
- Fatores que afetam a medição volumétrica (pressão, temperatura, composição)

Retorne como texto em português formal. NÃO referencie páginas, slides, blocos ou o documento.
Apresente como conhecimento técnico direto.""",
    },
    {
        "arquivo": "secao_3_pcs.md",
        "tema": "Poder Calorífico Superior e Cromatografia Gasosa",
        "prompt": """Extraia do documento anexo APENAS os conceitos teóricos sobre o PODER CALORÍFICO
SUPERIOR (PCS) e CROMATOGRAFIA GASOSA para gás natural.

Inclua:
- Definição do PCS e sua importância para comercialização de gás
- Método de determinação por cromatografia gasosa em linha
- Faixas típicas de PCS para gás natural brasileiro (kcal/m³)
- Portaria INMETRO e regulamentos aplicáveis à qualidade do gás
- Relação entre PCS e composição do gás
- Fatores que causam variação do PCS ao longo do tempo

Retorne como texto em português formal. NÃO referencie páginas, slides, blocos ou o documento.
Apresente como conhecimento técnico direto.""",
    },
    {
        "arquivo": "secao_4_energia.md",
        "tema": "Cálculo de Energia em Gás Natural",
        "prompt": """Extraia do documento anexo APENAS os conceitos teóricos sobre CÁLCULO DE ENERGIA
em gás natural.

Inclua:
- Fórmula fundamental: Energia = Volume × PCS
- Unidades de energia usadas (Gcal, MJ, kWh)
- Importância da energia como grandeza de faturamento ("Você fatura o que mede")
- Diferença entre medir volume e medir energia
- Papel do computador de vazão no cálculo de energia
- Validação e verificação do cálculo de energia

Retorne como texto em português formal. NÃO referencie páginas, slides, blocos ou o documento.
Apresente como conhecimento técnico direto.""",
    },
    {
        "arquivo": "secao_5_clientes.md",
        "tema": "Perfis de Consumo e Adequabilidade de Medidores",
        "prompt": """Extraia do documento anexo APENAS os conceitos teóricos sobre PERFIS DE CONSUMO
de clientes e ADEQUABILIDADE DE MEDIDORES em distribuição de gás natural.

Inclua:
- Importância da análise de perfis individuais de consumo
- Conceito de fator de carga e sua interpretação
- Verificação de condições operacionais (pressão, temperatura, vazão)
- Faixas de operação dos medidores e critérios de adequabilidade
- Consequências de operação fora da faixa do medidor
- Tratamento de dados faltantes e seu impacto na rastreabilidade

Retorne como texto em português formal. NÃO referencie páginas, slides, blocos ou o documento.
Apresente como conhecimento técnico direto.""",
    },
    {
        "arquivo": "secao_6_incertezas.md",
        "tema": "Incertezas de Medição conforme GUM",
        "prompt": """Extraia do documento anexo APENAS os conceitos teóricos sobre INCERTEZAS DE MEDIÇÃO
conforme o GUM (Guia para Expressão da Incerteza de Medição).

Inclua:
- Definição de incerteza de medição e sua importância regulatória
- Incerteza Tipo A (avaliação estatística) e Tipo B (certificados, especificações)
- Método RSS (raiz da soma dos quadrados) para combinar incertezas independentes
- Fator de abrangência k=2 para 95% de confiança
- Incerteza expandida e seu significado prático
- Limites de referência: limite fiscal e limite de apropriação
- Memorial de cálculo de incerteza para medição de gás

Retorne como texto em português formal. NÃO referencie páginas, slides, blocos ou o documento.
Apresente como conhecimento técnico direto.""",
    },
    {
        "arquivo": "secao_7_balanco.md",
        "tema": "Balanço de Massa com Bandas de Incerteza",
        "prompt": """Extraia do documento anexo APENAS os conceitos teóricos sobre BALANÇO DE MASSA
(ou balanço volumétrico) em distritos de distribuição de gás natural.

Inclua:
- Definição e fórmula do balanço de massa: Dif% = (Entrada - Soma_Saídas) / Entrada × 100
- Conceito de bandas de incerteza: V_min = V × (1 - U%), V_max = V × (1 + U%)
- Critério de aceitação: sobreposição das bandas de incerteza
- Interpretação: quando a diferença é coberta pelas incertezas
- Causas possíveis de diferenças no balanço (perdas técnicas, derivas instrumentais)
- Ações recomendadas quando o balanço não fecha

Retorne como texto em português formal. NÃO referencie páginas, slides, blocos ou o documento.
Apresente como conhecimento técnico direto.""",
    },
]

PROMPT_EQUACOES = """Extraia do documento anexo TODAS as equações e fórmulas matemáticas relevantes
para auditoria de medição de gás natural. Para cada equação, forneça:

1. Nome descritivo
2. A equação em formato LaTeX
3. Breve descrição do que calcula

Retorne EXCLUSIVAMENTE um JSON válido (sem markdown, sem blocos de código) com a seguinte estrutura:
[
    {"nome": "Nome da equação", "latex": "fórmula em LaTeX", "descricao": "O que calcula"},
    ...
]

Equações que DEVEM estar incluídas (se presentes no documento):
- Diferença percentual do balanço de massa
- Incerteza combinada RSS (raiz da soma dos quadrados)
- Incerteza expandida (U = k × u_c)
- Energia (E = V × PCS)
- Bandas de incerteza (V_min, V_max)
- Conversão de condições de referência (pressão, temperatura)
- Fator de compressibilidade
- Qualquer outra fórmula relevante para medição de gás

Retorne APENAS o JSON, sem texto adicional."""


def find_pdf():
    """Encontra o PDF da apostila."""
    if PDF_PATH.exists():
        return str(PDF_PATH)
    for f in BASE_DIR.glob("APOSTILA COMPLETA*.pdf"):
        return str(f)
    raise FileNotFoundError("Apostila PDF não encontrada")


def main():
    parser = argparse.ArgumentParser(description="Extrair metodologia da apostila ABAR")
    parser.add_argument("--api-key", required=True, help="Gemini API key")
    args = parser.parse_args()

    METODOLOGIA_DIR.mkdir(exist_ok=True)

    logger.info("=" * 60)
    logger.info("EXTRAÇÃO DE METODOLOGIA DA APOSTILA ABAR")
    logger.info("=" * 60)

    # Inicializar cliente e fazer upload do PDF
    client = GeminiAuditClient(api_key=args.api_key)
    pdf_path = find_pdf()
    logger.info(f"PDF: {pdf_path}")
    client.upload_pdf(pdf_path)

    # Extrair teoria para cada seção
    logger.info("")
    logger.info("EXTRAINDO TEORIA POR SEÇÃO...")
    logger.info("-" * 40)

    for i, secao in enumerate(SECOES_EXTRACAO, 1):
        arquivo = secao["arquivo"]
        output_path = METODOLOGIA_DIR / arquivo

        # Verificar se já existe
        if output_path.exists() and output_path.stat().st_size > 100:
            logger.info(f"  [{i}/7] {secao['tema']} -> já existe ({output_path.stat().st_size} bytes)")
            continue

        logger.info(f"  [{i}/7] {secao['tema']}")
        text = client.analyze_section(
            system_prompt="Você é um especialista em medição de gás natural. Extraia o conhecimento técnico solicitado.",
            section_prompt=secao["prompt"],
            image_paths=[],
            thinking_level="low",
            include_pdf=True,
        )

        output_path.write_text(text, encoding="utf-8")
        logger.info(f"    -> Salvo: {arquivo} ({len(text)} caracteres)")

    # Extrair equações
    logger.info("")
    logger.info("EXTRAINDO EQUAÇÕES...")
    logger.info("-" * 40)

    equacoes_path = METODOLOGIA_DIR / "equacoes.json"
    if equacoes_path.exists() and equacoes_path.stat().st_size > 50:
        logger.info("  Equações já extraídas.")
    else:
        text = client.analyze_section(
            system_prompt="Você é um especialista em medição de gás natural. Extraia as equações solicitadas.",
            section_prompt=PROMPT_EQUACOES,
            image_paths=[],
            thinking_level="low",
            include_pdf=True,
        )

        # Tentar parsear JSON (Gemini pode retornar com markdown)
        json_text = text.strip()
        if json_text.startswith("```"):
            lines = json_text.split("\n")
            json_text = "\n".join(lines[1:-1])

        try:
            equacoes = json.loads(json_text)
            equacoes_path.write_text(json.dumps(equacoes, ensure_ascii=False, indent=2), encoding="utf-8")
            logger.info(f"  -> {len(equacoes)} equações extraídas e salvas")
        except json.JSONDecodeError:
            # Salvar texto bruto para inspeção manual
            equacoes_path.write_text(json_text, encoding="utf-8")
            logger.warning(f"  -> JSON inválido. Texto bruto salvo para revisão ({len(json_text)} chars)")

    logger.info("")
    logger.info("=" * 60)
    logger.info("EXTRAÇÃO CONCLUÍDA!")
    files = list(METODOLOGIA_DIR.glob("*"))
    logger.info(f"  Arquivos em metodologia/: {len(files)}")
    for f in sorted(files):
        logger.info(f"    {f.name} ({f.stat().st_size / 1024:.1f} KB)")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()
