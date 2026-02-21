# -*- coding: utf-8 -*-
"""
Geração de diagramas de processo de auditoria usando Gemini Image API.
Gera 3 diagramas profissionais para a seção introdutória do relatório.

Uso:
    python gerar_diagramas.py --api-key SUA_CHAVE_API
"""
import argparse
import logging
from pathlib import Path

from gemini_client import GeminiAuditClient

from config import DIAGRAMAS_DIR

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

DIAGRAMAS = [
    {
        "arquivo": "fluxo_auditoria.png",
        "titulo": "Fluxo Geral da Auditoria",
        "prompt": """Crie um diagrama de fluxo profissional para um processo de auditoria de medição de gás natural.
IDIOMA: Todos os textos DEVEM estar em português brasileiro.

O diagrama deve mostrar 5 fases sequenciais conectadas por setas direcionais, da esquerda para a direita:

Fase 1: Retângulo arredondado com o texto "Coleta e Validação de Dados"
Fase 2: Retângulo arredondado com o texto "Verificação das Condições de Operação"
Fase 3: Retângulo arredondado com o texto "Análise de Incertezas de Medição"
Fase 4: Retângulo arredondado com o texto "Balanço de Massa"
Fase 5: Retângulo arredondado com o texto "Parecer Regulatório"

Acima de cada retângulo, coloque um número de fase: "Fase 1", "Fase 2", "Fase 3", "Fase 4", "Fase 5".

Título no topo: "Fluxo da Auditoria Técnica de Medição de Gás Natural"

Estilo: corporativo, limpo, profissional. Esquema de cores: tons de azul escuro e azul claro.
Fundo branco. Texto em azul escuro. Setas em azul médio.
Orientação horizontal (paisagem). Alta resolução.""",
    },
    {
        "arquivo": "processo_analise.png",
        "titulo": "Processo de Análise de Dados",
        "prompt": """Crie um infográfico profissional mostrando a metodologia de análise de dados para auditoria de distrito de gás natural.
IDIOMA: Todos os textos DEVEM estar em português brasileiro.

O infográfico deve mostrar as seguintes etapas conectadas por setas, de cima para baixo ou da esquerda para a direita:

Etapa 1: "Planilha Excel" com subtítulo "14 abas de dados"
Etapa 2: "Processamento Python" com subtítulo "Pandas e NumPy"
Etapa 3: "Validação Cruzada" com subtítulo "Concessionária vs Transportadora"
Etapa 4: "Cálculo de Energia" com subtítulo "E = V × PCS"
Etapa 5: "Perfis de Consumo" com subtítulo "7 clientes industriais"
Etapa 6: "Incertezas GUM" com subtítulo "Combinação RSS"
Etapa 7: "Balanço de Massa" com subtítulo "Bandas de incerteza"
Etapa 8: "Relatório Final" com subtítulo "Documento DOCX"

Título no topo: "Metodologia de Análise de Dados"

Estilo: limpo, minimalista, com ícones simples para cada etapa.
Esquema de cores: azul escuro para cabeçalhos, azul claro para acentos, fundo branco.
Profissional e corporativo. Alta resolução.""",
    },
    {
        "arquivo": "estrutura_distrito.png",
        "titulo": "Estrutura do Distrito",
        "prompt": """Crie um diagrama técnico profissional de um distrito de distribuição de gás natural.
IDIOMA: Todos os textos DEVEM estar em português brasileiro.

LADO ESQUERDO - PONTO DE ENTRADA:
- Retângulo grande com o texto "Estação de Recebimento" e subtítulo "City-Gate"
- Dentro, dois tramos de medição: "Tramo 101" e "Tramo 501"
- Abaixo: "Entrada: 182,9 milhões Nm³"

CENTRO:
- Linhas de tubulação ramificando da entrada para os 7 pontos de saída

LADO DIREITO - 7 PONTOS DE SAÍDA (clientes), mostrados como retângulos de tamanhos proporcionais ao volume:
1. "Empresa A" com texto "57,5%"
2. "Empresa B" com texto "24,1%"
3. "Empresa E" com texto "5,6%"
4. "Empresa G" com texto "5,6%"
5. "Empresa C" com texto "3,8%"
6. "Empresa F" com texto "3,3%"
7. "Empresa D" com texto "0,05%" e um ícone de alerta amarelo

PARTE INFERIOR:
- Barra resumo: "Saída Total: 180,9 milhões Nm³ | Diferença: 1,09%"

Título no topo: "Estrutura do Distrito de Distribuição de Gás Natural"

Estilo: técnico, profissional, tons de azul corporativo.
Linhas limpas, fundo branco. Alta resolução.""",
    },
]


def main():
    parser = argparse.ArgumentParser(description="Gerar diagramas de processo com Gemini")
    parser.add_argument("--api-key", required=True, help="Gemini API key")
    args = parser.parse_args()

    DIAGRAMAS_DIR.mkdir(exist_ok=True)

    logger.info("=" * 60)
    logger.info("GERAÇÃO DE DIAGRAMAS DE PROCESSO")
    logger.info(f"Modelo: {GeminiAuditClient.IMAGE_MODEL}")
    logger.info("=" * 60)

    client = GeminiAuditClient(api_key=args.api_key)

    for i, diag in enumerate(DIAGRAMAS, 1):
        output_path = DIAGRAMAS_DIR / diag["arquivo"]

        # Sempre regenerar (deletar existente)
        if output_path.exists():
            output_path.unlink()

        logger.info(f"  [{i}/3] {diag['titulo']}")
        success = client.generate_image(diag["prompt"], str(output_path))

        if success:
            logger.info(f"    -> Diagrama salvo com sucesso")
        else:
            logger.warning(f"    -> Falha na geração do diagrama")

    logger.info("")
    logger.info("=" * 60)
    logger.info("DIAGRAMAS CONCLUÍDOS!")
    files = list(DIAGRAMAS_DIR.glob("*.png"))
    logger.info(f"  Diagramas gerados: {len(files)}")
    for f in sorted(files):
        logger.info(f"    {f.name} ({f.stat().st_size / 1024:.0f} KB)")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()
