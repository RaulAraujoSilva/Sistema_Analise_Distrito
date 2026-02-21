# -*- coding: utf-8 -*-
"""
Templates de prompts para o relatório de auditoria — Versão 3 (segmentada).

Cada capítulo é gerado em 4 chamadas LLM separadas:
  A. Metodologia → B. Dados → C. Gráficos → D. Síntese

Cada função retorna (system_prompt, section_prompt, lista_de_graficos_png).
"""
# =====================================================================
# DIRETÓRIO DE DIAGRAMAS (para Capítulo 1)
# =====================================================================
from config import DIAGRAMAS_DIR

# =====================================================================
# SYSTEM PROMPTS ESPECIALIZADOS
# =====================================================================

SYSTEM_PROMPT_BASE = """Você é um auditor técnico especialista em medição de gás natural, atuando em uma Agência Regulatória.
Você está elaborando um relatório técnico de auditoria sobre as condições de operação de um
distrito de distribuição de gás natural, com base nos dados do período de abril a setembro de 2025.

REGRAS OBRIGATÓRIAS:
1. NÃO cite nomes de documentos, apostilas, slides, páginas, blocos ou fontes externas.
   Apresente todo conhecimento técnico como seu próprio parecer profissional de auditor.
2. Use terminologia técnica correta: Nm³, m³/h, PCS (kcal/m³), incerteza expandida, RSS, GUM, etc.
3. Escreva em português brasileiro formal, com acentuação correta, sem emojis.
4. Use formato Markdown com cabeçalhos hierárquicos (##, ###).
5. Seja específico: use números com separador de milhares e casas decimais adequadas.
6. Formate equações usando LaTeX: $$ para display e $ para inline.
7. Use **negrito** para termos técnicos importantes e conclusões-chave.
8. Inclua tabelas Markdown quando apropriado para organizar dados.
"""

SYSTEM_PROMPT_METODOLOGIA = SYSTEM_PROMPT_BASE + """
TAREFA ESPECÍFICA: Escreva APENAS a subseção "Fundamentação Teórica" deste capítulo.
- Explique os conceitos teóricos, normas técnicas e critérios regulatórios relevantes.
- Inclua equações fundamentais em LaTeX quando aplicável.
- NÃO analise dados numéricos específicos do distrito.
- NÃO discuta gráficos.
- NÃO emita parecer regulatório.
- Foco: base teórica para o leitor compreender a análise que virá a seguir.
"""

SYSTEM_PROMPT_DADOS = SYSTEM_PROMPT_BASE + """
TAREFA ESPECÍFICA: Escreva APENAS a subseção "Análise dos Dados" deste capítulo.
- Analise os dados estatísticos fornecidos. Compare com limites normativos quando existirem.
- Inclua tabelas Markdown para organizar valores importantes.
- NÃO repita a fundamentação teórica (já foi escrita).
- NÃO discuta gráficos (será feito em separado).
- NÃO emita parecer regulatório.
- Foco: o que os números revelam sobre o distrito.
"""

SYSTEM_PROMPT_GRAFICOS = SYSTEM_PROMPT_BASE + """
TAREFA ESPECÍFICA: Escreva APENAS a subseção "Discussão dos Gráficos" deste capítulo.
- Você receberá imagens dos gráficos gerados a partir dos dados do distrito.
- Descreva o que se observa em CADA gráfico, referenciando pelo número da figura.
  Exemplo: "Na Figura 2.1, observa-se que..."
- Aponte tendências, anomalias, padrões sazonais, outliers.
- NÃO repita a fundamentação teórica (já foi escrita).
- NÃO repita a análise numérica dos dados (já foi escrita).
- NÃO emita parecer regulatório.
- Foco: interpretação visual que complementa a análise numérica.
"""

SYSTEM_PROMPT_SINTESE = SYSTEM_PROMPT_BASE + """
TAREFA ESPECÍFICA: Com base nas subseções já escritas (que você receberá como contexto),
escreva DUAS partes para este capítulo:

1. **Introdução do Capítulo** (2-3 parágrafos): Resumo do que este capítulo aborda,
   sua importância no contexto da auditoria, e os principais achados.
   Comece diretamente com o texto, SEM cabeçalho "### Introdução".

2. **Parecer Regulatório**: Avaliação final do auditor sobre este aspecto específico.
   Indique se os valores/condições são NORMAIS, ACEITÁVEIS, PREOCUPANTES ou NÃO CONFORMES.
   Inclua ações recomendadas quando necessário.
   Comece com o cabeçalho "### Parecer Regulatório".

Separe as duas partes com uma linha contendo apenas: ---SEPARADOR---

NÃO repita o conteúdo das subseções. Sintetize e conclua.
"""

SYSTEM_PROMPT_CONCLUSOES = SYSTEM_PROMPT_BASE + """
TAREFA ESPECÍFICA: Escreva as Conclusões e Recomendações finais do relatório completo.
Você receberá o texto de TODOS os capítulos como contexto.
Use tom formal e definitivo. Este é o encerramento do relatório.
"""

SYSTEM_PROMPT_RESUMO = SYSTEM_PROMPT_BASE + """
TAREFA ESPECÍFICA: Escreva o Resumo Executivo do relatório.
Você receberá o texto de TODOS os capítulos e das Conclusões como contexto.
O Resumo Executivo deve ser compreensível por gestores não-técnicos, sem perder o rigor técnico.
Máximo 2 páginas.
"""

# =====================================================================
# CONFIGURAÇÃO DE CADA CAPÍTULO
# =====================================================================

CHAPTER_CONFIG = {
    1: {
        "titulo": "Visão Geral do Distrito e Dados Disponíveis",
        "titulo_docx": "1. Visão Geral do Distrito e Dados Disponíveis",
        "metodologia_file": "secao_1_visao_geral.md",
        "special": True,
        "tabela_key": None,
        "graph_files": [],
        "graph_captions": {},
        "diagram_files": [
            "estrutura_distrito.png",
            "fluxo_auditoria.png",
            "processo_analise.png",
        ],
        "diagram_captions": {
            "estrutura_distrito.png": "Diagrama: Estrutura do Distrito de Distribuição",
            "fluxo_auditoria.png": "Diagrama: Fluxo da Auditoria Técnica",
            "processo_analise.png": "Diagrama: Processo de Análise de Dados",
        },
        "tema_metodologia": "auditoria de condições de operação de um distrito de distribuição de gás natural, transferência de custódia, balanço de gás, rastreabilidade metrológica",
        "tema_dados": "",
        "graph_descriptions": "",
    },
    2: {
        "titulo": "Análise de Volumes de Entrada",
        "titulo_docx": "2. Análise de Volumes de Entrada",
        "metodologia_file": "secao_2_volumes.md",
        "special": False,
        "tabela_key": "secao_2_volumes",
        "graph_files": [
            "vol_entrada_serie.png",
            "vol_entrada_diferencas.png",
            "vol_entrada_histograma.png",
            "vol_entrada_boxplot.png",
        ],
        "graph_captions": {
            "vol_entrada_serie.png": "Figura 2.1: Série temporal de volumes diários de entrada",
            "vol_entrada_diferencas.png": "Figura 2.2: Diferenças entre Concessionária e Transportadora",
            "vol_entrada_histograma.png": "Figura 2.3: Distribuição das diferenças volumétricas",
            "vol_entrada_boxplot.png": "Figura 2.4: Boxplots mensais dos volumes de entrada",
        },
        "diagram_files": [],
        "diagram_captions": {},
        "tema_metodologia": "medição fiscal de gás natural, transferência de custódia, concordância entre sistemas de medição, limites de tolerância para diferenças volumétricas",
        "tema_dados": "volume médio diário, faixa de variação (mín/máx), diferença entre Concessionária e Transportadora, sazonalidade",
        "graph_descriptions": """1. Série temporal de volumes diários de entrada (Concessionária vs Transportadora) ao longo de 183 dias, com linha de média
2. Diferenças absolutas (Nm³) e percentuais (%) entre Concessionária e Transportadora ao longo do tempo
3. Histogramas da distribuição das diferenças absolutas e percentuais
4. Boxplots mensais dos volumes de entrada (abril a setembro)""",
    },
    3: {
        "titulo": "Análise do Poder Calorífico Superior (PCS)",
        "titulo_docx": "3. Análise do Poder Calorífico Superior (PCS)",
        "metodologia_file": "secao_3_pcs.md",
        "special": False,
        "tabela_key": "secao_3_pcs",
        "graph_files": [
            "pcs_serie.png",
            "pcs_histograma.png",
        ],
        "graph_captions": {
            "pcs_serie.png": "Figura 3.1: Série temporal do PCS diário",
            "pcs_histograma.png": "Figura 3.2: Distribuição do PCS",
        },
        "diagram_files": [],
        "diagram_captions": {},
        "tema_metodologia": "Poder Calorífico Superior (PCS), cromatografia gasosa, faixas típicas de PCS para gás natural brasileiro, importância para comercialização de energia",
        "tema_dados": "PCS médio, faixa de variação, desvio padrão, concordância entre Concessionária e Transportadora",
        "graph_descriptions": """1. Série temporal do PCS diário (Concessionária vs Transportadora) ao longo de 183 dias
2. Histograma da distribuição do PCS com linhas de média e mediana""",
    },
    4: {
        "titulo": "Cálculo e Validação de Energia",
        "titulo_docx": "4. Cálculo e Validação de Energia",
        "metodologia_file": "secao_4_energia.md",
        "special": False,
        "tabela_key": "secao_4_energia",
        "graph_files": [
            "energia_serie.png",
            "energia_diferencas.png",
            "energia_mensal.png",
            "energia_scatter.png",
        ],
        "graph_captions": {
            "energia_serie.png": "Figura 4.1: Série temporal da energia diária",
            "energia_diferencas.png": "Figura 4.2: Diferenças de energia Conc. vs Transp.",
            "energia_mensal.png": "Figura 4.3: Energia acumulada mensal",
            "energia_scatter.png": "Figura 4.4: Correlação Volume × Energia × PCS",
        },
        "diagram_files": [],
        "diagram_captions": {},
        "tema_metodologia": "cálculo de energia (E = V × PCS), validação computacional, correlação volume-energia, importância do faturamento energético",
        "tema_dados": "energia média diária, energia total no período, validação do cálculo contra a planilha, correlação r entre volume e energia",
        "graph_descriptions": """1. Série temporal da energia diária (Concessionária vs Transportadora) em Gcal/dia
2. Diferenças de energia (Gcal) entre Concessionária e Transportadora ao longo do tempo
3. Energia acumulada mensal — gráfico de barras agrupadas (Concessionária vs Transportadora)
4. Scatter plot Volume vs Energia colorido pelo PCS, com linha de tendência e r = 0,999999""",
    },
    5: {
        "titulo": "Perfis de Consumo dos Clientes",
        "titulo_docx": "5. Perfis de Consumo dos Clientes",
        "metodologia_file": "secao_5_clientes.md",
        "special": False,
        "tabela_key": "secao_5_clientes",
        "graph_files": [
            "clientes_participacao.png",
            "clientes_serie.png",
            "clientes_perfil_horario.png",
            "clientes_heatmap.png",
            "clientes_pressao_temp.png",
            "clientes_boxplot.png",
        ],
        "graph_captions": {
            "clientes_participacao.png": "Figura 5.1: Participação volumétrica dos clientes",
            "clientes_serie.png": "Figura 5.2: Séries temporais de consumo por cliente",
            "clientes_perfil_horario.png": "Figura 5.3: Perfis médios horários por cliente",
            "clientes_heatmap.png": "Figura 5.4: Heatmap de consumo da Empresa A",
            "clientes_pressao_temp.png": "Figura 5.5: Condições operacionais (pressão e temperatura)",
            "clientes_boxplot.png": "Figura 5.6: Distribuição comparativa de volumes por cliente",
        },
        "diagram_files": [],
        "diagram_captions": {},
        "tema_metodologia": "perfis de consumo individual, fator de carga, condições operacionais (pressão, temperatura), faixas de operação dos medidores, rangeabilidade",
        "tema_dados": "volume total e participação percentual de cada cliente, fator de carga, vazão média/mín/máx, condições de pressão e temperatura, dados faltantes do Empresa D",
        "graph_descriptions": """1. Dois painéis: barras horizontais de volume total por cliente + gráfico donut de participação
2. 7 painéis com séries temporais do volume horário por cliente (com média móvel 24h)
3. 7 painéis com perfil médio horário por cliente (hora do dia, com banda de ±1 desvio padrão)
4. Heatmap de consumo da Empresa A (Hora × Dia) — mostra padrões de operação contínua e paradas
5. 7 painéis com pressão (eixo esquerdo) e temperatura (eixo direito) por cliente
6. Boxplots comparativos de distribuição de volume por cliente""",
    },
    6: {
        "titulo": "Cálculo de Incertezas de Medição",
        "titulo_docx": "6. Cálculo de Incertezas de Medição",
        "metodologia_file": "secao_6_incertezas.md",
        "special": False,
        "tabela_key": "secao_6_incertezas",
        "graph_files": [
            "incertezas_barras.png",
            "incertezas_rss.png",
            "incertezas_contribuicao.png",
        ],
        "graph_captions": {
            "incertezas_barras.png": "Figura 6.1: Incerteza por ponto de medição",
            "incertezas_rss.png": "Figura 6.2: Incerteza combinada RSS — Entrada vs Saída",
            "incertezas_contribuicao.png": "Figura 6.3: Contribuição de cada cliente na incerteza total",
        },
        "diagram_files": [],
        "diagram_captions": {},
        "tema_metodologia": "incerteza de medição conforme GUM, Tipo A e Tipo B, método RSS (raiz da soma dos quadrados), fator de abrangência k=2, incerteza expandida, limites regulatórios (1% fiscal, 3% apropriação)",
        "tema_dados": "incerteza individual de cada ponto de medição, incerteza combinada de entrada (1,52%) vs saída (6,19%), contribuição de cada cliente na incerteza total",
        "graph_descriptions": """1. Barras horizontais de incerteza por ponto de medição, com linhas de referência para limite fiscal (1%) e limite de apropriação (3%), cores diferenciadas entrada/saída
2. Barras comparativas da incerteza combinada RSS: Entrada (1,52%) vs Saída (6,19%), com linha de referência do limite fiscal
3. Dois painéis: barras empilhadas + gráfico pizza mostrando contribuição de cada cliente para a incerteza total de saída""",
        "equacoes_extras": """
IMPORTANTE: Use equações LaTeX para as fórmulas de incerteza:
$$U_c = \\sqrt{u_1^2 + u_2^2 + \\cdots + u_n^2}$$
$$U = k \\cdot U_c \\quad (k=2, \\text{95\\% de confiança})$$""",
    },
    7: {
        "titulo": "Balanço de Massa com Bandas de Incerteza",
        "titulo_docx": "7. Balanço de Massa com Bandas de Incerteza",
        "metodologia_file": "secao_7_balanco.md",
        "special": False,
        "tabela_key": "secao_7_balanco",
        "graph_files": [
            "balanco_barras.png",
            "balanco_waterfall.png",
            "balanco_bandas.png",
            "balanco_dashboard.png",
        ],
        "graph_captions": {
            "balanco_barras.png": "Figura 7.1: Entrada vs Saída com bandas de incerteza",
            "balanco_waterfall.png": "Figura 7.2: Decomposição waterfall do balanço",
            "balanco_bandas.png": "Figura 7.3: Sobreposição das bandas de incerteza",
            "balanco_dashboard.png": "Figura 7.4: Dashboard do resultado do balanço",
        },
        "diagram_files": [],
        "diagram_captions": {},
        "tema_metodologia": "balanço de massa em sistemas de distribuição de gás, fórmula da diferença percentual, bandas de incerteza, critério de aceitação por sobreposição de bandas",
        "tema_dados": "diferença de 1,09% entre entrada e saída, volumes transferidos, bandas de incerteza de entrada e saída, sobreposição das bandas",
        "graph_descriptions": """1. Gráfico de barras Entrada vs Saída Total com barras de erro (bandas de incerteza)
2. Gráfico waterfall decompondo o balanço: Entrada menos cada cliente = Diferença
3. Visualização horizontal de sobreposição de bandas: banda de entrada vs banda de saída
4. Dashboard de 3 painéis: barras de volumes, gauge mostrando 1,09%, painel resultado""",
        "equacoes_extras": """
IMPORTANTE: Use equações LaTeX:
$$\\text{Dif\\%} = \\frac{V_{entrada} - \\sum V_{saída}}{V_{entrada}} \\times 100$$
$$V_{min} = V \\times (1 - U\\%), \\quad V_{max} = V \\times (1 + U\\%)$$

Esta é a seção mais crítica do relatório. O parecer sobre o balanço é a conclusão central.""",
    },
}


# =====================================================================
# FUNÇÕES DE GERAÇÃO DE PROMPTS — CHAMADAS SEGMENTADAS
# =====================================================================

def prompt_metodologia(cap_num: int, metodologia_text: str):
    """Chamada A: Gera subseção de Fundamentação Teórica."""
    config = CHAPTER_CONFIG[cap_num]
    equacoes = config.get("equacoes_extras", "")

    prompt = f"""## Capítulo {cap_num}: {config['titulo']}

### Contexto Metodológico (extraído previamente)
{metodologia_text}

### Tarefa
Escreva a subseção **Fundamentação Teórica** para este capítulo.
Tema central: {config['tema_metodologia']}.
{equacoes}
"""
    return SYSTEM_PROMPT_METODOLOGIA, prompt, []


def prompt_dados(cap_num: int, dados_texto: str, metodologia_text: str = ""):
    """Chamada B: Gera subseção de Análise dos Dados."""
    config = CHAPTER_CONFIG[cap_num]

    prompt = f"""## Capítulo {cap_num}: {config['titulo']}

### Dados Computados
{dados_texto}

### Tarefa
Escreva a subseção **Análise dos Dados** para este capítulo.
Foco: {config['tema_dados']}.
Compare os valores observados com limites e tolerâncias normativos quando existirem.
"""
    return SYSTEM_PROMPT_DADOS, prompt, []


def prompt_graficos(cap_num: int):
    """Chamada C: Gera subseção de Discussão dos Gráficos."""
    config = CHAPTER_CONFIG[cap_num]

    # Montar descrições com numeração de figuras
    desc_lines = []
    for i, (fname, caption) in enumerate(config["graph_captions"].items(), 1):
        desc_lines.append(f"{i}. {caption}")

    prompt = f"""## Capítulo {cap_num}: {config['titulo']}

### Gráficos Anexos ({len(config['graph_files'])} imagens)
{config['graph_descriptions']}

### Numeração das Figuras
{chr(10).join(desc_lines)}

### Tarefa
Descreva e interprete CADA gráfico anexo.
Use a numeração de figuras fornecida (ex: "Na Figura {cap_num}.1, observa-se...").
Aponte tendências, anomalias, padrões sazonais e outliers relevantes.
"""
    return SYSTEM_PROMPT_GRAFICOS, prompt, config["graph_files"]


def prompt_sintese(cap_num: int, met_text: str, dados_text: str, graf_text: str):
    """Chamada D: Gera Introdução do capítulo + Parecer Regulatório."""
    config = CHAPTER_CONFIG[cap_num]

    prompt = f"""## Capítulo {cap_num}: {config['titulo']}

Abaixo estão as subseções já escritas para este capítulo:

--- FUNDAMENTAÇÃO TEÓRICA ---
{met_text}

--- ANÁLISE DOS DADOS ---
{dados_text}

--- DISCUSSÃO DOS GRÁFICOS ---
{graf_text}

### Tarefa
Com base no conteúdo acima, escreva:

1. **Introdução do Capítulo** (2-3 parágrafos): Apresente o contexto e a importância
   deste aspecto da auditoria. Sintetize os principais achados sem repetir os detalhes.
   NÃO use cabeçalho — comece direto com o texto.

2. **Parecer Regulatório**: Comece com "### Parecer Regulatório".
   Emita sua avaliação técnica final: CONFORME, NÃO CONFORME, ou ATENÇÃO NECESSÁRIA.
   Inclua ações recomendadas quando necessário.

Separe as duas partes com uma linha contendo apenas: ---SEPARADOR---
"""
    return SYSTEM_PROMPT_SINTESE, prompt, []


# =====================================================================
# CAPÍTULO 1 — CASO ESPECIAL (conteúdo + síntese)
# =====================================================================

def prompt_secao1_conteudo(metodologia_text: str):
    """Capítulo 1: Conteúdo com diagramas passados como imagens."""
    prompt = f"""## Capítulo 1: Visão Geral do Distrito e Dados Disponíveis

### Contexto Metodológico
{metodologia_text}

### Contexto dos Dados Analisados
- Planilha Excel com 14 abas de dados
- Período: 01/04/2025 a 30/09/2025 (183 dias)
- 1 ponto de entrada: Estação de Recebimento (2 tramos de medição: Tramo 101 e Tramo 501)
- 7 clientes (pontos de saída): Empresa A, Empresa B, Empresa C,
  Empresa D, Empresa E, Empresa F, Empresa G
- Dados diários de entrada: volume (Nm³/d), PCS (kcal/m³), energia (kcal)
- Dados horários de clientes: volume (Nm³/h), pressão (bar_a), temperatura (°C)
- Empresa D: 57% de registros horários faltantes (NaN)
- Comparação Concessionária vs Transportadora em todas as medições de entrada

### Diagramas Anexos (3 imagens)
1. Diagrama da estrutura/topologia do distrito de distribuição de gás natural
2. Diagrama do fluxo geral da auditoria técnica (5 fases)
3. Diagrama do processo de análise de dados (8 etapas da metodologia)

### Tarefa
Elabore o conteúdo deste capítulo introdutório com:

1. **Fundamentação Teórica**: O que é uma análise de condições de operação de um distrito
   de distribuição de gás? Qual a importância regulatória? O que é o balanço de gás?

2. **Análise dos Dados**: Descreva e caracterize o distrito analisado. Comente a estrutura
   dos dados disponíveis, o período, o número de pontos de medição, a completude dos dados.

3. **Discussão dos Diagramas**: Descreva o que se observa em CADA diagrama anexo.
   Referencie como "Diagrama 1", "Diagrama 2", "Diagrama 3".
"""
    # Retorna lista de caminhos completos dos diagramas
    diagram_paths = []
    for fname in CHAPTER_CONFIG[1]["diagram_files"]:
        p = DIAGRAMAS_DIR / fname
        if p.exists():
            diagram_paths.append(str(p))
    return SYSTEM_PROMPT_BASE, prompt, diagram_paths


def prompt_secao1_sintese(conteudo_text: str):
    """Capítulo 1: Síntese (introdução + parecer)."""
    prompt = f"""## Capítulo 1: Visão Geral do Distrito e Dados Disponíveis

Abaixo está o conteúdo já escrito para este capítulo introdutório:

---
{conteudo_text}
---

### Tarefa
Com base no conteúdo acima, escreva:

1. **Introdução do Capítulo** (2-3 parágrafos): Apresente o objetivo da auditoria,
   o escopo (distrito, período), e uma visão geral dos dados disponíveis.
   NÃO use cabeçalho — comece direto com o texto.

2. **Parecer Regulatório**: Comece com "### Parecer Regulatório".
   Os dados são suficientes e adequados para uma auditoria regulatória?
   Comente sobre a qualidade dos dados (completude, dados faltantes do Empresa D).

Separe as duas partes com uma linha contendo apenas: ---SEPARADOR---
"""
    return SYSTEM_PROMPT_SINTESE, prompt, []


# =====================================================================
# SEÇÕES FINAIS — CONCLUSÕES E RESUMO EXECUTIVO
# =====================================================================

def prompt_conclusoes_recomendacoes(todas_secoes_texto: str):
    """Conclusões e Recomendações (recebe TODOS os capítulos)."""
    prompt = f"""## CONCLUSÕES E RECOMENDAÇÕES

Texto de todos os capítulos do relatório:

---
{todas_secoes_texto}
---

### Tarefa
Elabore a seção final com:
1. **Conclusões Técnicas**: Para cada área analisada, emita conclusão de 1-2 linhas com veredito
   (CONFORME / NÃO CONFORME / ATENÇÃO NECESSÁRIA)
2. **Pontos Positivos Identificados**: 3-5 aspectos positivos
3. **Pontos de Atenção**: 3-5 pontos que merecem atenção ou acompanhamento
4. **Recomendações**: 5-8 recomendações concretas priorizadas (Alta / Média / Baixa).
   Para cada uma: o que fazer, por que, e prioridade.
   Apresente em tabela Markdown com colunas: Prioridade | Recomendação | Fundamentação Técnica
5. **Parecer Final do Auditor**: Parecer final de 3-5 linhas sobre a regularidade geral
   das condições de operação do distrito.

Use tom formal e definitivo. Este é o encerramento do relatório.
"""
    return SYSTEM_PROMPT_CONCLUSOES, prompt, []


def prompt_resumo_executivo(todas_secoes_texto: str, conclusoes_texto: str):
    """Resumo Executivo (recebe TODOS os capítulos + Conclusões)."""
    prompt = f"""## RESUMO EXECUTIVO

Texto de todos os capítulos do relatório:

---
{todas_secoes_texto}
---

Conclusões e Recomendações:

---
{conclusoes_texto}
---

### Tarefa
Elabore um Resumo Executivo conciso (máximo 2 páginas):
1. **Objetivo e Escopo**: Descrever o objetivo da auditoria e o escopo
   (distrito, período, pontos de medição)
2. **Metodologia**: Resumir brevemente a metodologia
   (análise de dados, cálculo de incertezas conforme GUM, balanço de massa com bandas)
3. **Principais Achados**: Listar os 5-7 achados mais importantes em bullet points concisos
4. **Conclusão Geral**: Em 2-3 frases, qual o veredito da auditoria?

O Resumo Executivo deve ser compreensível por gestores não-técnicos, sem perder o rigor técnico.
NÃO inclua subseção "Discussão dos Gráficos" no Resumo Executivo.
"""
    return SYSTEM_PROMPT_RESUMO, prompt, []
