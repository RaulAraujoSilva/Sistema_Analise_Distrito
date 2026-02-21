# Auditoria de Medicao de Gas Natural - Curso ABAR

Sistema de analise de dados e geracao automatizada de relatorios de auditoria para distritos de distribuicao de gas natural, desenvolvido como parte do Curso ABAR de Medicoes Inteligentes e Gestao Integrada.

## Visao Geral

Este projeto transforma uma planilha Excel bruta (14 abas, 183 dias de dados) em um relatorio profissional de auditoria tecnica usando IA generativa. O pipeline completo roda em **menos de 2 minutos** gracias a execucao paralela.

### Entregas Geradas

| Entrega | Formato | Tamanho |
|---------|---------|---------|
| Relatorio de Auditoria (com apendice) | DOCX | 8.83 MB |
| Relatorio de Auditoria (sem apendice) | DOCX | 8.71 MB |
| Apresentacao | PPTX | 1.01 MB |
| Sumario Executivo | DOCX | 2.3 MB |
| Graficos de Analise | 23 PNGs | - |
| Diagramas de Processo (IA) | 3 PNGs | - |

## Arquitetura

```
Excel (14 abas) --> 7 Notebooks Jupyter --> 23 Graficos PNG + Dados estruturados
                                                     |
Apostila PDF --> Extracao IA --> 7 Metodologias .md   |
                                                     v
                                        Gemini Flash (28 chamadas paralelas)
                                                     |
                                                     v
                                          Cache (28 arquivos .md)
                                                     |
                                                     v
                                     DOCX Builder --> Relatorio 8.83 MB
                                     PPTX Builder --> Apresentacao 1.01 MB
```

### Pipeline Paralelo (v5)

O pipeline executa 28 chamadas ao Gemini em **5 waves**:

| Wave | Chamadas | Paralelismo | Descricao |
|------|----------|-------------|-----------|
| **Wave 1** | 19 | 19 threads | Secoes A/B/C de todos os capitulos |
| **Wave 2** | 7 | 7 threads | Sinteses (D) de todos os capitulos |
| **Wave 3** | 1 | sequencial | Conclusoes e Recomendacoes |
| **Wave 4** | 1 | sequencial | Resumo Executivo |
| **Wave 5** | 0 | - | Montagem DOCX + PPTX |

**Performance:**

| Versao | Modelo | Tempo Total |
|--------|--------|-------------|
| v4 (sequencial) | Gemini 3 Pro | 19:50 |
| v4 (sequencial) | Gemini 3 Flash | 19:21 |
| **v5 (paralelo)** | **Gemini 3 Flash** | **1:42** |

## Estrutura do Projeto

```
data/
  input/                  Planilha Excel de entrada
  apostila/               Material de referencia (PDF)
notebooks/                7 Jupyter notebooks de analise
  01_leitura_dados.ipynb    Leitura e exploracao
  02_volumes_entrada.ipynb  Volumes de entrada
  03_analise_pcs.ipynb      Poder Calorifico Superior
  04_calculo_energia.ipynb  Calculo de energia
  05_perfis_clientes.ipynb  Perfis de consumo
  06_incertezas.ipynb       Incertezas de medicao (GUM)
  07_balanco_massa.ipynb    Balanco de massa
src/
  config.py               Caminhos centralizados do projeto
  gemini_client.py        Wrapper da API Gemini (rate limit, retry, parallel)
  prompts_auditoria.py    6 system prompts + CHAPTER_CONFIG (7 capitulos)
  dados_distrito.py       Dataclasses com dados dos notebooks
  docx_builder.py         Construtor DOCX (equacoes OMML, TOC, graficos)
  gerar_relatorio_auditoria.py   Orquestrador do pipeline (ThreadPoolExecutor)
  gerar_apresentacao.py   Gerador do PPTX (13 slides)
  gerar_diagramas.py      Gerador de diagramas via Gemini Image
  extrair_metodologia.py  Extrator de metodologia da apostila
  gerar_sumario_executivo.py  Sumario executivo (Selenium + python-docx)
  web/
    app.py                FastAPI app (endpoints, SSE, upload)
    pipeline_runner.py    Runner com thread-safe progress events
    run_web.py            Entry point (uvicorn)
    templates/index.html  SPA (Vanilla JS, dark mode, responsive)
    static/               CSS, JS, favicon
outputs/
  graficos/               23 PNGs gerados pelos notebooks
  diagramas/              3 PNGs gerados pelo Gemini Image
  reports/                DOCX e PPTX finais
  cache/                  28 arquivos .md (cache do LLM)
  screenshots/            6 PNGs da interface web (Selenium)
metodologia/              7 textos .md + equacoes.json
docs/
  PIPELINE.md             Documentacao com diagramas Mermaid
```

## Tecnologias

| Componente | Tecnologia |
|-----------|-----------|
| Notebooks de Analise | Python 3.12, pandas, NumPy, matplotlib |
| Modelo de IA (Texto) | Google Gemini 3 Flash Preview (thinking: medium) |
| Modelo de IA (Imagem) | Google Gemini 3 Pro Image Preview |
| Construtor de Relatorio | python-docx, latex2mathml, lxml |
| Construtor de Apresentacao | python-pptx |
| Interface Web | FastAPI, Jinja2, JavaScript vanilla (SPA) |
| Streaming de Progresso | Server-Sent Events (SSE) |
| Paralelismo | concurrent.futures.ThreadPoolExecutor |
| Dados de Entrada | Excel / openpyxl |
| Screenshots | Selenium WebDriver (headless Chrome) |

## Como Usar

### 1. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 2. Interface Web (recomendado)

```bash
cd src/web
python run_web.py
```

Acesse `http://127.0.0.1:8000`:
1. Faca upload do Excel
2. Informe a API Key do Gemini
3. Selecione o modo (Gerar Completo / Retomar / Apenas Montar)
4. Acompanhe o progresso em tempo real (29 etapas)
5. Baixe DOCX e PPTX na aba Downloads

### 3. Linha de Comando

```bash
cd src

# Geracao completa (28 chamadas API, ~1:42)
python gerar_relatorio_auditoria.py --api-key SUA_CHAVE

# Retomar usando cache existente
python gerar_relatorio_auditoria.py --api-key SUA_CHAVE --resume

# Apenas montar DOCX a partir do cache (0 chamadas, ~7s)
python gerar_relatorio_auditoria.py --api-key SUA_CHAVE --montar
```

### 4. Sumario Executivo

```bash
python src/gerar_sumario_executivo.py
```

Gera DOCX com capa institucional, 10 secoes, 12 imagens (6 screenshots web + 4 graficos + 2 diagramas).

## Interface Web

A interface web oferece:

- **Configuracao**: Upload Excel, API Key, selecao de modo
- **Pipeline**: Barra de progresso, stepper de fases, cards por capitulo com badges A/B/C/D, log em tempo real
- **Graficos**: Galeria com 23 graficos, filtro por capitulo
- **Diagramas**: 3 diagramas de processo gerados pela IA
- **Textos**: Acordeao com 28 secoes em markdown renderizado
- **Downloads**: DOCX (2 versoes) + PPTX com botao de download
- **Extras**: Dark mode, cancel, favicon, responsive

## Conteudo do Relatorio

O relatorio de auditoria contem 8 capitulos:

1. **Visao Geral do Distrito** - Estrutura, diagramas, dados disponiveis
2. **Volumes de Entrada** - Series temporais, diferencas Conc. vs Transp.
3. **Poder Calorifico Superior (PCS)** - Distribuicao, faixas regulatorias
4. **Calculo de Energia** - E = V x PCS, validacao cruzada
5. **Perfis de Consumo** - 7 clientes industriais, participacao, sazonalidade
6. **Incertezas de Medicao** - GUM, combinacao RSS, bandas de confianca
7. **Balanco de Massa** - Entrada vs soma saidas, waterfall, diferenca %
8. **Conclusoes e Recomendacoes** - Parecer regulatorio consolidado

Cada capitulo segue a estrutura: Introducao > Metodologia > Dados > Graficos > Parecer Regulatorio.

Inclui Apendice A com codigo e resultados dos 7 notebooks Jupyter.

## Historico de Versoes

| Versao | Data | Descricao |
|--------|------|-----------|
| v1 | 2026-02-12 | Pipeline sequencial, Gemini 3 Pro, 9 secoes, DOCX 4.3MB |
| v2 | 2026-02-12 | Tabelas, diagramas, equacoes OMML, DOCX 5.9MB |
| v3 | 2026-02-12 | Pipeline segmentado (28 chamadas, 4/capitulo), cache granular |
| v4 | 2026-02-12 | Apendice notebooks, interface web, anonimizacao, DOCX 9.3MB |
| **v5** | **2026-02-12** | **Flash + paralelo (19+7 threads), 1:42, DOCX 8.83MB** |

## Creditos

- **Coordenacao**: Vladimir Paschoal Macedo
- **Orientacao**: Prof. Alexandre Beraldi Santos
- **Autor**: Raul Araujo da Silva
- **Parceria**: ABAR / LabDGE - UFF
- **IA**: Claude (Anthropic) + Gemini (Google)
