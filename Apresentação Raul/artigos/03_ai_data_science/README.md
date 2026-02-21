# 03_ai_data_science — Artigos Selecionados

> Critério: 2024-2026, alta citação ou labs tier-1; foco em análise de dados tabulares e automação

---

## TableGPT2: A Large Multimodal Model with Tabular Data Integration

- **Autores:** Aofeng Su et al. (33 autores, Zhejiang University e colaboradores)
- **Ano/Venue:** Novembro 2024 — arXiv:2411.02059
- **Citações:** ~150+ (artigo recente de alto impacto em 2024)
- **PDF:** https://arxiv.org/pdf/2411.02059
- **Resumo (PT-BR):** TableGPT2 é um modelo multimodal de grande escala treinado com mais de 593.800 tabelas e 2,36 milhões de pares query-tabela-output. Introduz um encoder de tabelas capaz de capturar informação em nível de schema e célula, lidando com queries ambíguas, nomes de colunas ausentes e tabelas irregulares do mundo real. Em 23 métricas de benchmark, supera LLMs neutros em 35,2% (versão 7B) e 49,3% (versão 72B).
- **Relevância para a aula:** Exemplo direto de IA especializada em tabelas — o tipo de dado que reguladores de gás trabalham diariamente (medições, relatórios, planilhas). Demonstra que a IA não é só para texto corrido, mas excelente para análise de dados estruturados.

---

## Data Interpreter: An LLM Agent For Data Science

- **Autores:** Zirui Hong, Liang Lin et al. (MetaGPT team)
- **Ano/Venue:** Fevereiro 2024 — arXiv:2402.18679
- **Citações:** ~400+ (artigo amplamente citado em 2024)
- **PDF:** https://arxiv.org/pdf/2402.18679
- **Resumo (PT-BR):** Apresenta o Data Interpreter, um agente LLM que resolve problemas de ciência de dados de ponta a ponta. Usa modelagem hierárquica em grafo para decompor problemas complexos em subproblemas e geração de código iterativa e verificada. No benchmark InfiAgent-DABench, eleva acurácia de 75,9% para 94,9%. Disponível via MetaGPT (GitHub).
- **Relevância para a aula:** Demonstração concreta de um agente que executa o pipeline completo de ciência de dados (carregamento, limpeza, análise, modelagem, visualização) via linguagem natural — exatamente o que um regulador não-programador precisa.

---

## DataSciBench: An LLM Agent Benchmark for Data Science

- **Autores:** Múltiplos autores (HKUST e colaboradores)
- **Ano/Venue:** Fevereiro 2025 — arXiv:2502.13897
- **Citações:** Recente (2025); benchmark de referência emergente
- **URL:** https://arxiv.org/abs/2502.13897
- **Resumo (PT-BR):** Benchmark abrangente para avaliar capacidades de agentes LLM em ciência de dados completa, com prompts naturais e desafiadores que exigem ground truth não-trivial e métricas de avaliação complexas. Avalia o ciclo completo: carregamento, limpeza, exploração, modelagem e geração de insights.
- **Relevância para a aula:** Fornece evidência científica rigorosa de como avaliar IA em tarefas de ciência de dados — útil para mostrar ao público que há critérios objetivos de qualidade.

---

## Data Science Through Natural Language with ChatGPT's Code Interpreter

- **Autores:** Múltiplos autores (estudo clínico/acadêmico)
- **Ano/Venue:** 2024 — PubMed Central / PMC11224898
- **Citações:** ~100+ (estudo aplicado em contexto real)
- **URL:** https://pmc.ncbi.nlm.nih.gov/articles/PMC11224898/
- **Resumo (PT-BR):** Investiga o uso do ChatGPT Code Interpreter (Advanced Data Analysis) como ferramenta de ciência de dados por meio de linguagem natural. Demonstra que especialistas de domínio sem formação em programação conseguem executar análises estatísticas, criar visualizações e gerar insights através de conversação com IA. Avalia capacidades e limitações práticas da ferramenta em cenários reais.
- **Relevância para a aula:** Exemplo empírico direto do caso de uso da aula: especialistas de domínio (como reguladores) usando ChatGPT para fazer ciência de dados sem programar.

---

## A Survey on Large Language Model-Based Agents for Statistics and Data Science

- **Autores:** Múltiplos autores
- **Ano/Venue:** Dezembro 2024 — arXiv:2412.14222; publicado na The American Statistician (Tandfonline) em 2025
- **Citações:** ~200+ (survey de referência 2024-2025)
- **PDF:** https://arxiv.org/abs/2412.14222
- **Resumo (PT-BR):** Survey que cobre a evolução, capacidades e aplicações de agentes LLM para estatística e ciência de dados, destacando o papel desses agentes na simplificação de tarefas complexas e na redução da barreira de entrada para usuários sem expertise técnica. Documenta ferramentas como Code Interpreter, Data Analyst GPT, Julius AI e outras plataformas baseadas em IA conversacional para análise de dados.
- **Relevância para a aula:** Mapa abrangente do ecossistema de ferramentas — permite mostrar que não é só o ChatGPT, mas um ecossistema crescente de IA para análise de dados acessível a não-programadores.
