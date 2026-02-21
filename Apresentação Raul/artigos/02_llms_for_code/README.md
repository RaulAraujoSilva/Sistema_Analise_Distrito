# 02_llms_for_code — Artigos Selecionados

> Critério: 2024-2026, alta citação ou labs tier-1; foco em produtividade e impacto real

---

## SWE-bench: Can Language Models Resolve Real-World GitHub Issues?

- **Autores:** Carlos E. Jimenez, John Yang, Alexander Wettig et al. (Princeton/Stanford)
- **Ano/Venue:** ICLR 2024 (oral) — arXiv:2310.06770
- **Citações:** ~800+ (benchmark de referência da área)
- **PDF:** https://arxiv.org/pdf/2310.06770
- **Resumo (PT-BR):** Introduz o SWE-bench, benchmark que avalia LLMs em cenários realistas de engenharia de software: dado um repositório GitHub real e uma issue descrevendo um bug ou funcionalidade, o modelo deve gerar um patch que resolva o problema. O dataset contém 2.294 pares issue-pull request de 12 repositórios Python populares. Em 2024, os melhores agentes resolviam ~76% dos problemas.
- **Relevância para a aula:** Demonstra com dados concretos que LLMs já resolvem problemas reais de programação em repositórios de produção — argumento central para mostrar que "IA que programa" não é ficção científica, mas realidade mensurável.

---

## Measuring GitHub Copilot's Impact on Productivity

- **Autores:** Albert Ziegler et al. (GitHub/Microsoft)
- **Ano/Venue:** Communications of the ACM, março 2024
- **Citações:** ~200+ (publicação tier-1 da área)
- **URL:** https://cacm.acm.org/research/measuring-github-copilots-impact-on-productivity/
- **Resumo (PT-BR):** Estudo de caso que correlaciona percepções de produtividade reportadas por usuários do Copilot com dados reais de uso. Os benefícios identificados cobrem tempo de execução de tarefas, qualidade do produto, carga cognitiva, satisfação e aprendizado. O paper oferece metodologia para medir impacto de IA em fluxos de trabalho reais de desenvolvimento.
- **Relevância para a aula:** Evidência empírica de que ferramentas de IA para código aumentam produtividade de desenvolvedores — o mesmo argumento se aplica a analistas de dados usando IA para scripts de análise sem precisar saber programar profundamente.

---

## The Impact of AI on Developer Productivity: Evidence from GitHub Copilot (Field Experiment)

- **Autores:** Sida Peng, Eirini Kalliamvakou, Peter Cihon, Mert Demirer (Microsoft/MIT)
- **Ano/Venue:** 2023 — arXiv:2302.06590 (amplamente citado em 2024)
- **Citações:** ~600+ (estudo experimental de referência)
- **PDF:** https://arxiv.org/pdf/2302.06590
- **Resumo (PT-BR):** Experimento controlado com programadores completando tarefas de código HTTP server em JavaScript com e sem Copilot. Resultado: programadores com Copilot completaram a tarefa **55,8% mais rápido**. O grupo com IA também reportou maior satisfação. É o estudo mais citado sobre ganhos de produtividade de LLMs para programação.
- **Relevância para a aula:** O número "55% mais rápido" é o slide de impacto — demonstra com experimento controlado que IA comprime o tempo de trabalho técnico. Aplicável diretamente ao contexto de reguladores que hoje levam horas para tarefas de análise que IA faria em minutos.

---

## A Survey on Large Language Models for Code Generation

- **Autores:** Jia Li et al. (múltiplas afiliações)
- **Ano/Venue:** Junho 2024 — arXiv:2406.00515
- **Citações:** ~300+ (survey de referência de 2024)
- **PDF:** https://arxiv.org/pdf/2406.00515
- **Resumo (PT-BR):** Survey abrangente que analisa o estado da arte em geração de código por LLMs, cobrindo benchmarks (HumanEval, MBPP, SWE-bench), técnicas de prompting, agentes de código e aplicações. Enfatiza que LLMs democratizaram a programação ao permitir que usuários sem experiência gerem código descrevendo seus requisitos em linguagem natural. Documenta a transição de "programação" para "especificação em linguagem natural".
- **Relevância para a aula:** Oferece a visão panorâmica de como LLMs transformam a programação de atividade especializada para habilidade acessível — argumento central do curso para reguladores sem background de programação.

---

## Developer Productivity With and Without GitHub Copilot: A Longitudinal Mixed-Methods Case Study

- **Autores:** Múltiplos autores (estudo longitudinal acadêmico)
- **Ano/Venue:** 2025 — arXiv:2509.20353
- **Citações:** Recente (2025); método robusto longitudinal
- **PDF:** https://arxiv.org/pdf/2509.20353
- **Resumo (PT-BR):** Estudo longitudinal que acompanha desenvolvedores com e sem Copilot, combinando métricas de atividade (commits, pull requests) com entrevistas qualitativas. Encontra que usuários do Copilot são consistentemente mais ativos. Identifica que os benefícios se manifestam em qualidade e satisfação, não apenas velocidade.
- **Relevância para a aula:** Nuança importante — IA para código melhora qualidade e satisfação, não apenas velocidade. Relevante para reguladores que se preocupam com qualidade das análises produzidas com auxílio de IA.
