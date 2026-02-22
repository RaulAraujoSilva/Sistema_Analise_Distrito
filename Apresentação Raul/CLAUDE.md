# Apresentação ABAR — Bloco 6: IA Aplicada a Dados Regulatórios

## Sobre
Página web interativa (React) para o Curso ABAR de Medições Inteligentes e Gestão Integrada.
Evento: 24-27 de fevereiro de 2026. Prof. Raul Araujo.

## Estrutura

```
apresentacao-react/
├── src/sections/            # Seções de conteúdo (01-Hero a 14-Resources)
├── src/sections/Quiz.tsx    # Quiz nível IA (após Hero)
├── src/sections/PollMethodology.tsx  # Poll IBM (após Methodology)
├── src/sections/WordCloud.tsx        # Feedback final (word cloud)
├── src/sections/ExerciseLimits.tsx   # Exercício 1: Violações normativas
├── src/sections/ExerciseData.tsx     # Exercício 2: CV de clientes
├── src/sections/ExerciseForecast.tsx # Exercício 3: Regressão e projeção
├── src/sections/ExerciseApplication.tsx # Exercício 4: Aplicação (word cloud)
├── src/components/          # Componentes reutilizáveis (Section, Navbar)
├── api/                     # Vercel serverless functions (Redis)
├── public/img/              # Screenshots, gráficos, diagramas
├── public/data/             # JSON + CSVs para exercícios
├── colab/                   # 7 notebooks Google Colab
└── dist/                    # Build estático (Vite)
```

## Fluxo da Apresentação

```
Hero → Quiz → Evolution → Models → Capabilities → DevTools → Demos
  → Exercício 1 (Limites Normativos)
→ Methodology → Poll Metodologia → DataSources → Pipeline
  → Exercício 2 (Análise de Dados — CV)
→ Results
  → Exercício 3 (Regressão e Projeção)
→ WebSystem → Dashboards → Deploy
  → Exercício 4 (Aplicação — Word Cloud)
→ Resources → Word Cloud Feedback
```

## Interativos e Redis

| Componente | API | Redis Key | Tipo | Descrição |
|-----------|-----|-----------|------|-----------|
| Quiz | `/api/quiz` | `quiz:votes` | Hash | Nível de IA (A-E) |
| PollMethodology | `/api/poll-methodology` | `poll:methodology` | Hash | Como analisa dados (A-E) |
| ExerciseLimits | `/api/exercise-limits` | `exercise:limits` | Hash | Meses fora ±1,5σ (A-E) |
| ExerciseData | `/api/exercise-data` | `exercise:data` | Hash | Cliente com maior CV (A-E) |
| ExerciseForecast | `/api/exercise-forecast` | `exercise:forecast` | List | Projeção numérica → histograma |
| ExerciseApplication | `/api/exercise-application` | `exercise:aplicacao` | List | Texto livre → word cloud |
| WordCloud | `/api/wordcloud` | `wordcloud:comments` | List | Feedback final → word cloud |

**Redis**: Upstash (`proven-whale-39354.upstash.io`)
**Env vars**: `UPSTASH_REDIS_REST_URL` e `UPSTASH_REDIS_REST_TOKEN` no Vercel

## CSVs dos Exercícios

| Arquivo | Linhas | Exercício | Resposta |
|---------|--------|-----------|----------|
| `public/data/exercicio-limites.csv` | 24 | Ex 1 | 3 meses (Abr/24, Ago/25, Nov/25) |
| `public/data/exercicio-clientes.csv` | 840 | Ex 2 | Cliente C (CV ≈ 63%) |
| `public/data/exercicio-projecao.csv` | 24 | Ex 3 | Log: ~1283 mil m³ (Dez/2027) |

## Comandos

```bash
cd apresentacao-react
npm run dev      # Dev server (localhost:5173+)
npm run build    # Build para dist/
```

## GitHub
- Repo: `RaulAraujoSilva/Aula_AI_Dados`
- Branch: `main`
- Deploy: Vercel (auto-deploy on push)

## Google Colab
- Notebooks em `colab/colab_01_*.ipynb` a `colab/colab_07_*.ipynb`
- Dados: Google Drive > `ABAR/data/Analise de Condições de Operação de Distrito.xlsx`
- Links Colab: `https://colab.research.google.com/github/RaulAraujoSilva/Aula_AI_Dados/blob/main/colab/`

## Credenciais Google (para Drive API)
- Client JSON: `INEA/INEA_Formularios para SIOP-AI/Acessos Google/client_secret_2_539906551194-*.json`
- Token: `INEA/INEA_Formularios para SIOP-AI/scripts/token.json`
- Projeto: `bap-sucesso-467013`
- Conta: `raularaujo@crie.coppe.ufrj.br`

## Decisões Importantes
- ANP Painéis Dinâmicos: usar imagem estática (iframe dá erro de renderização)
- Dashboard: chamar "Rede de Distribuição de Gás" (não AGENERSA)
- Seção de avaliação removida na v5.1
- Windsurf renomeado para Antigravity (rebranding Google)
- LMArena renomeado para Arena AI
- Word cloud: implementação custom com CSS flexbox (sem biblioteca externa)
- Exercícios usam mesmo Upstash Redis do quiz (keys separadas)
- Histograma do Exercício 3: bins de 50 mil m³, BarChart vertical
