# Apresentação ABAR — Bloco 6: IA Aplicada a Dados Regulatórios

## Sobre
Página web interativa (React) para o Curso ABAR de Medições Inteligentes e Gestão Integrada.
Evento: 24-27 de fevereiro de 2026. Prof. Raul Araujo.

## Estrutura

```
apresentacao-react/          # Projeto React principal
├── src/sections/            # 14 seções (01-Hero a 14-Resources)
├── src/components/          # Componentes reutilizáveis
├── public/img/              # Screenshots, gráficos, diagramas
├── public/data/             # JSON para gráficos Recharts
├── colab/                   # 7 notebooks Google Colab
└── dist/                    # Build estático (Vite)
```

## Comandos

```bash
cd apresentacao-react
npm run dev      # Dev server (localhost:5173+)
npm run build    # Build para dist/
```

## GitHub
- Repo: `RaulAraujoSilva/Aula_AI_Dados`
- Branch: `main`

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
