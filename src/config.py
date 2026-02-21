# -*- coding: utf-8 -*-
"""
Configuração centralizada de caminhos do projeto.
Todos os scripts e notebooks importam deste módulo.
"""
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent

# Dados de entrada
DATA_DIR = PROJECT_ROOT / "data" / "input"
APOSTILA_DIR = PROJECT_ROOT / "data" / "apostila"

# Código-fonte
SRC_DIR = PROJECT_ROOT / "src"

# Notebooks
NOTEBOOKS_DIR = PROJECT_ROOT / "notebooks"

# Outputs
OUTPUTS_DIR = PROJECT_ROOT / "outputs"
GRAFICOS_DIR = OUTPUTS_DIR / "graficos"
DIAGRAMAS_DIR = OUTPUTS_DIR / "diagramas"
REPORTS_DIR = OUTPUTS_DIR / "reports"
PRESENT_DIR = OUTPUTS_DIR / "presentations"
CACHE_DIR = OUTPUTS_DIR / "cache"

# Intermediários
METODOLOGIA_DIR = PROJECT_ROOT / "metodologia"

# Arquivo Excel padrão (para automação futura)
EXCEL_DEFAULT = "Analise de Condições de Operação de Distrito.xlsx"

# Notebooks (para apêndice do relatório)
NOTEBOOK_LIST = [
    {"file": "01_leitura_e_exploracao.ipynb",
     "titulo": "Leitura e Exploração dos Dados"},
    {"file": "02_analise_volumes_entrada.ipynb",
     "titulo": "Análise de Volumes de Entrada"},
    {"file": "03_analise_pcs.ipynb",
     "titulo": "Análise do Poder Calorífico Superior"},
    {"file": "04_calculo_energia.ipynb",
     "titulo": "Cálculo e Validação de Energia"},
    {"file": "05_perfis_clientes.ipynb",
     "titulo": "Perfis de Consumo dos Clientes"},
    {"file": "06_sumario_e_incertezas.ipynb",
     "titulo": "Cálculo de Incertezas de Medição"},
    {"file": "07_balanco_massa.ipynb",
     "titulo": "Balanço de Massa com Bandas de Incerteza"},
]
