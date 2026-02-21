"""
Script para corrigir bugs remanescentes e adicionar plt.savefig() em todos os notebooks.
Correções:
  BUG A: ≈ → ~ (encoding Windows)
  BUG C: "~6,86%" → "~6,19%" (vazão normal vs mínima)
  Adiciona import os + os.makedirs('graficos') nas células de import
  Adiciona plt.savefig() antes de cada plt.show() nos gráficos
"""
import json
import os
import re

BASE = os.path.dirname(os.path.abspath(__file__))

def load_nb(filename):
    path = os.path.join(BASE, filename)
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_nb(nb, filename):
    path = os.path.join(BASE, filename)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(nb, f, ensure_ascii=False, indent=1)

def get_source(cell):
    """Get cell source as a single string."""
    src = cell.get('source', [])
    if isinstance(src, list):
        return ''.join(src)
    return src

def set_source(cell, text):
    """Set cell source from a string (convert to list of lines)."""
    lines = text.split('\n')
    result = []
    for i, line in enumerate(lines):
        if i < len(lines) - 1:
            result.append(line + '\n')
        else:
            result.append(line)
    cell['source'] = result

def add_os_makedirs(cell):
    """Add import os and os.makedirs to an imports cell."""
    src = get_source(cell)
    if 'import os' not in src:
        src = src.rstrip() + '\nimport os\n\nos.makedirs(\'graficos\', exist_ok=True)\n'
        set_source(cell, src)
        return True
    return False

def add_savefig(cell, png_name):
    """Add plt.savefig() before plt.show() in a graph cell."""
    src = get_source(cell)
    savefig_line = f"plt.savefig('graficos/{png_name}', dpi=150, bbox_inches='tight')"
    if savefig_line in src:
        return False  # Already added
    # Insert savefig before plt.show()
    src = src.replace('plt.show()', f'{savefig_line}\nplt.show()', 1)
    set_source(cell, src)
    return True

def fix_approx_char(cell):
    """Replace ≈ with ~ to avoid Windows encoding issues."""
    src = get_source(cell)
    if '\u2248' in src:
        src = src.replace('\u2248', '~')
        set_source(cell, src)
        return True
    return False

def fix_text(cell, old, new):
    """Replace text in a cell."""
    src = get_source(cell)
    if old in src:
        src = src.replace(old, new)
        set_source(cell, src)
        return True
    return False


changes = []

# ============================================================
# NB02 - Análise de Volumes de Entrada
# ============================================================
nb = load_nb('02_analise_volumes_entrada.ipynb')

# cell-1: Add os import
if add_os_makedirs(nb['cells'][1]):
    changes.append('NB02 cell-1: Added import os + makedirs')

# cell-6: Fix ≈ character
if fix_approx_char(nb['cells'][6]):
    changes.append('NB02 cell-6: Fixed approx character')

# cell-8: savefig vol_entrada_serie
if add_savefig(nb['cells'][8], 'vol_entrada_serie.png'):
    changes.append('NB02 cell-8: Added savefig vol_entrada_serie.png')

# cell-10: savefig vol_entrada_diferencas
if add_savefig(nb['cells'][10], 'vol_entrada_diferencas.png'):
    changes.append('NB02 cell-10: Added savefig vol_entrada_diferencas.png')

# cell-12: savefig vol_entrada_histograma
if add_savefig(nb['cells'][12], 'vol_entrada_histograma.png'):
    changes.append('NB02 cell-12: Added savefig vol_entrada_histograma.png')

# cell-14: savefig vol_entrada_boxplot
if add_savefig(nb['cells'][14], 'vol_entrada_boxplot.png'):
    changes.append('NB02 cell-14: Added savefig vol_entrada_boxplot.png')

save_nb(nb, '02_analise_volumes_entrada.ipynb')

# ============================================================
# NB03 - Análise do PCS
# ============================================================
nb = load_nb('03_analise_pcs.ipynb')

# cell-1: Add os import
if add_os_makedirs(nb['cells'][1]):
    changes.append('NB03 cell-1: Added import os + makedirs')

# cell-8: savefig pcs_serie
if add_savefig(nb['cells'][8], 'pcs_serie.png'):
    changes.append('NB03 cell-8: Added savefig pcs_serie.png')

# cell-12: savefig pcs_histograma
if add_savefig(nb['cells'][12], 'pcs_histograma.png'):
    changes.append('NB03 cell-12: Added savefig pcs_histograma.png')

save_nb(nb, '03_analise_pcs.ipynb')

# ============================================================
# NB04 - Cálculo de Energia
# ============================================================
nb = load_nb('04_calculo_energia.ipynb')

# cell-1: Add os import
if add_os_makedirs(nb['cells'][1]):
    changes.append('NB04 cell-1: Added import os + makedirs')

# cell-8: Fix ≈ character
if fix_approx_char(nb['cells'][8]):
    changes.append('NB04 cell-8: Fixed approx character')

# cell-10: savefig energia_serie
if add_savefig(nb['cells'][10], 'energia_serie.png'):
    changes.append('NB04 cell-10: Added savefig energia_serie.png')

# cell-12: savefig energia_diferencas
if add_savefig(nb['cells'][12], 'energia_diferencas.png'):
    changes.append('NB04 cell-12: Added savefig energia_diferencas.png')

# cell-14: savefig energia_mensal
if add_savefig(nb['cells'][14], 'energia_mensal.png'):
    changes.append('NB04 cell-14: Added savefig energia_mensal.png')

# cell-16: savefig energia_scatter
if add_savefig(nb['cells'][16], 'energia_scatter.png'):
    changes.append('NB04 cell-16: Added savefig energia_scatter.png')

save_nb(nb, '04_calculo_energia.ipynb')

# ============================================================
# NB05 - Perfis dos Clientes
# ============================================================
nb = load_nb('05_perfis_clientes.ipynb')

# cell-1: Add os import
if add_os_makedirs(nb['cells'][1]):
    changes.append('NB05 cell-1: Added import os + makedirs')

# cell-7: savefig clientes_serie
if add_savefig(nb['cells'][7], 'clientes_serie.png'):
    changes.append('NB05 cell-7: Added savefig clientes_serie.png')

# cell-9: savefig clientes_perfil_horario
if add_savefig(nb['cells'][9], 'clientes_perfil_horario.png'):
    changes.append('NB05 cell-9: Added savefig clientes_perfil_horario.png')

# cell-11: savefig clientes_heatmap
if add_savefig(nb['cells'][11], 'clientes_heatmap.png'):
    changes.append('NB05 cell-11: Added savefig clientes_heatmap.png')

# cell-13: savefig clientes_pressao_temp
if add_savefig(nb['cells'][13], 'clientes_pressao_temp.png'):
    changes.append('NB05 cell-13: Added savefig clientes_pressao_temp.png')

# cell-15: savefig clientes_participacao
if add_savefig(nb['cells'][15], 'clientes_participacao.png'):
    changes.append('NB05 cell-15: Added savefig clientes_participacao.png')

# cell-17: savefig clientes_boxplot
if add_savefig(nb['cells'][17], 'clientes_boxplot.png'):
    changes.append('NB05 cell-17: Added savefig clientes_boxplot.png')

save_nb(nb, '05_perfis_clientes.ipynb')

# ============================================================
# NB06 - Sumário e Incertezas
# ============================================================
nb = load_nb('06_sumario_e_incertezas.ipynb')

# cell-1: Add os import
if add_os_makedirs(nb['cells'][1]):
    changes.append('NB06 cell-1: Added import os + makedirs')

# cell-3: Add pd.to_numeric for entry volumes
src = get_source(nb['cells'][3])
if "pd.to_numeric(df_vol_ent['Vol_Conc']" not in src:
    src = src.replace(
        "df_vol_ent = df_vol_ent.dropna(subset=['Data']).reset_index(drop=True)",
        "for col in ['Vol_Conc', 'Vol_Transp']:\n    df_vol_ent[col] = pd.to_numeric(df_vol_ent[col], errors='coerce')\ndf_vol_ent = df_vol_ent.dropna(subset=['Data']).reset_index(drop=True)"
    )
    set_source(nb['cells'][3], src)
    changes.append('NB06 cell-3: Added pd.to_numeric for entry volumes')

# cell-10: Fix "~6,86%" → "~6,19% (vazao normal)"
if fix_text(nb['cells'][10], '~6,86%', '~6,19% (vazao normal)'):
    changes.append('NB06 cell-10: Fixed U_saida expected text')
elif fix_text(nb['cells'][10], '** Valor esperado: ~6,86% **', '** Valor esperado: ~6,19% (vazao normal) **'):
    changes.append('NB06 cell-10: Fixed U_saida expected text (alt)')

# cell-12: savefig incertezas_barras
if add_savefig(nb['cells'][12], 'incertezas_barras.png'):
    changes.append('NB06 cell-12: Added savefig incertezas_barras.png')

# cell-14: savefig incertezas_rss
if add_savefig(nb['cells'][14], 'incertezas_rss.png'):
    changes.append('NB06 cell-14: Added savefig incertezas_rss.png')

# cell-16: savefig incertezas_contribuicao
if add_savefig(nb['cells'][16], 'incertezas_contribuicao.png'):
    changes.append('NB06 cell-16: Added savefig incertezas_contribuicao.png')

# cell-17: Fix markdown "~6,86%"
if fix_text(nb['cells'][17], '~6,86%', '~6,19%'):
    changes.append('NB06 cell-17: Fixed U_saida in conclusions markdown')

save_nb(nb, '06_sumario_e_incertezas.ipynb')

# ============================================================
# NB07 - Balanço de Massa
# ============================================================
nb = load_nb('07_balanco_massa.ipynb')

# cell-1: Add os import
if add_os_makedirs(nb['cells'][1]):
    changes.append('NB07 cell-1: Added import os + makedirs')

# cell-3: Add pd.to_numeric for entry volumes
src = get_source(nb['cells'][3])
if "pd.to_numeric(df_vol['Vol_Conc']" not in src:
    old = "df_vol = df_vol.dropna(subset=['Data'])\nvol_entrada = df_vol['Vol_Conc'].sum()"
    new = "for col in ['Vol_Conc', 'Vol_Transp']:\n    df_vol[col] = pd.to_numeric(df_vol[col], errors='coerce')\ndf_vol = df_vol.dropna(subset=['Data'])\nvol_entrada = df_vol['Vol_Conc'].sum()"
    if old in src:
        src = src.replace(old, new)
        set_source(nb['cells'][3], src)
        changes.append('NB07 cell-3: Added pd.to_numeric for entry volumes')

# cell-11: savefig balanco_barras
if add_savefig(nb['cells'][11], 'balanco_barras.png'):
    changes.append('NB07 cell-11: Added savefig balanco_barras.png')

# cell-13: savefig balanco_waterfall
if add_savefig(nb['cells'][13], 'balanco_waterfall.png'):
    changes.append('NB07 cell-13: Added savefig balanco_waterfall.png')

# cell-15: savefig balanco_bandas
if add_savefig(nb['cells'][15], 'balanco_bandas.png'):
    changes.append('NB07 cell-15: Added savefig balanco_bandas.png')

# cell-17: savefig balanco_dashboard
if add_savefig(nb['cells'][17], 'balanco_dashboard.png'):
    changes.append('NB07 cell-17: Added savefig balanco_dashboard.png')

# cell-21: Fix "~6,86%" expected text
if fix_text(nb['cells'][21], '~6,86%', '~6,19% (vazao normal)'):
    changes.append('NB07 cell-21: Fixed U_saida expected text')

# cell-22: Fix conclusions markdown
src_22 = get_source(nb['cells'][22])
if '~0,64%' in src_22:
    src_22 = src_22.replace(
        '- **Diferença:** ~0,64% (1,16 Mm³)',
        '- **Diferença:** ~1,09% (2,0 Mm³) - notebooks usam 183 dias vs 182 da planilha'
    )
    set_source(nb['cells'][22], src_22)
    changes.append('NB07 cell-22: Updated difference explanation in conclusions')

save_nb(nb, '07_balanco_massa.ipynb')

# ============================================================
# Summary
# ============================================================
print(f'\n=== {len(changes)} alteracoes realizadas ===\n')
for c in changes:
    print(f'  [OK] {c}')
print(f'\nPronto! Execute os notebooks para gerar os graficos.')
