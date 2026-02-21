"""
Converte os 7 notebooks locais para versão Google Colab.
Substitui o cell de config por montagem do Google Drive.
"""
import json
import copy
from pathlib import Path

NOTEBOOKS_DIR = Path(__file__).resolve().parent.parent / "notebooks"
OUTPUT_DIR = Path(__file__).resolve().parent
EXCEL_FILENAME = "Analise de Condições de Operação de Distrito.xlsx"

# Célula de setup para Colab (substitui a célula de config)
COLAB_HEADER_MD = {
    "cell_type": "markdown",
    "metadata": {},
    "source": [
        "## Configuração do Google Colab\n",
        "\n",
        "Este notebook foi adaptado para rodar no **Google Colab**.\n",
        "\n",
        "**Pré-requisito:** Coloque o arquivo Excel na pasta do Google Drive:\n",
        "```\n",
        "Google Drive / ABAR / data / Analise de Condições de Operação de Distrito.xlsx\n",
        "```\n",
        "\n",
        "> Se o arquivo estiver em outra pasta, altere `DRIVE_DATA_DIR` na célula abaixo."
    ]
}

COLAB_SETUP_CODE = {
    "cell_type": "code",
    "metadata": {},
    "outputs": [],
    "execution_count": None,
    "source": [
        "# === CONFIGURAÇÃO GOOGLE COLAB ===\n",
        "# Monte o Google Drive\n",
        "from google.colab import drive\n",
        "drive.mount('/content/drive')\n",
        "\n",
        "from pathlib import Path\n",
        "import os\n",
        "\n",
        "# Pasta no Google Drive onde está o arquivo Excel\n",
        "# Ajuste se necessário:\n",
        "DRIVE_DATA_DIR = Path('/content/drive/MyDrive/ABAR/data')\n",
        "\n",
        "# Pasta para salvar gráficos (no Colab)\n",
        "GRAFICOS_DIR = Path('/content/graficos')\n",
        "GRAFICOS_DIR.mkdir(parents=True, exist_ok=True)\n",
        "\n",
        f"EXCEL_DEFAULT = '{EXCEL_FILENAME}'\n",
        "EXCEL_PATH = DRIVE_DATA_DIR / EXCEL_DEFAULT\n",
        "\n",
        "# Verificar se o arquivo existe\n",
        "if EXCEL_PATH.exists():\n",
        "    print(f'Arquivo encontrado: {EXCEL_PATH}')\n",
        "    print(f'Tamanho: {EXCEL_PATH.stat().st_size / 1024:.0f} KB')\n",
        "else:\n",
        "    print(f'ERRO: Arquivo não encontrado em {EXCEL_PATH}')\n",
        "    print(f'Conteúdo de {DRIVE_DATA_DIR}:')\n",
        "    if DRIVE_DATA_DIR.exists():\n",
        "        for f in DRIVE_DATA_DIR.iterdir():\n",
        "            print(f'  {f.name}')\n",
        "    else:\n",
        "        print(f'  Pasta não existe! Crie: {DRIVE_DATA_DIR}')\n",
    ]
}

COLAB_IMPORTS_CODE = {
    "cell_type": "code",
    "metadata": {},
    "outputs": [],
    "execution_count": None,
    "source": [
        "import pandas as pd\n",
        "import numpy as np\n",
        "import matplotlib.pyplot as plt\n",
        "import seaborn as sns\n",
        "import warnings\n",
        "\n",
        "# Configurações gerais\n",
        "warnings.filterwarnings('ignore')\n",
        "pd.set_option('display.max_columns', 20)\n",
        "pd.set_option('display.float_format', '{:,.2f}'.format)\n",
        "plt.rcParams['figure.figsize'] = (14, 6)\n",
        "plt.rcParams['font.size'] = 12\n",
        "\n",
        "print('Bibliotecas carregadas com sucesso!')\n",
    ]
}


def is_config_cell(source_text: str) -> bool:
    """Detecta se uma célula é a célula de configuração/imports local."""
    indicators = [
        "from config import",
        "sys.path.insert",
        "PROJECT_ROOT",
        "EXCEL_DEFAULT",
        "from pathlib import Path",
    ]
    return sum(1 for ind in indicators if ind in source_text) >= 2


def get_source_text(cell: dict) -> str:
    """Extrai texto da célula."""
    src = cell.get("source", [])
    if isinstance(src, list):
        return "".join(src)
    return src


def convert_notebook(input_path: Path, output_path: Path):
    """Converte um notebook local para versão Colab."""
    with open(input_path, "r", encoding="utf-8") as f:
        nb = json.load(f)

    nb_out = copy.deepcopy(nb)

    # Adicionar metadata do Colab
    if "metadata" not in nb_out:
        nb_out["metadata"] = {}
    nb_out["metadata"]["colab"] = {
        "provenance": [],
        "toc_visible": True
    }

    new_cells = []
    config_replaced = False

    for cell in nb_out["cells"]:
        src = get_source_text(cell)

        # Substituir célula de config
        if not config_replaced and cell["cell_type"] == "code" and is_config_cell(src):
            new_cells.append(COLAB_HEADER_MD)
            new_cells.append(COLAB_SETUP_CODE)
            new_cells.append(COLAB_IMPORTS_CODE)
            config_replaced = True
            continue

        # Remover referências ao config em outras células de código
        if cell["cell_type"] == "code":
            # Substituir GRAFICOS_DIR paths que usam str()
            src_new = src.replace(
                "GRAFICOS_DIR.mkdir(parents=True, exist_ok=True)",
                "# GRAFICOS_DIR já criado no setup"
            )
            if src_new != src:
                if isinstance(cell["source"], list):
                    cell["source"] = src_new.splitlines(True)
                else:
                    cell["source"] = src_new

        new_cells.append(cell)

    nb_out["cells"] = new_cells

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(nb_out, f, ensure_ascii=False, indent=1)

    return config_replaced


def main():
    notebooks = sorted(NOTEBOOKS_DIR.glob("*.ipynb"))
    print(f"Encontrados {len(notebooks)} notebooks em {NOTEBOOKS_DIR}")
    print(f"Saída: {OUTPUT_DIR}\n")

    for nb_path in notebooks:
        out_name = f"colab_{nb_path.name}"
        out_path = OUTPUT_DIR / out_name
        replaced = convert_notebook(nb_path, out_path)
        status = "OK (config substituído)" if replaced else "AVISO (config não encontrado)"
        print(f"  {nb_path.name} -> {out_name} [{status}]")

    print(f"\nConversão concluída! {len(notebooks)} notebooks criados em {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
