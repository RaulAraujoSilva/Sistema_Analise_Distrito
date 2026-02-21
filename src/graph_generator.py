# -*- coding: utf-8 -*-
"""
Gerador de gráficos a partir do Excel de dados do distrito.
Extraído dos notebooks Jupyter 02-07 para execução standalone.

Uso:
    from graph_generator import gerar_todos_graficos
    gerados = gerar_todos_graficos(excel_path, output_dir)
"""
import logging
from pathlib import Path
from typing import Callable

import matplotlib
matplotlib.use("Agg")  # Non-interactive backend for server use

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.patches as mpatches
import numpy as np
import pandas as pd
import seaborn as sns

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants (same as notebooks)
# ---------------------------------------------------------------------------
CLIENTES = {
    "Cliente #1": "Empresa A",
    "Cliente #2": "Empresa B",
    "Cliente #3": "Empresa C",
    "Cliente #4": "Empresa D",
    "Cliente #5": "Empresa E",
    "Cliente #6": "Empresa F",
    "Cliente #7": "Empresa G",
}

VOLUMES_REFERENCIA = {"Empresa D": 88184}

INCERTEZAS = {
    "Entrada - Tramo 101 (Comgás 1)": 0.0106,
    "Entrada - Tramo 501 (Comgás 2)": 0.0109,
    "Empresa A": 0.0133,
    "Empresa B": 0.0161,
    "Empresa C": 0.0134,
    "Empresa D": 0.0358,
    "Empresa E": 0.0305,
    "Empresa F": 0.0148,
    "Empresa G": 0.028,
}

SAVE_KW = dict(dpi=150, bbox_inches="tight")


# ---------------------------------------------------------------------------
# Data loaders
# ---------------------------------------------------------------------------
def _load_volumes(excel_path: str | Path) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name="Vol Entrada Gas", header=1, usecols="B:F")
    df.columns = ["Data", "Concessionaria_Nm3d", "Transportadora_Nm3d", "Dif_Abs", "Dif_Pct"]
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    for col in ["Concessionaria_Nm3d", "Transportadora_Nm3d"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.dropna(subset=["Data"]).reset_index(drop=True)
    df["Dif_Abs_Calc"] = df["Concessionaria_Nm3d"] - df["Transportadora_Nm3d"]
    df["Dif_Pct_Calc"] = (df["Dif_Abs_Calc"] / df["Concessionaria_Nm3d"]) * 100
    df["Mes"] = df["Data"].dt.to_period("M").astype(str)
    return df


def _load_pcs(excel_path: str | Path) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name="PCS Ent ", header=1, usecols="B:F")
    df.columns = ["Data", "PCS_Conc_kcal", "PCS_Transp_kcal", "Dif_Abs", "Dif_Pct"]
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    for col in ["PCS_Conc_kcal", "PCS_Transp_kcal"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df.dropna(subset=["Data"]).reset_index(drop=True)


def _load_clientes(excel_path: str | Path) -> dict:
    dados = {}
    for aba, nome in CLIENTES.items():
        df = pd.read_excel(excel_path, sheet_name=aba, header=2, usecols="B:E")
        df.columns = ["Data", "Volume_Nm3h", "Pressao_bara", "Temperatura_C"]
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
        for col in ["Volume_Nm3h", "Pressao_bara", "Temperatura_C"]:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        df = df.dropna(subset=["Data"]).reset_index(drop=True)
        sem_dados = len(df) == 0 or df["Volume_Nm3h"].dropna().empty
        dados[aba] = {"nome": nome, "dados": df, "sem_dados": sem_dados}
    return dados


# ---------------------------------------------------------------------------
# 1. Volumes de Entrada (NB02) — 4 gráficos
# ---------------------------------------------------------------------------
def gerar_graficos_volumes(excel_path: str | Path, out: Path) -> list[str]:
    df = _load_volumes(excel_path)
    gerados = []

    # 1.1 Série temporal
    fig, ax = plt.subplots(figsize=(16, 6))
    ax.plot(df["Data"], df["Concessionaria_Nm3d"] / 1000,
            label="Concessionária", color="#2196F3", linewidth=1.5, alpha=0.9)
    ax.plot(df["Data"], df["Transportadora_Nm3d"] / 1000,
            label="Transportadora", color="#FF5722", linewidth=1.5, alpha=0.7, linestyle="--")
    media = df["Concessionaria_Nm3d"].mean() / 1000
    ax.axhline(y=media, color="gray", linestyle=":", alpha=0.5, label=f"Média: {media:,.0f} mil Nm³/d")
    ax.set_title("Volume de Entrada Diário - Concessionária vs Transportadora", fontsize=14, fontweight="bold")
    ax.set_xlabel("Data"); ax.set_ylabel("Volume (10³ Nm³/d)")
    ax.legend(loc="lower left"); ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b/%Y"))
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    plt.xticks(rotation=45); plt.tight_layout()
    fig.savefig(str(out / "vol_entrada_serie.png"), **SAVE_KW); plt.close(fig)
    gerados.append("vol_entrada_serie.png")

    # 1.2 Diferenças
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(16, 8), sharex=True)
    ax1.bar(df["Data"], df["Dif_Abs_Calc"], color="steelblue", alpha=0.7, width=1)
    ax1.axhline(y=0, color="red", linewidth=0.8)
    ax1.set_title("Diferença Absoluta (Concessionária - Transportadora)", fontsize=13, fontweight="bold")
    ax1.set_ylabel("Diferença (Nm³)"); ax1.grid(True, alpha=0.3)
    ax2.bar(df["Data"], df["Dif_Pct_Calc"], color="darkorange", alpha=0.7, width=1)
    ax2.axhline(y=0, color="red", linewidth=0.8)
    ax2.set_title("Diferença Percentual", fontsize=13, fontweight="bold")
    ax2.set_ylabel("Diferença (%)"); ax2.set_xlabel("Data"); ax2.grid(True, alpha=0.3)
    ax2.xaxis.set_major_formatter(mdates.DateFormatter("%b/%Y"))
    ax2.xaxis.set_major_locator(mdates.MonthLocator())
    plt.xticks(rotation=45); plt.tight_layout()
    fig.savefig(str(out / "vol_entrada_diferencas.png"), **SAVE_KW); plt.close(fig)
    gerados.append("vol_entrada_diferencas.png")

    # 1.3 Histograma
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 5))
    ax1.hist(df["Dif_Abs_Calc"], bins=30, color="steelblue", edgecolor="white", alpha=0.8)
    ax1.axvline(x=0, color="red", linewidth=1.5, linestyle="--")
    ax1.axvline(x=df["Dif_Abs_Calc"].mean(), color="orange", linewidth=1.5,
                linestyle="--", label=f'Média: {df["Dif_Abs_Calc"].mean():.3f}')
    ax1.set_title("Distribuição da Diferença Absoluta", fontweight="bold")
    ax1.set_xlabel("Diferença (Nm³)"); ax1.set_ylabel("Frequência"); ax1.legend()
    ax2.hist(df["Dif_Pct_Calc"], bins=30, color="darkorange", edgecolor="white", alpha=0.8)
    ax2.axvline(x=0, color="red", linewidth=1.5, linestyle="--")
    ax2.set_title("Distribuição da Diferença Percentual", fontweight="bold")
    ax2.set_xlabel("Diferença (%)"); ax2.set_ylabel("Frequência")
    plt.tight_layout()
    fig.savefig(str(out / "vol_entrada_histograma.png"), **SAVE_KW); plt.close(fig)
    gerados.append("vol_entrada_histograma.png")

    # 1.4 Boxplot mensal
    fig, ax = plt.subplots(figsize=(14, 6))
    meses = df["Mes"].unique()
    dados_box = [df[df["Mes"] == m]["Concessionaria_Nm3d"].values / 1000 for m in meses]
    ax.boxplot(dados_box, labels=meses, patch_artist=True,
               boxprops=dict(facecolor="lightblue", alpha=0.7),
               medianprops=dict(color="red", linewidth=2))
    ax.set_title("Distribuição Mensal dos Volumes de Entrada (Concessionária)", fontsize=14, fontweight="bold")
    ax.set_xlabel("Mês"); ax.set_ylabel("Volume (10³ Nm³/d)")
    ax.grid(True, alpha=0.3, axis="y"); plt.tight_layout()
    fig.savefig(str(out / "vol_entrada_boxplot.png"), **SAVE_KW); plt.close(fig)
    gerados.append("vol_entrada_boxplot.png")

    return gerados


# ---------------------------------------------------------------------------
# 2. PCS (NB03) — 2 gráficos
# ---------------------------------------------------------------------------
def gerar_graficos_pcs(excel_path: str | Path, out: Path) -> list[str]:
    df = _load_pcs(excel_path)
    gerados = []

    # 2.1 Série temporal
    fig, ax = plt.subplots(figsize=(16, 6))
    ax.plot(df["Data"], df["PCS_Conc_kcal"], label="Concessionária", color="#4CAF50", linewidth=1.5, alpha=0.9)
    ax.plot(df["Data"], df["PCS_Transp_kcal"], label="Transportadora", color="#FF9800", linewidth=1.5, alpha=0.7, linestyle="--")
    media = df["PCS_Conc_kcal"].mean()
    ax.axhline(y=media, color="gray", linestyle=":", alpha=0.5, label=f"Média: {media:,.0f} kcal/m³")
    ax.set_title("PCS de Entrada Diário - Concessionária vs Transportadora", fontsize=14, fontweight="bold")
    ax.set_xlabel("Data"); ax.set_ylabel("PCS (kcal/m³)")
    ax.legend(); ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b/%Y"))
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    plt.xticks(rotation=45); plt.tight_layout()
    fig.savefig(str(out / "pcs_serie.png"), **SAVE_KW); plt.close(fig)
    gerados.append("pcs_serie.png")

    # 2.2 Histograma
    fig, ax = plt.subplots(figsize=(12, 5))
    ax.hist(df["PCS_Conc_kcal"], bins=30, color="#4CAF50", edgecolor="white", alpha=0.7, label="Concessionária")
    ax.hist(df["PCS_Transp_kcal"], bins=30, color="#FF9800", edgecolor="white", alpha=0.5, label="Transportadora")
    ax.axvline(x=df["PCS_Conc_kcal"].mean(), color="green", linewidth=2,
               linestyle="--", label=f'Média Conc: {df["PCS_Conc_kcal"].mean():,.0f}')
    ax.axvline(x=df["PCS_Conc_kcal"].median(), color="darkgreen", linewidth=2,
               linestyle=":", label=f'Mediana Conc: {df["PCS_Conc_kcal"].median():,.0f}')
    ax.set_title("Distribuição do PCS de Entrada", fontsize=14, fontweight="bold")
    ax.set_xlabel("PCS (kcal/m³)"); ax.set_ylabel("Frequência (dias)")
    ax.legend(); plt.tight_layout()
    fig.savefig(str(out / "pcs_histograma.png"), **SAVE_KW); plt.close(fig)
    gerados.append("pcs_histograma.png")

    return gerados


# ---------------------------------------------------------------------------
# 3. Energia (NB04) — 4 gráficos
# ---------------------------------------------------------------------------
def gerar_graficos_energia(excel_path: str | Path, out: Path) -> list[str]:
    df_vol = _load_volumes(excel_path)[["Data", "Concessionaria_Nm3d", "Transportadora_Nm3d"]]
    df_vol.columns = ["Data", "Vol_Conc_Nm3d", "Vol_Transp_Nm3d"]
    df_pcs = _load_pcs(excel_path)[["Data", "PCS_Conc_kcal", "PCS_Transp_kcal"]]
    df_pcs.columns = ["Data", "PCS_Conc", "PCS_Transp"]

    df = pd.merge(df_vol, df_pcs, on="Data", how="inner")
    df["Energia_Conc_Calc"] = df["Vol_Conc_Nm3d"] * df["PCS_Conc"]
    df["Energia_Transp_Calc"] = df["Vol_Transp_Nm3d"] * df["PCS_Transp"]
    df["Dif_Energia_Abs"] = df["Energia_Conc_Calc"] - df["Energia_Transp_Calc"]
    df["Energia_Conc_Gcal"] = df["Energia_Conc_Calc"] / 1e6
    df["Energia_Transp_Gcal"] = df["Energia_Transp_Calc"] / 1e6
    df["Mes"] = df["Data"].dt.to_period("M").astype(str)
    gerados = []

    # 3.1 Série temporal
    fig, ax = plt.subplots(figsize=(16, 6))
    ax.plot(df["Data"], df["Energia_Conc_Gcal"], label="Concessionária", color="#9C27B0", linewidth=1.5, alpha=0.9)
    ax.plot(df["Data"], df["Energia_Transp_Gcal"], label="Transportadora", color="#FF9800", linewidth=1.5, alpha=0.7, linestyle="--")
    media = df["Energia_Conc_Gcal"].mean()
    ax.axhline(y=media, color="gray", linestyle=":", alpha=0.5, label=f"Média: {media:,.0f} Gcal/d")
    ax.set_title("Energia de Entrada Diária - Concessionária vs Transportadora", fontsize=14, fontweight="bold")
    ax.set_xlabel("Data"); ax.set_ylabel("Energia (Gcal/d)")
    ax.legend(); ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b/%Y"))
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    plt.xticks(rotation=45); plt.tight_layout()
    fig.savefig(str(out / "energia_serie.png"), **SAVE_KW); plt.close(fig)
    gerados.append("energia_serie.png")

    # 3.2 Diferenças
    fig, ax = plt.subplots(figsize=(16, 5))
    ax.bar(df["Data"], df["Dif_Energia_Abs"] / 1e6, color="purple", alpha=0.6, width=1)
    ax.axhline(y=0, color="red", linewidth=0.8)
    ax.set_title("Diferença de Energia (Concessionária - Transportadora)", fontweight="bold")
    ax.set_ylabel("Diferença (Gcal)"); ax.set_xlabel("Data"); ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b/%Y"))
    plt.xticks(rotation=45); plt.tight_layout()
    fig.savefig(str(out / "energia_diferencas.png"), **SAVE_KW); plt.close(fig)
    gerados.append("energia_diferencas.png")

    # 3.3 Mensal
    mensal = df.groupby("Mes").agg({"Energia_Conc_Gcal": "sum", "Energia_Transp_Gcal": "sum"}).reset_index()
    fig, ax = plt.subplots(figsize=(14, 6))
    x = np.arange(len(mensal)); width = 0.35
    bars1 = ax.bar(x - width / 2, mensal["Energia_Conc_Gcal"] / 1000, width, label="Concessionária", color="#9C27B0", alpha=0.8)
    ax.bar(x + width / 2, mensal["Energia_Transp_Gcal"] / 1000, width, label="Transportadora", color="#FF9800", alpha=0.8)
    for bar in bars1:
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.5,
                f"{bar.get_height():,.0f}", ha="center", va="bottom", fontsize=9)
    ax.set_title("Energia Mensal Acumulada (Tcal = 10³ Gcal)", fontsize=14, fontweight="bold")
    ax.set_xlabel("Mês"); ax.set_ylabel("Energia (Tcal)")
    ax.set_xticks(x); ax.set_xticklabels(mensal["Mes"])
    ax.legend(); ax.grid(True, alpha=0.3, axis="y"); plt.tight_layout()
    fig.savefig(str(out / "energia_mensal.png"), **SAVE_KW); plt.close(fig)
    gerados.append("energia_mensal.png")

    # 3.4 Scatter
    fig, ax = plt.subplots(figsize=(10, 8))
    scatter = ax.scatter(df["Vol_Conc_Nm3d"] / 1000, df["Energia_Conc_Gcal"],
                         c=df["PCS_Conc"], cmap="RdYlGn", alpha=0.7, s=30, edgecolors="gray", linewidth=0.5)
    z = np.polyfit(df["Vol_Conc_Nm3d"] / 1000, df["Energia_Conc_Gcal"], 1)
    p = np.poly1d(z)
    x_trend = np.linspace(df["Vol_Conc_Nm3d"].min() / 1000, df["Vol_Conc_Nm3d"].max() / 1000, 100)
    ax.plot(x_trend, p(x_trend), "r--", linewidth=2, alpha=0.7, label="Tendência Linear")
    plt.colorbar(scatter, ax=ax, label="PCS (kcal/m³)")
    ax.set_title("Relação Volume × Energia (colorido por PCS)", fontsize=14, fontweight="bold")
    ax.set_xlabel("Volume (10³ Nm³/d)"); ax.set_ylabel("Energia (Gcal/d)")
    ax.legend(); ax.grid(True, alpha=0.3)
    corr = df["Vol_Conc_Nm3d"].corr(df["Energia_Conc_Calc"])
    ax.text(0.05, 0.95, f"Correlação: {corr:.6f}", transform=ax.transAxes,
            fontsize=12, verticalalignment="top",
            bbox=dict(boxstyle="round", facecolor="wheat", alpha=0.5))
    plt.tight_layout()
    fig.savefig(str(out / "energia_scatter.png"), **SAVE_KW); plt.close(fig)
    gerados.append("energia_scatter.png")

    return gerados


# ---------------------------------------------------------------------------
# 4. Clientes (NB05) — 6 gráficos
# ---------------------------------------------------------------------------
def gerar_graficos_clientes(excel_path: str | Path, out: Path) -> list[str]:
    dados_clientes = _load_clientes(excel_path)
    cores = plt.cm.Set2(np.linspace(0, 1, 7))
    gerados = []

    # 4.1 Séries temporais
    fig, axes = plt.subplots(4, 2, figsize=(18, 20)); axes = axes.flatten()
    plot_idx = 0
    for aba, info in dados_clientes.items():
        ax = axes[plot_idx]
        if info["sem_dados"]:
            vol_ref = VOLUMES_REFERENCIA.get(info["nome"], 0)
            ax.text(0.5, 0.5, f'{info["nome"]}\nSem dados horários\nVol. referência: {vol_ref:,.0f} Nm³',
                    ha="center", va="center", fontsize=12, transform=ax.transAxes,
                    bbox=dict(boxstyle="round", facecolor="lightyellow"))
            ax.set_title(f'{info["nome"]} (sem dados)', fontweight="bold", fontsize=11)
        else:
            df = info["dados"]
            ax.plot(df["Data"], df["Volume_Nm3h"], color=cores[plot_idx], alpha=0.5, linewidth=0.3)
            mm = df["Volume_Nm3h"].rolling(window=24, center=True).mean()
            ax.plot(df["Data"], mm, color="red", alpha=0.8, linewidth=1, label="MM 24h")
            vol_total = df["Volume_Nm3h"].sum() / 1e6
            ax.set_title(f'{info["nome"]} ({vol_total:.1f} Mm³)', fontweight="bold", fontsize=11)
            ax.set_ylabel("Nm³/h"); ax.grid(True, alpha=0.3); ax.legend(fontsize=8)
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%b"))
        plot_idx += 1
    axes[-1].set_visible(False)
    fig.suptitle("Séries Temporais de Volume por Cliente (horário)", fontsize=16, fontweight="bold", y=1.01)
    plt.tight_layout()
    fig.savefig(str(out / "clientes_serie.png"), **SAVE_KW); plt.close(fig)
    gerados.append("clientes_serie.png")

    # 4.2 Perfil horário
    fig, axes = plt.subplots(2, 4, figsize=(20, 10)); axes = axes.flatten()
    plot_idx = 0
    for aba, info in dados_clientes.items():
        ax = axes[plot_idx]
        if info["sem_dados"]:
            ax.text(0.5, 0.5, "Sem dados", ha="center", va="center", fontsize=12, transform=ax.transAxes)
            ax.set_title(info["nome"] + " *", fontweight="bold", fontsize=10)
        else:
            df = info["dados"].copy()
            df["Hora"] = df["Data"].dt.hour
            perfil = df.groupby("Hora")["Volume_Nm3h"].agg(["mean", "std"])
            ax.fill_between(perfil.index, perfil["mean"] - perfil["std"],
                            perfil["mean"] + perfil["std"], alpha=0.2, color=cores[plot_idx])
            ax.plot(perfil.index, perfil["mean"], color=cores[plot_idx], linewidth=2, marker="o", markersize=3)
            ax.set_title(info["nome"], fontweight="bold", fontsize=10)
            ax.set_xlabel("Hora do dia"); ax.set_ylabel("Vol médio (Nm³/h)")
            ax.set_xticks(range(0, 24, 4)); ax.grid(True, alpha=0.3)
        plot_idx += 1
    axes[-1].set_visible(False)
    fig.suptitle("Perfil Horário Médio por Cliente (faixa = ±1 desvio padrão)", fontsize=14, fontweight="bold")
    plt.tight_layout()
    fig.savefig(str(out / "clientes_perfil_horario.png"), **SAVE_KW); plt.close(fig)
    gerados.append("clientes_perfil_horario.png")

    # 4.3 Heatmap (Empresa A)
    df_a = dados_clientes["Cliente #1"]["dados"].copy()
    if not dados_clientes["Cliente #1"]["sem_dados"] and len(df_a) > 0:
        df_a["Dia"] = df_a["Data"].dt.date
        df_a["Hora"] = df_a["Data"].dt.hour
        pivot = df_a.pivot_table(values="Volume_Nm3h", index="Hora", columns="Dia", aggfunc="mean")
        fig, ax = plt.subplots(figsize=(20, 8))
        step = max(1, len(pivot.columns) // 30)
        sns.heatmap(pivot, ax=ax, cmap="YlOrRd", xticklabels=step, yticklabels=1,
                    cbar_kws={"label": "Volume (Nm³/h)"})
        ax.set_title("Empresa A - Mapa de Calor do Consumo (Hora × Dia)", fontsize=14, fontweight="bold")
        ax.set_xlabel("Dia"); ax.set_ylabel("Hora do Dia")
        plt.xticks(rotation=90, fontsize=7); plt.tight_layout()
        fig.savefig(str(out / "clientes_heatmap.png"), **SAVE_KW); plt.close(fig)
        gerados.append("clientes_heatmap.png")

    # 4.4 Pressão e temperatura
    fig, axes = plt.subplots(4, 2, figsize=(18, 20)); axes = axes.flatten()
    plot_idx = 0
    for aba, info in dados_clientes.items():
        ax = axes[plot_idx]
        if info["sem_dados"]:
            ax.text(0.5, 0.5, "Sem dados", ha="center", va="center", fontsize=12, transform=ax.transAxes)
            ax.set_title(info["nome"] + " *", fontweight="bold", fontsize=10)
        else:
            df = info["dados"]
            c1 = "#2196F3"
            ax.plot(df["Data"], df["Pressao_bara"], color=c1, alpha=0.4, linewidth=0.3)
            ax.plot(df["Data"], df["Pressao_bara"].rolling(window=24).mean(), color=c1, linewidth=1.5, label="Pressão (MM 24h)")
            ax.set_ylabel("Pressão (bara)", color=c1); ax.tick_params(axis="y", labelcolor=c1)
            ax2 = ax.twinx()
            c2 = "#F44336"
            ax2.plot(df["Data"], df["Temperatura_C"], color=c2, alpha=0.4, linewidth=0.3)
            ax2.plot(df["Data"], df["Temperatura_C"].rolling(window=24).mean(), color=c2, linewidth=1.5, label="Temp (MM 24h)")
            ax2.set_ylabel("Temperatura (°C)", color=c2); ax2.tick_params(axis="y", labelcolor=c2)
            ax.set_title(info["nome"], fontweight="bold", fontsize=11)
            ax.grid(True, alpha=0.2); ax.xaxis.set_major_formatter(mdates.DateFormatter("%b"))
            lines1, labels1 = ax.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax.legend(lines1 + lines2, labels1 + labels2, fontsize=7, loc="upper right")
        plot_idx += 1
    axes[-1].set_visible(False)
    fig.suptitle("Condições Operacionais: Pressão e Temperatura por Cliente", fontsize=16, fontweight="bold", y=1.01)
    plt.tight_layout()
    fig.savefig(str(out / "clientes_pressao_temp.png"), **SAVE_KW); plt.close(fig)
    gerados.append("clientes_pressao_temp.png")

    # 4.5 Participação
    vol_total = {}
    for aba, info in dados_clientes.items():
        nome = info["nome"]
        if info["sem_dados"]:
            vol_total[nome] = VOLUMES_REFERENCIA.get(nome, 0) / 1e6
        else:
            vol_total[nome] = info["dados"]["Volume_Nm3h"].sum() / 1e6
    nomes = list(vol_total.keys()); volumes = list(vol_total.values())
    total = sum(volumes); pcts = [v / total * 100 for v in volumes]
    ordem = np.argsort(volumes)[::-1]
    nomes_ord = [nomes[i] for i in ordem]; volumes_ord = [volumes[i] for i in ordem]; pcts_ord = [pcts[i] for i in ordem]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 7))
    cores_bar = plt.cm.Set2(np.linspace(0, 1, len(nomes_ord)))
    bars = ax1.barh(nomes_ord[::-1], volumes_ord[::-1], color=cores_bar)
    for bar, vol, pct in zip(bars, volumes_ord[::-1], pcts_ord[::-1]):
        ax1.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height() / 2,
                 f"{vol:.1f} Mm³ ({pct:.1f}%)", va="center", fontsize=10)
    ax1.set_title("Volume Total por Cliente", fontweight="bold", fontsize=13)
    ax1.set_xlabel("Volume (Mm³)")
    explode = [0.05 if v == max(volumes_ord) else 0 for v in volumes_ord]
    wedges, texts, autotexts = ax2.pie(
        volumes_ord, labels=nomes_ord, autopct="%1.1f%%",
        colors=plt.cm.Set2(np.linspace(0, 1, len(nomes_ord))),
        explode=explode, pctdistance=0.75, startangle=90)
    ax2.add_patch(plt.Circle((0, 0), 0.5, fc="white"))
    ax2.text(0, 0, f"Total\n{total:.1f} Mm³", ha="center", va="center", fontsize=14, fontweight="bold")
    ax2.set_title("Participação no Distrito", fontweight="bold", fontsize=13)
    for t in autotexts: t.set_fontsize(9)
    for t in texts: t.set_fontsize(8)
    plt.tight_layout()
    fig.savefig(str(out / "clientes_participacao.png"), **SAVE_KW); plt.close(fig)
    gerados.append("clientes_participacao.png")

    # 4.6 Boxplot
    fig, ax = plt.subplots(figsize=(14, 7))
    box_data, box_labels, box_cores = [], [], []
    color_idx = 0
    for aba, info in dados_clientes.items():
        if not info["sem_dados"]:
            box_data.append(info["dados"]["Volume_Nm3h"].dropna().values)
            box_labels.append(info["nome"])
            box_cores.append(cores[color_idx])
        color_idx += 1
    bp = ax.boxplot(box_data, labels=box_labels, patch_artist=True, vert=True,
                    showfliers=False, boxprops=dict(alpha=0.7),
                    medianprops=dict(color="red", linewidth=2))
    for patch, color in zip(bp["boxes"], box_cores):
        patch.set_facecolor(color)
    ax.set_title("Distribuição de Volumes por Cliente (sem outliers)", fontsize=14, fontweight="bold")
    ax.set_ylabel("Volume (Nm³/h)"); ax.grid(True, alpha=0.3, axis="y")
    plt.xticks(rotation=30, ha="right"); plt.tight_layout()
    fig.savefig(str(out / "clientes_boxplot.png"), **SAVE_KW); plt.close(fig)
    gerados.append("clientes_boxplot.png")

    return gerados


# ---------------------------------------------------------------------------
# 5. Incertezas (NB06) — 3 gráficos
# ---------------------------------------------------------------------------
def gerar_graficos_incertezas(excel_path: str | Path, out: Path) -> list[str]:
    inc_entrada = [INCERTEZAS["Entrada - Tramo 101 (Comgás 1)"],
                   INCERTEZAS["Entrada - Tramo 501 (Comgás 2)"]]
    u_entrada = np.sqrt(np.sum(np.array(inc_entrada) ** 2))

    inc_saidas_keys = [k for k in INCERTEZAS if "Entrada" not in k]
    inc_saidas = [INCERTEZAS[k] for k in inc_saidas_keys]
    u_saida = np.sqrt(np.sum(np.array(inc_saidas) ** 2))
    gerados = []

    # 5.1 Barras horizontais
    fig, ax = plt.subplots(figsize=(12, 7))
    pontos = list(INCERTEZAS.keys()); valores = [v * 100 for v in INCERTEZAS.values()]
    cores_inc = ["#2196F3" if "Entrada" in p else "#FF9800" for p in pontos]
    bars = ax.barh(pontos[::-1], valores[::-1], color=cores_inc[::-1], alpha=0.8, edgecolor="gray")
    for bar, val in zip(bars, valores[::-1]):
        ax.text(bar.get_width() + 0.05, bar.get_y() + bar.get_height() / 2,
                f"{val:.2f}%", va="center", fontsize=11, fontweight="bold")
    ax.axvline(x=1.0, color="red", linestyle="--", alpha=0.7, label="Limite Fiscal (1%)")
    ax.axvline(x=3.0, color="orange", linestyle="--", alpha=0.7, label="Limite Apropriação (3%)")
    ax.set_title("Incerteza de Medição por Ponto", fontsize=14, fontweight="bold")
    ax.set_xlabel("Incerteza (%)")
    legend_elements = [mpatches.Patch(facecolor="#2196F3", label="Entrada"),
                       mpatches.Patch(facecolor="#FF9800", label="Saída")]
    ax.legend(handles=legend_elements + ax.get_legend_handles_labels()[0][:2], loc="lower right")
    ax.grid(True, alpha=0.3, axis="x"); plt.tight_layout()
    fig.savefig(str(out / "incertezas_barras.png"), **SAVE_KW); plt.close(fig)
    gerados.append("incertezas_barras.png")

    # 5.2 RSS combinada
    fig, ax = plt.subplots(figsize=(10, 6))
    cats = ["Entrada\n(combinada)", "Saídas\n(combinada)"]
    u_vals = [u_entrada * 100, u_saida * 100]; cores_rss = ["#2196F3", "#FF9800"]
    bars = ax.bar(cats, u_vals, color=cores_rss, alpha=0.8, width=0.5, edgecolor="gray", linewidth=1.5)
    for bar, val in zip(bars, u_vals):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.1,
                f"{val:.2f}%", ha="center", va="bottom", fontsize=16, fontweight="bold")
    ax.axhline(y=1.0, color="red", linestyle="--", alpha=0.5, label="Limite Fiscal (1%)")
    ax.set_title("Incerteza Combinada (RSS) - Entrada vs Saídas", fontsize=14, fontweight="bold")
    ax.set_ylabel("Incerteza (%)"); ax.legend()
    ax.grid(True, alpha=0.3, axis="y"); ax.set_ylim(0, max(u_vals) * 1.3)
    plt.tight_layout()
    fig.savefig(str(out / "incertezas_rss.png"), **SAVE_KW); plt.close(fig)
    gerados.append("incertezas_rss.png")

    # 5.3 Contribuição
    contribuicoes = [(inc ** 2) / sum(x ** 2 for x in inc_saidas) * 100 for inc in inc_saidas]
    nomes_saida = inc_saidas_keys
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
    cores_contrib = plt.cm.Set2(np.linspace(0, 1, len(nomes_saida)))
    left = 0
    for i, (nome, contrib) in enumerate(zip(nomes_saida, contribuicoes)):
        ax1.barh("Incerteza\nSaída", contrib, left=left, color=cores_contrib[i],
                 label=f"{nome} ({contrib:.1f}%)", edgecolor="white")
        if contrib > 5:
            ax1.text(left + contrib / 2, 0, f"{contrib:.1f}%", ha="center", va="center", fontsize=9)
        left += contrib
    ax1.set_title("Contribuição na Incerteza (u²/Σu²)", fontweight="bold")
    ax1.set_xlabel("Contribuição (%)")
    ax1.legend(bbox_to_anchor=(1.02, 1), loc="upper left", fontsize=9)
    ax2.pie(contribuicoes, labels=nomes_saida, autopct="%1.1f%%", colors=cores_contrib, pctdistance=0.75)
    ax2.set_title("Contribuição na Incerteza Total das Saídas", fontweight="bold")
    plt.tight_layout()
    fig.savefig(str(out / "incertezas_contribuicao.png"), **SAVE_KW); plt.close(fig)
    gerados.append("incertezas_contribuicao.png")

    return gerados


# ---------------------------------------------------------------------------
# 6. Balanço de Massa (NB07) — 4 gráficos
# ---------------------------------------------------------------------------
def gerar_graficos_balanco(excel_path: str | Path, out: Path) -> list[str]:
    # Load data
    df_vol = _load_volumes(excel_path)
    vol_entrada = df_vol["Concessionaria_Nm3d"].sum()
    dados_clientes = _load_clientes(excel_path)

    volumes_clientes = {}
    for aba, info in dados_clientes.items():
        nome = info["nome"]
        if info["sem_dados"]:
            volumes_clientes[nome] = VOLUMES_REFERENCIA.get(nome, 0)
        else:
            volumes_clientes[nome] = info["dados"]["Volume_Nm3h"].sum()
    vol_saida_total = sum(volumes_clientes.values())
    diferenca = vol_entrada - vol_saida_total
    diferenca_pct = (diferenca / vol_entrada) * 100

    inc_entrada = [INCERTEZAS["Entrada - Tramo 101 (Comgás 1)"],
                   INCERTEZAS["Entrada - Tramo 501 (Comgás 2)"]]
    u_entrada = np.sqrt(sum(x ** 2 for x in inc_entrada))
    incertezas_clientes = {k: v for k, v in INCERTEZAS.items() if "Entrada" not in k}
    u_saida = np.sqrt(sum(v ** 2 for v in incertezas_clientes.values()))

    entrada_min = vol_entrada * (1 - u_entrada)
    entrada_max = vol_entrada * (1 + u_entrada)
    saida_min = vol_saida_total * (1 - u_saida)
    saida_max = vol_saida_total * (1 + u_saida)
    sobrepoe = entrada_min <= saida_max and saida_min <= entrada_max
    gerados = []

    # 6.1 Barras com erro
    fig, ax = plt.subplots(figsize=(12, 8))
    cats = ["Entrada", "Saída Total"]; vals = [vol_entrada / 1e6, vol_saida_total / 1e6]
    erros_b = [(vol_entrada - entrada_min) / 1e6, (vol_saida_total - saida_min) / 1e6]
    erros_c = [(entrada_max - vol_entrada) / 1e6, (saida_max - vol_saida_total) / 1e6]
    cores_b = ["#2196F3", "#FF9800"]
    bars = ax.bar(cats, vals, color=cores_b, alpha=0.8, width=0.5, edgecolor="gray", linewidth=1.5)
    ax.errorbar(cats, vals, yerr=[erros_b, erros_c], fmt="none", color="black", linewidth=2, capsize=15, capthick=2)
    for bar, val, inc in zip(bars, vals, [u_entrada, u_saida]):
        ax.text(bar.get_x() + bar.get_width() / 2, val * 1.01,
                f"{val:,.1f} Mm³\n(±{inc * 100:.2f}%)",
                ha="center", va="bottom", fontsize=13, fontweight="bold")
    ax.annotate(f"Diferença: {diferenca_pct:.2f}%\n({diferenca / 1e6:.2f} Mm³)",
                xy=(0.5, (vals[0] + vals[1]) / 2), fontsize=14, ha="center", va="center",
                bbox=dict(boxstyle="round,pad=0.5", facecolor="lightyellow", edgecolor="orange"))
    ax.set_title("Balanço de Massa - Entrada vs Saída Total\n(com bandas de incerteza)", fontsize=14, fontweight="bold")
    ax.set_ylabel("Volume (Mm³)"); ax.grid(True, alpha=0.3, axis="y"); plt.tight_layout()
    fig.savefig(str(out / "balanco_barras.png"), **SAVE_KW); plt.close(fig)
    gerados.append("balanco_barras.png")

    # 6.2 Waterfall
    fig, ax = plt.subplots(figsize=(16, 8))
    clientes_ord = sorted(volumes_clientes.items(), key=lambda x: x[1], reverse=True)
    labels = ["Entrada"] + [c[0] for c in clientes_ord] + ["Diferença"]
    valores_wf = [vol_entrada / 1e6] + [-c[1] / 1e6 for c in clientes_ord] + [diferenca / 1e6]
    running = []; total_run = 0
    for val in valores_wf: running.append(total_run); total_run += val
    cores_wf = ["#2196F3"] + ["#FF9800"] * len(clientes_ord) + ["#4CAF50" if diferenca > 0 else "#F44336"]
    for i, (label, val) in enumerate(zip(labels, valores_wf)):
        if i == 0: bottom, height = 0, val
        elif i == len(labels) - 1: bottom, height = 0, running[i]
        else: bottom = running[i] + val; height = abs(val)
        ax.bar(i, height, bottom=bottom, color=cores_wf[i], alpha=0.8, edgecolor="gray", linewidth=0.5)
        ax.text(i, bottom + height / 2, f"{abs(val):.1f}", ha="center", va="center",
                fontsize=9, fontweight="bold", color="white" if abs(val) > 5 else "black")
    for i in range(len(labels) - 1):
        y = valores_wf[0] if i == 0 else running[i] + valores_wf[i]
        ax.plot([i + 0.4, i + 0.6], [y, y], color="gray", linewidth=0.8, linestyle="--")
    ax.set_xticks(range(len(labels))); ax.set_xticklabels(labels, rotation=30, ha="right", fontsize=10)
    ax.set_title("Waterfall Chart - Decomposição do Balanço de Massa", fontsize=14, fontweight="bold")
    ax.set_ylabel("Volume (Mm³)"); ax.grid(True, alpha=0.3, axis="y"); plt.tight_layout()
    fig.savefig(str(out / "balanco_waterfall.png"), **SAVE_KW); plt.close(fig)
    gerados.append("balanco_waterfall.png")

    # 6.3 Bandas
    fig, ax = plt.subplots(figsize=(14, 6))
    y_ent, y_sai, altura = 1.5, 0.5, 0.6
    ax.barh(y_ent, (entrada_max - entrada_min) / 1e6, left=entrada_min / 1e6, height=altura,
            color="#2196F3", alpha=0.3, edgecolor="#2196F3", linewidth=2, label="Banda Entrada")
    ax.plot(vol_entrada / 1e6, y_ent, "D", color="#2196F3", markersize=12, zorder=5)
    ax.barh(y_sai, (saida_max - saida_min) / 1e6, left=saida_min / 1e6, height=altura,
            color="#FF9800", alpha=0.3, edgecolor="#FF9800", linewidth=2, label="Banda Saída")
    ax.plot(vol_saida_total / 1e6, y_sai, "D", color="#FF9800", markersize=12, zorder=5)
    if sobrepoe:
        ov_min = max(entrada_min, saida_min) / 1e6; ov_max = min(entrada_max, saida_max) / 1e6
        ax.axvspan(ov_min, ov_max, alpha=0.2, color="green", label="Sobreposição")
    ax.text(vol_entrada / 1e6, y_ent + 0.4,
            f"Entrada: {vol_entrada / 1e6:,.1f} Mm³ (±{u_entrada * 100:.2f}%)",
            ha="center", fontsize=11, fontweight="bold", color="#2196F3")
    ax.text(vol_saida_total / 1e6, y_sai - 0.4,
            f"Saída: {vol_saida_total / 1e6:,.1f} Mm³ (±{u_saida * 100:.2f}%)",
            ha="center", fontsize=11, fontweight="bold", color="#FF9800")
    ax.set_yticks([y_sai, y_ent]); ax.set_yticklabels(["Saídas", "Entrada"], fontsize=13)
    ax.set_xlabel("Volume (Mm³)", fontsize=12)
    ax.set_title("Bandas de Incerteza - Entrada vs Saídas", fontsize=14, fontweight="bold")
    ax.legend(loc="upper right"); ax.grid(True, alpha=0.3, axis="x"); ax.set_ylim(-0.2, 2.5)
    plt.tight_layout()
    fig.savefig(str(out / "balanco_bandas.png"), **SAVE_KW); plt.close(fig)
    gerados.append("balanco_bandas.png")

    # 6.4 Dashboard
    fig, axes = plt.subplots(1, 3, figsize=(18, 6))
    # Panel 1: Volumes
    ax = axes[0]
    ax.bar(["Entrada", "Saída"], [vol_entrada / 1e6, vol_saida_total / 1e6],
           color=["#2196F3", "#FF9800"], alpha=0.8, width=0.5)
    ax.set_title("Volumes (Mm³)", fontweight="bold", fontsize=13); ax.set_ylabel("Volume (Mm³)")
    for i, val in enumerate([vol_entrada / 1e6, vol_saida_total / 1e6]):
        ax.text(i, val + 1, f"{val:,.1f}", ha="center", fontsize=13, fontweight="bold")
    ax.grid(True, alpha=0.3, axis="y")
    # Panel 2: Gauge
    ax = axes[1]
    cor_dif = "#4CAF50" if abs(diferenca_pct) < 5 else "#F44336"
    theta = np.linspace(0, np.pi, 100)
    ax.plot(np.cos(theta), np.sin(theta), "lightgray", linewidth=20, solid_capstyle="round")
    angulo = np.pi * (1 - min(abs(diferenca_pct), 10) / 10)
    ax.plot([0, 0.8 * np.cos(angulo)], [0, 0.8 * np.sin(angulo)],
            color=cor_dif, linewidth=4, solid_capstyle="round")
    ax.plot(0, 0, "ko", markersize=10)
    ax.text(0, 0.4, f"{diferenca_pct:.2f}%", ha="center", va="center", fontsize=28, fontweight="bold", color=cor_dif)
    ax.text(0, 0.15, f"({diferenca / 1e6:,.2f} Mm³)", ha="center", va="center", fontsize=11)
    ax.text(-1, -0.05, "0%", fontsize=10); ax.text(0.85, -0.05, "10%", fontsize=10)
    ax.set_xlim(-1.3, 1.3); ax.set_ylim(-0.2, 1.2)
    ax.set_title("Diferença (%)", fontweight="bold", fontsize=13); ax.set_aspect("equal"); ax.axis("off")
    # Panel 3: Result
    ax = axes[2]; ax.axis("off")
    res_cor = "#4CAF50" if sobrepoe else "#F44336"
    res_txt = "ACEITÁVEL" if sobrepoe else "NÃO ACEITÁVEL"
    res_emoji = "APROVADO" if sobrepoe else "REPROVADO"
    ax.add_patch(plt.Rectangle((0.05, 0.1), 0.9, 0.8, facecolor=res_cor, alpha=0.15,
                                edgecolor=res_cor, linewidth=3, transform=ax.transAxes))
    ax.text(0.5, 0.7, res_emoji, ha="center", va="center", fontsize=28, fontweight="bold",
            color=res_cor, transform=ax.transAxes)
    ax.text(0.5, 0.5, f"Balanço {res_txt}", ha="center", va="center", fontsize=14,
            fontweight="bold", transform=ax.transAxes)
    ax.text(0.5, 0.35, f"Incerteza Entrada: ±{u_entrada * 100:.2f}%", ha="center",
            fontsize=11, transform=ax.transAxes)
    ax.text(0.5, 0.22, f"Incerteza Saída: ±{u_saida * 100:.2f}%", ha="center",
            fontsize=11, transform=ax.transAxes)
    fig.suptitle("DASHBOARD - BALANÇO DE MASSA DO DISTRITO", fontsize=16, fontweight="bold", y=1.02)
    plt.tight_layout()
    fig.savefig(str(out / "balanco_dashboard.png"), **SAVE_KW); plt.close(fig)
    gerados.append("balanco_dashboard.png")

    return gerados


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------
GENERATORS = [
    ("volumes", "Volumes de Entrada", gerar_graficos_volumes),
    ("pcs", "Poder Calorífico Superior", gerar_graficos_pcs),
    ("energia", "Cálculo de Energia", gerar_graficos_energia),
    ("clientes", "Perfis de Clientes", gerar_graficos_clientes),
    ("incertezas", "Incertezas de Medição", gerar_graficos_incertezas),
    ("balanco", "Balanço de Massa", gerar_graficos_balanco),
]


def gerar_todos_graficos(
    excel_path: str | Path,
    output_dir: str | Path,
    on_progress: Callable | None = None,
) -> list[str]:
    """Generate all 23 graphs from Excel data.

    Args:
        excel_path: Path to the district Excel file.
        output_dir: Directory to save PNG files.
        on_progress: Optional callback(step_name, files_generated).

    Returns:
        List of generated PNG filenames.
    """
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    all_generated = []

    for group_id, label, gen_func in GENERATORS:
        logger.info(f"Gerando gráficos: {label}")
        try:
            files = gen_func(excel_path, out)
            all_generated.extend(files)
            logger.info(f"  -> {len(files)} gráficos gerados")
            if on_progress:
                on_progress({"group": group_id, "label": label, "files": files})
        except Exception as e:
            logger.error(f"Erro ao gerar gráficos de {label}: {e}")
            if on_progress:
                on_progress({"group": group_id, "label": label, "files": [], "error": str(e)})

    logger.info(f"Total: {len(all_generated)} gráficos gerados")
    return all_generated


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import argparse
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")

    parser = argparse.ArgumentParser(description="Gerar gráficos do distrito")
    parser.add_argument("--excel", type=str, help="Caminho do Excel")
    parser.add_argument("--output", type=str, default=None, help="Diretório de saída")
    args = parser.parse_args()

    from config import DATA_DIR, GRAFICOS_DIR, EXCEL_DEFAULT

    excel = args.excel or str(DATA_DIR / EXCEL_DEFAULT)
    output = args.output or str(GRAFICOS_DIR)

    gerados = gerar_todos_graficos(excel, output)
    print(f"\n{len(gerados)} gráficos gerados em {output}")
