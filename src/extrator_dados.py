# -*- coding: utf-8 -*-
"""
Extração dinâmica de dados estatísticos a partir do Excel do distrito.
Reutiliza os loaders de graph_generator.py e preenche as dataclasses de dados_distrito.py.

Uso:
    from extrator_dados import extrair_todos
    dados = extrair_todos("caminho/para/excel.xlsx")
    # dados = {2: VolumesEntrada(...), 3: PCSData(...), ...}
"""
import json
import logging
import math
from dataclasses import asdict
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

from graph_generator import (
    _load_volumes, _load_pcs, _load_clientes,
    CLIENTES, INCERTEZAS, VOLUMES_REFERENCIA,
)
from dados_distrito import (
    DistritoConfig, VolumesEntrada, PCSData, EnergiaData,
    PerfisClientes, ClienteInfo, IncertezasData, BalancoMassa,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Extração por seção
# ---------------------------------------------------------------------------

def extrair_config(df_vol: pd.DataFrame) -> DistritoConfig:
    """Extrai metadados do período a partir do DataFrame de volumes."""
    data_min = df_vol["Data"].min()
    data_max = df_vol["Data"].max()
    dias = (data_max - data_min).days + 1
    return DistritoConfig(
        periodo_inicio=data_min.strftime("%d/%m/%Y"),
        periodo_fim=data_max.strftime("%d/%m/%Y"),
        dias=dias,
        n_clientes=len(CLIENTES),
        entrada="Estação Comgás (Tramo 101 + Tramo 501)",
        arquivo_excel="Analise de Condições de Operação de Distrito.xlsx",
        n_abas_excel=14,
    )


def extrair_volumes(excel_path: str | Path) -> VolumesEntrada:
    """Extrai estatísticas de volumes de entrada."""
    df = _load_volumes(excel_path)
    conc = df["Concessionaria_Nm3d"]
    return VolumesEntrada(
        vol_medio_nm3d=float(conc.mean()),
        vol_min_nm3d=float(conc.min()),
        vol_max_nm3d=float(conc.max()),
        vol_desvio_nm3d=float(conc.std()),
        vol_medio_m3h=float(conc.mean() / 24),
        vol_min_m3h=float(conc.min() / 24),
        vol_max_m3h=float(conc.max() / 24),
        vol_total_nm3=float(conc.sum()),
        dif_conc_transp_media_pct=float(df["Dif_Pct_Calc"].mean() / 100),
        dif_conc_transp_max_pct=float(df["Dif_Pct_Calc"].abs().max() / 100),
    )


def extrair_pcs(excel_path: str | Path) -> PCSData:
    """Extrai estatísticas do PCS."""
    df = _load_pcs(excel_path)
    pcs_conc = df["PCS_Conc_kcal"].dropna()
    pcs_transp = df["PCS_Transp_kcal"].dropna()
    # Diferença percentual entre concessionária e transportadora
    df_clean = df.dropna(subset=["PCS_Conc_kcal", "PCS_Transp_kcal"])
    dif_pct = ((df_clean["PCS_Conc_kcal"] - df_clean["PCS_Transp_kcal"]) / df_clean["PCS_Conc_kcal"] * 100)
    return PCSData(
        media_kcal=float(pcs_conc.mean()),
        min_kcal=float(pcs_conc.min()),
        max_kcal=float(pcs_conc.max()),
        desvio_padrao=float(pcs_conc.std()),
        dif_conc_transp_media_pct=float(dif_pct.mean()),
        dif_conc_transp_max_pct=float(dif_pct.abs().max()),
    )


def extrair_energia(excel_path: str | Path) -> EnergiaData:
    """Extrai estatísticas de energia (E = Volume × PCS)."""
    df_vol = _load_volumes(excel_path)[["Data", "Concessionaria_Nm3d", "Transportadora_Nm3d"]]
    df_vol.columns = ["Data", "Vol_Conc", "Vol_Transp"]
    df_pcs = _load_pcs(excel_path)[["Data", "PCS_Conc_kcal", "PCS_Transp_kcal"]]
    df_pcs.columns = ["Data", "PCS_Conc", "PCS_Transp"]

    df = pd.merge(df_vol, df_pcs, on="Data", how="inner")
    df["E_Conc"] = df["Vol_Conc"] * df["PCS_Conc"]  # kcal
    df["E_Transp"] = df["Vol_Transp"] * df["PCS_Transp"]
    df["E_Conc_Gcal"] = df["E_Conc"] / 1e6

    # Tentar ler a aba de energia da planilha para validação
    # Colunas: Data, Vol Conc m3, PC Conc kcal/m3, Energia Conc kcal, Vol Transp m3
    dif_calc_plan_kcal = 0.0
    dif_calc_plan_pct = 0.0
    try:
        df_ene = pd.read_excel(excel_path, sheet_name="Energia Ent", header=1, usecols="B:F")
        cols = list(df_ene.columns)
        df_ene.columns = ["Data", "Vol_Plan", "PCS_Plan", "E_Plan", "Vol_Transp_Plan"]
        df_ene["Data"] = pd.to_datetime(df_ene["Data"], errors="coerce")
        df_ene["E_Plan"] = pd.to_numeric(df_ene["E_Plan"], errors="coerce")
        df_ene = df_ene.dropna(subset=["Data", "E_Plan"])
        if len(df_ene) > 0:
            merged = pd.merge(
                df[["Data", "E_Conc"]],
                df_ene[["Data", "E_Plan"]],
                on="Data", how="inner",
            )
            if len(merged) > 0:
                dif_calc_plan_kcal = float((merged["E_Conc"] - merged["E_Plan"]).mean())
                e_mean = merged["E_Plan"].mean()
                if e_mean != 0:
                    dif_calc_plan_pct = float(dif_calc_plan_kcal / e_mean * 100)
    except Exception:
        pass

    # Correlação entre volume e energia calculada
    corr = float(df["Vol_Conc"].corr(df["E_Conc_Gcal"]))

    return EnergiaData(
        media_gcal_dia=float(df["E_Conc_Gcal"].mean()),
        min_gcal_dia=float(df["E_Conc_Gcal"].min()),
        max_gcal_dia=float(df["E_Conc_Gcal"].max()),
        total_gcal=float(df["E_Conc_Gcal"].sum()),
        dif_calculado_planilha_kcal=dif_calc_plan_kcal,
        dif_calculado_planilha_pct=dif_calc_plan_pct,
        correlacao_vol_energia=corr,
    )


def extrair_perfis(excel_path: str | Path) -> PerfisClientes:
    """Extrai perfis estatísticos dos clientes."""
    dados_clientes = _load_clientes(excel_path)

    # Calcular volumes totais para participação
    volumes_totais = {}
    for aba, info in dados_clientes.items():
        nome = info["nome"]
        if info["sem_dados"]:
            volumes_totais[nome] = VOLUMES_REFERENCIA.get(nome, 0)
        else:
            volumes_totais[nome] = float(info["dados"]["Volume_Nm3h"].sum())
    soma_total = sum(volumes_totais.values())

    clientes_info = []
    for aba, info in dados_clientes.items():
        nome = info["nome"]
        df = info["dados"]
        vol_total = volumes_totais[nome]

        if info["sem_dados"] or df["Volume_Nm3h"].dropna().empty:
            clientes_info.append(ClienteInfo(
                nome=nome,
                vol_total_mm3=round(vol_total / 1e6, 2),
                vol_medio_nm3h=0,
                vol_min_nm3h=0,
                vol_max_nm3h=0,
                press_media_bara=None,
                temp_media_c=None,
                fator_carga=None,
                participacao_pct=round(vol_total / soma_total * 100, 2) if soma_total else 0,
                incerteza_pct=INCERTEZAS.get(nome, 0) * 100,
            ))
        else:
            vol = df["Volume_Nm3h"].dropna()
            press = df["Pressao_bara"].dropna()
            temp = df["Temperatura_C"].dropna()
            vol_max = float(vol.max()) if len(vol) > 0 else 1
            fator = float(vol.mean() / vol_max) if vol_max > 0 else None
            clientes_info.append(ClienteInfo(
                nome=nome,
                vol_total_mm3=round(vol_total / 1e6, 2),
                vol_medio_nm3h=round(float(vol.mean())),
                vol_min_nm3h=round(float(vol.min())),
                vol_max_nm3h=round(float(vol.max())),
                press_media_bara=round(float(press.mean()), 2) if len(press) > 0 else None,
                temp_media_c=round(float(temp.mean()), 2) if len(temp) > 0 else None,
                fator_carga=round(fator, 3) if fator is not None else None,
                participacao_pct=round(vol_total / soma_total * 100, 2) if soma_total else 0,
                incerteza_pct=INCERTEZAS.get(nome, 0) * 100,
            ))

    # Ordenar por participação decrescente
    clientes_info.sort(key=lambda c: c.participacao_pct, reverse=True)
    return PerfisClientes(clientes=clientes_info)


def extrair_incertezas() -> IncertezasData:
    """Retorna dados de incerteza (constantes de calibração)."""
    t101 = INCERTEZAS["Entrada - Tramo 101 (Comgás 1)"] * 100
    t501 = INCERTEZAS["Entrada - Tramo 501 (Comgás 2)"] * 100
    u_entrada = math.sqrt(t101**2 + t501**2)

    inc_clientes = [
        (nome, INCERTEZAS[nome] * 100)
        for nome in ["Empresa A", "Empresa B", "Empresa C", "Empresa D",
                     "Empresa E", "Empresa F", "Empresa G"]
        if nome in INCERTEZAS
    ]
    u_saida = math.sqrt(sum(u**2 for _, u in inc_clientes))

    return IncertezasData(
        tramo_101_pct=t101,
        tramo_501_pct=t501,
        u_entrada_rss_pct=round(u_entrada, 2),
        u_saida_rss_pct=round(u_saida, 2),
        limite_fiscal_pct=1.0,
        limite_apropriacao_pct=3.0,
        incertezas_clientes=inc_clientes,
    )


def extrair_balanco(excel_path: str | Path, incertezas: IncertezasData) -> BalancoMassa:
    """Extrai dados do balanço de massa."""
    df_vol = _load_volumes(excel_path)
    vol_entrada = float(df_vol["Concessionaria_Nm3d"].sum())
    dados_clientes = _load_clientes(excel_path)

    volumes_clientes = {}
    for aba, info in dados_clientes.items():
        nome = info["nome"]
        if info["sem_dados"]:
            volumes_clientes[nome] = VOLUMES_REFERENCIA.get(nome, 0)
        else:
            volumes_clientes[nome] = float(info["dados"]["Volume_Nm3h"].sum())

    vol_saida_total = sum(volumes_clientes.values())
    diferenca = vol_entrada - vol_saida_total
    diferenca_pct = (diferenca / vol_entrada * 100) if vol_entrada else 0

    u_entrada = incertezas.u_entrada_rss_pct / 100
    u_saida = incertezas.u_saida_rss_pct / 100

    entrada_min = vol_entrada * (1 - u_entrada)
    entrada_max = vol_entrada * (1 + u_entrada)
    saida_min = vol_saida_total * (1 - u_saida)
    saida_max = vol_saida_total * (1 + u_saida)
    sobrepoe = entrada_min <= saida_max and saida_min <= entrada_max

    volumes_saida = [
        (nome, round(vol), round(vol / vol_saida_total * 100, 2) if vol_saida_total else 0)
        for nome, vol in sorted(volumes_clientes.items(), key=lambda x: x[1], reverse=True)
    ]

    return BalancoMassa(
        vol_entrada_nm3=round(vol_entrada),
        vol_saida_total_nm3=round(vol_saida_total),
        diferenca_nm3=round(diferenca),
        diferenca_pct=round(diferenca_pct, 2),
        u_entrada_pct=incertezas.u_entrada_rss_pct,
        u_saida_pct=incertezas.u_saida_rss_pct,
        banda_entrada_min=round(entrada_min),
        banda_entrada_max=round(entrada_max),
        banda_saida_min=round(saida_min),
        banda_saida_max=round(saida_max),
        bandas_sobrepoem=sobrepoe,
        resultado="ACEITAVEL" if sobrepoe else "INACEITAVEL",
        volumes_saida=volumes_saida,
    )


# ---------------------------------------------------------------------------
# Orquestrador
# ---------------------------------------------------------------------------

def extrair_todos(excel_path: str | Path) -> dict:
    """Extrai todos os dados do Excel e retorna dict de dataclasses.

    Returns:
        {
            "config": DistritoConfig,
            2: VolumesEntrada,
            3: PCSData,
            4: EnergiaData,
            5: PerfisClientes,
            6: IncertezasData,
            7: BalancoMassa,
        }
    """
    logger.info("Extraindo dados do Excel...")
    excel_path = str(excel_path)

    df_vol = _load_volumes(excel_path)
    config = extrair_config(df_vol)
    logger.info(f"  Período: {config.periodo_inicio} a {config.periodo_fim} ({config.dias} dias)")

    volumes = extrair_volumes(excel_path)
    logger.info(f"  Volumes: total={volumes.vol_total_nm3:,.0f} Nm³")

    pcs = extrair_pcs(excel_path)
    logger.info(f"  PCS: média={pcs.media_kcal:,.2f} kcal/m³")

    energia = extrair_energia(excel_path)
    logger.info(f"  Energia: total={energia.total_gcal:,.0f} Gcal")

    perfis = extrair_perfis(excel_path)
    logger.info(f"  Clientes: {len(perfis.clientes)} perfis")

    incertezas = extrair_incertezas()
    logger.info(f"  Incertezas: entrada={incertezas.u_entrada_rss_pct:.2f}%, saída={incertezas.u_saida_rss_pct:.2f}%")

    balanco = extrair_balanco(excel_path, incertezas)
    logger.info(f"  Balanço: diferença={balanco.diferenca_pct:.2f}%, resultado={balanco.resultado}")

    return {
        "config": config,
        2: volumes,
        3: pcs,
        4: energia,
        5: perfis,
        6: incertezas,
        7: balanco,
    }


# ---------------------------------------------------------------------------
# Serialização JSON (para cache entre fases)
# ---------------------------------------------------------------------------

def _serialize(obj):
    """Converte dataclass para dict serializável."""
    if hasattr(obj, "__dataclass_fields__"):
        d = asdict(obj)
        d["__class__"] = type(obj).__name__
        return d
    return obj


def salvar_json(dados: dict, path: str | Path):
    """Salva dados extraídos como JSON."""
    out = {}
    for key, obj in dados.items():
        out[str(key)] = _serialize(obj)
    Path(path).write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info(f"Dados salvos em {path}")


def carregar_json(path: str | Path) -> dict:
    """Carrega dados extraídos do JSON."""
    raw = json.loads(Path(path).read_text(encoding="utf-8"))

    CLASS_MAP = {
        "DistritoConfig": DistritoConfig,
        "VolumesEntrada": VolumesEntrada,
        "PCSData": PCSData,
        "EnergiaData": EnergiaData,
        "PerfisClientes": PerfisClientes,
        "IncertezasData": IncertezasData,
        "BalancoMassa": BalancoMassa,
    }

    result = {}
    for key, d in raw.items():
        cls_name = d.pop("__class__", None)
        if cls_name == "PerfisClientes":
            # Reconstruct ClienteInfo list
            d["clientes"] = [ClienteInfo(**c) for c in d["clientes"]]
            result[key if key == "config" else int(key)] = PerfisClientes(**d)
        elif cls_name == "IncertezasData":
            d["incertezas_clientes"] = [tuple(x) for x in d["incertezas_clientes"]]
            result[key if key == "config" else int(key)] = IncertezasData(**d)
        elif cls_name == "BalancoMassa":
            d["volumes_saida"] = [tuple(x) for x in d["volumes_saida"]]
            result[key if key == "config" else int(key)] = BalancoMassa(**d)
        elif cls_name in CLASS_MAP:
            result[key if key == "config" else int(key)] = CLASS_MAP[cls_name](**d)

    logger.info(f"Dados carregados de {path}")
    return result


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import argparse
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")

    parser = argparse.ArgumentParser(description="Extrair dados do Excel do distrito")
    parser.add_argument("--excel", required=True, help="Caminho do Excel")
    parser.add_argument("--output", default=None, help="Caminho JSON de saída")
    args = parser.parse_args()

    dados = extrair_todos(args.excel)

    if args.output:
        salvar_json(dados, args.output)
    else:
        from dados_distrito import formatar_dados_secao
        for key in [2, 3, 4, 5, 6, 7]:
            print(f"\n{'='*60}")
            print(formatar_dados_secao(dados[key]))
