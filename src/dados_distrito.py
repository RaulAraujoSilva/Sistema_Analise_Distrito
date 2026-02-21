# -*- coding: utf-8 -*-
"""
Dados estaticos do distrito de distribuicao de gas natural.
Valores pre-computados pelos 7 notebooks de analise.
Periodo: abril a setembro de 2025 (183 dias), 7 clientes, 1 entrada.
"""
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class DistritoConfig:
    periodo_inicio: str = "01/04/2025"
    periodo_fim: str = "30/09/2025"
    dias: int = 183
    n_clientes: int = 7
    entrada: str = "Estacao Comgas (Tramo 101 + Tramo 501)"
    arquivo_excel: str = "Analise de Condicoes de Operacao de Distrito.xlsx"
    n_abas_excel: int = 14


@dataclass
class VolumesEntrada:
    """NB02: Analise de Volumes de Entrada"""
    vol_medio_nm3d: float = 999_562.0
    vol_min_nm3d: float = 505_965.0
    vol_max_nm3d: float = 1_240_865.0
    vol_desvio_nm3d: float = 168_609.2
    vol_medio_m3h: float = 41_648.42
    vol_min_m3h: float = 21_081.87
    vol_max_m3h: float = 51_702.70
    vol_total_nm3: float = 182_919_850.0
    dif_conc_transp_media_pct: float = -0.000009
    dif_conc_transp_max_pct: float = 0.000083
    graficos: List[str] = field(default_factory=lambda: [
        "vol_entrada_serie.png",
        "vol_entrada_diferencas.png",
        "vol_entrada_histograma.png",
        "vol_entrada_boxplot.png",
    ])


@dataclass
class PCSData:
    """NB03: Analise do PCS (Poder Calorifico Superior)"""
    media_kcal: float = 9_538.92
    min_kcal: float = 9_167.85
    max_kcal: float = 9_785.96
    desvio_padrao: float = 104.34
    dif_conc_transp_media_pct: float = 0.000046
    dif_conc_transp_max_pct: float = 0.005208
    graficos: List[str] = field(default_factory=lambda: [
        "pcs_serie.png",
        "pcs_histograma.png",
    ])


@dataclass
class EnergiaData:
    """NB04: Calculo de Energia (E = Volume x PCS)"""
    media_gcal_dia: float = 9_536.23
    min_gcal_dia: float = 4_779.25
    max_gcal_dia: float = 11_730.53
    total_gcal: float = 1_745_130.0
    dif_calculado_planilha_kcal: float = 0.0
    dif_calculado_planilha_pct: float = 0.0
    correlacao_vol_energia: float = 0.999999
    graficos: List[str] = field(default_factory=lambda: [
        "energia_serie.png",
        "energia_diferencas.png",
        "energia_mensal.png",
        "energia_scatter.png",
    ])


@dataclass
class ClienteInfo:
    nome: str
    vol_total_mm3: float
    vol_medio_nm3h: float
    vol_min_nm3h: float
    vol_max_nm3h: float
    press_media_bara: Optional[float]
    temp_media_c: Optional[float]
    fator_carga: Optional[float]
    participacao_pct: float
    incerteza_pct: float


@dataclass
class PerfisClientes:
    """NB05: Perfis dos Clientes"""
    clientes: List[ClienteInfo] = field(default_factory=lambda: [
        ClienteInfo("Empresa A", 104.10, 23_965, 1_359, 31_245, 15.47, 23.49, 0.767, 57.5, 1.33),
        ClienteInfo("Empresa B", 43.66, 10_052, 178, 17_113, 15.96, 23.41, 0.587, 24.1, 1.61),
        ClienteInfo("Empresa E", 10.18, 2_345, 300, 4_244, 4.93, 16.70, 0.552, 5.6, 3.05),
        ClienteInfo("Empresa G", 10.08, 2_321, 418, 7_085, 7.34, 18.52, 0.328, 5.6, 2.80),
        ClienteInfo("Empresa C", 6.84, 1_567, 0, 4_959, 5.15, 17.84, 0.316, 3.8, 1.34),
        ClienteInfo("Empresa F", 5.96, 1_372, 0, 3_509, 7.55, 20.49, 0.391, 3.3, 1.48),
        ClienteInfo("Empresa D", 0.09, 47, 0, 187, 18.57, 23.64, 0.253, 0.05, 3.58),
    ])
    graficos: List[str] = field(default_factory=lambda: [
        "clientes_serie.png",
        "clientes_perfil_horario.png",
        "clientes_heatmap.png",
        "clientes_pressao_temp.png",
        "clientes_participacao.png",
        "clientes_boxplot.png",
    ])


@dataclass
class IncertezasData:
    """NB06: Sumario e Incertezas de Medicao"""
    tramo_101_pct: float = 1.06
    tramo_501_pct: float = 1.09
    u_entrada_rss_pct: float = 1.52
    u_saida_rss_pct: float = 6.19
    limite_fiscal_pct: float = 1.0
    limite_apropriacao_pct: float = 3.0
    incertezas_clientes: List[tuple] = field(default_factory=lambda: [
        ("Empresa A", 1.33),
        ("Empresa B", 1.61),
        ("Empresa C", 1.34),
        ("Empresa D", 3.58),
        ("Empresa E", 3.05),
        ("Empresa F", 1.48),
        ("Empresa G", 2.80),
    ])
    graficos: List[str] = field(default_factory=lambda: [
        "incertezas_barras.png",
        "incertezas_rss.png",
        "incertezas_contribuicao.png",
    ])


@dataclass
class BalancoMassa:
    """NB07: Balanco de Massa com Bandas de Incerteza"""
    vol_entrada_nm3: float = 182_919_850.0
    vol_saida_total_nm3: float = 180_923_440.0
    diferenca_nm3: float = 1_996_410.0
    diferenca_pct: float = 1.09
    u_entrada_pct: float = 1.52
    u_saida_pct: float = 6.19
    banda_entrada_min: float = 180_138_686.0
    banda_entrada_max: float = 185_701_014.0
    banda_saida_min: float = 169_725_770.0
    banda_saida_max: float = 192_121_110.0
    bandas_sobrepoem: bool = True
    resultado: str = "ACEITAVEL"
    volumes_saida: List[tuple] = field(default_factory=lambda: [
        ("Empresa A", 104_104_553, 57.5),
        ("Empresa B", 43_664_475, 24.1),
        ("Empresa E", 10_184_645, 5.6),
        ("Empresa G", 10_081_924, 5.6),
        ("Empresa C", 6_841_747, 3.8),
        ("Empresa F", 5_957_912, 3.3),
        ("Empresa D", 88_184, 0.05),
    ])
    graficos: List[str] = field(default_factory=lambda: [
        "balanco_barras.png",
        "balanco_waterfall.png",
        "balanco_bandas.png",
        "balanco_dashboard.png",
    ])


def gerar_tabelas_resumo(dados: dict = None) -> dict:
    """Retorna tabelas de dados estatísticos para cada seção do relatório.
    Cada entrada: {"headers": [...], "rows": [[...], ...]}

    Args:
        dados: Dict com dataclasses dinâmicas {2: VolumesEntrada, 3: PCSData, ...}.
               Se None, usa valores hardcoded (retrocompatibilidade).
    """
    _cfg = dados.get("config", DistritoConfig()) if dados else DistritoConfig()
    if dados:
        vol = dados.get(2, VolumesEntrada())
        pcs = dados.get(3, PCSData())
        ene = dados.get(4, EnergiaData())
        perf = dados.get(5, PerfisClientes())
        inc = dados.get(6, IncertezasData())
        bal = dados.get(7, BalancoMassa())
    else:
        vol = VolumesEntrada()
        pcs = PCSData()
        ene = EnergiaData()
        perf = PerfisClientes()
        inc = IncertezasData()
        bal = BalancoMassa()

    tabelas = {}

    # Seção 2: Volumes de Entrada
    tabelas["secao_2_volumes"] = {
        "titulo": "Tabela 2.1: Estatísticas Descritivas dos Volumes de Entrada",
        "headers": ["Métrica", "Concessionária (Nm³/d)", "Observação"],
        "rows": [
            ["Volume Médio Diário", f"{vol.vol_medio_nm3d:,.0f}", f"{vol.vol_medio_m3h:,.2f} m³/h"],
            ["Volume Mínimo Diário", f"{vol.vol_min_nm3d:,.0f}", f"{vol.vol_min_m3h:,.2f} m³/h"],
            ["Volume Máximo Diário", f"{vol.vol_max_nm3d:,.0f}", f"{vol.vol_max_m3h:,.2f} m³/h"],
            ["Desvio Padrão", f"{vol.vol_desvio_nm3d:,.1f}", ""],
            ["Volume Total (período)", f"{vol.vol_total_nm3:,.0f}", f"{_cfg.dias} dias"],
            ["Dif. Média Conc. vs Transp.", f"{vol.dif_conc_transp_media_pct:.6f}%", "Excelente"],
            ["Dif. Máxima Conc. vs Transp.", f"{vol.dif_conc_transp_max_pct:.6f}%", "< 0,01%"],
        ],
    }

    # Seção 3: PCS
    tabelas["secao_3_pcs"] = {
        "titulo": "Tabela 3.1: Estatísticas Descritivas do PCS",
        "headers": ["Métrica", "Valor (kcal/m³)", "Observação"],
        "rows": [
            ["PCS Médio", f"{pcs.media_kcal:,.2f}", ""],
            ["PCS Mínimo", f"{pcs.min_kcal:,.2f}", ""],
            ["PCS Máximo", f"{pcs.max_kcal:,.2f}", ""],
            ["Desvio Padrão", f"{pcs.desvio_padrao:,.2f}", ""],
            ["Amplitude", f"{pcs.max_kcal - pcs.min_kcal:,.2f}", ""],
            ["Dif. Média Conc. vs Transp.", f"{pcs.dif_conc_transp_media_pct:.6f}%", ""],
        ],
    }

    # Seção 4: Energia
    tabelas["secao_4_energia"] = {
        "titulo": "Tabela 4.1: Validação do Cálculo de Energia",
        "headers": ["Métrica", "Valor", "Observação"],
        "rows": [
            ["Energia Média Diária", f"{ene.media_gcal_dia:,.2f} Gcal/d", ""],
            ["Energia Mínima Diária", f"{ene.min_gcal_dia:,.2f} Gcal/d", ""],
            ["Energia Máxima Diária", f"{ene.max_gcal_dia:,.2f} Gcal/d", ""],
            ["Energia Total", f"{ene.total_gcal:,.0f} Gcal", f"~{ene.total_gcal/1000:,.0f} TJ"],
            ["Dif. Calculado vs Planilha", f"{ene.dif_calculado_planilha_kcal:.2f} kcal", "Erro zero"],
            ["Correlação Vol × Energia", f"{ene.correlacao_vol_energia:.6f}", "r ≈ 1"],
        ],
    }

    # Seção 5: Clientes
    tabelas["secao_5_clientes"] = {
        "titulo": "Tabela 5.1: Resumo dos Clientes do Distrito",
        "headers": ["Cliente", "Vol. Total (Mm³)", "Média (Nm³/h)", "Fator Carga", "Participação (%)"],
        "rows": [
            [c.nome, f"{c.vol_total_mm3:.2f}", f"{c.vol_medio_nm3h:,}", f"{c.fator_carga:.3f}", f"{c.participacao_pct:.1f}%"]
            for c in perf.clientes
        ],
    }

    # Seção 6: Incertezas
    tabelas["secao_6_incertezas"] = {
        "titulo": "Tabela 6.1: Incertezas de Medição por Ponto",
        "headers": ["Ponto de Medição", "Incerteza (%)", "Classificação"],
        "rows": [
            ["Tramo 101 (Entrada 1)", f"{inc.tramo_101_pct:.2f}", "Entrada"],
            ["Tramo 501 (Entrada 2)", f"{inc.tramo_501_pct:.2f}", "Entrada"],
            ["Combinada RSS (Entrada)", f"{inc.u_entrada_rss_pct:.2f}", "Entrada"],
        ] + [
            [nome, f"{u:.2f}", "Saída"] for nome, u in inc.incertezas_clientes
        ] + [
            ["Combinada RSS (Saída)", f"{inc.u_saida_rss_pct:.2f}", "Saída"],
        ],
    }

    # Seção 7: Balanço
    tabelas["secao_7_balanco"] = {
        "titulo": "Tabela 7.1: Resultado do Balanço de Massa",
        "headers": ["Item", "Volume (Nm³)", "Incerteza (%)", "Banda Mín.", "Banda Máx."],
        "rows": [
            ["Entrada Total", f"{bal.vol_entrada_nm3:,.0f}", f"{bal.u_entrada_pct:.2f}",
             f"{bal.banda_entrada_min:,.0f}", f"{bal.banda_entrada_max:,.0f}"],
            ["Saída Total", f"{bal.vol_saida_total_nm3:,.0f}", f"{bal.u_saida_pct:.2f}",
             f"{bal.banda_saida_min:,.0f}", f"{bal.banda_saida_max:,.0f}"],
            ["Diferença", f"{bal.diferenca_nm3:,.0f}", f"{bal.diferenca_pct:.2f}%", "", bal.resultado],
        ],
    }

    return tabelas


def formatar_dados_secao(obj, config: DistritoConfig = None) -> str:
    """Converte qualquer dataclass de dados em texto legivel para prompt.

    Args:
        obj: Dataclass de dados (VolumesEntrada, PCSData, etc.)
        config: DistritoConfig com período dinâmico. Se None, usa valores padrão.
    """
    _cfg = config or DistritoConfig()
    if isinstance(obj, VolumesEntrada):
        return (
            "VOLUMES DE ENTRADA DO DISTRITO\n"
            f"- Periodo: {_cfg.dias} dias ({_cfg.periodo_inicio} a {_cfg.periodo_fim})\n"
            f"- Volume total: {obj.vol_total_nm3:,.0f} Nm3\n"
            f"- Volume medio diario: {obj.vol_medio_nm3d:,.1f} Nm3/d ({obj.vol_medio_m3h:,.2f} m3/h)\n"
            f"- Volume minimo diario: {obj.vol_min_nm3d:,.1f} Nm3/d ({obj.vol_min_m3h:,.2f} m3/h)\n"
            f"- Volume maximo diario: {obj.vol_max_nm3d:,.1f} Nm3/d ({obj.vol_max_m3h:,.2f} m3/h)\n"
            f"- Desvio padrao: {obj.vol_desvio_nm3d:,.1f} Nm3/d\n"
            f"- Diferenca media Concessionaria vs Transportadora: {obj.dif_conc_transp_media_pct:.6f}%\n"
            f"- Diferenca maxima Concessionaria vs Transportadora: {obj.dif_conc_transp_max_pct:.6f}%\n"
            f"- Concordancia: < 0,01% (excelente)"
        )
    elif isinstance(obj, PCSData):
        return (
            "PCS (PODER CALORIFICO SUPERIOR) NA ENTRADA\n"
            f"- PCS medio: {obj.media_kcal:,.2f} kcal/m3\n"
            f"- PCS minimo: {obj.min_kcal:,.2f} kcal/m3\n"
            f"- PCS maximo: {obj.max_kcal:,.2f} kcal/m3\n"
            f"- Desvio padrao: {obj.desvio_padrao:,.2f} kcal/m3\n"
            f"- Amplitude: {obj.max_kcal - obj.min_kcal:,.2f} kcal/m3\n"
            f"- Diferenca media Conc vs Transp: {obj.dif_conc_transp_media_pct:.6f}%\n"
            f"- Diferenca maxima Conc vs Transp: {obj.dif_conc_transp_max_pct:.6f}%"
        )
    elif isinstance(obj, EnergiaData):
        return (
            "ENERGIA DIARIA (E = Volume x PCS)\n"
            f"- Energia media diaria: {obj.media_gcal_dia:,.2f} Gcal/dia\n"
            f"- Energia minima diaria: {obj.min_gcal_dia:,.2f} Gcal/dia\n"
            f"- Energia maxima diaria: {obj.max_gcal_dia:,.2f} Gcal/dia\n"
            f"- Energia total no periodo: {obj.total_gcal:,.0f} Gcal (~{obj.total_gcal/1000:,.0f} Tcal)\n"
            f"- Validacao: diferenca calculado vs planilha = {obj.dif_calculado_planilha_kcal:.2f} kcal ({obj.dif_calculado_planilha_pct:.6f}%)\n"
            f"- Correlacao Volume-Energia: r = {obj.correlacao_vol_energia:.6f}"
        )
    elif isinstance(obj, PerfisClientes):
        lines = ["PERFIS DE CONSUMO DOS 7 CLIENTES\n"]
        lines.append("| Cliente | Vol Total (Mm3) | Vol Medio (Nm3/h) | Faixa (Nm3/h) | Pressao (bara) | Temp (C) | Fator Carga | Participacao |")
        lines.append("|---------|----------------|-------------------|---------------|----------------|----------|-------------|-------------|")
        for c in obj.clientes:
            lines.append(
                f"| {c.nome} | {c.vol_total_mm3:.2f} | {c.vol_medio_nm3h:,.0f} | "
                f"{c.vol_min_nm3h:,.0f}-{c.vol_max_nm3h:,.0f} | {c.press_media_bara:.2f} | "
                f"{c.temp_media_c:.2f} | {c.fator_carga:.3f} | {c.participacao_pct:.1f}% |"
            )
        lines.append(f"\nNota: Empresa D possui 57% de dados horarios faltantes.")
        lines.append(f"Fator de carga = Vol Medio / Vol Maximo (mais proximo de 1 = consumo mais constante)")
        return "\n".join(lines)
    elif isinstance(obj, IncertezasData):
        lines = [
            "INCERTEZAS DE MEDICAO\n",
            "Entrada:",
            f"  - Tramo 101 (Comgas 1): {obj.tramo_101_pct:.2f}%",
            f"  - Tramo 501 (Comgas 2): {obj.tramo_501_pct:.2f}%",
            f"  - Combinada RSS: sqrt({obj.tramo_101_pct/100:.4f}^2 + {obj.tramo_501_pct/100:.4f}^2) = {obj.u_entrada_rss_pct:.2f}%\n",
            "Saida (por cliente):",
        ]
        for nome, u in obj.incertezas_clientes:
            lines.append(f"  - {nome}: {u:.2f}%")
        lines.append(f"  - Combinada RSS (saida): {obj.u_saida_rss_pct:.2f}%\n")
        lines.append(f"Limites de referencia:")
        lines.append(f"  - Limite fiscal: {obj.limite_fiscal_pct:.1f}%")
        lines.append(f"  - Limite de apropriacao: {obj.limite_apropriacao_pct:.1f}%")
        return "\n".join(lines)
    elif isinstance(obj, BalancoMassa):
        lines = [
            "BALANCO DE MASSA DO DISTRITO\n",
            f"Volume de entrada: {obj.vol_entrada_nm3:,.0f} Nm3",
            f"Volume total de saida: {obj.vol_saida_total_nm3:,.0f} Nm3",
            f"Diferenca: {obj.diferenca_nm3:,.0f} Nm3 ({obj.diferenca_pct:.2f}%)\n",
            "Volumes por cliente de saida:",
        ]
        for nome, vol, pct in obj.volumes_saida:
            lines.append(f"  - {nome}: {vol:,.0f} Nm3 ({pct:.1f}%)")
        lines.append(f"\nIncertezas:")
        lines.append(f"  - U_entrada: {obj.u_entrada_pct:.2f}%")
        lines.append(f"  - U_saida: {obj.u_saida_pct:.2f}%")
        lines.append(f"\nBandas de incerteza:")
        lines.append(f"  - Entrada: [{obj.banda_entrada_min:,.0f} ; {obj.banda_entrada_max:,.0f}] Nm3")
        lines.append(f"  - Saida:   [{obj.banda_saida_min:,.0f} ; {obj.banda_saida_max:,.0f}] Nm3")
        lines.append(f"\nSobreposicao das bandas: {'SIM' if obj.bandas_sobrepoem else 'NAO'}")
        lines.append(f"Resultado: {obj.resultado}")
        return "\n".join(lines)
    else:
        return str(obj)
