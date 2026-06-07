"""
Gerador de Relatórios Automatizados — Excel para PDF
Lê uma planilha .xlsx e exporta um relatório profissional em PDF.
"""

from __future__ import annotations

import os
import io
import sys
from pathlib import Path
from datetime import datetime
from typing import Optional

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from fpdf import FPDF


# ─── Constantes de Design ────────────────────────────────────────────────────
PRIMARY     = (26, 60, 94)      # azul escuro
ACCENT      = (41, 182, 246)    # azul claro
SUCCESS     = (46, 160, 67)     # verde
DANGER      = (203, 36, 49)     # vermelho
LIGHT_BG    = (240, 246, 255)   # fundo alternado das células
WHITE       = (255, 255, 255)
DARK_TEXT   = (30, 30, 30)
GRAY        = (120, 120, 120)

CHART_DIR   = Path("__charts_temp__")


# ─── Utilitários ─────────────────────────────────────────────────────────────

def _rgb(color: tuple[int, int, int]) -> tuple[int, int, int]:
    return color


def _ensure_chart_dir() -> None:
    CHART_DIR.mkdir(exist_ok=True)


def _cleanup_charts() -> None:
    if CHART_DIR.exists():
        for f in CHART_DIR.iterdir():
            f.unlink()
        CHART_DIR.rmdir()


def _formatar_brl(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ─── Geração de Gráficos ─────────────────────────────────────────────────────

def _gerar_grafico_barras(df: pd.DataFrame, coluna_grupo: str, coluna_valor: str, titulo: str, nome_arquivo: str) -> str:
    agrupado = df.groupby(coluna_grupo)[coluna_valor].sum().sort_values(ascending=False)

    fig, ax = plt.subplots(figsize=(8, 4))
    colors = [f"#{PRIMARY[0]:02x}{PRIMARY[1]:02x}{PRIMARY[2]:02x}"] * len(agrupado)

    bars = ax.bar(agrupado.index, agrupado.values, color=colors, edgecolor="white", linewidth=0.8)

    ax.set_title(titulo, fontsize=13, fontweight="bold", color="#1A3C5E", pad=12)
    ax.set_xlabel(coluna_grupo, fontsize=10, color="#555555")
    ax.set_ylabel(coluna_valor, fontsize=10, color="#555555")
    ax.tick_params(colors="#555555", labelsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.set_facecolor("#F8FBFF")
    fig.patch.set_facecolor("white")

    for bar in bars:
        height = bar.get_height()
        ax.annotate(
            _formatar_brl(height),
            xy=(bar.get_x() + bar.get_width() / 2, height),
            xytext=(0, 4),
            textcoords="offset points",
            ha="center",
            va="bottom",
            fontsize=8,
            color="#1A3C5E",
            fontweight="bold",
        )

    plt.tight_layout()
    caminho = str(CHART_DIR / nome_arquivo)
    plt.savefig(caminho, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return caminho


def _gerar_grafico_pizza(df: pd.DataFrame, coluna_grupo: str, coluna_valor: str, titulo: str, nome_arquivo: str) -> str:
    agrupado = df.groupby(coluna_grupo)[coluna_valor].sum()

    palette = ["#1A3C5E", "#29B6F6", "#2EA043", "#E3B341", "#CB2431", "#8E44AD"]
    colors = [palette[i % len(palette)] for i in range(len(agrupado))]

    fig, ax = plt.subplots(figsize=(6, 4))
    wedges, texts, autotexts = ax.pie(
        agrupado.values,
        labels=None,
        autopct="%1.1f%%",
        colors=colors,
        startangle=140,
        wedgeprops={"edgecolor": "white", "linewidth": 2},
        pctdistance=0.82,
    )
    for at in autotexts:
        at.set_fontsize(8)
        at.set_color("white")
        at.set_fontweight("bold")

    legend_patches = [mpatches.Patch(color=colors[i], label=label) for i, label in enumerate(agrupado.index)]
    ax.legend(handles=legend_patches, loc="center left", bbox_to_anchor=(1, 0, 0.5, 1), fontsize=8, frameon=False)
    ax.set_title(titulo, fontsize=12, fontweight="bold", color="#1A3C5E", pad=10)
    fig.patch.set_facecolor("white")

    plt.tight_layout()
    caminho = str(CHART_DIR / nome_arquivo)
    plt.savefig(caminho, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return caminho


# ─── Classe PDF ───────────────────────────────────────────────────────────────

def _sanitize(texto: str) -> str:
    """Remove/substitui caracteres fora do latin-1 para compatibilidade."""
    return (
        texto
        .replace("\u2014", "-")   # em-dash
        .replace("\u2013", "-")   # en-dash
        .replace("\u2019", "'")   # aspas tipograficas
        .replace("\u201c", '"')
        .replace("\u201d", '"')
        .encode("latin-1", errors="replace")
        .decode("latin-1")
    )


class RelatorioPDF(FPDF):

    def __init__(self, titulo_relatorio: str, subtitulo: str = ""):
        super().__init__(orientation="L", unit="mm", format="A4")
        self.titulo_relatorio = _sanitize(titulo_relatorio)
        self.subtitulo = _sanitize(subtitulo)
        self.set_auto_page_break(auto=True, margin=15)
        # Usa latin-1 nativo — caracteres acentuados funcionam via encode
        self.add_page()

    # ── Cabeçalho ──────────────────────────────────────────────────────────

    def header(self):
        self.set_fill_color(*PRIMARY)
        self.rect(0, 0, 297, 22, "F")

        self.set_fill_color(*ACCENT)
        self.rect(0, 22, 297, 3, "F")

        self.set_xy(10, 4)
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(*WHITE)
        self.cell(180, 8, self.titulo_relatorio, align="L")

        self.set_xy(10, 13)
        self.set_font("Helvetica", "", 9)
        self.set_text_color(200, 220, 255)
        self.cell(180, 5, self.subtitulo, align="L")

        now_str = datetime.now().strftime("%d/%m/%Y  %H:%M")
        self.set_xy(200, 8)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(180, 210, 255)
        self.cell(87, 5, _sanitize(f"Gerado em: {now_str}"), align="R")

        self.set_y(30)

    # ── Rodapé ─────────────────────────────────────────────────────────────

    def footer(self):
        self.set_y(-12)
        self.set_fill_color(*PRIMARY)
        self.rect(0, self.get_y(), 297, 15, "F")
        self.set_font("Helvetica", "", 8)
        self.set_text_color(180, 210, 255)
        rodape = _sanitize(f"Pagina {self.page_no()} - Relatorio gerado automaticamente")
        self.cell(0, 10, rodape, align="C")

    # ── Seção ──────────────────────────────────────────────────────────────

    def secao(self, titulo: str) -> None:
        self.ln(4)
        self.set_fill_color(*LIGHT_BG)
        self.set_draw_color(*ACCENT)
        self.rect(10, self.get_y(), 277, 9, "FD")

        self.set_fill_color(*ACCENT)
        self.rect(10, self.get_y(), 3, 9, "F")

        self.set_xy(16, self.get_y() + 1)
        self.set_font("Helvetica", "B", 11)
        self.set_text_color(*PRIMARY)
        self.cell(270, 7, _sanitize(titulo), align="L")
        self.ln(11)



    def cards_kpi(self, kpis: list[dict]) -> None:
        """kpis: lista de {label, value, color?}"""
        card_w = 60
        card_h = 20
        gap = 8
        start_x = 10
        base_y = self.get_y()
        x = start_x

        for kpi in kpis:
            color = kpi.get("color", PRIMARY)
            self.set_fill_color(*color)
            self.set_draw_color(*WHITE)
            self.set_line_width(0)
            self.rect(x, base_y, card_w, card_h, "F")

            self.set_xy(x, base_y + 3)
            self.set_font("Helvetica", "B", 12)
            self.set_text_color(*WHITE)
            self.cell(card_w, 7, _sanitize(str(kpi["value"])), align="C")

            self.set_xy(x, base_y + 12)
            self.set_font("Helvetica", "", 7)
            self.set_text_color(200, 230, 255)
            self.cell(card_w, 5, _sanitize(kpi["label"]), align="C")

            x += card_w + gap

        self.set_y(base_y + card_h + 6)

    # ── Tabela de Dados ────────────────────────────────────────────────────

    def tabela(self, df: pd.DataFrame, max_rows: int = 30) -> None:
        if df.empty:
            return

        colunas = list(df.columns)
        page_width = 277
        col_width = page_width / len(colunas)

        # Cabeçalho da tabela
        self.set_fill_color(*PRIMARY)
        self.set_text_color(*WHITE)
        self.set_font("Helvetica", "B", 8)
        self.set_draw_color(*WHITE)
        self.set_line_width(0.3)

        for col in colunas:
            self.cell(col_width, 8, _sanitize(str(col)[:18]), border=1, align="C", fill=True)
        self.ln()

        # Linhas
        self.set_font("Helvetica", "", 7.5)
        self.set_draw_color(210, 220, 235)

        for i, (_, row) in enumerate(df.head(max_rows).iterrows()):
            if i % 2 == 0:
                self.set_fill_color(*LIGHT_BG)
            else:
                self.set_fill_color(*WHITE)
            self.set_text_color(*DARK_TEXT)

            for val in row:
                texto = _sanitize(str(val)[:22])
                self.cell(col_width, 7, texto, border=1, align="C", fill=True)
            self.ln()

        if len(df) > max_rows:
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(*GRAY)
            self.cell(0, 6, _sanitize(f"  * Exibindo {max_rows} de {len(df)} registros."), align="L")
            self.ln(6)

    # ── Imagem de Gráfico ──────────────────────────────────────────────────

    def inserir_grafico(self, caminho_img: str, largura: int = 130) -> None:
        if not os.path.exists(caminho_img):
            return
        x = self.get_x()
        y = self.get_y()
        self.image(caminho_img, x=x, y=y, w=largura)
        self.ln(largura * 0.5 + 4)


# ─── Orquestrador Principal ───────────────────────────────────────────────────

def _calcular_kpis(df: pd.DataFrame) -> list[dict]:
    total_vendas = df["Total (R$)"].sum() if "Total (R$)" in df.columns else 0
    total_registros = len(df)
    ticket_medio = total_vendas / total_registros if total_registros else 0
    meta_atingida = (df["Meta Atingida"] == "Sim").sum() if "Meta Atingida" in df.columns else 0

    return [
        {"label": "Total em Vendas", "value": _formatar_brl(total_vendas), "color": PRIMARY},
        {"label": "Nr de Registros",  "value": str(total_registros),       "color": (41, 100, 160)},
        {"label": "Ticket Medio",     "value": _formatar_brl(ticket_medio), "color": SUCCESS},
        {"label": "Metas Atingidas",  "value": str(meta_atingida),          "color": (180, 120, 0)},
    ]


def gerar_relatorio(
    caminho_excel: str,
    caminho_saida: Optional[str] = None,
    aba: Optional[str] = None,
) -> str:
    """
    Lê `caminho_excel` e gera um relatório PDF em `caminho_saida`.
    Retorna o caminho do PDF gerado.
    """
    caminho_excel = Path(caminho_excel)
    if not caminho_excel.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_excel}")

    if caminho_saida is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        caminho_saida = caminho_excel.parent / f"relatorio_{timestamp}.pdf"

    print(f"[>>] Lendo planilha: {caminho_excel} ...")
    df = pd.read_excel(caminho_excel, sheet_name=aba or 0)
    df.columns = [str(c).strip() for c in df.columns]
    print(f"    {len(df)} linhas carregadas. Colunas: {list(df.columns)}")

    _ensure_chart_dir()

    # ── Gráficos ──────────────────────────────────────────────────────────
    graficos: dict[str, str] = {}

    if "Vendedor" in df.columns and "Total (R$)" in df.columns:
        graficos["vendedor"] = _gerar_grafico_barras(
            df, "Vendedor", "Total (R$)", "Vendas por Vendedor", "chart_vendedor.png"
        )
    if "Região" in df.columns and "Total (R$)" in df.columns:
        graficos["regiao"] = _gerar_grafico_pizza(
            df, "Região", "Total (R$)", "Distribuição por Região", "chart_regiao.png"
        )
    if "Produto" in df.columns and "Total (R$)" in df.columns:
        graficos["produto"] = _gerar_grafico_barras(
            df, "Produto", "Total (R$)", "Vendas por Produto", "chart_produto.png"
        )

    # ── Construção do PDF ─────────────────────────────────────────────────
    print("[>>] Gerando PDF ...")
    pdf = RelatorioPDF(
        titulo_relatorio="Relatório de Vendas",
        subtitulo=f"Fonte: {caminho_excel.name}",
    )

    # KPIs
    pdf.secao("Resumo Executivo")
    pdf.cards_kpi(_calcular_kpis(df))

    # Gráficos lado a lado (vendedor + região)
    if "vendedor" in graficos or "regiao" in graficos:
        pdf.secao("Análise Gráfica")
        y_graf = pdf.get_y()

        if "vendedor" in graficos:
            pdf.set_xy(10, y_graf)
            pdf.image(graficos["vendedor"], x=10, y=y_graf, w=145)

        if "regiao" in graficos:
            pdf.set_xy(160, y_graf)
            pdf.image(graficos["regiao"], x=160, y=y_graf, w=127)

        pdf.ln(75)

    if "produto" in graficos:
        pdf.secao("Vendas por Produto")
        pdf.inserir_grafico(graficos["produto"], largura=200)

    # Tabela de dados
    pdf.add_page()
    pdf.secao(f"Dados Detalhados ({len(df)} registros)")
    pdf.tabela(df, max_rows=40)

    pdf.output(str(caminho_saida))
    _cleanup_charts()

    print(f"[OK] Relatorio gerado com sucesso: '{caminho_saida}'")
    return str(caminho_saida)


# ─── CLI ─────────────────────────────────────────────────────────────────────

def _encontrar_excel() -> str:
    """Detecta automaticamente planilhas .xlsx na pasta atual."""
    arquivos = sorted(Path(".").glob("*.xlsx"))

    if not arquivos:
        print("[ERRO] Nenhum arquivo .xlsx encontrado na pasta atual.")
        print("       Coloque sua planilha Excel nesta pasta e tente novamente.")
        sys.exit(1)

    if len(arquivos) == 1:
        print(f"[>>] Planilha encontrada: {arquivos[0].name}")
        return str(arquivos[0])

    print("Multiplas planilhas encontradas. Escolha uma:")
    for i, arq in enumerate(arquivos, start=1):
        print(f"  [{i}] {arq.name}")

    while True:
        escolha = input("Digite o numero: ").strip()
        if escolha.isdigit() and 1 <= int(escolha) <= len(arquivos):
            return str(arquivos[int(escolha) - 1])
        print("  Opcao invalida, tente novamente.")


def _main() -> None:
    import argparse

    parser = argparse.ArgumentParser(
        description="Gera relatorio PDF a partir de uma planilha Excel."
    )
    parser.add_argument(
        "excel",
        nargs="?",
        default=None,
        help="(Opcional) Caminho do arquivo .xlsx. Se omitido, detecta automaticamente.",
    )
    parser.add_argument("-o", "--output", default=None, help="Caminho do PDF de saida")
    parser.add_argument("--aba", default=None, help="Nome da aba da planilha (padrao: primeira)")
    args = parser.parse_args()

    caminho_excel = args.excel if args.excel else _encontrar_excel()

    try:
        gerar_relatorio(caminho_excel, args.output, args.aba)
    except FileNotFoundError as exc:
        print(f"[ERRO] {exc}")
        sys.exit(1)


if __name__ == "__main__":
    _main()
