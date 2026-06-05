import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

try:
    import matplotlib
    matplotlib.use("TkAgg")
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.backends.backend_pdf import PdfPages
except ImportError:
    plt = None
    FigureCanvasTkAgg = None
    PdfPages = None

try:
    from openpyxl import Workbook
except ImportError:
    Workbook = None

DATA_FILE = Path(__file__).with_name("vendas.csv")


def carregar_dados(path):
    if not path.exists():
        raise FileNotFoundError(f"Arquivo de dados não encontrado: {path}")

    df = pd.read_csv(path, parse_dates=["data"], encoding="utf-8")
    df.rename(columns={"data": "data_venda"}, inplace=True)
    df["receita"] = df["quantidade"] * df["preco_unitario"]
    df["periodo"] = df["data_venda"].dt.to_period("M").astype(str)
    df["data_venda"] = df["data_venda"].dt.strftime("%Y-%m-%d")
    return df


def calcular_métricas(df):
    total_registros = len(df)
    receita_total = df["receita"].sum()
    vendas_eletronicos = df[df["categoria"] == "Eletrônicos"].copy()

    receita_categoria = df.groupby("categoria")["receita"].sum().sort_values(ascending=False)
    top_produtos_quantidade = df.groupby("produto")["quantidade"].sum().sort_values(ascending=False).head(10)
    receita_regiao = df.groupby("regiao")["receita"].sum().sort_values(ascending=False)
    receita_mensal = df.groupby("periodo")["receita"].sum()
    pivot_regiao_categoria = pd.pivot_table(
        df,
        values="receita",
        index="regiao",
        columns="categoria",
        aggfunc="sum",
        fill_value=0,
    )

    produto_top = top_produtos_quantidade.index[0] if not top_produtos_quantidade.empty else None
    quantidade_top = int(top_produtos_quantidade.iloc[0]) if not top_produtos_quantidade.empty else 0
    regiao_top = receita_regiao.index[0] if not receita_regiao.empty else None
    valor_regiao_top = float(receita_regiao.iloc[0]) if not receita_regiao.empty else 0.0

    return {
        "total_registros": total_registros,
        "receita_total": receita_total,
        "vendas_eletronicos": vendas_eletronicos,
        "produto_top": produto_top,
        "quantidade_top": quantidade_top,
        "regiao_top": regiao_top,
        "valor_regiao_top": valor_regiao_top,
        "receita_categoria": receita_categoria,
        "top_produtos_quantidade": top_produtos_quantidade,
        "receita_regiao": receita_regiao,
        "receita_mensal": receita_mensal,
        "pivot_regiao_categoria": pivot_regiao_categoria,
    }


def criar_treeview(frame, colunas, largura_coluna, dados):
    tree = ttk.Treeview(frame, columns=colunas, show="headings", height=6)
    for coluna, largura in zip(colunas, largura_coluna):
        tree.heading(coluna, text=coluna)
        tree.column(coluna, width=largura, anchor=tk.W)
    for item in dados:
        tree.insert("", tk.END, values=item)
    scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    tree.grid(row=0, column=0, sticky="nsew")
    scrollbar.grid(row=0, column=1, sticky="ns")
    frame.grid_rowconfigure(0, weight=1)
    frame.grid_columnconfigure(0, weight=1)
    return tree


def criar_grafico(frame, dados, titulo, xlabel, ylabel, rotacao=0, tipo="bar"):
    if plt is None:
        label = ttk.Label(frame, text="Matplotlib não está disponível.")
        label.pack(fill=tk.BOTH, expand=True)
        return

    figura, ax = plt.subplots(figsize=(5, 4), dpi=100)
    chaves = list(dados.keys())
    valores = list(dados.values())
    if tipo == "line":
        ax.plot(chaves, valores, marker="o", color="#4c72b0")
    else:
        ax.bar(chaves, valores, color="#4c72b0")
    ax.set_title(titulo)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.tick_params(axis="x", rotation=rotacao)
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    figura.tight_layout()

    canvas = FigureCanvasTkAgg(figura, master=frame)
    canvas.draw()
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(fill=tk.BOTH, expand=True)


def montar_pivot_rows(pivot):
    categorias = list(pivot.columns)
    colunas = ["Região"] + categorias
    rows = []
    ordenado = pivot.sort_values(by=list(pivot.columns), ascending=False)
    for regiao in ordenado.index:
        linha = [regiao] + [f"R$ {pivot.loc[regiao, categoria]:,.2f}" for categoria in categorias]
        rows.append(tuple(linha))
    return colunas, rows


def exportar_para_xlsx(metricas, caminho):
    if Workbook is None:
        raise ImportError("openpyxl não está instalado.")

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        pd.DataFrame(
            [
                ["Total de registros", metricas["total_registros"]],
                ["Receita total", f"R$ {metricas['receita_total']:,.2f}"],
                ["Produto mais vendido", f"{metricas['produto_top']} ({metricas['quantidade_top']} unidades)"],
                ["Região com maior valor", f"{metricas['regiao_top']} (R$ {metricas['valor_regiao_top']:,.2f})"],
            ],
            columns=["Métrica", "Valor"],
        ).to_excel(writer, index=False, sheet_name="Resumo")

        metricas["pivot_regiao_categoria"].to_excel(writer, sheet_name="Pivot Região x Categoria")
        metricas["receita_categoria"].rename("Receita").to_frame().to_excel(writer, sheet_name="Receita por Categoria")
        metricas["receita_mensal"].rename("Receita").to_frame().to_excel(writer, sheet_name="Receita Mensal")

    workbook.save(caminho)


def exportar_para_pdf(metricas, caminho):
    if plt is None or PdfPages is None:
        raise ImportError("matplotlib não está disponível para gerar PDF.")

    with PdfPages(caminho) as pdf:
        figura = plt.figure(figsize=(8.27, 11.69), dpi=100)
        figura.clf()
        texto = (
            f"Resumo de Vendas\n\n"
            f"Total de registros: {metricas['total_registros']}\n"
            f"Receita total: R$ {metricas['receita_total']:,.2f}\n"
            f"Produto mais vendido: {metricas['produto_top']} ({metricas['quantidade_top']} unidades)\n"
            f"Região com maior valor: {metricas['regiao_top']} (R$ {metricas['valor_regiao_top']:,.2f})\n"
        )
        figura.text(0.1, 0.95, texto, fontsize=12, va="top")
        pdf.savefig(figura)

        def salvar_figura(dados, titulo, xlabel, ylabel, tipo="bar"):
            fig, ax = plt.subplots(figsize=(8, 4.5), dpi=100)
            if hasattr(dados, "index"):
                chaves = list(dados.index)
                valores = list(dados.values)
            else:
                chaves = list(dados.keys())
                valores = list(dados.values())
            if tipo == "line":
                ax.plot(chaves, valores, marker="o", color="#4c72b0")
            else:
                ax.bar(chaves, valores, color="#4c72b0")
            ax.set_title(titulo)
            ax.set_xlabel(xlabel)
            ax.set_ylabel(ylabel)
            ax.tick_params(axis="x", rotation=45)
            ax.grid(axis="y", linestyle="--", alpha=0.4)
            fig.tight_layout()
            pdf.savefig(fig)
            plt.close(fig)

        salvar_figura(metricas["receita_categoria"], "Receita por categoria", "Categoria", "Receita (R$)")
        salvar_figura(metricas["receita_mensal"], "Evolução mensal de receita", "Período", "Receita (R$)", tipo="line")

        pivot = metricas["pivot_regiao_categoria"]
        categorias = list(pivot.columns)
        linhas = [tuple(["Região"] + categorias)]
        for regiao in pivot.index:
            linhas.append(tuple([regiao] + [pivot.loc[regiao, categoria] for categoria in categorias]))
        fig, ax = plt.subplots(figsize=(8.27, 6), dpi=100)
        ax.axis("tight")
        ax.axis("off")
        tabela = ax.table(cellText=linhas, loc="center", cellLoc="center")
        tabela.auto_set_font_size(False)
        tabela.set_fontsize(8)
        tabela.scale(1, 1.5)
        pdf.savefig(fig)
        plt.close(fig)


def salvar_relatorio_xlsx(root, metricas):
    if Workbook is None:
        messagebox.showwarning(
            "Dependência ausente",
            "Instale openpyxl para exportar para XLSX: python -m pip install openpyxl",
        )
        return
    caminho = filedialog.asksaveasfilename(
        parent=root,
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")],
    )
    if not caminho:
        return
    try:
        exportar_para_xlsx(metricas, caminho)
        messagebox.showinfo("Exportação concluída", f"Relatório XLSX salvo em:\n{caminho}")
    except Exception as exc:
        messagebox.showerror("Erro ao exportar", str(exc))


def salvar_relatorio_pdf(root, metricas):
    if plt is None:
        messagebox.showwarning(
            "Dependência ausente",
            "Instale matplotlib para exportar para PDF.",
        )
        return
    caminho = filedialog.asksaveasfilename(
        parent=root,
        defaultextension=".pdf",
        filetypes=[("PDF", "*.pdf")],
    )
    if not caminho:
        return
    try:
        exportar_para_pdf(metricas, caminho)
        messagebox.showinfo("Exportação concluída", f"Relatório PDF salvo em:\n{caminho}")
    except Exception as exc:
        messagebox.showerror("Erro ao exportar", str(exc))


def criar_interface(registros, métricas):
    root = tk.Tk()
    root.title("Análise de Vendas - TechStore")
    root.geometry("1200x840")

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True)

    resumo_frame = ttk.Frame(notebook, padding=12)
    dados_frame = ttk.Frame(notebook, padding=12)
    graficos_frame = ttk.Frame(notebook, padding=12)
    relatorio_frame = ttk.Frame(notebook, padding=12)

    notebook.add(resumo_frame, text="Resumo")
    notebook.add(dados_frame, text="Dados")
    notebook.add(graficos_frame, text="Gráficos")
    notebook.add(relatorio_frame, text="Relatório")

    # Resumo
    metricas_frame = ttk.Frame(resumo_frame)
    metricas_frame.pack(fill=tk.X, pady=4)

    labels = [
        ("Total de registros:", métricas["total_registros"]),
        ("Receita total:", f"R$ {métricas['receita_total']:,.2f}"),
        ("Produto mais vendido:", f"{métricas['produto_top']} ({métricas['quantidade_top']} unidades)"),
        ("Região com maior valor:", f"{métricas['regiao_top']} (R$ {métricas['valor_regiao_top']:,.2f})"),
    ]

    for idx, (texto, valor) in enumerate(labels):
        label = ttk.Label(metricas_frame, text=f"{texto} {valor}", font=(None, 12, "bold"))
        label.grid(row=idx, column=0, sticky="w", pady=2)

    eletronicos_frame = ttk.LabelFrame(resumo_frame, text="Vendas Eletrônicos")
    eletronicos_frame.pack(fill=tk.BOTH, expand=True, pady=8)

    colunas_elec = ["data_venda", "produto", "quantidade", "preco_unitario", "receita", "regiao"]
    dados_eletronicos = [
        (
            row["data_venda"],
            row["produto"],
            int(row["quantidade"]),
            f"R$ {row['preco_unitario']:,.2f}",
            f"R$ {row['receita']:,.2f}",
            row["regiao"],
        )
        for _, row in métricas["vendas_eletronicos"].iterrows()
    ]
    criar_treeview(eletronicos_frame, colunas_elec, [120, 180, 90, 120, 120, 120], dados_eletronicos)

    # Dados
    primeiros_frame = ttk.LabelFrame(dados_frame, text="Primeiras 5 linhas")
    primeiros_frame.pack(fill=tk.BOTH, expand=False, pady=8)

    colunas = ["data_venda", "produto", "categoria", "quantidade", "preco_unitario", "receita", "regiao"]
    dados_principais = [
        (
            row["data_venda"],
            row["produto"],
            row["categoria"],
            int(row["quantidade"]),
            f"R$ {row['preco_unitario']:,.2f}",
            f"R$ {row['receita']:,.2f}",
            row["regiao"],
        )
        for _, row in registros.head(5).iterrows()
    ]
    criar_treeview(primeiros_frame, colunas, [100, 170, 120, 90, 120, 120, 120], dados_principais)

    # Relatório
    export_frame = ttk.Frame(relatorio_frame)
    export_frame.pack(fill=tk.X, pady=4)

    xlsx_button = ttk.Button(
        export_frame,
        text="Exportar para XLSX",
        command=lambda: salvar_relatorio_xlsx(root, métricas),
    )
    pdf_button = ttk.Button(
        export_frame,
        text="Exportar para PDF",
        command=lambda: salvar_relatorio_pdf(root, métricas),
    )
    xlsx_button.pack(side=tk.LEFT, padx=4)
    pdf_button.pack(side=tk.LEFT, padx=4)

    pivot_frame = ttk.LabelFrame(relatorio_frame, text="Pivot Região × Categoria")
    pivot_frame.pack(fill=tk.BOTH, expand=True, pady=8)
    colunas_pivot, linhas_pivot = montar_pivot_rows(métricas["pivot_regiao_categoria"])
    criar_treeview(pivot_frame, colunas_pivot, [140] + [130] * (len(colunas_pivot) - 1), linhas_pivot)

    # Gráficos
    graficos_grid = ttk.Frame(graficos_frame)
    graficos_grid.pack(fill=tk.BOTH, expand=True)
    graficos_grid.columnconfigure(0, weight=1)
    graficos_grid.columnconfigure(1, weight=1)
    graficos_grid.rowconfigure(0, weight=1)
    graficos_grid.rowconfigure(1, weight=1)

    chart1_frame = ttk.LabelFrame(graficos_grid, text="Receita por categoria")
    chart2_frame = ttk.LabelFrame(graficos_grid, text="Top 10 produtos por quantidade")
    chart3_frame = ttk.LabelFrame(graficos_grid, text="Receita por região")
    chart4_frame = ttk.LabelFrame(graficos_grid, text="Receita mensal")

    chart1_frame.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
    chart2_frame.grid(row=0, column=1, sticky="nsew", padx=6, pady=6)
    chart3_frame.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
    chart4_frame.grid(row=1, column=1, sticky="nsew", padx=6, pady=6)

    criar_grafico(chart1_frame, métricas["receita_categoria"], "Receita por categoria", "Categoria", "Receita (R$)", rotacao=30)
    criar_grafico(chart2_frame, métricas["top_produtos_quantidade"], "Top 10 produtos mais vendidos", "Produto", "Quantidade", rotacao=45)
    criar_grafico(chart3_frame, métricas["receita_regiao"], "Receita por região", "Região", "Receita (R$)")
    criar_grafico(chart4_frame, métricas["receita_mensal"], "Evolução mensal de receita", "Período", "Receita (R$)", rotacao=45, tipo="line")

    root.mainloop()


def main():
    try:
        registros = carregar_dados(DATA_FILE)
    except FileNotFoundError as exc:
        messagebox.showerror("Erro", str(exc))
        return

    métrica = calcular_métricas(registros)
    criar_interface(registros, métrica)


if __name__ == "__main__":
    main()
