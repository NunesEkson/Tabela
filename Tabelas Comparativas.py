import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog
import matplotlib.pyplot as plt

def exibir_planilhas():
    # Abre a caixa de diálogo para selecionar o arquivo 1
    arquivo1 = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")], initialdir="/", parent=root)

    # Verifica se o arquivo 1 foi selecionado
    if arquivo1:
        # Abre a caixa de diálogo para selecionar o arquivo 2
        arquivo2 = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")], initialdir="/", parent=root)

        # Verifica se o arquivo 2 foi selecionado
        if arquivo2:
            # Lê as planilhas Excel
            df1 = pd.read_excel(arquivo1)
            df2 = pd.read_excel(arquivo2)

            # Cria a janela das planilhas
            janela_planilhas = tk.Toplevel(root)
            janela_planilhas.title("Planilhas")

            # Cria as tabelas usando Treeview do tkinter
            tabela1 = ttk.Treeview(janela_planilhas)
            tabela2 = ttk.Treeview(janela_planilhas)

            # Define as colunas
            colunas1 = list(df1.columns)
            colunas2 = list(df2.columns)

            tabela1["columns"] = colunas1
            tabela2["columns"] = colunas2

            # Formata as colunas
            for col in colunas1:
                tabela1.column(col, width=100)
                tabela1.heading(col, text=col)

            for col in colunas2:
                tabela2.column(col, width=100)
                tabela2.heading(col, text=col)

            # Insere os dados nas tabelas
            for i, row in df1.iterrows():
                tabela1.insert("", "end", text=i, values=list(row))

            for i, row in df2.iterrows():
                tabela2.insert("", "end", text=i, values=list(row))

            # Posiciona os elementos na janela
            tabela1.pack(side=tk.LEFT, padx=10, pady=10)
            tabela2.pack(side=tk.RIGHT, padx=10, pady=10)

            # Cria o botão de gerar gráfico
            btn_gerar_grafico = ttk.Button(janela_planilhas, text="Gerar Gráfico", command=lambda: exibir_selecao_colunas(df1, df2, tabela1, tabela2))
            btn_gerar_grafico.pack(pady=10)

def exibir_selecao_colunas(df1, df2, tabela1, tabela2):
    # Cria a janela de seleção de colunas
    janela_colunas = tk.Toplevel(root)
    janela_colunas.title("Selecionar Colunas")

    # Obtém as colunas disponíveis para cada tabela
    colunas1 = tabela1["columns"]
    colunas2 = tabela2["columns"]

    # Cria os widgets de seleção de colunas
    colunas_var1 = [tk.StringVar() for _ in colunas1]
    colunas_var2 = [tk.StringVar() for _ in colunas2]

    for i, col in enumerate(colunas1):
        checkbox = ttk.Checkbutton(janela_colunas, text=col, variable=colunas_var1[i])
        checkbox.grid(row=i, column=0, sticky="w")

    for i, col in enumerate(colunas2):
        checkbox = ttk.Checkbutton(janela_colunas, text=col, variable=colunas_var2[i])
        checkbox.grid(row=i, column=1, sticky="w")

    # Cria o botão de confirmar seleção de colunas
    btn_confirmar_colunas = ttk.Button(janela_colunas, text="Confirmar", command=lambda: gerar_grafico(df1, df2, colunas1, colunas2, colunas_var1, colunas_var2))
    btn_confirmar_colunas.grid(row=max(len(colunas1), len(colunas2)), columnspan=2, pady=10)

def gerar_grafico(df1, df2, colunas1, colunas2, colunas_var1, colunas_var2):
    # Obtém as colunas selecionadas para cada tabela
    colunas_selecionadas1 = [colunas1[i] for i, col in enumerate(colunas_var1) if col.get()]
    colunas_selecionadas2 = [colunas2[i] for i, col in enumerate(colunas_var2) if col.get()]

    # Cria os dataframes com as colunas selecionadas
    df_selecionado1 = df1[colunas_selecionadas1] if colunas_selecionadas1 else None
    df_selecionado2 = df2[colunas_selecionadas2] if colunas_selecionadas2 else None

    # Verifica se há colunas selecionadas para gerar o gráfico
    if df_selecionado1 is not None or df_selecionado2 is not None:
        # Cria a figura para exibir o gráfico
        fig = plt.figure(figsize=(8, 6))

        # Gera o gráfico de barras para cada dataframe selecionado
        if df_selecionado1 is not None:
            ax1 = fig.add_subplot(2, 1, 1)
            for col in df_selecionado1.columns:
                valores = df_selecionado1[col].value_counts()
                ax1.bar(valores.index.astype(str), valores.values, label=col)
            ax1.set_title("Gráfico de Colunas - Tabela 1")
            ax1.set_xlabel("Valores")
            ax1.set_ylabel("Contagem")
            ax1.legend()

        if df_selecionado2 is not None:
            ax2 = fig.add_subplot(2, 1, 2)
            for col in df_selecionado2.columns:
                valores = df_selecionado2[col].value_counts()
                ax2.bar(valores.index.astype(str), valores.values, label=col)
            ax2.set_title("Gráfico de Colunas - Tabela 2")
            ax2.set_xlabel("Valores")
            ax2.set_ylabel("Contagem")
            ax2.legend()

        # Ajusta os espaçamentos e exibe o gráfico
        plt.tight_layout()
        plt.xticks(rotation=45)  # Inclinação dos rótulos do eixo x
        plt.show()

# Cria a janela principal
root = tk.Tk()
root.title("Exibir Planilhas")

# Centraliza a janela principal
largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()
largura_janela = 300
altura_janela = 200
pos_x = (largura_tela - largura_janela) // 2
pos_y = (altura_tela - altura_janela) // 2
root.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")

# Cria o botão para selecionar os arquivos
botao_arquivos = ttk.Button(root, text="Selecionar Arquivos", command=exibir_planilhas)
botao_arquivos.pack(pady=70)

# Inicia o loop da interface gráfica
root.mainloop()