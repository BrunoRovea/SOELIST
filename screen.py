import tkinter as tk
from tkinter import filedialog
import pandas as pd
import glob
import os


caminho_arquivo_selecionado = ""  # Variável global para armazenar o caminho do arquivo

def obter_caminho():
    global caminho_arquivo_selecionado
    caminho_arquivo_selecionado = filedialog.askopenfilename(filetypes=[('Arquivos Excel', '*.xlsx')])
    campo_caminho.delete(0, tk.END)
    campo_caminho.insert(tk.END, caminho_arquivo_selecionado)
    caminho_arquivo_selecionado = os.path.dirname(caminho_arquivo_selecionado)
#    janela.destroy()
#    janela.quit()
    
def executar():
    caminho_arquivo = caminho_arquivo_selecionado
    caminho_arquivo = os.path.normpath(caminho_arquivo)  # Normalize the path
    for file in glob.glob(caminho_arquivo + '/*.xlsx'):
        cenario = pd.read_excel(file)  # realiza a leitura dos arquivos .xlsx
        cenario_aux = cenario.copy()
        cenario_aux['Element'] = cenario_aux['Element'].astype(str)
        cenario_aux['msg'] = cenario_aux['msg'].astype(str)
        cenario_aux.loc[cenario_aux['Info'] != 'MvMoment', 'mensagem'] = 'TA,' + cenario_aux['B1'] + '   ,' + cenario_aux['B2'] + '   ,' + cenario_aux['B3'] + '   ,' + cenario_aux['Element'] + '   ,' + cenario_aux['Info'] + '   ,' + '\nsta ' + cenario_aux['msg'] + ' tra'
        cenario_aux.loc[cenario_aux['Info'] == 'MvMoment', 'mensagem'] = 'TA,' + cenario_aux['B1'] + '   ,' + cenario_aux['B2'] + '   ,' + cenario_aux['B3'] + '   ,' + cenario_aux['Element'] + '   ,' + cenario_aux['Info'] + '   ,' + '\nval ' + cenario_aux['msg'] + ' tra'
        output_filename = os.path.splitext(file)[0] + '.txt'
        print(file)
        print(output_filename)
        cenario_aux.to_csv(output_filename, sep='\t', quotechar=' ', index=False, header=False)

janela = tk.Tk()
janela.title("Selecionar Arquivo .xlsx")
janela.geometry("500x420")

#base_frame = tk.Frame(janela)
#base_frame.pack()

# Campo para exibir o caminho do arquivo
campo_caminho = tk.Entry(janela, width=50)
campo_caminho.pack(pady=10)

# Botão para obter o caminho do arquivo
botao_obter = tk.Button(janela, text="Selecionar Arquivo", command=obter_caminho)
botao_obter.pack(pady=5)

# Botão para executar o código
botao_executar = tk.Button(janela, text="Executar", command=executar)
botao_executar.pack(pady=5)

# Campo para descrever os passos do código
campo_passos = tk.Text(janela, height=10, width=50)
campo_passos.pack(pady=10)

janela.mainloop()

 
 
