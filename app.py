import openpyxl
from urllib.parse import quote
from time import sleep
import pyautogui
import os
import webbrowser
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox  # Adicionado para usar messagebox


def enviar_mensagens(arquivo_excel, mensagem_padrao):
    webbrowser.open('https://web.whatsapp.com/')
    sleep(15)

    workbook = openpyxl.load_workbook(arquivo_excel)
    pagina_clientes = workbook['Sheet1']

    contador_mensagens_enviadas = 0

    for linha in pagina_clientes.iter_rows(min_row=2):
        if contador_mensagens_enviadas >= 70:  # Verifica se atingiu o limite de 70 mensagens
            break

        if all(cell.value is not None for cell in linha):
            nome = linha[0].value
            telefone = int(linha[1].value)
            vencimento = linha[2].value

            print(telefone)

            mensagem_personalizada = f'Olá {nome} seu boleto vence no dia {vencimento}. Favor pagar no link https://www.link_do_pagamento.com\n{mensagem_padrao}'

            try:
                link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem_personalizada)}'
                print(link_mensagem_whatsapp)
                webbrowser.open(link_mensagem_whatsapp)
                sleep(15)
                pyautogui.moveTo(900, 965)
                sleep(5)
                pyautogui.click()
                sleep(5)
                pyautogui.press('enter')
                sleep(2)
                pyautogui.hotkey('ctrl', 'w')
                contador_mensagens_enviadas += 1 # Contador de mensagens enviadas
                sleep(40)
            except:
                print(f'Não foi possível enviar mensagem para {nome}')
                with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                    arquivo.write(f'{nome},{telefone}{os.linesep}')
            else:
                continue

    messagebox.showinfo("Disparo Finalizado", "Disparo finalizado com sucesso!")

def comecar_disparo():
    if arquivo_excel:
        enviar_mensagens(arquivo_excel, entrada_mensagem.get())
    else:
        messagebox.showwarning("Aviso", "Selecione um arquivo Excel antes de começar o disparo.")

def abrir_arquivo():
    global arquivo_excel
    arquivo_excel = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if arquivo_excel:
        lbl_nome_arquivo.config(text=f"Arquivo selecionado: {os.path.basename(arquivo_excel)}")
        btn_comecar_disparo.config(state=tk.NORMAL)
        entrada_mensagem.config(state=tk.NORMAL)

def remover_arquivo():
    global arquivo_excel
    arquivo_excel = None
    lbl_nome_arquivo.config(text="Nenhum arquivo selecionado")
    btn_comecar_disparo.config(state=tk.DISABLED)
    entrada_mensagem.config(state=tk.DISABLED)

root = tk.Tk()
root.title("Envio de Mensagens")
root.geometry("800x500")

label_arquivo = tk.Label(root, text="Selecione o arquivo Excel:", font=("Arial", 16))
label_arquivo.pack(pady=10)

btn_abrir_arquivo = tk.Button(root, text="Abrir Arquivo", command=abrir_arquivo, width=15, height=1, font=("Arial", 16, "bold"), background="#79b6c9", foreground="white")
btn_abrir_arquivo.pack()

lbl_nome_arquivo = tk.Label(root, text="Nenhum arquivo selecionado", font=("Arial", 14))
lbl_nome_arquivo.pack(pady=10, padx=40)
lbl_nome_arquivo.place(x=90, y=130)

btn_remover_arquivo = tk.Button(root, text="Remover Arquivo", command=remover_arquivo, width=15, height=1, font=("Arial", 12, "bold"), background="#ff6347", foreground="white")
btn_remover_arquivo.place(x=600, y=120)

label_mensagem = tk.Label(root, text="Digite a mensagem padrão:", font=("Arial", 16))
label_mensagem.pack(pady=60)


entrada_mensagem = tk.Entry(root, width=30, font=("Arial", 16), background="#fff")
entrada_mensagem.pack()

btn_comecar_disparo = tk.Button(root, text="Começar Disparo", command=comecar_disparo, state=tk.DISABLED, font=("Arial", 16, "bold"), background="#90ee90", foreground="#000")
btn_comecar_disparo.pack(pady=10)

root.mainloop()