# bot-mensagem-automatica

O projeto descrito é um bot de automação para envio de mensagens via WhatsApp, utilizando uma interface gráfica feita com a biblioteca CustomTkinter. Aqui está um resumo das principais funcionalidades e componentes do projeto:

Componentes e Funcionalidades
Bibliotecas Utilizadas:

openpyxl: Para manipulação de planilhas Excel.
os: Para operações com o sistema de arquivos.
urllib.parse.quote: Para codificação de URLs.
webbrowser: Para abrir o navegador web.
time.sleep: Para pausas durante a execução.
pyautogui: Para automação de interações com a interface gráfica do usuário.
tkinter e customtkinter: Para a criação da interface gráfica.
Interface Gráfica:

A interface gráfica permite ao usuário selecionar uma planilha Excel contendo os contatos e iniciar o processo de envio de mensagens.
A interface contém botões para selecionar o arquivo, iniciar o envio de mensagens e abrir um arquivo de texto para edição.
Funcionamento:

Seleção de Arquivo: O usuário seleciona uma planilha Excel que contém os nomes, telefones e datas de vencimento dos clientes.
Leitura de Arquivo: O conteúdo de um arquivo de texto é lido e armazenado para ser usado como parte da mensagem enviada.
Automação do WhatsApp: O bot abre o WhatsApp Web, envia mensagens personalizadas para cada contato listado na planilha e fecha a aba após cada envio.
Detalhes da Implementação:

Arquivo Excel: A planilha deve conter os nomes, telefones e datas de vencimento dos clientes. O bot lê essa informação e gera uma mensagem personalizada para cada cliente.
Mensagens Personalizadas: A mensagem enviada inclui o nome do cliente e a data de vencimento formatada.
Tratamento de Erros: Se ocorrer um erro durante o envio de uma mensagem, uma mensagem de erro é exibida.
Aparência e Tema:

A interface permite alternar entre os modos claro e escuro usando um switch.
Fluxo de Trabalho do Bot
Iniciar Interface: O usuário abre a interface gráfica.
Selecionar Planilha: O usuário seleciona a planilha Excel com os dados dos clientes.
Ler Arquivo de Texto: O bot lê o conteúdo de um arquivo de texto, que será incluído na mensagem enviada.
Abrir WhatsApp Web: O bot abre o WhatsApp Web no navegador.
Enviar Mensagens: Para cada linha da planilha Excel, o bot gera e envia uma mensagem personalizada.
Feedback ao Usuário: Se uma mensagem não puder ser enviada, o usuário é notificado com uma mensagem de erro.
Interface Gráfica
CustomTkinter: Utilizada para criar uma interface moderna e personalizável.
Elementos da Interface:
Label para instruções.
Entry para mostrar o caminho do arquivo selecionado.
Botões para selecionar o arquivo, iniciar o envio de mensagens e abrir o arquivo de texto.
Switch para alternar entre temas claro e escuro.
Código Resumido
Abaixo está um trecho simplificado do código principal:

python
Copiar código
import openpyxl
import os
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import tkinter as tk
from tkinter import messagebox, filedialog
import customtkinter as ctk

def start():
    # Leitura do arquivo de texto e abertura do WhatsApp Web
    with open(caminho_arquivo, 'r') as arquivo:
        mensagem2 = arquivo.read()
    webbrowser.open('https://web.whatsapp.com/')
    sleep(10)
    pyautogui.hotkey('ctrl', 'w')
    sleep(5)

    # Leitura da planilha Excel
    workbook = openpyxl.load_workbook(filename)
    pagina_clientes = workbook['Planilha1']

    for linha in pagina_clientes.iter_rows(min_row=2):
        nome = linha[0].value
        telefone = linha[1].value
        vencimento = linha[2].value
        mensagem = f'Ola {nome}, aqui é da empresa XXXX estou entrando em contato para o envio do boleto no vencimento {vencimento.strftime("%d/%m/%Y")}.\n\n{mensagem2}'
        try:
            link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
            webbrowser.open(link_mensagem_whatsapp)
            sleep(7)
            pyautogui.hotkey('enter')
            sleep(5)
            pyautogui.hotkey('ctrl', 'w')
            sleep(5)
        except:
            tk.messagebox.showinfo("erro", f"Não foi possivel enviar mensagem para {nome}")

# Interface gráfica com CustomTkinter
app = ctk.CTk()
ctk.set_appearance_mode("dark")
app.title("WhatsApp Bot")
app.geometry("400x250+500+300")
app.resizable(0, 0)

def iniciar():
    global filename
    if entry.get() and os.path.exists(filename):
        start()
    else:
        tk.messagebox.showinfo("erro", "Arquivo não localizado")

def selecionar_arquivo():
    global filename
    filename = filedialog.askopenfilename(filetypes=[("Planilhas Excel", ".xlsx;.xls")])
    entry_arquivo.set(filename)

def txt():
    global caminho_arquivo
    caminho_arquivo = 'Arquivo.txt'
    os.startfile(caminho_arquivo)

# Configuração da interface
frame = ctk.CTkFrame(app, fg_color="black")
frame.pack(padx=15, pady=15, side="top", fill="both", expand=True)
label = ctk.CTkLabel(frame, text="Selecione a planilha com os telefones")
label.pack()
entry_arquivo = tk.StringVar()
entry = ctk.CTkEntry(frame, textvariable=entry_arquivo, width=250)
entry.pack(pady=5)
botao = ctk.CTkButton(frame, text="Selecionar Arquivo", command=selecionar_arquivo)
botao.pack(pady=5)
btn = ctk.CTkButton(frame, text="Iniciar", command=iniciar)
btn.pack()
btn2 = ctk.CTkButton(frame, text="txt para envio", command=txt)
btn2.pack(pady=5)
swtVar = tk.IntVar()
switch = ctk.CTkSwitch(frame, text="Theme", command=lambda: ctk.set_appearance_mode("light" if swtVar.get() else "dark"), variable=swtVar)
switch.pack(pady=10)

app.mainloop()
Conclusão
Este projeto é uma solução prática para automatizar o envio de mensagens via WhatsApp, economizando tempo e esforço, especialmente em situações onde mensagens semelhantes precisam ser enviadas a múltiplos destinatários. A interface gráfica torna o uso do bot acessível para usuários com pouca experiência em programação.






