import customtkinter as ctk
import tkinter as Tk
from tkinter import messagebox
from random import randint
import win32com.client as win32


janela = Tk.Tk()
janela.title('Gerador de senhas')
janela.configure(background='#bce6fe')
janela.geometry('550x250')
janela.iconbitmap('cadeado.ico')
x = ''


def gerarsenha(tam='', cara='', email=''):

    global x

    try:
        if tam.isnumeric():
            contador = int(tam)
            if tam == 0 or int(tam) < 0:
                messagebox.showwarning('Erro', 'Tamanho da senha definida como 8')
                contador = 8

        else:
            messagebox.showwarning('Erro', 'Valor não reconhecido, 8 definido como tamanho padrão')
            contador = 8
    except:
        messagebox.showwarning('ERRO', 'Algo de errado aconteceu')


    lista = cara.strip()

    if lista == '' or ' ':
        lista = 'N'

    mail = email.strip()

    if mail == "":
        messagebox.showwarning('Falha ao enviar', 'Email não reconhecido, digite um válido')

    if lista[0] == 'N' or 'n':
        try:

            lista_simples = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
                                 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm',
                                 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z',
                                 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
                                 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']


            for i in range(0, contador):
                carc = randint(0, len(lista_simples))
                x += lista_simples[carc-1]

                outlook = win32.Dispatch('outlook.application')


            email = outlook.CreateItem(0)


            email.To = mail
            email.Subject = "Gerador de Senhas"
            email.HTMLBody = f"""
            <p>Olá, sua senha é: {x}</p>


            <p>Abs,</p>
            <p>L.Felipe</p>
            """


            email.Send()
            resp.configure(text="Senha enviada")
            return

        except:
            resp.configure(text='Algo deu errado tente novamente')



    if lista[0] == 'Ss':
        try:
            lista_composta = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
                             'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm',
                             'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z',
                             'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
                             'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
                              '@', '#', '&', '*', '%']

            for i in range(0, contador):
                carc = randint(0, len(lista_composta))
                x += lista_composta[carc-1]

                outlook = win32.Dispatch('outlook.application')


            email = outlook.CreateItem(0)


            email.To = mail
            email.Subject = "Gerador de Senhas"
            email.HTMLBody = f"""
            <p>Olá, sua senha é: {x}</p>


            <p>Abs,</p>
            <p>Gerador de Senhas</p>
            """


            email.Send()

            resp.configure(text="Senha enviada")
            return

        except:
            resp.configure(text='Algo deu errado tente novamente')

    x = ''


texto1 = ctk.CTkLabel(janela, text='Total de caractéres: ',font=('Bodoni', 12, 'bold'),text_color='white',
                      fg_color='black', bg_color='#bce6fe',corner_radius=1000, anchor='center', width=100)
texto1.grid(row=0, column=0, columnspan=1, padx=7, pady=12,sticky='nwse')

etexto1 = ctk.CTkEntry(janela, corner_radius=100, bg_color='#bce6fe',width=50)
etexto1.grid(row=0, column=1, columnspan=1, padx=7, pady=12)


texto2 = ctk.CTkLabel(janela, text='Contém caractéres especiais? ',font=('Bodoni', 12, 'bold'),text_color='white',
                      fg_color='black', bg_color='#bce6fe',corner_radius=1000, anchor='center')
texto2.grid(row=0, column=2, columnspan=1, padx=7, pady=12)

etexto2 = ctk.CTkEntry(janela, corner_radius=100, bg_color='#bce6fe',width=50)
etexto2.grid(row=0, column=4, columnspan=1, padx=7, pady=12)


email = ctk.CTkLabel(janela, text='Email: ', font=('Bodoni', 12, 'bold'),text_color='white',
                      fg_color='black', bg_color='#bce6fe',corner_radius=1000, anchor='center')
email.grid(row=2, column=0, columnspan=1, padx=12, pady=12, sticky='nwse')

eemail = ctk.CTkEntry(janela, corner_radius=100, bg_color='#bce6fe',width=300)
eemail.grid(row=2, column=1, columnspan=2, padx=12, pady=12)


button = ctk.CTkButton(janela, text='Gerar e enviar', bg_color='#bce6fe', fg_color='red', hover_color='green',
                       command=lambda: gerarsenha(etexto1.get(), etexto2.get(), eemail.get()))
button.grid(row=3, column=0, columnspan=1, padx=1, pady=12)


resp = ctk.CTkLabel(janela, text='', font=('Bodoni', 12, 'bold'),text_color='white',
                      fg_color='black', bg_color='#bce6fe',corner_radius=1000, anchor='center')
resp.grid(row=3, column=2, columnspan=2, padx=12, pady=12)


janela.mainloop()