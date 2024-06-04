# pip install openpyxl
# pip install pandas
# pip install pyinstaller
import tkinter as tk
import pandas as pd


class Main:

    def __init__(self):
        self.root = tk.Tk()
        self.root.config(width=230, height=180)
        self.root.resizable(width=False, height=False)
        self.root.title('Gerador')
        self.root.config(bg='#ffffff')

        # instancias definidas fora de __init__
        self.primeiro_aviso = None
        self.segundo_aviso = None

        # Labels
        label = tk.Label(self.root, text="Gerador de tabela\n informativa de atendimento")
        label.config(bg='#ffffff', font=("Arial", 10))
        label.place(x=27, y=25)

        # Botões
        self.botao_gerar = tk.Button(self.root, text="Gerar Excel", command=self.executar)
        self.botao_gerar.config(width=16, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.botao_gerar.place(x=50, y=100)
        self.botao_sair = tk.Button(self.root, text="Sair", command=self.root.destroy)
        self.botao_sair.config(width=16, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.botao_sair.place(x=50, y=130)

        self.root.mainloop()

    def executar(self):
        try:
            df = pd.read_csv('Base/calendario.CSV')
            df = df[df['Organizador da reunião'] != 'Gustavo Silva de Oliveira']
            lista_col = list(df.columns)
            lista_col.remove('Assunto')
            lista_col.remove('Data de início')
            lista_col.remove('Hora de início')
            lista_col.remove('Descrição')
            df.drop(lista_col, axis=1, inplace=True)
            df = df[(df['Assunto'].str.contains('Atendimento personalizado')) |
                    (df['Assunto'].str.contains('Atendimento Opções'))]
            df.reset_index(drop=True, inplace=True)
            lista1 = []
            lista2 = []
            lista3 = []
            for i in df.index:
                lista = df['Descrição'][i].replace('\r', '').split('\n')
                lista4 = []
                for j in lista:
                    if 'Nome:' in j or 'Name:' in j:
                        lista1.append(j)
                    if 'Resposta' in j or 'Answer' in j:
                        lista4.append(j)
                lista2.append(lista4[0])
                lista3.append(lista4[1])
            df['nome'] = lista1
            df['respostas 1'] = lista2
            df['respostas 2'] = lista3
            df.drop('Descrição', axis=1, inplace=True)
            try:
                df.to_excel('Base/calendario.xlsx', index=False)
                self.primeiro_aviso = 'Comunicados gerados com'
                self.segundo_aviso = ' sucesso!'
                self.aviso()
            except Exception:
                self.primeiro_aviso = 'Problemas ao salvar'
                self.segundo_aviso = ' a tabela'
                self.aviso()

        except Exception:
            self.primeiro_aviso = 'Problemas ao importar'
            self.segundo_aviso = ' a tabela'
            self.aviso()

    def aviso(self):
        # Janela que gera os avisos
        aviso_janela = tk.Toplevel()
        # Janela
        aviso_janela.title("Gerador")
        aviso_janela.config(width=300, height=200, bg='#ffffff')
        aviso_janela.resizable(width=False, height=False)
        # Botão
        botao_aviso = tk.Button(aviso_janela, text="Fechar", command=aviso_janela.destroy)
        botao_aviso.place(x=105, y=150)
        botao_aviso.config(width=10)
        botao_aviso.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        # Label
        label_aviso = tk.Label(aviso_janela, text=str(self.primeiro_aviso) + '\n' + str(self.segundo_aviso))
        label_aviso.config(font=("Courier", 10))
        label_aviso.place(x=50, y=60)
        label_aviso.config(bg='#ffffff')


Main()
