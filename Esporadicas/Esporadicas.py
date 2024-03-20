import tkinter as tk
from tkinter import *
from pandastable import Table
import pandas as pd
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import locale
import seaborn as sns


class Main:
    def __init__(self):

        # Root principal
        self.root = tk.Tk()
        self.root.geometry("1350x670")
        p = PhotoImage(file='Base/logo.png')
        self.root.iconphoto(False, p)
        self.root.title('Contribuições Esporádicas Visão Prev')
        self.root.config(bg='#ffffff')
        self.table = None
        self.ind = pd.read_excel('Base/primeira.xlsx')
        self.mostrar_tabela()

        # Instancias definidas fora de __init__
        self.como = None
        self.mes = None
        self.ano_inicial = None
        self.ano_final = None
        self.valor_inicial_int = None
        self.valor_final_int = None
        self.primeiro_aviso = None
        self.segundo_aviso = None

        # Caixa
        self.caixa_status = tk.LabelFrame(self.root, text="Status", bd=5, width=90, height=200)
        self.caixa_status.place(x=30, y=220)
        self.caixa_status.config(bg='#ffffff')
        self.caixa_plano = tk.LabelFrame(self.root, text="Plano", bd=5, width=90, height=200)
        self.caixa_plano.place(x=160, y=220)
        self.caixa_plano.config(bg='#ffffff')
        self.caixa_cpf = tk.LabelFrame(self.root, text="Analisar CPF", bd=5, width=90, height=50)
        self.caixa_cpf.place(x=30, y=500)
        self.caixa_cpf.config(bg='#ffffff')

        # Botões
        self.botao_gerar = tk.Button(self.root, text="Gerar Tabela", command=self.coletar_valores)
        self.botao_gerar.place(x=30, y=415)
        self.botao_gerar.config(width=16, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.botao_exportar = tk.Button(self.root, text="Exportar para o Excel", command=self.exportar_excel)
        self.botao_exportar.place(x=30, y=450)
        self.botao_exportar.config(width=16, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.botao_CPF = tk.Button(self.caixa_cpf, text="Analisar CPF", command=self.cpf_analise)
        self.botao_CPF.grid(row=2, pady=5, sticky="W")
        self.botao_CPF.config(width=15, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.botao_sair = tk.Button(self.root, text="Sair", command=self.root.destroy)
        self.botao_sair.config(width=16, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.botao_sair.place(x=30, y=600)

        # Labels
        self.label_posicao = tk.Label(self.root, text="Faixa")
        self.label_posicao.config(font=("Arial", 10))
        self.label_posicao.place(x=30, y=25)
        self.label_posicao.config(bg='#ffffff')
        self.label_mes = tk.Label(self.root, text="Mês")
        self.label_mes.config(font=("Arial", 10))
        self.label_mes.place(x=160, y=25)
        self.label_mes.config(bg='#ffffff')
        self.label_ano_inicial = tk.Label(self.root, text="Ano Inicial")
        self.label_ano_inicial.config(font=("Arial", 10))
        self.label_ano_inicial.place(x=30, y=90)
        self.label_ano_inicial.config(bg='#ffffff')
        self.label_ano_final = tk.Label(self.root, text="Ano Final")
        self.label_ano_final.config(font=("Arial", 10))
        self.label_ano_final.place(x=160, y=90)
        self.label_ano_final.config(bg='#ffffff')
        self.label_valor_inicial = tk.Label(self.root, text="Valor Inicial")
        self.label_valor_inicial.config(font=("Arial", 10))
        self.label_valor_inicial.place(x=30, y=155)
        self.label_valor_inicial.config(bg='#ffffff')
        self.label_valor_final = tk.Label(self.root, text="Valor Final")
        self.label_valor_final.config(font=("Arial", 10))
        self.label_valor_final.place(x=160, y=155)
        self.label_valor_final.config(bg='#ffffff')

        # Caixas de seleção
        self.variavel_posicao = tk.StringVar()
        self.variavel_posicao.set("Opcional")
        self.menu_posicao = tk.OptionMenu(self.root, self.variavel_posicao, 'Opcional', 'Obrigatório')
        self.menu_posicao.place(x=30, y=45)
        self.menu_posicao.config(width=10, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.menu_posicao['menu'].config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")

        self.variavel_mes = tk.StringVar(self.root)
        self.variavel_mes.set("janeiro")
        self.menu_mes = tk.OptionMenu(self.root, self.variavel_mes, 'janeiro', 'fevereiro', 'março',
                                      'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro',
                                      'novembro', 'dezembro')
        self.menu_mes.place(x=160, y=45)
        self.menu_mes.config(width=10, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.menu_mes['menu'].config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")

        self.variavel_ano_inicial = tk.StringVar(self.root)
        self.variavel_ano_inicial.set('2016')
        self.menu_ano_inicial = tk.OptionMenu(self.root, self.variavel_ano_inicial, "2016", "2017",
                                              "2018", "2019", "2020", "2021", "2022", "2023")
        self.menu_ano_inicial.place(x=30, y=110)
        self.menu_ano_inicial.config(width=10, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.menu_ano_inicial['menu'].config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")

        self.variavel_ano_final = tk.StringVar(self.root)
        self.variavel_ano_final.set("2023")
        self.menu_ano_final = tk.OptionMenu(self.root, self.variavel_ano_final, "2016", "2017",
                                            "2018", "2019", "2020", "2021", "2022", "2023")
        self.menu_ano_final.place(x=160, y=110)
        self.menu_ano_final.config(width=10, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.menu_ano_final['menu'].config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")

        # Caixas de texto
        self.valor_inicial = tk.Text(self.root, height=1.2, width=12, bg="#f2f2f2")
        self.valor_inicial.insert(tk.END, 'min')
        self.valor_inicial.place(x=30, y=175)
        self.valor_final = tk.Text(self.root, height=1.2, width=12, bg="#f2f2f2")
        self.valor_final.insert(tk.END, 'max')
        self.valor_final.place(x=160, y=175)
        self.valor_CPF = tk.Text(self.caixa_cpf, height=1.2, width=14, bg="#f2f2f2")
        self.valor_CPF.insert(tk.END, '')
        self.valor_CPF.grid(row=0, pady=5)

        # Checkbox
        self.status_todos_var = tk.IntVar()
        self.status_todos = tk.Checkbutton(self.caixa_status, text='Todos', variable=self.status_todos_var,
                                           command=self.outros_status)
        self.status_todos.grid(row=0,sticky="W")
        self.status_todos.select()
        self.status_todos.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.status_ativo_var = tk.IntVar()
        self.status_ativo = tk.Checkbutton(self.caixa_status, text='Ativo', variable=self.status_ativo_var,
                                           command=self.todos_status)
        self.status_ativo.grid(row=1,sticky="W")
        self.status_ativo.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.status_beneficio_var = tk.IntVar()
        self.status_beneficio = tk.Checkbutton(self.caixa_status, text='Em Benefício',
                                               variable=self.status_beneficio_var, command=self.todos_status)
        self.status_beneficio.grid(row=2,sticky="W")
        self.status_beneficio.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.status_auto_var = tk.IntVar()
        self.status_auto = tk.Checkbutton(self.caixa_status, text='Autopatrocinado', variable=self.status_auto_var,
                                           command=self.todos_status)
        self.status_auto.grid(row=3,sticky="W")
        self.status_auto.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.status_bpd_var = tk.IntVar()
        self.status_bpd = tk.Checkbutton(self.caixa_status, text='BPD', variable=self.status_bpd_var,
                                           command=self.todos_status)
        self.status_bpd.grid(row=4,sticky="W")
        self.status_bpd.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.status_aop_var = tk.IntVar()
        self.status_aop= tk.Checkbutton(self.caixa_status, text='AOP', variable=self.status_aop_var,
                                           command=self.todos_status)
        self.status_aop.grid(row=6, sticky="W")
        self.status_aop.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.plano_todos_var = tk.IntVar()
        self.plano_todos = tk.Checkbutton(self.caixa_plano, text='Todos', variable=self.plano_todos_var,
                                          command=self.outros_plano)
        self.plano_todos.grid(row=0, sticky="W")
        self.plano_todos.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.plano_todos.select()
        self.plano_vt_var = tk.IntVar()
        self.plano_vt = tk.Checkbutton(self.caixa_plano, text='Visão Telefônica', variable=self.plano_vt_var,
                                          command=self.todos_plano)
        self.plano_vt.grid(row=1, sticky="W")
        self.plano_vt.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.plano_vm_var = tk.IntVar()
        self.plano_vm = tk.Checkbutton(self.caixa_plano, text='Visão Multi', variable=self.plano_vm_var,
                                          command=self.todos_plano)
        self.plano_vm.grid(row=2, sticky="W")
        self.plano_vm.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.plano_mv_var = tk.IntVar()
        self.plano_mv = tk.Checkbutton(self.caixa_plano, text='Mais Visão', variable=self.plano_mv_var,
                                          command=self.todos_plano)
        self.plano_mv.grid(row=3, sticky="W")
        self.plano_mv.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")

        self.root.mainloop()
        
    def coletar_valores(self):
        # importando a base total
        dt_total = pd.read_excel('Base/total.xlsx')

        # Coletando valores minimos e maximos
        minimo = dt_total['Valor'].unique().min()
        maximo = dt_total['Valor'].unique().max()

        # Definindo a tabela opcional ou obrigatoria
        posicao_label = self.variavel_posicao.get()
        if posicao_label == 'Opcional':
            self.como = 'outer'
        else:
            self.como = 'inner'

        # Definindo o mês escolhido
        mes_label = self.variavel_mes.get()
        if mes_label == 'janeiro':
            self.mes = 'jan'
        elif mes_label == 'fevereiro':
            self.mes = 'fev'
        elif mes_label == 'março':
            self.mes = 'mar'
        elif mes_label == 'abril':
            self.mes = 'abr'
        elif mes_label == 'maio':
            self.mes = 'mai'
        elif mes_label == 'junho':
            self.mes = 'jun'
        elif mes_label == 'julho':
            self.mes = 'jul'
        elif mes_label == 'agosto':
            self.mes = 'ago'
        elif mes_label == 'setembro':
            self.mes = 'set'
        elif mes_label == 'outubro':
            self.mes = 'out'
        elif mes_label == 'novembro':
            self.mes = 'nov'
        elif mes_label == 'dezembro':
            self.mes = 'dez'

        # Definindo o ano inicial e o ano final
        ano_inicial_label = self.variavel_ano_inicial.get()
        self.ano_inicial = int(ano_inicial_label)
        ano_final_label = self.variavel_ano_final.get()
        self.ano_final = int(ano_final_label)

        # Definindo os status marcados no checkbox
        if self.status_todos_var.get() == 1:
            self.ativo = 'Ativo'
            self.auto = 'Autopatrocinado'
            self.beneficio = 'Em Benefício'
            self.bpd = 'BPD'
            self.aop = 'AOP'
        else:
            if self.status_ativo_var.get() == 1:
                self.ativo = 'Ativo'
            else:
                self.ativo = 'nada'
            if self.status_auto_var.get() == 1:
                self.auto = 'Autopatrocinado'
            else:
                self.auto = 'nada'
            if self.status_beneficio_var.get() == 1:
                self.beneficio = 'Em Benefício'
            else:
                self.beneficio = 'nada'
            if self.status_bpd_var.get() == 1:
                self.bpd = 'BPD'
            else:
                self.bpd = 'nada'
            if self.status_aop_var.get() == 1:
                self.aop = 'AOP'
            else:
                self.aop = 'nada'

        # Definindo os planos marcados no checkbox
        if self.plano_todos_var.get() == 1:
            self.vt = 'Visão Telefônica'
            self.vm = 'Visão Multi'
            self.mv = 'Mais Visão'
        else:
            if self.plano_vt_var.get() == 1:
                self.vt = 'Visão Telefônica'
            else:
                self.vt = 'nada'
            if self.plano_vm_var.get() == 1:
                self.vm = 'Visão Multi'
            else:
                self.vm = 'nada'
            if self.plano_mv_var.get() == 1:
                self.mv = 'Mais Visão'
            else:
                self.mv = 'nada'

        # Definindo os valores iniciais e os valores finais
        valor_inicial_label = self.valor_inicial.get('1.0', tk.END)
        valor_final_label = self.valor_final.get('1.0', tk.END)
        if str.strip(valor_inicial_label).lower() == 'min':
            self.valor_inicial_int = float(minimo)
            try:
                self.valor_final_int = float(valor_final_label)
                self.gerador_tabela()
            except Exception:
                if str.strip(valor_final_label).lower() == 'max':
                    self.valor_final_int = float(maximo)
                    self.gerador_tabela()
                else:
                    self.primeiro_aviso = 'Erro nos valores!'
                    self.segundo_aviso = ''
                    self.aviso()
        else:
            try:
                self.valor_inicial_int = float(valor_inicial_label)
                if str.strip(valor_final_label).lower() == 'max':
                    self.valor_final_int = float(maximo)
                    self.gerador_tabela()
                else:
                    try:
                        self.valor_final_int = float(valor_final_label)
                        self.gerador_tabela()
                    except Exception:
                        self.primeiro_aviso = 'Erro nos valores!'
                        self.segundo_aviso = ''
                        self.aviso()
            except Exception:
                self.primeiro_aviso = 'Erro nos valores!'
                self.segundo_aviso = ''
                self.aviso()

    def gerador_tabela(self):
        # definindo as variaveis do valor inicial e final
        com = self.valor_inicial_int
        fim = self.valor_final_int

        # importando as bases
        df16 = pd.read_excel('Base/2016.xlsx')
        df17 = pd.read_excel('Base/2017.xlsx')
        df18 = pd.read_excel('Base/2018.xlsx')
        df19 = pd.read_excel('Base/2019.xlsx')
        df20 = pd.read_excel('Base/2020.xlsx')
        df21 = pd.read_excel('Base/2021.xlsx')
        df22 = pd.read_excel('Base/2022.xlsx')
        df23 = pd.read_excel('Base/2023.xlsx')
        df_total = pd.read_excel('Base/total.xlsx')

        # Definindo os planos selecionados para as bases
        df16 = df16.loc[(df16['Plano_2016'] == self.vt) | (df16['Plano_2016'] == self.vm) |
                        (df16['Plano_2016'] == self.mv)]
        df17 = df17.loc[(df17['Plano_2017'] == self.vt) | (df17['Plano_2017'] == self.vm) |
                        (df17['Plano_2017'] == self.mv)]
        df18 = df18.loc[(df18['Plano_2018'] == self.vt) | (df18['Plano_2018'] == self.vm) |
                        (df18['Plano_2018'] == self.mv)]
        df19 = df19.loc[(df19['Plano_2019'] == self.vt) | (df19['Plano_2019'] == self.vm) |
                        (df19['Plano_2019'] == self.mv)]
        df20 = df20.loc[(df20['Plano_2020'] == self.vt) | (df20['Plano_2020'] == self.vm) |
                        (df20['Plano_2020'] == self.mv)]
        df21 = df21.loc[(df21['Plano_2021'] == self.vt) | (df21['Plano_2021'] == self.vm) |
                        (df21['Plano_2021'] == self.mv)]
        df22 = df22.loc[(df22['Plano_2022'] == self.vt) | (df22['Plano_2022'] == self.vm) |
                        (df22['Plano_2022'] == self.mv)]
        df23 = df23.loc[(df23['Plano_2023'] == self.vt) | (df23['Plano_2023'] == self.vm) |
                        (df23['Plano_2023'] == self.mv)]
        df_total = df_total.loc[(df_total['Plano'] == self.vt) | (df_total['Plano'] == self.vm) |
                                (df_total['Plano'] == self.mv)]

        # Definindo os status selecionados para as bases
        df16 = df16.loc[(df16['Status_2016'] == self.ativo) | (df16['Status_2016'] == self.auto) |
                        (df16['Status_2016'] == self.beneficio) | (df16['Status_2016'] == self.bpd) |
                        (df16['Status_2016'] == self.aop)]
        df17 = df17.loc[(df17['Status_2017'] == self.ativo) | (df17['Status_2017'] == self.auto) |
                        (df17['Status_2017'] == self.beneficio) | (df17['Status_2017'] == self.bpd) |
                        (df17['Status_2017'] == self.aop)]
        df18 = df18.loc[(df18['Status_2018'] == self.ativo) | (df18['Status_2018'] == self.auto) |
                        (df18['Status_2018'] == self.beneficio) | (df18['Status_2018'] == self.bpd) |
                        (df18['Status_2018'] == self.aop)]
        df19 = df19.loc[(df19['Status_2019'] == self.ativo) | (df19['Status_2019'] == self.auto) |
                        (df19['Status_2019'] == self.beneficio) | (df19['Status_2019'] == self.bpd) |
                        (df19['Status_2019'] == self.aop)]
        df20 = df20.loc[(df20['Status_2020'] == self.ativo) | (df20['Status_2020'] == self.auto) |
                        (df20['Status_2020'] == self.beneficio) | (df20['Status_2020'] == self.bpd) |
                        (df20['Status_2020'] == self.aop)]
        df21 = df21.loc[(df21['Status_2021'] == self.ativo) | (df21['Status_2021'] == self.auto) |
                        (df21['Status_2021'] == self.beneficio) | (df21['Status_2021'] == self.bpd) |
                        (df21['Status_2021'] == self.aop)]
        df22 = df22.loc[(df22['Status_2022'] == self.ativo) | (df22['Status_2022'] == self.auto) |
                        (df22['Status_2022'] == self.beneficio) | (df22['Status_2022'] == self.bpd) |
                        (df22['Status_2022'] == self.aop)]
        df23 = df23.loc[(df23['Status_2023'] == self.ativo) | (df23['Status_2023'] == self.auto) |
                        (df23['Status_2023'] == self.beneficio) | (df23['Status_2023'] == self.bpd) |
                        (df23['Status_2023'] == self.aop)]
        df_total = df_total.loc[(df_total['Status'] == self.ativo) | (df_total['Status'] == self.auto) |
                                (df_total['Status'] == self.beneficio) | (df_total['Status'] == self.bpd) |
                                (df_total['Status'] == self.aop)]

        # Definindo as faixas de valores para as bases
        df2016 = df16[(df16['Mês_2016'] == self.mes) & (df16['Valor_2016'] >= com) & (df16['Valor_2016'] <= fim)]
        df2017 = df17[(df17['Mês_2017'] == self.mes) & (df17['Valor_2017'] >= com) & (df17['Valor_2017'] <= fim)]
        df2018 = df18[(df18['Mês_2018'] == self.mes) & (df18['Valor_2018'] >= com) & (df18['Valor_2018'] <= fim)]
        df2019 = df19[(df19['Mês_2019'] == self.mes) & (df19['Valor_2019'] >= com) & (df19['Valor_2019'] <= fim)]
        df2020 = df20[(df20['Mês_2020'] == self.mes) & (df20['Valor_2020'] >= com) & (df20['Valor_2020'] <= fim)]
        df2021 = df21[(df21['Mês_2021'] == self.mes) & (df21['Valor_2021'] >= com) & (df21['Valor_2021'] <= fim)]
        df2022 = df22[(df22['Mês_2022'] == self.mes) & (df22['Valor_2022'] >= com) & (df22['Valor_2022'] <= fim)]
        df2023 = df23[(df23['Mês_2023'] == self.mes) & (df23['Valor_2023'] >= com) & (df23['Valor_2023'] <= fim)]

        # Checando se o ano inicial é menor ou igual ao ano final
        if self.ano_inicial > self.ano_final:
            self.primeiro_aviso = 'Erro na faixa dos anos!'
            self.segundo_aviso = ''
            self.aviso()

        # Checando se o valor inicial é menor ou igual ao valor final
        elif com > fim:
            self.primeiro_aviso = '   Erro nos valores!'
            self.segundo_aviso = ''
            self.aviso()

        # Criando a tabela quando o ano inicial é igual ao final
        elif self.ano_inicial == self.ano_final:
            self.ind = df_total.loc[(df_total['Valor'] >= com) & (df_total['Valor'] <= fim) &
                                    (df_total['Mês'] == self.mes) & (df_total['Ano'] == self.ano_inicial)]
            self.ind = self.ind.sort_values('Repetição_Total', ascending=False)
            self.mostrar_tabela()

        # Criando a tabela quando a diferença entre o ano inicial e final é de 1
        elif self.ano_final - self.ano_inicial == 1:
            self.ind = pd.merge(locals()['df'+str(self.ano_inicial)], locals()['df'+str(
                self.ano_final)], on='CPF', how=self.como)
            self.calcula_media()

        # Criando a tabela quando a diferença entre o ano inicial e final é de 2
        elif self.ano_final - self.ano_inicial == 2:
            self.ind = pd.merge(locals()['df'+str(self.ano_inicial)], locals()['df'+str(
                self.ano_inicial + 1)], on='CPF', how=self.como).merge(locals()['df'+str(
                self.ano_final)], on='CPF', how=self.como)
            self.calcula_media()

        # Criando a tabela quando a diferença entre o ano inicial e final é de 3
        elif self.ano_final - self.ano_inicial == 3:
            self.ind = pd.merge(locals()['df'+str(self.ano_inicial)], locals()['df'+str(
                self.ano_inicial + 1)], on='CPF', how=self.como).merge(locals()['df'+str(
                self.ano_final - 1)], on='CPF', how=self.como).merge(locals()['df'+str(
                self.ano_final)], on='CPF', how=self.como)
            self.calcula_media()

        # Criando a tabela quando a diferença entre o ano inicial e final é de 4
        elif self.ano_final - self.ano_inicial == 4:
            self.ind = pd.merge(locals()['df'+str(self.ano_inicial)], locals()['df'+str(
                self.ano_inicial + 1)], on='CPF', how=self.como).merge(locals()['df'+str(
                self.ano_inicial + 2)], on='CPF', how=self.como).merge(locals()['df'+str(
                self.ano_final - 1)], on='CPF', how=self.como).merge(locals()['df'+str(
                self.ano_final)], on='CPF', how=self.como)
            self.calcula_media()

        # Criando a tabela quando a diferença entre o ano inicial e final é de 5
        elif self.ano_final - self.ano_inicial == 5:
            self.ind = pd.merge(locals()['df' + str(self.ano_inicial)], locals()['df' + str(
                self.ano_inicial + 1)], on='CPF', how=self.como).merge(locals()['df' + str(
                self.ano_inicial + 2)], on='CPF', how=self.como).merge(locals()['df' + str(
                self.ano_final - 2)], on='CPF', how=self.como).merge(locals()['df' + str(
                self.ano_final - 1)], on='CPF', how=self.como).merge(locals()['df'+str(
                self.ano_final)], on='CPF', how=self.como)
            self.calcula_media()

        # Criando a tabela quando a diferença entre o ano inicial e final é de 6
        elif self.ano_final - self.ano_inicial == 6:
            self.ind = pd.merge(locals()['df' + str(self.ano_inicial)], locals()['df' + str(
                self.ano_inicial + 1)], on='CPF', how=self.como).merge(locals()['df' + str(
                self.ano_inicial + 2)], on='CPF', how=self.como).merge(locals()['df' + str(
                self.ano_inicial + 3)], on='CPF', how=self.como).merge(locals()['df' + str(
                self.ano_final - 2)], on='CPF', how=self.como).merge(locals()['df' + str(
                self.ano_final - 1)], on='CPF', how=self.como).merge(locals()['df'+str(
                self.ano_final)], on='CPF', how=self.como)
            self.calcula_media()

        # Criando a tabela quando a diferença entre o ano inicial e final é de 7
        elif self.ano_final - self.ano_inicial == 7:
            self.ind = pd.merge(locals()['df' + str(self.ano_inicial)], locals()['df' + str(
                self.ano_inicial + 1)], on='CPF', how=self.como).merge(locals()['df' + str(
                self.ano_inicial + 2)], on='CPF', how=self.como).merge(locals()['df' + str(
                self.ano_inicial + 3)], on='CPF', how=self.como).merge(locals()['df' + str(
                self.ano_final - 3)], on='CPF', how=self.como).merge(locals()['df' + str(
                self.ano_final - 2)], on='CPF', how=self.como).merge(locals()['df' + str(
                self.ano_final - 1)], on='CPF', how=self.como).merge(locals()['df'+str(
                self.ano_final)], on='CPF', how=self.como)
            self.calcula_media()

    def calcula_media(self):
        acumulado = []
        lista = 0
        for i in self.ind.columns.str.contains('Acumulado'):
            if i:
                acumulado.append(list(self.ind.iloc[:, [lista]].columns)[0])
                lista += 1
            else:
                lista += 1
        valor = []
        lista = 0
        for i in self.ind.columns.str.contains('Valor'):
            if i:
                valor.append(list(self.ind.iloc[:, [lista]].columns)[0])
                lista += 1
            else:
                lista += 1
        acu_final = []
        for i in self.ind.index:
            calculadora = []
            t = 0
            for j in acumulado:
                self.ind[self.ind.loc[:, [j]].columns] = self.ind.loc[:, [j]].fillna(0)
                if self.ind.loc[i][j] != 0:
                    t += 1
                    calculadora.append(self.ind.loc[i][j])
            acu_final.append(np.mean(calculadora))
        val_final = []
        contador = []
        for i in self.ind.index:
            calculadora = []
            t = 0
            for j in valor:
                self.ind[self.ind.loc[:, [j]].columns] = self.ind.loc[:, [j]].fillna(0)
                if self.ind.loc[i][j] != 0:
                    t += 1
                    calculadora.append(self.ind.loc[i][j])
            contador.append(t)
            val_final.append(np.mean(calculadora))
        self.ind['Media_Valor'] = list(np.around(np.array(val_final), 2))
        self.ind['Media_Acumulado'] = list(np.around(np.array(acu_final), 2))
        self.ind['Repeticao_Total'] = contador
        self.ind = self.ind[['CPF', 'Repeticao_Total', 'Media_Valor', 'Media_Acumulado'] + list(self.ind.columns)[1:-3]]
        self.ind = self.ind.sort_values('Repeticao_Total', ascending=False)
        self.mostrar_tabela()

    def mostrar_tabela(self):
        frame = Frame(self.root)
        frame.place(x=300, y=40)
        self.table = Table(frame, dataframe=self.ind, showtoolbar=False, showstatusbar=True, height=550, width=970)
        self.table.show()

    def exportar_excel(self):
        self.coletar_valores()
        if self.ano_inicial <= self.ano_final:
            if self.valor_inicial_int <= self.valor_final_int:
                try:
                    if self.variavel_ano_inicial.get() != self.variavel_ano_final.get():
                        self.ind.to_excel(f'tabela_ano_{self.variavel_ano_inicial.get()}_a_{
                        self.variavel_ano_final.get()}_mes_{self.variavel_mes.get()}.xlsx', index=False)
                        self.primeiro_aviso = 'Arquivo salvo com sucesso!'
                        self.segundo_aviso = ''
                        self.aviso()
                    else:
                        self.ind.to_excel(f'tabela_ano_{self.variavel_ano_inicial.get()}_mes_{
                        self.variavel_mes.get()}.xlsx', index=False)
                        self.primeiro_aviso = 'Arquivo salvo com sucesso!'
                        self.segundo_aviso = ''
                        self.aviso()
                except Exception:
                    self.primeiro_aviso = 'Erro com a pasta Base'
                    self.segundo_aviso = 'Problemas ao salvar o arquivo'
                    self.aviso()
            else:
                pass
        else:
            pass

    def cpf_analise(self):
        df_total = pd.read_excel('Base/total.xlsx')
        try:
            valor_cpf = int(round(float(str.strip(self.valor_CPF.get('1.0', tk.END)))))
            if valor_cpf in list(df_total['CPF']):
                df_full = df_total[df_total['CPF'] == valor_cpf]
                grafico_janela = tk.Toplevel()
                p = PhotoImage(file='Base/logo.png')
                grafico_janela.iconphoto(False, p)
                grafico_janela.title("Contribuições Esporádicas Visão Prev")
                grafico_janela.config(width=1300, height=650, bg='#ffffff')

                locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
                valor = locale.currency(df_full['Valor'].sum(), grouping=True, symbol=True)
                media = locale.currency(df_full['Valor'].mean(), grouping=True, symbol=True)
                media_acumulado = locale.currency(df_full['Acumulado_Ano'].mean(), grouping=True, symbol=True)

                label_nome = tk.Label(grafico_janela,
                                             text="Nome: " + str(list(df_full['Participante'])[0]))
                label_nome.config(font=("Arial", 10))
                label_nome.place(x=25, y=5)
                label_nome.config(bg='#ffffff')
                label_cpf = tk.Label(grafico_janela,
                                             text="CPF: " + str(list(df_full['CPF'])[0]))
                label_cpf.config(font=("Arial", 10))
                label_cpf.place(x=25, y=30)
                label_cpf.config(bg='#ffffff')
                label_valor_total = tk.Label(grafico_janela,
                                             text="Valor Total das Contribuições: " + str(valor))
                label_valor_total.config(font=("Arial", 10))
                label_valor_total.place(x=400, y=5)
                label_valor_total.config(bg='#ffffff')
                label_valor_medio = tk.Label(grafico_janela,
                                             text="Valor Médio das Contribuições: " + str(media))
                label_valor_medio.config(font=("Arial", 10))
                label_valor_medio.place(x=400, y=30)
                label_valor_medio.config(bg='#ffffff')
                label_valor_anual = tk.Label(grafico_janela,
                                             text="Valor Médio do Acumulado Anual: " + str(media_acumulado))
                label_valor_anual.config(font=("Arial", 10))
                label_valor_anual.place(x=800, y=5)
                label_valor_anual.config(bg='#ffffff')
                label_qtd_total = tk.Label(grafico_janela,
                                             text="Quantidade Total de Contribuições: " + str(
                                                 list(df_full['Repetição_Total'])[0]))
                label_qtd_total.config(font=("Arial", 10))
                label_qtd_total.place(x=800, y=30)
                label_qtd_total.config(bg='#ffffff')

                months_dict = {'jan': 0, 'fev': 1, 'mar': 2, 'abr': 3, 'mai': 4, 'jun': 5, 'jul': 6, 'ago': 7,
                               'set': 8, 'out': 9, 'nov': 10, 'dez': 11}
                months = sorted(list(df_full['Mês'].unique()), key=lambda x: months_dict[x.lower()])
                years = sorted(list(df_full['Ano'].unique()))

                fig_ano = Figure(figsize=(6, 2.8), dpi=100)
                ax_ano = fig_ano.add_subplot()
                sns.countplot(x='Ano', data=df_full, order=years, ax=ax_ano)
                ax_ano.bar_label(ax_ano.containers[0], fontsize=8, label_type='center')
                ax_ano.set_ylabel('')
                ax_ano.set_xlabel('')
                ax_ano.set_title('Quantidade de contribuições por ano\n', fontsize=8)
                canvas_ano = FigureCanvasTkAgg(fig_ano, master=grafico_janela)
                canvas_ano.draw()
                canvas_ano.get_tk_widget().place(x=25, y=355)

                fig_mes = Figure(figsize=(6, 2.8), dpi=100)
                ax_mes = fig_mes.add_subplot()
                sns.barplot(x=df_full['Mês'], y=df_full['Valor'], ax=ax_mes, hue=df_full['Plano'], errorbar=None,
                            order=months)
                for c_mes in ax_mes.containers:
                    labels_mes = [locale.currency(h, grouping=True, symbol=True) if (h := v.get_height()) != 0 else ''
                                  for v in c_mes]
                    ax_mes.bar_label(c_mes, labels=labels_mes, fontsize=8, label_type='center', rotation=90, padding=30)
                ax_mes.ticklabel_format(axis="y", style='plain')
                ax_mes.set_xlabel('')
                ax_mes.set_ylabel('')
                ax_mes.set_title('Valores das contribuições por mês pelo plano\n', fontsize=8)
                ax_mes.legend(fontsize=6)
                canvas_mes = FigureCanvasTkAgg(fig_mes, master=grafico_janela)
                canvas_mes.draw()
                canvas_mes.get_tk_widget().place(x=650, y=65)

                fig_status = Figure(figsize=(6, 2.8), dpi=100)
                ax_status = fig_status.add_subplot()
                sns.countplot(x='Mês', data=df_full, ax=ax_status, order=months, hue=df_full['Status'])
                for c_status in ax_status.containers:
                    labels = [int(h) if (h := v.get_height()) != 0 else '' for v in c_status]
                    ax_status.bar_label(c_status, labels=labels, fontsize=8)
                ax_status.set_xlabel('')
                ax_status.set_ylabel('')
                ax_status.set_title('Quantidade de contribuições por mês pelo status\n', fontsize=8)
                ax_status.legend(fontsize=6)
                canvas_status = FigureCanvasTkAgg(fig_status, master=grafico_janela)
                canvas_status.draw()
                canvas_status.get_tk_widget().place(x=25, y=65)

                lista = []
                for i in years:
                    lista_interna = []
                    for j in months:
                        if j in df_full[df_full['Ano'] == i]['Mês'].values:
                            lista_interna.append(True)
                        else:
                            lista_interna.append(False)
                    lista.append(lista_interna)
                matriz = pd.DataFrame(lista, index=years, columns=months)
                fig_matriz = Figure(figsize=(6, 2.8), dpi=100)
                ax_mat = fig_matriz.add_subplot()
                sns.heatmap(matriz, cmap='Blues', annot=False, cbar=False, ax=ax_mat)
                ax_mat.set_yticks(np.arange(len(years)), years)
                ax_mat.set_xticks(np.arange(len(months)), months)
                ax_mat.set_yticklabels(ax_mat.get_yticklabels(), rotation=0, ha='right')
                ax_mat.set_title('Meses de cada ano que houveram contribuições\n', fontsize=8)
                canvas_matriz = FigureCanvasTkAgg(fig_matriz, master=grafico_janela)
                canvas_matriz.draw()
                canvas_matriz.get_tk_widget().place(x=650, y=355)

                botao_cpf_fechar = tk.Button(grafico_janela, text="     Fechar    ",
                                             command=grafico_janela.destroy)
                botao_cpf_fechar.place(x=1200, y=20)
                botao_cpf_fechar.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")

            else:
                self.primeiro_aviso = 'Erro no CPF!'
                self.segundo_aviso = ''
                self.aviso()

        except Exception:
            self.primeiro_aviso = 'Erro no CPF!'
            self.segundo_aviso = ''
            self.aviso()

    def aviso(self):
        # Janela que gera os avisos
        aviso_janela = tk.Toplevel()
        p = PhotoImage(file='Base/logo.png')
        # Janela
        aviso_janela.iconphoto(False, p)
        aviso_janela.title("Contribuições Esporádicas Visão Prev")
        aviso_janela.config(width=300, height=200)
        aviso_janela.resizable(width=False, height=False)
        aviso_janela.config(bg='#ffffff')
        # Botão
        botao_aviso = tk.Button(aviso_janela, text="Fechar", command=aviso_janela.destroy)
        botao_aviso.place(x=120, y=150)
        botao_aviso.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        # Label
        label_aviso = tk.Label(aviso_janela, text=str(self.primeiro_aviso) + '\n' + str(self.segundo_aviso))
        label_aviso.config(font=("Courier", 10))
        label_aviso.place(x=60, y=60)
        label_aviso.config(bg='#ffffff')

    def todos_status(self):
        # Tira a seleção de todos no status
        self.status_todos.deselect()

    def outros_status(self):
        # Tira a seleção dos status
        self.status_ativo.deselect()
        self.status_beneficio.deselect()
        self.status_auto.deselect()
        self.status_bpd.deselect()
        self.status_aop.deselect()

    def todos_plano(self):
        # Tira a seleção de todos no plano
        self.plano_todos.deselect()

    def outros_plano(self):
        # Tira a seleção dos planos
        self.plano_vt.deselect()
        self.plano_vm.deselect()
        self.plano_mv.deselect()


Main()
