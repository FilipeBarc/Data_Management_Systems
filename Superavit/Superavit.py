# pip install openpyxl
# pip install fsspec
# pip install Pyarrow
# pip install pyinstaller
import tkinter as tk
from tkinter import *
import pandas as pd
import numpy as np
from datetime import datetime


class Main:

    def __init__(self):
        self.root = tk.Tk()
        self.root.config(width=320, height=450)
        self.root.resizable(width=False, height=False)
        p = PhotoImage(file='Base//1.Logo//logo.png')
        self.root.iconphoto(False, p)
        self.root.title('Superavit Visão Prev')
        self.root.config(bg='#ffffff')
        self.hoje = datetime.today()

        # instancias definidas fora de __init__
        self.status = None
        self.condicao_patrocinadora = None
        self.df_cotas = None
        self.df_patrocinadora = None
        self.df_saldo = None
        self.df_abatimento = None
        self.df_retirada = None
        self.saldo_cotas_antes = None
        self.saldo_cotas_depois = None
        self.saldo_real_antes = None
        self.saldo_real_depois = None
        self.abatimento_real = None
        self.abatimento_cotas = None
        self.data = None
        self.data_janela = None
        self.valor_data = None
        self.primeiro_aviso = None
        self.segundo_aviso = None
        self.plano = None

        # Caixas de seleção
        self.variavel_status = tk.StringVar()
        self.variavel_status.set("Autopatrocinado")
        self.menu_status = tk.OptionMenu(self.root, self.variavel_status, 'Autopatrocinado', 'Ativo',
                                         command=self.autopatrocinado)
        self.menu_status.config(width=15, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.menu_status['menu'].config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black",)
        self.menu_status.place(x=50, y=50)
        self.variavel_plano = tk.StringVar()
        self.variavel_plano.set("Visão Telefônica")
        self.menu_plano = tk.OptionMenu(self.root, self.variavel_plano, 'Visão Telefônica', 'Visão Multi')
        self.menu_plano.place(x=50, y=120)
        self.menu_plano.config(width=15, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.menu_plano['menu'].config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")

        # Labels
        self.label_status = tk.Label(self.root, text="Status do participante")
        self.label_status.config(font=("Arial", 10))
        self.label_status.place(x=50, y=25)
        self.label_status.config(bg='#ffffff')
        self.label_plano = tk.Label(self.root, text="Tipo de plano")
        self.label_plano.config(font=("Arial", 10))
        self.label_plano.place(x=50, y=98)
        self.label_plano.config(bg='#ffffff')

        # Botões
        self.botao_gerar = tk.Button(self.root, text="       Gerar Tabela        ", command=self.importar_tabelas)
        self.botao_gerar.place(x=50, y=360)
        self.botao_gerar.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.botao_sair = tk.Button(self.root, text="               Sair               ", command=self.root.destroy)
        self.botao_sair.place(x=50, y=390)
        self.botao_sair.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")

        # Checkbox
        self.var_patrocinadora = tk.IntVar()
        self.check_patrocinadora = tk.Checkbutton(self.root,
                                                  text='A tabela de abatimento dos ativos\nnão possui patrocinadora',
                                                  variable=self.var_patrocinadora, command=self.check_data,
                                                  state=DISABLED)
        self.check_patrocinadora.place(x=50, y=190)
        self.check_patrocinadora.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.var_data = tk.IntVar()
        self.check_data = tk.Checkbutton(self.root,
                                         text='A tabela de abatimento dos ativos\nnão possui data',
                                         variable=self.var_data, state=DISABLED)
        self.check_data.place(x=50, y=240)
        self.check_data.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")

        self.root.mainloop()

    def autopatrocinado(self, *args):
        # Tira a seleção e bloqueia os checkboxes quando seleciona o Autopatrocinado
        self.check_patrocinadora.deselect()
        self.check_data.deselect()
        if self.variavel_status.get() == 'Ativo':
            self.check_patrocinadora.config(state=ACTIVE)
        elif self.variavel_status.get() == 'Autopatrocinado':
            self.check_patrocinadora.config(state=DISABLED)
            self.check_data.config(state=DISABLED)

    def check_data(self):
        # Tira a seleção e bloqueia os checkbox data quando o checkbox autopatrocinado não é selecionado
        self.check_data.deselect()
        if self.var_patrocinadora.get() == 1:
            self.check_data.config(state=ACTIVE)
        elif self.var_patrocinadora.get() == 0:
            self.check_data.config(state=DISABLED)

    def importar_tabelas(self):

        # Coleta os valores selecionados na primeira interface
        self.status = self.variavel_status.get()
        self.condicao_patrocinadora = self.var_patrocinadora.get()
        self.plano = self.variavel_plano.get()
        condicao_data = self.var_data.get()

        if self.plano == 'Visão Telefônica':
            plano_importar = self.plano[:11]
        else:
            plano_importar = self.plano

        try:
            # Importando a tabela cotas em Y:
            xls1 = pd.ExcelFile("Y://CORPORATIVO//INVESTIMENTOS//Cotas Perfis e Planos.xlsx")
            df1 = pd.read_excel(xls1, 'Cotas ' + plano_importar)
            df1.columns = df1.iloc[1]
            xls2 = pd.ExcelFile("Y://CORPORATIVO//INVESTIMENTOS//Cotas Perfis e Planos.xlsx")
            df2 = pd.read_excel(xls2, 'Previstas ' + plano_importar)
            df2.columns = df2.iloc[1]
            df1.drop([df1.index[0], df1.index[1]], inplace=True)
            df1.columns.name = None
            df1.reset_index(drop=True, inplace=True)
            df2.drop([df2.index[0], df2.index[1]], inplace=True)
            df2.columns.name = None
            df2.reset_index(drop=True, inplace=True)
            df1 = df1[['Data', 'Conservador']]
            df1 = df1.dropna()
            df1['Data'] = pd.to_datetime(df1['Data'])
            df_cotas_reais = df1[df1['Data'] >= pd.Timestamp('2023-12-29')]
            df_cotas_reais = df_cotas_reais.rename(columns={'Conservador': 'Cotas'}).copy()
            df2 = df2[['Data', 'Conservador']]
            df2 = df2.dropna()
            df2['Data'] = pd.to_datetime(df2['Data'])
            df_cotas_previstas = df2[df2['Data'] > df1['Data'].iloc[-1]]
            df_cotas_previstas = df_cotas_previstas.rename(columns={'Conservador': 'Cotas'}).copy()
            data_repetida = []
            dias_diferenca = (pd.to_datetime(self.hoje) - df_cotas_previstas['Data'].iloc[-1]).days
            for i in range(dias_diferenca - 1, -1, -1):
                data_repetida.append(str((pd.to_datetime(self.hoje) - pd.DateOffset(days=i)).date()))
            df_cotas_repetidas = pd.DataFrame(data_repetida, columns=['Data'])
            df_cotas_repetidas['Cotas'] = df_cotas_previstas['Cotas'].iloc[-1]
            df_cotas_inc = pd.concat([df_cotas_reais, df_cotas_previstas, df_cotas_repetidas])
            df_cotas_inc.reset_index(drop=True, inplace=True)
            df_cotas_inc['Data'] = pd.to_datetime(df_cotas_inc['Data'])
            data_repetida = []
            dias_diferenca = (pd.to_datetime(self.hoje) - df_cotas_inc['Data'].iloc[0]).days
            for i in range(dias_diferenca, -1, -1):
                data_repetida.append(str((pd.to_datetime(self.hoje) - pd.DateOffset(days=i)).date()))
            barra = []
            falta = 0
            for i in data_repetida:
                if i in str(df_cotas_inc['Data']):
                    barra.append(pd.Series(df_cotas_inc[df_cotas_inc['Data'] == i]['Cotas'].values).to_list()[0])
                    falta = i
                else:
                    barra.append(pd.Series(df_cotas_inc[df_cotas_inc['Data'] == falta]['Cotas'].values).to_list()[0])
            self.df_cotas = pd.DataFrame(data_repetida, columns=['Data'])
            self.df_cotas['Cotas'] = barra
            self.df_cotas['Data'] = pd.to_datetime(self.df_cotas['Data'])

            try:
                # verificando se a escolha não foi de ativos do Visão Multi
                if self.plano == 'Visão Multi' and self.status == 'Ativo':
                    self.primeiro_aviso = ' Não existe Ativos'
                    self.segundo_aviso = ' no plano Visão Multi'
                    self.aviso()

                else:

                    if self.plano == 'Visão Telefônica' and self.status == 'Autopatrocinado':
                        # Importando o saldo da pasta Visão Multi Autopatrocinado
                        xls8 = pd.ExcelFile(
                            'Base//Saldo//Autopatrocinado//Visão Multi//Saldo_' + str(
                                self.hoje.month - 1).zfill(2) + '_' + str(self.hoje.year) + '.xlsx')
                        self.df_saldo = pd.read_excel(xls8)

                    elif self.plano == 'Visão Multi' and self.status == 'Autopatrocinado':
                        # Importando o saldo da pasta Visão Telefônica Autopatrocinado
                        xls8 = pd.ExcelFile(
                            'Base//Saldo//Autopatrocinado//Visão Telefônica//Saldo_' + str(
                                self.hoje.month - 1).zfill(2) + '_' + str(self.hoje.year) + '.xlsx')
                        self.df_saldo = pd.read_excel(xls8)

                    elif self.plano == 'Visão Telefônica' and self.status == 'Ativo':
                        # Importando o saldo da pasta Visão Telefônica Ativo
                        xls8 = pd.ExcelFile(
                            'Base//Saldo//Ativo//Visão Telefônica//Saldo_' + str(
                                self.hoje.month - 1).zfill(2) + '_' + str(self.hoje.year) + '.xlsx')
                        self.df_saldo = pd.read_excel(xls8)

                    try:
                        # Importando a data de pagamento da patrocinadora
                        xls3 = pd.ExcelFile(
                            'Base//Pagamentos_patrocinadora//Pagamentos_Patrocinadora_' + str(self.hoje.month - 1).zfill(
                                2) + '_' + str(
                                self.hoje.year) + '.xlsx')
                        self.df_patrocinadora = pd.read_excel(xls3)

                        # Importando o abatimento dos ativos e autopatrocinados
                        abatimento_nome = 'Abatimento_' + str(self.hoje.month - 1).zfill(2) + '_' + str(self.hoje.year)
                        xls4 = pd.ExcelFile(
                            'Base//Abatimentos//' + self.status + '//' + self.plano + '//' + abatimento_nome + '.xlsx')
                        self.df_abatimento = pd.read_excel(xls4)
                        # Importando a retirada dos ativos e autopatrocinados
                        xls6 = pd.ExcelFile(
                            'Base//Retirada//' + self.status + '//' + self.plano + '//Retirada_' + str(
                                self.hoje.month - 1).zfill(2) + '_' + str(self.hoje.year) + '.xlsx')
                        self.df_retirada = pd.read_excel(xls6)

                        # Definindo o nome das colunas
                        self.saldo_cotas_antes = 'Saldo_cotas_' + str(self.hoje.month - 1).zfill(2) + '_' + str(
                            self.hoje.year)
                        self.saldo_cotas_depois = 'Saldo_cotas_' + str(self.hoje.month).zfill(2) + '_' + str(self.hoje.year)
                        self.saldo_real_antes = 'Saldo_real_' + str(self.hoje.month - 1).zfill(2) + '_' + str(
                            self.hoje.year)
                        self.saldo_real_depois = 'Saldo_real_' + str(self.hoje.month).zfill(2) + '_' + str(self.hoje.year)
                        self.abatimento_cotas = 'Abatimento_cotas_' + str(self.hoje.month - 1).zfill(2) + '_' + str(
                            self.hoje.year)
                        self.abatimento_real = 'Abatimento_real_' + str(self.hoje.month - 1).zfill(2) + '_' + str(
                            self.hoje.year)

                        # Caso a tabela não possua data e foi marcado o checkbox data o programa abre a entrada da data
                        if condicao_data == 1:
                            self.mensagem_data()
                        # Caso a tabela possua data e não foi marcado o checkbox data o programa vai criar as tabelas
                        else:
                            self.verifica_coluna()

                    except Exception:
                        self.primeiro_aviso = 'Erro na pasta Base'
                        self.segundo_aviso = 'Problemas com as tabelas'
                        self.aviso()

            except Exception:
                self.primeiro_aviso = 'Erro no Saldo'
                self.segundo_aviso = 'Inicie com o Visão Telefônica'
                self.aviso()

        except Exception:
            self.primeiro_aviso = 'Erro de conexão com a rede Y:'
            self.segundo_aviso = ''
            self.aviso()

    def verifica_coluna(self):

        # Verifica se a seleção da interface foi "Ativo" e se o check "Patrocinadora" esta desmarcado
        if self.status == 'Ativo' and self.condicao_patrocinadora == 0:
            # Verifica se tem valores nulos no CPF
            if len(self.df_abatimento[self.df_abatimento['CPF'].isnull()]['CPF']) == 0:
                # Verifica se o CPF é formado por números inteiros
                if self.df_abatimento['CPF'].dtype == 'int64':
                    # Verifica se tem valores nulos no Status
                    if len(self.df_abatimento[self.df_abatimento['Status'].isnull()]['Status']) == 0:
                        # Verifica se tem algum valor diferente de "ativos" na coluna status
                        passe = []
                        for i in self.df_abatimento['Status']:
                            if i.lower() == 'ativo':
                                passe.append(1)
                            else:
                                passe.append(0)
                        if 0 in passe:
                            self.primeiro_aviso = 'Erro na coluna Status'
                            self.segundo_aviso = 'Contém valores incorretos'
                            self.aviso()
                        else:
                            # Verifica se a coluna abatimento tem valores nulos
                            if len(self.df_abatimento[
                                       self.df_abatimento[self.abatimento_real].isnull()][self.abatimento_real]) == 0:
                                # Verifica se as colunas são números flutuantes
                                if self.df_abatimento[self.abatimento_real].dtype == 'float64':
                                    try:
                                        if len(self.df_abatimento[
                                                   self.df_abatimento['Patrocinadora'].isnull()]['Patrocinadora']) == 0:
                                            # Verifica se os códigos da patrocinadora no base abatimento são os mesmos
                                            # da base patrocinadora
                                            if len(set(self.df_abatimento['Patrocinadora']).intersection(
                                                    self.df_patrocinadora['Codigo'])) == len(set(
                                                        self.df_abatimento['Patrocinadora'])):

                                                # Chama a função gerar tabela
                                                self.gerar_tabela()
                                            else:
                                                self.primeiro_aviso = 'Erro na coluna Patrocinadora'
                                                self.segundo_aviso = 'Contém valores incorretos'
                                                self.aviso()
                                        else:
                                            self.primeiro_aviso = 'Erro na coluna Patrocinadora'
                                            self.segundo_aviso = 'Contém valores nulos'
                                            self.aviso()
                                    except Exception:
                                        self.primeiro_aviso = 'Erro na coluna Patrocinadora'
                                        self.segundo_aviso = ''
                                        self.aviso()
                                else:
                                    self.primeiro_aviso = 'Erro na coluna Abatimento'
                                    self.segundo_aviso = 'Contém valores incorretos'
                                    self.aviso()
                            else:
                                self.primeiro_aviso = 'Erro na coluna Abatimento'
                                self.segundo_aviso = 'Contém valores nulos'
                                self.aviso()
                    else:
                        self.primeiro_aviso = 'Erro na coluna Status'
                        self.segundo_aviso = 'Contém valores nulos'
                        self.aviso()
                else:
                    self.primeiro_aviso = 'Erro na coluna CPF'
                    self.segundo_aviso = 'Contém valores incorretos'
                    self.aviso()
            else:
                self.primeiro_aviso = 'Erro na coluna CPF'
                self.segundo_aviso = 'Contém valores nulos'
                self.aviso()

        # Verifica se a seleção da interface foi "Ativo" sem patrocinadora ou "Autopatrocinado"
        else:
            # Verifica se tem valores nulos no CPF
            if len(self.df_abatimento[self.df_abatimento['CPF'].isnull()]['CPF']) == 0:
                # Verifica se o CPF é formado por números inteiros
                if self.df_abatimento['CPF'].dtype == 'int64':
                    # Verifica se tem valores nulos no Status
                    if len(self.df_abatimento[self.df_abatimento['Status'].isnull()]['Status']) == 0:
                        passe = []
                        for i in self.df_abatimento['Status']:
                            # Verifica se tem algum valor diferente de "ativos" ou "autopatrocinado" na coluna status
                            if i.lower() == 'ativo' or i.lower() == 'autopatrocinado':
                                passe.append(1)
                            else:
                                passe.append(0)
                        if 0 in passe:
                            self.primeiro_aviso = 'Erro na coluna Status'
                            self.segundo_aviso = 'Contém valores incorretos'
                            self.aviso()
                        else:
                            # Verifica se tem valores nulos no Abatimento
                            if len(self.df_abatimento[
                                       self.df_abatimento[self.abatimento_real].isnull()][self.abatimento_real]) == 0:
                                # Verifica se as colunas são números flutuantes
                                if self.df_abatimento[self.abatimento_real].dtype == 'float64':

                                    try:
                                        self.data = pd.to_datetime(self.df_abatimento['Data'])
                                        if len(self.df_abatimento[self.df_abatimento['Data'].isnull()]['Data']) == 0:
                                            # Verifica se as datas na base abatimento são as mesmas das na base cotas
                                            if len(set(self.df_abatimento['Data']).intersection(
                                                self.df_cotas['Data'])) == len(set(
                                                    self.df_abatimento['Data'])):

                                                # Chama a função gerar tabela

                                                self.gerar_tabela()
                                            else:
                                                self.primeiro_aviso = 'Erro na coluna Data'
                                                self.segundo_aviso = 'Contém datas incorretas'
                                                self.aviso()
                                        else:
                                            self.primeiro_aviso = 'Erro na coluna Data'
                                            self.segundo_aviso = 'Contém valores nulos'
                                            self.aviso()
                                    except Exception:
                                        self.primeiro_aviso = 'Erro na coluna Data'
                                        self.segundo_aviso = 'Formato ou valores incorretos'
                                        self.aviso()
                                else:
                                    self.primeiro_aviso = 'Erro na coluna Abatimento'
                                    self.segundo_aviso = 'Contém valores incorretos'
                                    self.aviso()
                            else:
                                self.primeiro_aviso = 'Erro na coluna Abatimento'
                                self.segundo_aviso = 'Contém valores nulos'
                                self.aviso()
                    else:
                        self.primeiro_aviso = 'Erro na coluna Status'
                        self.segundo_aviso = 'Contém valores nulos'
                        self.aviso()
                else:
                    self.primeiro_aviso = 'Erro na coluna CPF'
                    self.segundo_aviso = 'Contém valores não numéricos'
                    self.aviso()
            else:
                self.primeiro_aviso = 'Erro na coluna CPF'
                self.segundo_aviso = 'Contém valores nulos'
                self.aviso()

    def gerar_tabela(self):

        # remove os CPFs dos participantes ativos que precisam serem retirados da tabela saldo
        for i in self.df_retirada['CPF']:
            for j in self.df_saldo.index:
                if self.df_saldo['CPF'][j] == i:
                    self.df_saldo = self.df_saldo.drop(j)
        self.df_saldo.reset_index(drop=True, inplace=True)

        # Verifica se a seleção da interface foi "Ativo" e se o check "Patrocinadora" esta desmarcado
        if self.status == 'Ativo' and self.condicao_patrocinadora == 0:
            # Cria a coluna data na tabela abatimento usando o código das patrocinadoras
            data_abatimento = []
            for i in self.df_abatimento.index:
                for j in self.df_patrocinadora.index:
                    if self.df_abatimento['Patrocinadora'][i] == self.df_patrocinadora['Codigo'][j]:
                        data_abatimento.append(self.df_patrocinadora['Data'][j])

            self.df_abatimento['Data'] = data_abatimento

        # Calcula as cotas usando a coluna data na tabela abatimento e a tabela cotas

        cota_abatimento = []
        for i in self.df_abatimento.index:
            for j in self.df_cotas.index:
                if self.df_abatimento['Data'][i] == self.df_cotas['Data'][j]:
                    cota_abatimento.append(self.df_abatimento[self.abatimento_real][i] / self.df_cotas['Cotas'][j])

        self.df_abatimento[self.abatimento_cotas] = cota_abatimento

        # Cria a tabela análise
        cota_para_abater = []
        cota_resultado = []
        data_final = []
        abatimento_real = []

        for i in list(self.df_saldo['CPF']):
            if i in list(self.df_abatimento['CPF']):
                abatimento_real.append(pd.Series(self.df_abatimento[
                                                     self.df_abatimento['CPF'] == i][
                                                     self.abatimento_real].values).to_list()[0])
                cota_para_abater.append(pd.Series(self.df_abatimento[
                                                      self.df_abatimento['CPF'] == i][
                                                      self.abatimento_cotas].values).to_list()[0])
                cota_resultado.append(pd.Series(self.df_saldo[self.df_saldo['CPF'] == i][
                                                    self.saldo_cotas_antes].values).to_list()[0] -
                                      pd.Series(self.df_abatimento[
                                                    self.df_abatimento['CPF'] == i][
                                                    self.abatimento_cotas].values).to_list()[0])
                data_final.append(self.df_abatimento[self.df_abatimento['CPF'] == i]['Data'].iloc[0])
            else:
                abatimento_real.append(np.nan)
                cota_para_abater.append(np.nan)
                cota_resultado.append(pd.Series(self.df_saldo[self.df_saldo['CPF'] == i][
                                                    self.saldo_cotas_antes].values).to_list()[0])
                data_final.append(np.nan)

        self.df_saldo[self.abatimento_cotas] = cota_para_abater
        self.df_saldo[self.saldo_cotas_depois] = cota_resultado
        self.df_saldo['Data_cota'] = data_final
        self.df_saldo[self.abatimento_real] = abatimento_real

        saldo_real_antes = []
        saldo_real_depois = []
        cota_usada = []

        for i in list(self.df_saldo['CPF']):
            if i in list(self.df_abatimento['CPF']):
                data = self.df_saldo[self.df_saldo['CPF'] == i]['Data_cota'].iloc[0]
                cota_usada.append(pd.Series(self.df_cotas[self.df_cotas['Data'] == data]['Cotas'].values).to_list()[0])
                saldo_real_antes.append(pd.Series(self.df_saldo[self.df_saldo['CPF'] == i][
                                self.saldo_cotas_antes].values).to_list()[0] * pd.Series(self.df_cotas[self.df_cotas[
                                                                        'Data'] == data]['Cotas'].values).to_list()[0])
                saldo_real_depois.append(pd.Series(self.df_saldo[self.df_saldo['CPF'] == i][
                                self.saldo_cotas_depois].values).to_list()[0] * pd.Series(self.df_cotas[self.df_cotas[
                                                                        'Data'] == data]['Cotas'].values).to_list()[0])
            else:
                cota_usada.append(np.nan)
                saldo_real_antes.append(np.nan)
                saldo_real_depois.append(np.nan)

        self.df_saldo[self.saldo_real_antes] = saldo_real_antes
        self.df_saldo[self.saldo_real_depois] = saldo_real_depois
        self.df_saldo['Cota_utilizada'] = cota_usada

        if self.status == 'Ativo' and self.condicao_patrocinadora == 0:
            patrocinadora_final = []
            for i in list(self.df_saldo['CPF']):
                if i in list(self.df_abatimento['CPF']):

                    patrocinadora_final.append(pd.Series(self.df_abatimento[
                                                         self.df_abatimento['CPF'] == i][
                                                         'Patrocinadora'].values).to_list()[0])
                else:
                    patrocinadora_final.append(np.nan)
            self.df_saldo['Patrocinadora'] = patrocinadora_final
            analise = self.df_saldo[self.df_saldo['Status'] != 'Autopatrocinado'][['CPF', 'Status',
                                                                                   self.saldo_cotas_antes,
                                                                                   self.saldo_real_antes,
                                                                                   self.abatimento_cotas,
                                                                                   self.abatimento_real,
                                                                                   self.saldo_cotas_depois,
                                                                                   self.saldo_real_depois,
                                                                                   'Cota_utilizada', 'Data_cota',
                                                                                   'Patrocinadora']].sort_values(
                    self.saldo_real_antes, ascending=False)

        else:
            if self.status == 'Ativo':
                analise = self.df_saldo[self.df_saldo['Status'] != 'Autopatrocinado'][['CPF', 'Status',
                                                                                       self.saldo_cotas_antes,
                                                                                       self.saldo_real_antes,
                                                                                       self.abatimento_cotas,
                                                                                       self.abatimento_real,
                                                                                       self.saldo_cotas_depois,
                                                                                       self.saldo_real_depois,
                                                                                       'Cota_utilizada',
                                                                                       'Data_cota']].sort_values(
                    self.saldo_real_antes, ascending=False)
            else:
                analise = self.df_saldo[self.df_saldo['Status'] == 'Autopatrocinado'][['CPF', 'Status',
                                                                                       self.saldo_cotas_antes,
                                                                                       self.saldo_real_antes,
                                                                                       self.abatimento_cotas,
                                                                                       self.abatimento_real,
                                                                                       self.saldo_cotas_depois,
                                                                                       self.saldo_real_depois,
                                                                                       'Cota_utilizada',
                                                                                       'Data_cota']].sort_values(
                    self.saldo_real_antes, ascending=False)

        try:
            # Exporta a tabela análise, saldo e retirada para o Excel
            xls10 = 'Base//Analise//' + self.status + '//' + self.plano + '//Analise_' + str(self.hoje.month).zfill(
                2) + '_' + str(self.hoje.year) + '.xlsx'
            analise.to_excel(xls10, index=False)
            df_retirar = analise[(analise[self.saldo_cotas_depois] < 0) & (
                ~analise['Data_cota'].isnull())][['CPF', 'Status', self.saldo_cotas_depois]]
            xls12 = 'Base//Retirada//' + self.status + '//' + self.plano + '//Retirada_' + str(
                self.hoje.month).zfill(2) + '_' + str(self.hoje.year) + '.xlsx'
            df_retirar.to_excel(xls12, index=False)
            if self.plano == 'Visão Telefônica' and self.status == 'Autopatrocinado':
                df_saldo_final = self.df_saldo[['CPF', 'Status', self.saldo_cotas_depois]]
                df_saldo_final = df_saldo_final.rename(columns={self.saldo_cotas_depois: self.saldo_cotas_antes})
                xls11 = 'Base//Saldo//Autopatrocinado//Visão Telefônica//Saldo_' + str(self.hoje.month - 1).zfill(
                    2) + '_' + str(self.hoje.year) + '.xlsx'
                df_saldo_final.to_excel(xls11, index=False)

            elif self.plano == 'Visão Multi' and self.status == 'Autopatrocinado':
                df_saldo_final = self.df_saldo[['CPF', 'Status', self.saldo_cotas_depois]]
                xls11 = 'Base//Saldo//Autopatrocinado//Visão Multi//Saldo_' + str(self.hoje.month).zfill(
                    2) + '_' + str(self.hoje.year) + '.xlsx'
                df_saldo_final.to_excel(xls11, index=False)

                df_saldo_final = self.df_saldo[['CPF', 'Status', self.saldo_cotas_depois]]
                df_saldo_final = df_saldo_final.rename(columns={self.saldo_cotas_depois: self.saldo_cotas_antes})
                xls11 = 'Base//Saldo//Ativo//Visão Telefônica//Saldo_' + str(self.hoje.month - 1).zfill(
                    2) + '_' + str(self.hoje.year) + '.xlsx'
                df_saldo_final.to_excel(xls11, index=False)

            elif self.plano == 'Visão Telefônica' and self.status == 'Ativo':
                df_saldo_final = self.df_saldo[['CPF', 'Status', self.saldo_cotas_depois]]
                df_saldo_final = df_saldo_final.rename(columns={self.saldo_cotas_depois: self.saldo_cotas_antes})
                xls11 = 'Base//Saldo//Ativo//Visão Telefônica//Saldo_' + str(self.hoje.month).zfill(
                    2) + '_' + str(self.hoje.year) + '.xlsx'
                df_saldo_final.to_excel(xls11, index=False)

                df_saldo_final = self.df_saldo[['CPF', 'Status', self.saldo_cotas_depois]]
                xls11 = 'Base//Saldo//Autopatrocinado//Visão Multi//Saldo_' + str(self.hoje.month).zfill(
                    2) + '_' + str(self.hoje.year) + '.xlsx'
                df_saldo_final.to_excel(xls11, index=False)

            # Abre a janela de exportado com sucesso
            self.primeiro_aviso = 'Arquivo gerado com sucesso!'
            self.segundo_aviso = ''
            self.aviso()

        except Exception:
            # Fecha a janela da data
            self.primeiro_aviso = 'Problemas com a exportação!'
            self.segundo_aviso = ''
            self.aviso()

    def mensagem_data(self):
        # Abre a janela para inserir a data desejada
        self.data_janela = tk.Toplevel()
        p = PhotoImage(file='Base//1.Logo//logo.png')
        self.data_janela.iconphoto(False, p)
        self.data_janela.title("Superavit Visão Prev")
        self.data_janela.config(width=300, height=200)
        self.data_janela.resizable(width=False, height=False)
        # Botões
        botao_data = tk.Button(self.data_janela, text="     Fechar     ", command=self.data_janela.destroy)
        botao_data.place(x=100, y=150)
        # Chama a função calcular tabela usando a data de entrada
        botao_gerar = tk.Button(self.data_janela, text="Gerar Tabela", command=self.calcular_tabela_data)
        botao_gerar.place(x=100, y=110)
        # Caixas de texto
        self.valor_data = tk.Text(self.data_janela, height=1.2, width=15)
        self.valor_data.insert(tk.END, '')
        self.valor_data.place(x=80, y=50)
        # Labels
        label_plano = tk.Label(self.data_janela, text="Data")
        label_plano.config(font=("Courier", 10))
        label_plano.place(x=80, y=25)

    def calcular_tabela_data(self):

        try:
            # Testa se o valor inserido é uma data válida
            data_raw = str.strip(self.valor_data.get('1.0', tk.END))
            self.data = pd.to_datetime(data_raw[:6] + str(pd.to_datetime(data_raw).year), format='%d/%m/%Y')

            # testa se a data inserida está dentro das datas das cotas
            if self.data in list(pd.to_datetime(self.df_cotas['Data'])):

                # Verifica se tem valores nulos no CPF
                if len(self.df_abatimento[self.df_abatimento['CPF'].isnull()]['CPF']) == 0:
                    # Verifica se o CPF é formado por números inteiros
                    if self.df_abatimento['CPF'].dtype == 'int64':
                        # Verifica se tem valores nulos no Status
                        if len(self.df_abatimento[self.df_abatimento['Status'].isnull()]['Status']) == 0:
                            passe = []

                            for i in self.df_abatimento['Status']:
                                # Verifica se algum valor é diferente de "ativos" ou "autopatrocinado" na coluna status
                                if i.lower() == 'ativo':
                                    passe.append(1)
                                else:
                                    passe.append(0)
                            if 0 in passe:
                                self.primeiro_aviso = 'Erro na coluna Status'
                                self.segundo_aviso = 'Contém valores incorretos'
                                self.aviso()
                            else:
                                # Verifica se tem valores nulos no Abatimento
                                if len(self.df_abatimento[
                                           self.df_abatimento[self.abatimento_real].isnull()][
                                           self.abatimento_real]) == 0:
                                    # Verifica se as colunas são números flutuantes

                                    if self.df_abatimento[self.abatimento_real].dtype == 'float64':

                                        # Chama a função gerar tabela com data definida
                                        self.gerar_tabela_data()
                                    else:
                                        self.primeiro_aviso = 'Erro na coluna Abatimento'
                                        self.segundo_aviso = 'Contém valores incorretos'
                                        self.aviso()
                                else:
                                    self.primeiro_aviso = 'Erro na coluna Abatimento'
                                    self.segundo_aviso = 'Contém valores nulos'
                                    self.aviso()
                        else:
                            self.primeiro_aviso = 'Erro na coluna Status'
                            self.segundo_aviso = 'Contém valores nulos'
                            self.aviso()
                    else:
                        self.primeiro_aviso = 'Erro na coluna CPF'
                        self.segundo_aviso = 'Contém valores não numéricos'
                        self.aviso()
                else:
                    self.primeiro_aviso = 'Erro na coluna CPF'
                    self.segundo_aviso = 'Contém valores nulos'
                    self.aviso()
            else:
                self.primeiro_aviso = 'Data fora da tabela cotas!'
                self.segundo_aviso = ''
                self.aviso()
        except Exception:
            self.primeiro_aviso = 'Erro no formato da data!'
            self.segundo_aviso = 'Use: 01/01/2001 ou 01/01/01'
            self.aviso()

    def gerar_tabela_data(self):

        # Remove os CPFs dos participantes ativos e autopatrocinados que precisam serem retirados da tabela saldo
        for i in self.df_retirada['CPF']:
            for j in self.df_saldo.index:
                if self.df_saldo['CPF'][j] == i:
                    self.df_saldo = self.df_saldo.drop(j)
        self.df_saldo.reset_index(drop=True, inplace=True)

        # Cria o cálculo das cotas para a data definida
        cota_abatimento = []
        cota_definida = float(''.join(map(str, self.df_cotas[self.df_cotas['Data'] == self.data]['Cotas'].values)))

        for i in self.df_abatimento.index:
            cota_abatimento.append(self.df_abatimento[self.abatimento_real][i] / cota_definida)
        self.df_abatimento[self.abatimento_cotas] = cota_abatimento

        # Cria a tabela análise
        cota_para_abater = []
        cota_resultado = []
        data_final = []
        abatimento_real = []

        for i in list(self.df_saldo['CPF']):
            if i in list(self.df_abatimento['CPF']):
                abatimento_real.append(pd.Series(self.df_abatimento[
                                                     self.df_abatimento['CPF'] == i][
                                                     self.abatimento_real].values).to_list()[0])
                cota_para_abater.append(pd.Series(self.df_abatimento[
                                                      self.df_abatimento['CPF'] == i][
                                                      self.abatimento_cotas].values).to_list()[0])
                cota_resultado.append(pd.Series(self.df_saldo[self.df_saldo['CPF'] == i][
                                                    self.saldo_cotas_antes].values).to_list()[0] -
                                      pd.Series(self.df_abatimento[
                                                    self.df_abatimento['CPF'] == i][
                                                    self.abatimento_cotas].values).to_list()[0])
                data_final.append(self.data)
            else:
                abatimento_real.append(np.nan)
                cota_para_abater.append(np.nan)
                cota_resultado.append(pd.Series(self.df_saldo[self.df_saldo['CPF'] == i][
                                                    self.saldo_cotas_antes].values).to_list()[0])
                data_final.append(np.nan)

        self.df_saldo[self.abatimento_cotas] = cota_para_abater
        self.df_saldo[self.saldo_cotas_depois] = cota_resultado
        self.df_saldo['Data_cota'] = data_final
        self.df_saldo[self.abatimento_real] = abatimento_real

        saldo_real_antes = []
        saldo_real_depois = []
        cota_usada = []

        for i in list(self.df_saldo['CPF']):
            if i in list(self.df_abatimento['CPF']):
                cota_usada.append(cota_definida)
                saldo_real_antes.append(pd.Series(self.df_saldo[self.df_saldo['CPF'] == i][
                                                  self.saldo_cotas_antes].values).to_list()[0] * cota_definida)
                saldo_real_depois.append(pd.Series(self.df_saldo[self.df_saldo['CPF'] == i][
                                                  self.saldo_cotas_depois].values).to_list()[0] * cota_definida)
            else:
                cota_usada.append(np.nan)
                saldo_real_antes.append(np.nan)
                saldo_real_depois.append(np.nan)

        self.df_saldo[self.saldo_real_antes] = saldo_real_antes
        self.df_saldo[self.saldo_real_depois] = saldo_real_depois
        self.df_saldo['Cota_utilizada'] = cota_usada

        if self.status == 'Ativo':
            analise = self.df_saldo[self.df_saldo['Status'] != 'Autopatrocinado'][['CPF', 'Status',
                                                                                   self.saldo_cotas_antes,
                                                                                   self.saldo_real_antes,
                                                                                   self.abatimento_cotas,
                                                                                   self.abatimento_real,
                                                                                   self.saldo_cotas_depois,
                                                                                   self.saldo_real_depois,
                                                                                   'Cota_utilizada',
                                                                                   'Data_cota']].sort_values(
                    self.saldo_real_antes, ascending=False)
        else:
            analise = self.df_saldo[self.df_saldo['Status'] == 'Autopatrocinado'][['CPF', 'Status',
                                                                                   self.saldo_cotas_antes,
                                                                                   self.saldo_real_antes,
                                                                                   self.abatimento_cotas,
                                                                                   self.abatimento_real,
                                                                                   self.saldo_cotas_depois,
                                                                                   self.saldo_real_depois,
                                                                                   'Cota_utilizada',
                                                                                   'Data_cota']].sort_values(
                    self.saldo_real_antes, ascending=False)

        try:
            # Exporta a tabela análise, saldo e retirada dos ativos e autopatrocinados para o Excel
            xls16 = 'Base//Analise//' + self.status + '//' + self.plano + '//Analise_' + str(self.hoje.month).zfill(
                2) + '_' + str(self.hoje.year) + '.xlsx'
            analise.to_excel(xls16, index=False)
            df_retirar = analise[(analise[self.saldo_cotas_depois] < 0) & (
                ~analise['Data_cota'].isnull())][['CPF', 'Status', self.saldo_cotas_depois]]
            xls18 = 'Base//Retirada//' + self.status + '//' + self.plano + '//Retirada_' + str(
                self.hoje.month).zfill(2) + '_' + str(self.hoje.year) + '.xlsx'
            df_retirar.to_excel(xls18, index=False)


            df_saldo_final = self.df_saldo[['CPF', 'Status', self.saldo_cotas_depois]]
            df_saldo_final = df_saldo_final.rename(columns={self.saldo_cotas_depois: self.saldo_cotas_antes})
            xls11 = 'Base//Saldo//Ativo//Visão Telefônica//Saldo_' + str(self.hoje.month).zfill(
                2) + '_' + str(self.hoje.year) + '.xlsx'
            df_saldo_final.to_excel(xls11, index=False)

            df_saldo_final = self.df_saldo[['CPF', 'Status', self.saldo_cotas_depois]]
            xls11 = 'Base//Saldo//Autopatrocinado//Visão Multi//Saldo_' + str(self.hoje.month).zfill(
                2) + '_' + str(self.hoje.year) + '.xlsx'
            df_saldo_final.to_excel(xls11, index=False)

            # Fecha a janela da data
            self.data_janela.destroy()
            # Gera a janela de exportado com sucesso
            self.primeiro_aviso = 'Arquivo gerado com sucesso!'
            self.segundo_aviso = ''
            self.aviso()

        except Exception:
            # Fecha a janela da data
            self.data_janela.destroy()
            self.primeiro_aviso = 'Problemas com a exportação!'
            self.segundo_aviso = ''
            self.aviso()

    def aviso(self):
        # Janela que gera os avisos
        aviso_janela = tk.Toplevel()
        p = PhotoImage(file='Base//1.Logo//logo.png')
        # Janela
        aviso_janela.iconphoto(False, p)
        aviso_janela.title("Superavit Visão Prev")
        aviso_janela.config(width=300, height=200, bg='#ffffff')
        aviso_janela.resizable(width=False, height=False)
        # Botão
        botao_aviso = tk.Button(aviso_janela, text="Fechar", command=aviso_janela.destroy)
        botao_aviso.place(x=120, y=150)
        botao_aviso.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        # Label
        label_aviso = tk.Label(aviso_janela, text=str(self.primeiro_aviso) + '\n' + str(self.segundo_aviso))
        label_aviso.config(font=("Courier", 10))
        label_aviso.place(x=40, y=60)
        label_aviso.config(bg='#ffffff')


Main()
