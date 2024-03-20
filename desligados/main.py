# pip install openpyxl
# pip install fsspec
# pip install Pyarrow
# pip install pyinstaller
import tkinter as tk
from tkinter import *
import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pickle


class Main:

    def __init__(self):
        self.root = tk.Tk()
        self.root.config(width=230, height=450)
        self.root.resizable(width=False, height=False)
        p = PhotoImage(file='Base//logo.png')
        self.root.iconphoto(False, p)
        self.root.title('Oráculo')
        self.root.config(bg='#ffffff')
        self.hoje = datetime.today()

        # instancias definidas fora de __init__
        self.data_janela = None
        self.primeiro_aviso = None
        self.segundo_aviso = None
        self.df = None

        # Labels
        label = tk.Label(self.root,
                        text=" Previsão das escolhas\ndos participantes demitidos\nbaseado no perfil de cada um.")
        label.config(font=("Arial", 10))
        label.place(x=25, y=25)
        label.config(bg='#ffffff')

        # Botões
        self.botao_gerar = tk.Button(self.root, text="       Gerar Tabela        ", command=self.importar)
        self.botao_gerar.place(x=50, y=360)
        self.botao_gerar.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.botao_sair = tk.Button(self.root, text="               Sair               ", command=self.root.destroy)
        self.botao_sair.place(x=50, y=390)
        self.botao_sair.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")

        # Caixa
        self.caixa = tk.LabelFrame(self.root, text="Status", bd=5, width=90, height=200)
        self.caixa.place(x=30, y=120)
        self.caixa.config(bg='#ffffff')

        # Checkbox
        self.checkbox_todos_var = tk.IntVar()
        self.checkbox_todos = tk.Checkbutton(self.caixa, text='Todos', variable=self.checkbox_todos_var,
                                           command=self.outros)
        self.checkbox_todos.grid(row=0, sticky="W")
        self.checkbox_todos.select()
        self.checkbox_todos.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")

        self.bpd_var = tk.IntVar()
        self.bpd = tk.Checkbutton(self.caixa, text='BPD presumido', variable=self.bpd_var,
                                           command=self.todos)
        self.bpd.grid(row=2, sticky="W")
        self.bpd.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")

        self.demitido_var = tk.IntVar()
        self.demitido = tk.Checkbutton(self.caixa, text='Demitido aguard. opção',
                                               variable=self.demitido_var, command=self.todos)
        self.demitido.grid(row=1, sticky="W")
        self.demitido.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")

        self.root.mainloop()

    def importar(self):
        # importando a tabela da mala direta
        try:
            xls1 = pd.ExcelFile(
                "Y://DIRETORIA_DE_PREVIDENCIA_E_RELACIONAMENTO//CIENCIA_DE_DADOS//Mala_Direta_Atual.xlsx")
            self.df = pd.read_excel(xls1)
            self.transformar()
        except Exception:
            self.primeiro_aviso = 'Erro ao importar a'
            self.segundo_aviso = ' mala direta'
            self.aviso()

    def transformar(self):
        # Removendo as colunas dipensáveis da mala direta
        try:
            self.df.drop(['contratoparticipante', 'Matricula', 'Endereco', 'Bairro', 'Cidade', 'Pais', 'CEP',
                          'Telefone',
                          'TelefoneCelular', 'TelefoneComercial', 'EMailComercial', 'EMailPessoal', 'DestinoEmail',
                          'ParticipanteSA', 'Sms', 'DataInicioBeneficio', 'DataObitoParticipante', 'DataOpcaoIR',
                          'Naturalidade', 'Nacionalidade', 'PoliticamenteExposto', 'NomeEmpregador', 'TipoBeneficio',
                          'DataDemissao', 'Cargo', 'Titular', 'JornalMirante', 'Fundador', 'NomeConjuge',
                          'DataNascimentoConjuge',
                          'NomeMae', 'NomePai', 'Identidade', 'OrgaoExpedidor', 'DataExpedicaoIdentidade', 'Local',
                          'NumeroINSS',
                          'Banco', 'Agencia', 'NumeroContaBancaria', 'DigitoVerificadorContaBancaria', 'FormaPagamento',
                          'GrupoParticipante', 'DataAdmissaoAnterior', 'DataAdesaoAnterior', 'DataInicialBPD',
                          'DataFinalBPD',
                          'DataInicialAutopatrocinio', 'DataFinalAutopatrocinio', 'DepartamentoPessoal', 'CentroCusto',
                          'TipoAfastamento', 'DataAfastamento', 'AssuncaoContratual', 'DataRecadastramento',
                          'InformeRendimento',
                          'ContraCheque', 'Ocupacao', 'MolestiaGrave', 'DataInicioMolestiaGrave',
                          'DataFinalMolestiaGrave',
                          'UFIdentidade', 'CpfResponsavelFinanceiro', 'NomeResponsavelFinanceiro',
                          'DataNascimentoResponsavelFinanceiro', 'EMailResponsavelFinanceiro', 'AnoMes'], axis=1,
                         inplace=True)
            self.df = self.df[(self.df['DescricaoParticipante'] == 'Demitido - Aguardando Opção') |
                              (self.df['DescricaoParticipante'] == 'BPD-Presumido')]
            # Transformando a coluna data de nascimento
            try:
                self.df['DataNascimento'] = pd.to_datetime(self.df['DataNascimento'], format="%d/%m/%Y")
                self.df['Idade'] = self.df['DataNascimento'].apply(lambda x:
                                                                   (datetime.now() - relativedelta(years=x.year)).year)
                # Transfomando a coluna data de adesão
                try:
                    self.df['DataAdesao'] = pd.to_datetime(self.df['DataAdesao'], format="%d/%m/%Y")
                    self.df['Anos_Adesao'] = self.df['DataAdesao'].apply(lambda x: (datetime.now() - relativedelta(
                                                                             years=x.year)).year)
                    # Transformando a coluna data de admissão
                    try:
                        self.df['DataAdmissao'] = pd.to_datetime(self.df['DataAdmissao'], format="%d/%m/%Y")
                        self.df['Anos_Admissao'] = self.df['DataAdmissao'].apply(lambda x:
                                                                                 (datetime.now() - relativedelta(
                                                                                     years=x.year)).year)
                        self.df.drop(['DataNascimento', 'DataAdesao', 'DataAdmissao'], axis=1, inplace=True)
                        # Transformando a coluna estados
                        try:
                            self.df['UF'] = self.df['UF'].apply(lambda x: 'Outros' if x != 'SP' else x)
                            # Transformando a coluna estado civil
                            try:
                                self.df['EstadoCivil'] = self.df['EstadoCivil'].replace('Desquitado', 'Divorciado')
                                self.df['EstadoCivil'] = self.df['EstadoCivil'].replace('Separado', 'Divorciado')
                                self.df['EstadoCivil'] = self.df['EstadoCivil'].replace('Marital', 'UniaoEstavel')
                                self.df['EstadoCivil'] = self.df['EstadoCivil'].replace('Outros', 'UniaoEstavel')
                                self.df['EstadoCivil'] = self.df['EstadoCivil'].replace('Companheiro', 'UniaoEstavel')
                                self.df['EstadoCivil'] = self.df['EstadoCivil'].replace('Amasiado', 'UniaoEstavel')
                                self.df['EstadoCivil'] = self.df['EstadoCivil'].replace('NaoExigido', 'Solteiro')
                                self.df['EstadoCivil'] = self.df['EstadoCivil'].replace('VIUVO', 'Viuvo')
                                # Transformando a coluna perfil de investimento
                                try:
                                    self.df['PerfilInvestimento'] = self.df['PerfilInvestimento'].replace(
                                        'TELEFÔNICA BD', None)
                                    chave = list(self.df['PerfilInvestimento'].value_counts().index)
                                    valor = list((self.df['PerfilInvestimento'].value_counts() / self.df[
                                        'PerfilInvestimento'].count()))
                                    self.df['PerfilInvestimento'] = self.df['PerfilInvestimento'].fillna(
                                        pd.Series(np.random.choice(
                                            chave, p=valor, size=len(self.df)), index=self.df.index))
                                    # Transformando a coluna dependentes
                                    try:
                                        chave = list(self.df['DependentesIRF'].value_counts().index)
                                        valor = list((self.df['DependentesIRF'].value_counts() / self.df[
                                            'DependentesIRF'].count()))
                                        self.df['DependentesIRF'] = self.df['DependentesIRF'].fillna(
                                            pd.Series(np.random.choice(
                                                chave, p=valor, size=len(self.df)), index=self.df.index))
                                        self.df = self.df.rename(columns={'NomePlano': 'Plano'})
                                        self.numeros()
                                    except Exception:
                                        self.primeiro_aviso = 'Erro na coluna'
                                        self.segundo_aviso = ' DependentesIRF'
                                        self.aviso()
                                except Exception:
                                    self.primeiro_aviso = 'Erro na coluna'
                                    self.segundo_aviso = 'PerfilInvestimento'
                                    self.aviso()
                            except Exception:
                                self.primeiro_aviso = 'Erro na coluna'
                                self.segundo_aviso = ' EstadoCivil'
                                self.aviso()
                        except Exception:
                            self.primeiro_aviso = 'Erro na coluna'
                            self.segundo_aviso = '      UF'
                            self.aviso()
                    except Exception:
                        self.primeiro_aviso = 'Erro na coluna'
                        self.segundo_aviso = ' DataAdmissao'
                        self.aviso()
                except Exception:
                    self.primeiro_aviso = 'Erro na coluna'
                    self.segundo_aviso = '   DataAdesao'
                    self.aviso()
            except Exception:
                self.primeiro_aviso = 'Erro na coluna'
                self.segundo_aviso = ' DataNascimento'
                self.aviso()
        except Exception:
            self.primeiro_aviso = 'Colunas faltando na'
            self.segundo_aviso = ' mala direta'
            self.aviso()

    def numeros(self):
        # transformando as variaveis classificatórias de valores strings para valores numéricos
        self.df['TipoVinculo'] = self.df['TipoVinculo'].replace('Participante', 0)
        self.df['TipoVinculo'] = self.df['TipoVinculo'].replace('Beneficiário', 1)
        self.df['UF'] = self.df['UF'].replace('SP', 0)
        self.df['UF'] = self.df['UF'].replace('Outros', 1)
        self.df['OpcaoIR'] = self.df['OpcaoIR'].replace('Regime Progressivo', 0)
        self.df['OpcaoIR'] = self.df['OpcaoIR'].replace('Regime Regressivo', 1)
        self.df['PerfilInvestimento'] = self.df['PerfilInvestimento'].replace('CONSERVADOR', 0)
        self.df['PerfilInvestimento'] = self.df['PerfilInvestimento'].replace('MODERADO', 1)
        self.df['PerfilInvestimento'] = self.df['PerfilInvestimento'].replace('AGRESSIVO', 2)
        self.df['PerfilInvestimento'] = self.df['PerfilInvestimento'].replace('SUPER CONSER', 3)
        self.df['PerfilInvestimento'] = self.df['PerfilInvestimento'].replace('AGRESSIVO RF LP', 4)
        self.df['Sexo'] = self.df['Sexo'].replace('M', 0)
        self.df['Sexo'] = self.df['Sexo'].replace('F', 1)
        self.df['Plano'] = self.df['Plano'].replace('Visão Telefônica', 0)
        self.df['Plano'] = self.df['Plano'].replace('Visão Multi', 1)
        self.df['Plano'] = self.df['Plano'].replace('Telefônica BD', 2)
        self.df['Plano'] = self.df['Plano'].replace('PreVisão', 3)
        self.df['Plano'] = self.df['Plano'].replace('TCOPREV', 4)
        self.df['Plano'] = self.df['Plano'].replace('Mais Visão', 5)
        self.df['EstadoCivil'] = self.df['EstadoCivil'].replace('Casado', 0)
        self.df['EstadoCivil'] = self.df['EstadoCivil'].replace('Solteiro', 1)
        self.df['EstadoCivil'] = self.df['EstadoCivil'].replace('Divorciado', 2)
        self.df['EstadoCivil'] = self.df['EstadoCivil'].replace('Viuvo', 3)
        self.df['EstadoCivil'] = self.df['EstadoCivil'].replace('UniaoEstavel', 4)
        self.df = self.df[['CPF', 'TipoVinculo', 'Nome', 'UF', 'DependentesIRF', 'DescricaoParticipante', 'OpcaoIR',
                           'PerfilInvestimento', 'Sexo', 'Plano', 'EstadoCivil', 'Idade', 'Anos_Adesao',
                           'Anos_Admissao']]
        self.model()

    def model(self):
       try:
            # importando o algoritmo e usando para prever a mala direta
            todos = self.checkbox_todos_var.get()
            bpd_pres = self.bpd_var.get()
            demitido_aop = self.demitido_var.get()
            tabela = self.df.copy()
            if todos == 0:
                if bpd_pres == 1 and demitido_aop == 0:
                    tabela = tabela[tabela['DescricaoParticipante'] == 'BPD-Presumido']
                elif demitido_aop == 1 and bpd_pres == 0:
                    tabela = tabela[tabela['DescricaoParticipante'] == 'Demitido - Aguardando Opção']
                else:
                    pass
            else:
                pass
            tabela = tabela.reset_index(drop=True)
            tabela_prev = tabela.drop(['CPF', 'DescricaoParticipante', 'Nome'], axis=1)
            try:
                nome = 'Base//model.sav'
                model = pickle.load(open(nome, 'rb'))
                prev_final = model.predict(tabela_prev)
                prev_proba_final = model.predict_proba(tabela_prev)
                tabela = tabela[['CPF', 'Nome', 'DescricaoParticipante']]
                tabela['Previsão'] = prev_final
                df_proba = pd.DataFrame(prev_proba_final, columns=['Probabilidade BPD', 'Probabilidade Em Benefício',
                                                                   'Probabilidade Portabilidade',
                                                                   'Probabilidade Resgate'])
                df = pd.concat([tabela, df_proba], axis=1)
                try:
                    if todos == 0:
                        if bpd_pres == 1 and demitido_aop == 0:
                            df.to_excel(
                                'previsão_desligados_bdp_presum_' + str(self.hoje.month).zfill(2) + '_' + str(
                                    self.hoje.year) + '.xlsx', index=False)
                        elif demitido_aop == 1 and bpd_pres == 0:
                            df.to_excel(
                                'previsão_desligados_demitidos_aop_' + str(self.hoje.month).zfill(2) + '_' + str(
                                    self.hoje.year) + '.xlsx', index=False)
                        else:
                            df.to_excel(
                                'previsão_desligados_todos_' + str(self.hoje.month).zfill(2) + '_' + str(
                                    self.hoje.year) + '.xlsx', index=False)
                    else:
                        df.to_excel(
                            'previsão_desligados_todos_' + str(self.hoje.month).zfill(2) + '_' + str(
                                self.hoje.year) + '.xlsx', index=False)
                    self.primeiro_aviso = 'Tabela gerada com'
                    self.segundo_aviso = ' sucesso!'
                    self.aviso()
                except Exception:
                    self.primeiro_aviso = 'Problema ao salvar'
                    self.segundo_aviso = ' o arquivo'
                    self.aviso()
            except Exception:
                self.primeiro_aviso = 'Problema com o'
                self.segundo_aviso = ' algoritmo'
                self.aviso()
       except Exception:
           self.primeiro_aviso = ' Erro na coluna'
           self.segundo_aviso = 'DescricaoParticipante'
           self.aviso()

    def aviso(self):
        # Janela que gera os avisos
        aviso_janela = tk.Toplevel()
        p = PhotoImage(file='Base//logo.png')
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
        label_aviso.place(x=75, y=60)
        label_aviso.config(bg='#ffffff')

    def todos(self):
        # Tira a seleção de todos no status
        self.checkbox_todos.deselect()

    def outros(self):
        # Tira a seleção dos status
        self.bpd.deselect()
        self.demitido.deselect()


Main()
