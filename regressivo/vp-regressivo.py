# pip install pandas
# pip install openpyxl
# pip install fsspec
# pip install Pyarrow
# pip install requests
# pip install selenium
# pip install msedge-selenium-tools
# pip install pyinstaller

import tkinter as tk
from tkinter import *
import pandas as pd
import numpy as np
import os
import requests
from selenium.webdriver.common.by import By
from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import zipfile
from zipfile import ZipFile
from dateutil.relativedelta import relativedelta
from datetime import datetime, timedelta
import locale
import threading
from io import BytesIO
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')


class Main:

    def __init__(self):
        self.root = tk.Tk()
        self.root.config(width=230, height=450)
        self.root.resizable(width=False, height=False)
        p = PhotoImage(file='Base//logo.png')
        self.root.iconphoto(False, p)
        self.root.title('Regressívo')
        self.root.config(bg='#ffffff')
        self.hoje = datetime.today()

        # instancias definidas fora de __init__
        self.df = None
        self.parar = None
        self.primeiro_aviso = None
        self.segundo_aviso = None
        self.quant = None

        # Caixas
        self.caixa = tk.LabelFrame(self.root, text="Reporte", bd=5, width=20, height=20)
        self.caixa.config(bg='#ffffff')
        self.caixa.place(x=55, y=80)

        # Labels
        label = tk.Label(self.root, text="  Calculadora de cotas\n  do perfil regressívo")
        label.config(font=("Arial", 10), bg='#ffffff')
        label.place(x=40, y=25)
        user_label = tk.Label(self.root, text="Usuário")
        user_label.config(font=("Arial", 10), bg='#ffffff')
        user_label.place(x=50, y=223)
        senha_label = tk.Label(self.root, text="Senha")
        senha_label.config(font=("Arial", 10), bg='#ffffff')
        senha_label.place(x=50, y=268)
        self.quantidade = tk.StringVar()
        self.quantidade.set("Quantidade: 0")
        self.label1 = tk.Label(self.caixa, textvariable=self.quantidade)
        self.label1.config(font=("Arial", 10), bg='#ffffff')
        self.label1.grid(row=0)
        self.counter = datetime(2000, 1, 1, 0, 0, 0)
        self.string = self.counter.strftime("Tempo: " + "%H:%M:%S")
        self.tempo = tk.Label(self.caixa)
        self.tempo.config(font=("Arial", 10), bg='#ffffff', text=self.string)
        self.tempo.grid(row=1, sticky="W")

        # Caixas de texto
        self.user = StringVar()
        self.login = tk.Entry(self.root, bg="#f2f2f2", textvariable=self.user)
        self.login.place(x=50, y=245)
        self.password = StringVar()
        self.senha = tk.Entry(self.root, bg="#f2f2f2", textvariable=self.password)
        self.senha.config(show="*")
        self.senha.place(x=50, y=290)
        self.taxa = StringVar()

        # Botões
        self.root.bind("<Return>", (lambda event: self.funcao()))
        self.botao_gerar = tk.Button(self.root, text="Calcular", command=self.funcao)
        self.botao_gerar.config(width=16, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.botao_gerar.place(x=50, y=330)
        self.botao_reiniciar = tk.Button(self.root, text="Reiniciar", command=self.ativar)
        self.botao_reiniciar.config(width=16, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.botao_reiniciar.place(x=50, y=360)
        self.botao_sair = tk.Button(self.root, text="Sair", command=self.root.destroy)
        self.botao_sair.config(width=16, bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.botao_sair.place(x=50, y=390)

        self.root.mainloop()

    def time(self):
        # Faz o contador rodar
        self.string = self.counter.strftime("Tempo: " + "%H:%M:%S")
        self.tempo.config(text=self.string)
        self.parar = self.root.after(1000, self.time)
        self.counter += timedelta(seconds=1)

    def funcao(self):
        # Verificar se o programa esta rodando pela primeira vez e pede para reiniciar
        if self.quantidade.get() == "Quantidade: 0":
            self.senha.config(state='disabled')
            self.login.config(state='disabled')
            self.time()
            threading.Thread(target=self.contador).start()
        else:
            self.primeiro_aviso = 'Clique no botão Reiniciar'
            self.segundo_aviso = ' para gerar novamente'
            self.aviso()

    def contador(self):
        # Zera o contador e inicia a escolha do caminho
        self.quant = 0
        self.gerar()

    def stop(self):
        # Finaliza o programa
        self.root.after_cancel(self.parar)

    def ativar(self):
        # Reinicia o programa depois de rodar a primeira vez
        self.quantidade.set("Quantidade: 0")
        self.label1.config(text=self.quantidade)
        self.counter = datetime(2000, 1, 1, 0, 0, 0)
        self.string = self.counter.strftime("Tempo: " + "%H:%M:%S")
        self.tempo.config(text=self.string)
        self.senha.config(state='normal')
        self.login.config(state='normal')

    # Seleciona o caminho que o programa deve seguir baseado na escolha do template
    def gerar(self):
        lista = list(pd.read_excel("Base/Participantes.xlsx")['Participantes SA'])
        try:
            path = f'{os.path.abspath("edgedriver")}'
            previous_version = open(os.path.join(path, "edgeversion.txt")).read().strip()
            installed_version = \
            os.popen("reg query HKCU\\Software\\Microsoft\\Edge\\BLBeacon /v version").read().split()[-1]
            if previous_version == installed_version:
                print('ja att')
            else:
                print('precisa att')
                with open(os.path.join(path, "edgeversion.txt"), "w") as version_file:
                    version_file.write(installed_version)
                os.remove(os.path.join(path, "msedgedriver.exe"))
                download = f"https://msedgedriver.azureedge.net/{installed_version}/edgedriver_win64.zip"
                response = requests.get(download)
                zipp = "edgedriver//edgedriver.zip"
                with open(zipp, "wb") as f:
                    f.write(response.content)
                with ZipFile(zipp, "r") as zip_ref:
                    zip_ref.extractall("edgedriver")
                os.remove(zipp)

            try:
                direcao = f'{os.path.abspath("edgedriver")}\\msedgedriver.exe'
                options = EdgeOptions()
                options.use_chromium = True
                #options.add_argument("headless")
                options.add_argument("disable-gpu")
                directory = f'{os.path.abspath("Resultado")}\\'
                if not os.path.exists(directory):
                    os.makedirs(directory)
                options.add_experimental_option("prefs", {"download.default_directory": directory})
                browser = Edge(executable_path=direcao, options=options)
                browser.get('https://ww1.sinqiaprevidencia.com.br/')
                sleep(2)
                username = browser.find_element(By.XPATH, '/html/body/div[1]/div[1]/form[1]/div[3]/input')
                username.send_keys(self.user.get())
                password = browser.find_element(By.XPATH, '/html/body/div[1]/div[1]/form[1]/div[4]/input')
                password.send_keys(self.password.get())
                selector = browser.find_element(By.XPATH, '/html/body/div[1]/div[1]/form[1]/select')
                drop = Select(selector)
                drop.select_by_visible_text("Previdência")
                sleep(1)
                enter = browser.find_element(By.XPATH, '/html/body/div[1]/div[1]/form[1]/table/tbody/tr[3]/td[1]/input')
                enter.click()
                try:
                    incorreto = '/html/body/div[1]/div[1]/form[1]/center/div'
                    acesso = browser.find_element(By.XPATH, incorreto)
                    self.primeiro_aviso = acesso.text
                    self.segundo_aviso = ' '
                    self.aviso()
                    self.stop()
                    sleep(2)
                    browser.quit()
                except Exception:
                    iframe = browser.find_element(By.CSS_SELECTOR, '#Iframe1')
                    browser.switch_to.frame(iframe)
                    lupa = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                    "/html/body/table/tbody/tr/td[4]/table/tbody/tr/td[14]/a/img")))
                    lupa.click()
                    sleep(2)
                    for i in lista:
                        try:
                            browser.switch_to.default_content()
                            iform = browser.find_element(By.CSS_SELECTOR, '#iform')
                            browser.switch_to.frame(iform)
                            psa = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                            '/html/body/form/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td[6]/input')))
                            psa.clear()
                            psa.send_keys(i)
                            search = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                            '/html/body/form/table/tbody/tr[3]/td/table/tbody/tr/td[3]/a/input')))
                            search.click()
                            browser.switch_to.default_content()
                            ibusca = browser.find_element(By.CSS_SELECTOR, '#ibusca')
                            browser.switch_to.frame(ibusca)
                            participante = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                            '/html/body/div[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/a[2]')))
                            participante.click()
                            browser.switch_to.default_content()
                            idetalhe = browser.find_element(By.CSS_SELECTOR, '#iDetalhe')
                            browser.switch_to.frame(idetalhe)
                            cadastro = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                            '/html/body/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr/td[2]/a/img')))
                            cadastro.click()
                            ipcinco = browser.find_element(By.CSS_SELECTOR, '#p5')
                            browser.switch_to.frame(ipcinco)
                            ipquatroa1 = browser.find_element(By.CSS_SELECTOR, '#p4a')
                            browser.switch_to.frame(ipquatroa1)
                            descer = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                            '/html/body/table/tbody/tr[1]/td/table/tbody/tr/td[4]/table/tbody/tr/td[2]/a/img')))
                            descer.click()
                            sleep(4)
                            plano = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                            '/html/body/div/table/tbody/tr/td[1]/table/tbody/tr[1]/td[2]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]')))
                            plano_nome = plano.text
                            data_admissao = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                            '/html/body/div/table/tbody/tr/td[1]/table/tbody/tr[1]/td[2]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[1]/table/tbody/tr[2]/td[2]')))
                            data_admissao_nome = data_admissao.text
                            data_desligamento = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                            '/html/body/div/table/tbody/tr/td[1]/table/tbody/tr[1]/td[2]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[1]/table/tbody/tr[3]/td[2]')))
                            data_desligamento_nome = data_desligamento.text
                            browser.switch_to.default_content()
                            idetalhe = browser.find_element(By.CSS_SELECTOR, '#iDetalhe')
                            browser.switch_to.frame(idetalhe)
                            reserva = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                            '/html/body/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr/td[7]/a/img')))
                            reserva.click()
                            perfil = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                            '/html/body/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr/td[8]/table/tbody/tr[1]/td/table/tbody/tr[2]/td[2]/a/img')))
                            perfil.click()
                            ipcinco = browser.find_element(By.CSS_SELECTOR, '#p5')
                            browser.switch_to.frame(ipcinco)
                            ipquatrob = browser.find_element(By.CSS_SELECTOR, '#p4b')
                            browser.switch_to.frame(ipquatrob)
                            tabela = browser.find_elements(By.XPATH,
                            '/html/body/div/div/table/tbody/tr/td[2]/table/tbody/tr/td/a')
                            counter = len(tabela) + 1
                            for j in range(1, counter):
                                tipo = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                                f'/html/body/div/div/table/tbody/tr/td[2]/table/tbody/tr[{j}]/td/a')))
                                tipo_nome = tipo.text
                                tipo.click()
                                browser.switch_to.default_content()
                                idetalhe = browser.find_element(By.CSS_SELECTOR, '#iDetalhe')
                                browser.switch_to.frame(idetalhe)
                                saldo = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                                '/html/body/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr/td[8]/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/a/img')))
                                saldo.click()
                                ipcinco = browser.find_element(By.CSS_SELECTOR, '#p5')
                                browser.switch_to.frame(ipcinco)
                                ipquatroa1 = browser.find_element(By.CSS_SELECTOR, '#p4a')
                                browser.switch_to.frame(ipquatroa1)
                                tamanho = len(os.listdir(directory))
                                download = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                                '/html/body/table/tbody/tr[2]/td/table/tbody/tr/td[2]/div[2]/table/tbody/tr/td[2]/a/img')))
                                download.click()
                                while True:
                                    arq = os.listdir(directory)
                                    arquivos = len(list((pal for pal in arq if '.xlsx' or '.xls' in pal)))
                                    if tamanho < arquivos:
                                        sleep(2)
                                        files = os.listdir(directory)
                                        files.sort(key=lambda x: os.path.getctime(os.path.join(directory, x)))
                                        last_file = files[-1]
                                        original = directory + '\\' + last_file
                                        new = directory + f'\\{i}_{j}.xlsx'
                                        os.rename(original, new)
                                        break
                                df = pd.read_excel(directory + f'\\{i}_{j}.xlsx')
                                df['perfil'] = tipo_nome
                                df['plano'] = plano_nome
                                df['data_admissao'] = data_admissao_nome
                                df['data_desligamento'] = data_desligamento_nome
                                df.to_excel(directory + f'\\{i}_{j}.xlsx', index=False)
                                browser.switch_to.default_content()
                                idetalhe = browser.find_element(By.CSS_SELECTOR, '#iDetalhe')
                                browser.switch_to.frame(idetalhe)
                                perfil = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH,
                                '/html/body/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr/td[8]/table/tbody/tr[1]/td/table/tbody/tr[2]/td[2]/a/img')))
                                perfil.click()
                                ipcinco = browser.find_element(By.CSS_SELECTOR, '#p5')
                                browser.switch_to.frame(ipcinco)
                                ipquatrob = browser.find_element(By.CSS_SELECTOR, '#p4b')
                                browser.switch_to.frame(ipquatrob)
                            df_total = []
                            for x in range(1, counter):
                                df = pd.read_excel(directory + f'\\{i}_{x}.xlsx')
                                df = df[1:]
                                df.rename(columns={'Saldo Participante': 'Saldo Participante Real',
                                                   'Unnamed: 3': 'Saldo Participante Qtd. (Cotas)',
                                                   'Saldo Patrocinadora': 'Saldo Patrocinadora Real',
                                                   'Unnamed: 5': 'Saldo Patrocinadora Qtd. (Cotas)',
                                                   'Saldo Total': 'Saldo Total Real',
                                                   'Unnamed: 7': 'Saldo Total Qtd. (Cotas)'}, inplace=True)
                                df_total.append(df)
                                os.remove(directory + f'\\{i}_{x}.xlsx')
                            df_total = pd.concat(df_total)
                            df_total = df_total.sort_values(['Data', 'Valor da Cota'])
                            df_total['Data'] = pd.to_datetime(df_total['Data'], format="%d/%m/%Y")
                            df_total = df_total.sort_values(['Data', 'Valor da Cota'])
                            df_total = df_total.reset_index(drop=True)
                            df_total['data_admissao'] = pd.to_datetime(df_total['data_admissao'], format="%d/%m/%Y")
                            df_total['data_desligamento'] = pd.to_datetime(df_total['data_desligamento'], format="%d/%m/%Y")

                            df_base = pd.read_excel("Base/Participantes.xlsx")
                            status = df_base[df_base['Participantes SA'] == i]['Status'].values[0]
                            if status == 'AOP':
                                dias = []
                                for x in df_total.index:
                                    novo = relativedelta(df_total['data_desligamento'][x], df_total['data_admissao'][x])
                                    dias.append(novo.years * 12 + novo.months)
                                df_total['meses'] = dias
                            else:
                                dias = []
                                if datetime.today().month == 1:
                                    anterior = f'25/12/{datetime.today().year - 1}'
                                else:
                                    anterior = f'25/{str(datetime.today().month - 1).zfill(2)}/{datetime.today().year}'

                                for a in df_total.index:
                                    novo = relativedelta(pd.to_datetime(anterior, format='%d/%m/%Y'),
                                                         pd.to_datetime(df_total['data_admissao'][a],
                                                                        format='%d/%m/%Y'))
                                    dias.append(novo.years * 12 + novo.months)
                                df_total['meses'] = dias

                            perc = []
                            for x in df_total.index:
                                if df_total['plano'][x] == 'VISAO MULTI':
                                    perc.append(
                                        60 if df_total['meses'][x] * 0.25 >= 60 else df_total['meses'][x] * 0.25)
                                else:
                                    perc.append(3 if df_total['meses'][x] <= 12 else (
                                        6 if df_total['meses'][x] > 12 and df_total['meses'][x] <= 24 else (
                                            9 if df_total['meses'][x] > 24 and df_total['meses'][x] <= 36 else (
                                                12 if df_total['meses'][x] > 36 and df_total['meses'][x] <= 48 else (
                                                    60 if df_total['meses'][x] > 48 and df_total['meses'][
                                                        x] <= 60 else (
                                                        67.5 if df_total['meses'][x] > 60 and df_total['meses'][
                                                            x] <= 72 else (
                                                            75 if df_total['meses'][x] > 72 and df_total['meses'][
                                                                x] <= 84 else (
                                                                82.5 if df_total['meses'][x] > 84 and df_total['meses'][
                                                                    x] <= 96 else (
                                                                    90 if df_total['meses'][x] > 96 else
                                                                    df_total['meses'][x])))))))))
                            df_total['percentual'] = perc
                            print(list(df_total['Saldo Patrocinadora Qtd. (Cotas)']))
                            print(perc)

                            cota_nova = []
                            for x in df_total.index:
                                cota_nova.append(
                                    float(''.join(
                                        ['.' if i == ',' else i for i in list(str(df_total['Valor da Cota'][x]))])))
                            df_total['Valor da Cota'] = cota_nova
                            print(cota_nova)

                            lista = []
                            for x in df_total.index:
                                lista.append(
                                    df_total['Saldo Patrocinadora Qtd. (Cotas)'][x] * df_total['percentual'][x] / 100)
                            print(lista)

                            lista2 = []
                            for x in df_total['Saldo Participante Qtd. (Cotas)'].index:
                                lista2.append(df_total['Saldo Participante Qtd. (Cotas)'][x] + lista[x])
                            print(lista2)

                            lista3 = []
                            for x in range(0, len(lista2)):
                                if x == 0:
                                    lista3.append(round(lista2[x], 4))
                                else:
                                    if lista2[x - 1] == 0:
                                        lista3.append(round(lista2[x] - lista2[x - 2], 4))
                                    elif lista2[x] == 0:
                                        lista3.append(0)
                                    else:
                                        lista3.append(round(lista2[x] - lista2[x - 1], 4))
                            df_total['part_cotas'] = lista3
                            print(lista3)

                            cotas_mudanca = []
                            soma_cotas = []
                            df_total = df_total.sort_values(['Data', 'Saldo Participante Real'])
                            df_total = df_total.reset_index(drop=True)
                            if len(df_total['perfil'].unique()) == 1:
                                resultado = sum(df_total['part_cotas'])
                                penultima = df_total['Valor da Cota'].iloc[-1]
                                segunda = pd.DataFrame([round(resultado, 2),
                                                        locale.currency(round(resultado * penultima, 2), symbol=True,
                                                                        grouping=True,
                                                                        international=False)], index=['Cotas', 'Real'],
                                                       columns=['Valor'])

                                tempo = []
                                for y in df_total['Data']:
                                    tempo.append(relativedelta(datetime.today(), y).years)
                                df_total['tempo em anos'] = tempo

                                aliquota = []
                                for k in df_total['tempo em anos']:
                                    if k >= 10:
                                        aliquota.append('10%')
                                    elif k >= 8 and k < 10:
                                        aliquota.append('15%')
                                    elif k >= 6 and k < 8:
                                        aliquota.append('20%')
                                    elif k >= 4 and k < 6:
                                        aliquota.append('25%')
                                    elif k >= 2 and k < 4:
                                        aliquota.append('30%')
                                    else:
                                        aliquota.append('35%')
                                df_total['aliquota'] = aliquota

                                come_cota = []
                                come = abs(sum(df_total[df_total['part_cotas'] < 0]['part_cotas']))
                                for e in df_total['part_cotas']:
                                    if come != 0:
                                        if e >= come:
                                            come_cota.append(e - come)
                                            come = 0
                                        else:
                                            come_cota.append(-e)
                                            come = come - e
                                    else:
                                        come_cota.append(e)
                                df_total['come_cota'] = come_cota
                                dez = sum(df_total[(df_total['aliquota'] == '10%') & (df_total['come_cota'] > 0)]['come_cota'])
                                quinze = sum(
                                    df_total[(df_total['aliquota'] == '15%') & (df_total['come_cota'] > 0)]['come_cota'])
                                vinte = sum(df_total[(df_total['aliquota'] == '20%') & (df_total['come_cota'] > 0)]['come_cota'])
                                vintecinco = sum(
                                    df_total[(df_total['aliquota'] == '25%') & (df_total['come_cota'] > 0)]['come_cota'])
                                trinta = sum(df_total[(df_total['aliquota'] == '30%') & (df_total['come_cota'] > 0)]['come_cota'])
                                trintacinco = sum(
                                    df_total[(df_total['aliquota'] == '35%') & (df_total['come_cota'] > 0)]['come_cota'])
                                quarta = pd.DataFrame(['10%', '15%', '20%', '25%', '30%', '35%'], columns=['Aliquota'])
                                quarta['Cotas'] = [dez, quinze, vinte, vintecinco, trinta, trintacinco]
                                quarta['Saldo'] = [dez * penultima, quinze * penultima, vinte * penultima,
                                                   vintecinco * penultima,
                                                   trinta * penultima, trintacinco * penultima]
                                quarta['IR'] = [dez * penultima * 0.1, quinze * penultima * 0.15, vinte * penultima * 0.2,
                                                vintecinco * penultima * 0.25, trinta * penultima * 0.3,
                                                trintacinco * penultima * 0.35]
                                result = pd.DataFrame([np.nan, 'Total', sum(quarta['Saldo']), sum(quarta['IR'])],
                                                      index=['Aliquota', 'Cotas', 'Saldo', 'IR'], columns=[6]).transpose()
                                quarta = pd.concat([quarta, result])
                                with pd.ExcelWriter(directory + f'\\calculo_final_{i}.xlsx', engine='xlsxwriter') as writer:
                                    df_total.to_excel(writer, sheet_name='Histórico', index=False)
                                    segunda.to_excel(writer, sheet_name='Resultado')
                                    quarta.to_excel(writer, sheet_name='IR Regressívo', index=False)
                                self.quant += 1
                                self.quantidade.set(f"Quantidade: {self.quant}")

                            else:
                                for x in df_total.index:
                                    if x != df_total.index[-1] and x != df_total.index[0]:
                                        if df_total['perfil'][x] != df_total['perfil'][x + 1] and \
                                                df_total['Saldo Participante Real'][x] == 0:
                                            cotas_mudanca.append(round(df_total['Valor da Cota'][x], 4))
                                            cotas_mudanca.append(round(df_total['Valor da Cota'][x + 1], 4))
                                            df_total.loc[x, 'part_cotas'] = sum(soma_cotas)
                                            soma_cotas = []
                                        elif df_total['perfil'][x] != df_total['perfil'][x - 1] and \
                                                df_total['Saldo Participante Real'][x - 1] == 0:
                                            pass
                                        else:
                                            cotas_mudanca.append(np.nan)
                                            soma_cotas.append(round(df_total['part_cotas'][x], 4))
                                    else:
                                        cotas_mudanca.append(np.nan)
                                        soma_cotas.append(round(df_total['part_cotas'][x], 4))
                                df_total['cotas_mudanca'] = cotas_mudanca
                                var = np.nan
                                for x in df_total.index:
                                    if x != df_total.index[-1] and x != df_total.index[0]:
                                        if (df_total['Data'][x] == df_total['Data'][x + 1] and
                                                df_total['perfil'][x] == df_total['perfil'][x + 1] and
                                                df_total['Saldo Participante Real'][x] == 0 or
                                                df_total['Data'][x] == df_total['Data'][x - 1] and
                                                df_total['perfil'][x] == df_total['perfil'][x - 1] and
                                                df_total['Saldo Participante Real'][x] == 0):
                                            var = x
                                df_total = df_total[df_total.index != var].reset_index(drop=True)
                                tag = 0
                                cot_mudanca = list(df_total[
                                    ~df_total['cotas_mudanca'].isnull()].reset_index(drop=True)['cotas_mudanca'].values)
                                v1 = 0
                                v2 = 1
                                prim = 0
                                sec = 0
                                inde = 0
                                mass = 0
                                posic = 0
                                value = 0
                                coluna_final = 0
                                resultado = 0
                                nova_tx = []
                                mist = []
                                calculo_tx = []
                                for x in df_total.index:
                                    if v2 <= len(cot_mudanca):
                                        if x != df_total.index[0] and x != df_total.index[-1]:
                                            if df_total['perfil'][x] != df_total['perfil'][x + 1] and \
                                                    df_total['Saldo Participante Real'][x] == 0:
                                                if prim == 0:
                                                    value = round(df_total['part_cotas'][x + 1], 2)
                                                    mist.append(value)
                                                    df_total.loc[x, 'part_cotas'] = sum(df_total['part_cotas'][:x])
                                                    df_total.loc[x + 1, 'part_cotas'] = sum(nova_tx)
                                                    nova_tx.append(sum(nova_tx))
                                                    calculo_tx = nova_tx.copy()
                                                    contador = list(range(0, len(calculo_tx)))
                                                    posic = 2
                                                    for m in contador:
                                                        v11 = 2
                                                        v22 = 3
                                                        for n in range(posic, len(cot_mudanca), 2):
                                                            calculo_tx[m] = calculo_tx[m] * cot_mudanca[v11] / \
                                                                            cot_mudanca[v22]
                                                            v11 = v11 + 2
                                                            v22 = v22 + 2
                                                        if m == contador[-1]:
                                                            calculo_tx[m] = 0
                                                            calculo_tx.append(0)
                                                    posic += 2
                                                    prim = 1
                                                    inde = x + 1
                                                    v1 = v1 + 2
                                                    v2 = v2 + 2
                                                else:
                                                    value = round(df_total['part_cotas'][x + 1], 2)
                                                    mist.append(value)
                                                    df_total.loc[x, 'part_cotas'] = (
                                                                sum(df_total['part_cotas'][inde:inde + len(
                                                                    nova_tx[inde:])]) + df_total['part_cotas'][
                                                                    inde - 1] -
                                                                df_total['part_cotas'][inde] + mist[tag])
                                                    df_total.loc[x + 1, 'part_cotas'] = (sum(nova_tx[mass:]) +
                                                                                         (df_total['part_cotas'][
                                                                                              inde - 1] -
                                                                                          df_total['part_cotas'][inde] +
                                                                                          mist[tag]) * cot_mudanca[v1] /
                                                                                         cot_mudanca[v2])
                                                    nova_tx.append(
                                                        sum(nova_tx[mass:]) + (df_total['part_cotas'][inde - 1] -
                                                                               df_total['part_cotas'][inde] + mist[
                                                                                   tag]) * cot_mudanca[v1] /
                                                        cot_mudanca[v2])
                                                    proxima_tx = ([(df_total['part_cotas'][inde - 1] -
                                                                    df_total['part_cotas'][inde] + mist[tag]) *
                                                                   cot_mudanca[v1] /
                                                                   cot_mudanca[v2]] + nova_tx[mass:])
                                                    contador = list(range(0, len(proxima_tx)))
                                                    for m in contador:
                                                        v11 = v1 + 2
                                                        v22 = v2 + 2
                                                        for n in range(posic, len(cot_mudanca), 2):
                                                            proxima_tx[m] = proxima_tx[m] * cot_mudanca[v11] / \
                                                                            cot_mudanca[v22]
                                                            v11 = v11 + 2
                                                            v22 = v22 + 2
                                                        if m == contador[-1]:
                                                            proxima_tx[m] = 0
                                                            proxima_tx.append(0)
                                                    posic += 2
                                                    calculo_tx = (calculo_tx + proxima_tx).copy()
                                                    tag += 1
                                                    inde = x + 1
                                                    v1 = v1 + 2
                                                    v2 = v2 + 2
                                                    mass = x - 1
                                            elif df_total['perfil'][x - 1] != df_total['perfil'][x] and \
                                                    df_total['Saldo Participante Real'][x - 1] == 0:
                                                nova_tx.append(value)
                                                mass = x + 1
                                            else:
                                                nova_tx.append(
                                                    df_total['part_cotas'][x] * cot_mudanca[v1] / cot_mudanca[v2])
                                        else:
                                            nova_tx.append(
                                                df_total['part_cotas'][x] * cot_mudanca[v1] / cot_mudanca[v2])
                                    else:
                                        if sec == 0:
                                            df_total.loc[x, 'part_cotas'] = sum(calculo_tx)
                                            resultado = sum(calculo_tx + [df_total['part_cotas'][x - 1] -
                                                                          df_total['part_cotas'][x] + mist[tag]] + list(
                                                df_total['part_cotas'][x + 1:]))
                                            coluna_final = list(
                                                calculo_tx + [
                                                    df_total['part_cotas'][x - 1] - df_total['part_cotas'][x] + mist[
                                                        tag]] +
                                                list(df_total['part_cotas'][x + 1:]))
                                            sec = 1
                                final_coluna = []
                                for z in range(0, len(coluna_final)):
                                    if coluna_final[z - 1] == 0 and coluna_final[z] == 0:
                                        pass
                                    else:
                                        final_coluna.append(coluna_final[z])
                                penultima = df_total['Valor da Cota'].iloc[-1]
                                segunda = pd.DataFrame([list(df_total['Data']), final_coluna],
                                                       index=['Data', 'Cotas Calculadas']).transpose()
                                terceira = pd.DataFrame([round(resultado, 2), locale.currency(round(resultado *
                                penultima, 2), symbol=True, grouping=True, international=False)], index=['Cotas', 'Real'],
                                columns=['Valor'])
                                tempo = []
                                for y in segunda['Data']:
                                    tempo.append(relativedelta(datetime.today(), y).years)
                                segunda['tempo em anos'] = tempo
                                aliquota = []
                                for k in segunda['tempo em anos']:
                                    if k >= 10:
                                        aliquota.append('10%')
                                    elif k >= 8 and k < 10:
                                        aliquota.append('15%')
                                    elif k >= 6 and k < 8:
                                        aliquota.append('20%')
                                    elif k >= 4 and k < 6:
                                        aliquota.append('25%')
                                    elif k >= 2 and k < 4:
                                        aliquota.append('30%')
                                    else:
                                        aliquota.append('35%')
                                segunda['aliquota'] = aliquota
                                come_cota = []
                                come = abs(sum(segunda[segunda['Cotas Calculadas'] < 0]['Cotas Calculadas']))
                                for e in segunda['Cotas Calculadas']:
                                    if come != 0:
                                        if e >= come:
                                            come_cota.append(e - come)
                                            come = 0
                                        else:
                                            come_cota.append(-e)
                                            come = come - e
                                    else:
                                        come_cota.append(e)
                                segunda['come_cota'] = come_cota
                                dez = sum(segunda[(segunda['aliquota'] == '10%') & (segunda['come_cota'] > 0)]['come_cota'])
                                quinze = sum(segunda[(segunda['aliquota'] == '15%') & (segunda['come_cota'] > 0)]['come_cota'])
                                vinte = sum(segunda[(segunda['aliquota'] == '20%') & (segunda['come_cota'] > 0)]['come_cota'])
                                vintecinco = sum(
                                    segunda[(segunda['aliquota'] == '25%') & (segunda['come_cota'] > 0)]['come_cota'])
                                trinta = sum(segunda[(segunda['aliquota'] == '30%') & (segunda['come_cota'] > 0)]['come_cota'])
                                trintacinco = sum(
                                    segunda[(segunda['aliquota'] == '35%') & (segunda['come_cota'] > 0)]['come_cota'])
                                quarta = pd.DataFrame(['10%', '15%', '20%', '25%', '30%', '35%'], columns=['Aliquota'])
                                quarta['Cotas'] = [dez, quinze, vinte, vintecinco, trinta, trintacinco]
                                quarta['Saldo'] = [dez * penultima, quinze * penultima, vinte * penultima,
                                                   vintecinco * penultima,
                                                   trinta * penultima, trintacinco * penultima]
                                quarta['IR'] = [dez * penultima * 0.1, quinze * penultima * 0.15, vinte * penultima * 0.2,
                                                vintecinco * penultima * 0.25, trinta * penultima * 0.3,
                                                trintacinco * penultima * 0.35]
                                result = pd.DataFrame([np.nan, 'Total', sum(quarta['Saldo']), sum(quarta['IR'])],
                                                      index=['Aliquota', 'Cotas', 'Saldo', 'IR'], columns=[6]).transpose()
                                quarta = pd.concat([quarta, result])
                                with pd.ExcelWriter(directory + f'\\calculo_final_{i}.xlsx', engine='xlsxwriter') as writer:
                                    df_total.to_excel(writer, sheet_name='Histórico', index=False)
                                    segunda.to_excel(writer, sheet_name='Cotas Calculadas', index=False)
                                    terceira.to_excel(writer, sheet_name='Resultado')
                                    quarta.to_excel(writer, sheet_name='IR Regressívo', index=False)
                                self.quant += 1
                                self.quantidade.set(f"Quantidade: {self.quant}")
                        except Exception:
                            pass
                    browser.switch_to.default_content()
                    self.primeiro_aviso = 'Arquivos calculados com'
                    self.segundo_aviso = ' sucesso!'
                    self.aviso()
                    self.stop()
                    sleep(1)
                    browser.quit()
            except Exception:
                self.primeiro_aviso = '  Erro no programa'
                self.segundo_aviso = ' '
                self.aviso()
                self.stop()
        except Exception:
            self.primeiro_aviso = ' Erro no download do'
            self.segundo_aviso = '  driver'
            self.aviso()
            self.stop()

    def aviso(self):
        # Janela que gera os avisos
        aviso_janela = tk.Toplevel()
        p = PhotoImage(file='Base//logo.png')

        # Janela
        aviso_janela.iconphoto(False, p)
        aviso_janela.title("Regressívo")
        aviso_janela.config(width=300, height=200, bg='#ffffff')
        aviso_janela.resizable(width=False, height=False)

        # Botão
        botao_aviso = tk.Button(aviso_janela, text="Fechar", command=aviso_janela.destroy)
        botao_aviso.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black", width=10)
        botao_aviso.place(x=105, y=150)

        # Label
        label_aviso = tk.Label(aviso_janela, text=str(self.primeiro_aviso) + '\n' + str(self.segundo_aviso))
        label_aviso.config(font=("Courier", 10), bg='#ffffff')
        label_aviso.place(x=50, y=60)


Main()
