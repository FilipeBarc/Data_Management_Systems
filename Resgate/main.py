# pip install openpyxl
# pip install fsspec
# pip install Pyarrow
# pip install pyinstaller
# pip install matplotlib
# pip install pypiwin32
# pip install python-pptx
# pip install comtypes

import tkinter as tk
from tkinter import *
import os
import pandas as pd
import matplotlib.pyplot as plt
import comtypes.client
import locale
import time
from datetime import datetime, timedelta
from pptx import Presentation
from pptx.util import Pt, Inches
from dateutil.relativedelta import relativedelta
import threading
import pythoncom
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')


class Main:

    def __init__(self):
        self.root = tk.Tk()
        self.root.config(width=230, height=450)
        self.root.resizable(width=False, height=False)
        p = PhotoImage(file='Base//logo.png')
        self.root.iconphoto(False, p)
        self.root.title('Gerador')
        self.root.config(bg='#ffffff')
        self.hoje = datetime.today()
        pythoncom.CoInitialize()

        # instancias definidas fora de __init__
        self.df = None

        # Caixa
        self.caixa = tk.LabelFrame(self.root, text="Reporte", bd=5, width=90, height=200)
        self.caixa.place(x=55, y=120)
        self.caixa.config(bg='#ffffff')

        # Labels
        label = tk.Label(self.root,
                        text="Gerador de informativo\npersonalizado de resgate")
        label.config(font=("Arial", 10))
        label.place(x=40, y=25)
        label.config(bg='#ffffff')
        self.quantidade = tk.StringVar()
        self.quantidade.set("Quantidade: 0")
        self.label1 = tk.Label(self.caixa, textvariable=self.quantidade)
        self.label1.config(font=("Arial", 10))
        self.label1.grid(row=0, sticky="W")
        self.label1.config(bg='#ffffff')

        self.counter = datetime(2000, 1, 1, 0, 0, 0)
        self.tempo = tk.Label(self.caixa)
        self.tempo.config(font=("Arial", 10))
        self.tempo.grid(row=1, sticky="W")
        self.tempo.config(bg='#ffffff')
        self.string = self.counter.strftime("Tempo: " + "%H:%M:%S")
        self.tempo.config(text=self.string)

        # Botões
        self.botao_gerar = tk.Button(self.root, text="     Gerar Template     ", command=self.funcao)
        self.botao_gerar.place(x=50, y=360)
        self.botao_gerar.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        self.botao_sair = tk.Button(self.root, text="               Sair               ", command=self.root.destroy)
        self.botao_sair.place(x=50, y=390)
        self.botao_sair.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")

        self.root.mainloop()

    def time(self):
        global parar
        self.string = self.counter.strftime("Tempo: " + "%H:%M:%S")
        self.tempo.config(text=self.string)
        parar = self.root.after(1000, self.time)
        self.counter += timedelta(seconds=1)

    def funcao(self):
        self.time()
        threading.Thread(target=self.contador).start()

    def contador(self):
        self.quant = 0
        self.editar()


    def stop(self):
        self.root.after_cancel(parar)

    def editar(self):
        self.df = pd.read_excel('Base//Cópia de Retenção de resgate_022024 (002).xlsx',sheet_name='Todos')
        lista = list(self.df.iloc[5])
        for i in range(0, len(lista)):
            if self.df.iloc[5][i] != self.df.iloc[5][i]:
                lista[i] = list(self.df.iloc[6])[i]
        self.df.columns = lista
        self.df = self.df.drop([0, 1, 2, 3, 4, 5, 6])
        self.df = self.df.reset_index(drop=True)
        self.df['SALDO PATROCINADORA BRUTO'] = self.df['SALDO PATROCINADORA BRUTO'].replace('-', 0)

        self.df = self.df[~self.df['TOTAL BRUTO'].isnull()]
        self.df = self.df[self.df['SOLICITADO VIA'] != 'RESGATE PARCIAL']
        self.df = self.df[self.df['PLANO'] != 'Mais Visão']
        dias = []
        for i in self.df.index:
            novo = relativedelta(pd.to_datetime(self.df['DATA FIM DO PLANO'][i], format='%d/%m/%Y'),
                                 pd.to_datetime(self.df['DATA ADMISSÃO ATUAL'][i], format='%d/%m/%Y'))
            dias.append(novo.years * 12 + novo.months)
        self.df['Dias em Meses'] = dias
        perc = []
        for i in self.df.index:
            if self.df['PLANO'][i] == 'Visão Multi':
                perc.append(60 if self.df['Dias em Meses'][i] * 0.25 >= 60 else self.df['Dias em Meses'][i] * 0.25)
            else:
                perc.append(3 if self.df['Dias em Meses'][i] <= 12 else (
                    6 if self.df['Dias em Meses'][i] > 12 and self.df['Dias em Meses'][i] <= 24 else (
                        9 if self.df['Dias em Meses'][i] > 24 and self.df['Dias em Meses'][i] <= 36 else (
                            12 if self.df['Dias em Meses'][i] > 36 and self.df['Dias em Meses'][i] <= 48 else (
                                60 if self.df['Dias em Meses'][i] > 48 and self.df['Dias em Meses'][i] <= 60 else (
                                    67.5 if self.df['Dias em Meses'][i] > 60 and self.df['Dias em Meses'][
                                        i] <= 72 else (
                                        75 if self.df['Dias em Meses'][i] > 72 and self.df['Dias em Meses'][
                                            i] <= 84 else (
                                            82.5 if self.df['Dias em Meses'][i] > 84 and self.df['Dias em Meses'][
                                                i] <= 96 else (
                                                90 if self.df['Dias em Meses'][i] > 96 else self.df['Dias em Meses'][
                                                    i])))))))))
        self.df['Percentual de resgate'] = perc

        val = []
        for i in self.df.index:
            val.append(self.df['SALDO PATROCINADORA BRUTO'][i] * self.df['Percentual de resgate'][i] / 100)
        self.df['Valor resgatavel'] = val
        bruto = []
        for i in self.df.index:
            bruto.append(self.df['SALDO PARTICIPANTE BRUTO'][i] + self.df['Valor resgatavel'][i])
        self.df['Valor bruto de resgate'] = bruto
        nomes = []
        for i in self.df['PARTICIPANTE'].index:
            nomes.append(self.df['PARTICIPANTE'][i].split(' ')[0].capitalize())
        self.df['nome'] = nomes

        self.df = self.df[['CPF', 'nome', 'PLANO', 'SALDO PARTICIPANTE BRUTO', 'SALDO PATROCINADORA BRUTO',
                           'TOTAL BRUTO', 'Percentual de resgate', 'Valor resgatavel', 'Valor bruto de resgate',
                           'Dias em Meses']]

        self.df = self.df[:5]
        self.gerar()

    def gerar(self):
        template = 'Base//Retenção de resgate.pptx'
        caminho = f'{os.path.abspath("Templates")}\\'
        self.gerador(template, caminho)
        self.primeiro_aviso = 'Comunicados gerados com'
        self.segundo_aviso = '   Sucesso!'
        self.aviso()
        self.stop()

    def convert(self, secs):
        seconds = secs % (24 * 3600)
        hour = seconds // 3600
        seconds %= 3600
        minutes = seconds // 60
        seconds %= 60

        return "%d:%02d:%02d" % (hour, minutes, seconds)

    def manter_formatacao_original(self, paragrafo, fonte_original):
        for run in paragrafo.runs:
            run.font.size = fonte_original.size
            run.font.bold = fonte_original.bold
            run.font.italic = fonte_original.italic
            run.font.color.rgb = fonte_original.color.rgb
            run.font.name = fonte_original.name
            run.font.underline = fonte_original.underline

    def substituir_texto(self, apresentacao, texto_antigo, texto_novo):
        for slide in apresentacao.slides:
            for shape in slide.shapes:
                if hasattr(shape, 'text'):
                    for paragrafo in shape.text_frame.paragraphs:
                        for run in paragrafo.runs:
                            if texto_antigo in run.text:
                                fonte_original = run.font
                                run.text = run.text.replace(texto_antigo, texto_novo)
                                novo_paragrafo = shape.text_frame.add_paragraph()
                                self.manter_formatacao_original(novo_paragrafo, fonte_original)

    def principal(self, CPF, nome, renda, patrocina, valor, percentual, resgate, bruto,
                  ppt, saida, dez, quinze, vinte):
        apresentacao = Presentation(ppt)
        self.substituir_texto(apresentacao, '{nome}', nome)
        self.substituir_texto(apresentacao, '{renda}', renda)
        self.substituir_texto(apresentacao, '{patrocinadora}', patrocina)
        self.substituir_texto(apresentacao, '{valor}', valor)
        self.substituir_texto(apresentacao, '{dez}', dez)
        self.substituir_texto(apresentacao, '{quinze}', quinze)
        self.substituir_texto(apresentacao, '{vintes}', vinte)
        self.substituir_texto(apresentacao, '{percentual}', percentual)
        self.substituir_texto(apresentacao, '{resgate}', resgate)
        self.substituir_texto(apresentacao, '{brutos}', bruto)
        apresentacao.save(f'{saida}{CPF}.pptx')
        self.imagens(f'{saida}{CPF}.pptx', CPF, saida)

    def imagens(self, ppt, CPF, saida):
        prs = Presentation(ppt)
        slide = prs.slides[1]
        img_path = "Templates//foto.png"
        left = Inches(6.5)
        top = Inches(1.5)
        width = Inches(13.5)
        height = Inches(10)
        slide.shapes.add_picture(img_path, left, top, width, height)
        slide = prs.slides[2]
        img_path1 = "Templates//foto1.png"
        left1 = Inches(0.2)
        top1 = Inches(3.4)
        width1 = Inches(8.2)
        height1 = Inches(6)
        img_path2 = "Templates//foto2.png"
        left2 = Inches(5.8)
        img_path3 = "Templates//foto3.png"
        left3 = Inches(11.5)
        slide.shapes.add_picture(img_path1, left1, top1, width1, height1)
        slide.shapes.add_picture(img_path2, left2, top1, width1, height1)
        slide.shapes.add_picture(img_path3, left3, top1, width1, height1)
        prs.save(f'{saida}{CPF}.pptx')
        time.sleep(1)
        self.PPTtoPDF(f'{saida}{CPF}.pptx', f'{saida}{CPF}.pdf', formatType=32)

    def PPTtoPDF(self, inputFileName, outputFileName, formatType=32):
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application",pythoncom.CoInitialize())
        deck = powerpoint.Presentations.Open(inputFileName)
        deck.ExportAsFixedFormat(outputFileName, 32)  # 32 corresponds to PDF format
        deck.Close()
        powerpoint.Quit()

        os.remove(inputFileName)
        os.remove("Templates//foto.png")
        os.remove("Templates//foto1.png")
        os.remove("Templates//foto2.png")
        os.remove("Templates//foto3.png")

    def gerador(self, template_pptx, caminho_base):
        for i in self.df.index:
            CPF = str(self.df['CPF'][i])
            nome = self.df['nome'][i]
            renda = locale.currency(self.df['SALDO PARTICIPANTE BRUTO'][i], symbol=True, grouping=True,
                                    international=False)
            patrocina = locale.currency(self.df['SALDO PATROCINADORA BRUTO'][i], symbol=True, grouping=True,
                                        international=False)
            valor = locale.currency(self.df['TOTAL BRUTO'][i], symbol=True, grouping=True,
                                    international=False)
            percentual = str(self.df['Percentual de resgate'][i]) + '%'

            resgate = locale.currency(self.df['Valor resgatavel'][i], symbol=True, grouping=True,
                                      international=False)
            bruto = locale.currency(self.df['Valor bruto de resgate'][i], symbol=True, grouping=True,
                                    international=False)
            fig = plt.figure()
            ax = fig.add_subplot(111)
            explode = (0.01, 0.01)
            v1 = list(self.df[['SALDO PARTICIPANTE BRUTO', 'SALDO PATROCINADORA BRUTO']].iloc[i])[0]
            v2 = list(self.df[['SALDO PARTICIPANTE BRUTO', 'SALDO PATROCINADORA BRUTO']].iloc[i])[1]
            ax = self.df[['SALDO PARTICIPANTE BRUTO', 'SALDO PATROCINADORA BRUTO']].iloc[i].plot.pie(
                colors=["#58aa7a", "#065d70"], shadow=True, explode=explode, labels=None,
                autopct=lambda p: locale.currency(p * (v1 + v2) / 100, symbol=True, grouping=True,
                                                  international=False), textprops={'color': "w", 'weight': 'bold'})
            ax.yaxis.set_visible(False)
            ax.legend(loc='lower right', fontsize='x-small', labels=['Participante', 'Patrocinadora'])
            plt.savefig('Templates//foto.png', transparent=True)
            plt.close()
            v10 = (v1 * (1 + 0.085) ** 10) - v1
            v15 = (v1 * (1 + 0.085) ** 15) - v1
            v20 = (v1 * (1 + 0.085) ** 20) - v1
            ax = pd.DataFrame([v1, v10])[0].plot.pie(
                colors=["#58aa7a", "#065d70"], shadow=True, explode=explode, labels=None,
                autopct=lambda p: locale.currency(p * (v1 + v10) / 100, symbol=True, grouping=True,
                                                  international=False), textprops={'color': "w", 'weight': 'bold'})
            ax.yaxis.set_visible(False)
            ax.legend(loc='lower right', fontsize='x-small', labels=['Contribuição', 'Rentabilidade'])
            plt.savefig('Templates//foto1.png', transparent=True)
            plt.close()
            ax = pd.DataFrame([v1, v15])[0].plot.pie(
                colors=["#58aa7a", "#065d70"], shadow=True, explode=explode, labels=None,
                autopct=lambda p: locale.currency(p * (v1 + v15) / 100, symbol=True, grouping=True,
                                                  international=False), textprops={'color': "w", 'weight': 'bold'})
            ax.xaxis.set_visible(False)
            ax.legend(loc='lower right', fontsize='x-small', labels=['Contribuição', 'Rentabilidade'])
            plt.savefig('Templates//foto2.png', transparent=True)
            plt.close()
            ax = pd.DataFrame([v1, v20])[0].plot.pie(
                colors=["#58aa7a", "#065d70"], shadow=True, explode=explode, labels=None,
                autopct=lambda p: locale.currency(p * (v1 + v20) / 100, symbol=True, grouping=True,
                                                  international=False), textprops={'color': "w", 'weight': 'bold'})
            ax.yaxis.set_visible(False)
            ax.legend(loc='lower right', fontsize='x-small', labels=['Contribuição', 'Rentabilidade'])
            plt.savefig('Templates//foto3.png', transparent=True)
            plt.close()
            time.sleep(1)
            dez = locale.currency(v1 + v10, symbol=True, grouping=True, international=False)
            quinze = locale.currency(v1 + v15, symbol=True, grouping=True, international=False)
            vinte = locale.currency(v1 + v20, symbol=True, grouping=True, international=False)

            self.principal(CPF, nome, renda, patrocina, valor, percentual, resgate, bruto,
                      template_pptx, caminho_base, dez, quinze, vinte)

            self.quant += 1
            self.quantidade.set(f"Quantidade: {self.quant}")
            time.sleep(2)

    def aviso(self):
        # Janela que gera os avisos
        aviso_janela = tk.Toplevel()
        p = PhotoImage(file='Base//logo.png')
        # Janela
        aviso_janela.iconphoto(False, p)
        aviso_janela.title("Gerador")
        aviso_janela.config(width=300, height=200, bg='#ffffff')
        aviso_janela.resizable(width=False, height=False)
        # Botão
        botao_aviso = tk.Button(aviso_janela, text="Fechar", command=aviso_janela.destroy)
        botao_aviso.place(x=120, y=150)
        botao_aviso.config(bg='#ffffff', activebackground="#e6e6e6", activeforeground="Black")
        # Label
        label_aviso = tk.Label(aviso_janela, text=str(self.primeiro_aviso) + '\n' + str(self.segundo_aviso))
        label_aviso.config(font=("Courier", 10))
        label_aviso.place(x=50, y=60)
        label_aviso.config(bg='#ffffff')


Main()