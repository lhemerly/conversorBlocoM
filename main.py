import pandas as pd
import numpy as np
import warnings
import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
from calendar import monthrange
from datetime import datetime

warnings.simplefilter(action='ignore', category=FutureWarning)

"""
class selectSheet(tk.Frame):

    def __init__(self):
        super().__init__(master)
        self.master = main
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        for sheet in main.m300:
            self.button.append(Button(self, text='Sheet ' + str(i + 1), command=lambda sheet=sheet: self.open_this(sheet)))
            self.button[sheet].grid(column=4, row=i + 1, sticky=W)
"""
# TODO: JANELA DE SELEÇÃO DE SHEET


class Main(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):

        self.winfo_toplevel().title("Conversor Bloco M")

        self.anoLabel = tk.Label(self)
        self.anoLabel["text"] = "Ano"
        self.anoLabel.pack()

        self.anoEntry = tk.Entry(self)
        self.anoEntry.pack()
        self.anoEntry.focus()

        self.m300button = tk.Button(self,
                              text="M300",
                              command=self.importM300)
        self.m300button.pack()

        self.m305button = tk.Button(self,
                              text="M305",
                              command=self.importM305)
        self.m305button.pack()

        self.m310button = tk.Button(self,
                              text="M310",
                              command=self.importM310)
        self.m310button.pack()

        self.converterbutton = tk.Button(self,
                                   text="Converter",
                                   command=self.validador)
        self.converterbutton.pack()

        self.quit = tk.Button(self,
                              text="Sair",
                              fg="red",
                              command=self.master.destroy)
        self.quit.pack(side="bottom")



    def importM300(self):
        #try:
            path = tk.filedialog.askopenfilename(initialdir="/",
                                                 title="Selecione o arquivo do bloco M300",
                                                 filetypes=[("Arquivos Suportados",
                                                            "*.csv *.xls *.xlsx")])  # "*.csv *.xls *.xlsx"

            if "xls" not in path:
                self.m300 = pd.read_csv(path, decimal=',')

            else:
                self.m300 = pd.read_excel(path,
                                          dtype={"Bloco":str,
                                          "Cod":str,
                                          "Descricao":str,
                                          "Tipo": str,
                                          "Relacao": str,
                                          "Valor": float,
                                          "Historico": str,
                                          "Data": str})
                # TODO: Chamada de janela de seleção de sheet e tratamento para DF ser único

            print(self.m300.head())

        #except:
        #    self.errorMsg = tk.messagebox.showerror("Erro",
        #                                            "Não foi possível carregar a tabela")

    def importM305(self):
        try:
            path = tk.filedialog.askopenfilename(initialdir="/",
                                                 title="Selecione o arquivo do bloco M305",
                                                 filetypes=[("Arquivos Suportados",
                                                            "*.csv *.xls *.xlsx")])  # "*.csv *.xls *.xlsx"

            if "xls" not in path:
                self.m305 = pd.read_csv(path, decimal=',')

            else:
                self.m305 = pd.read_excel(path,
                                          dtype={"Bloco":str,
                                          "Conta":str,
                                          "Saldo":float,
                                          "DC":str,
                                          "Lalur":str,
                                          "Data":str})
                # TODO: Chamada de janela de seleção de sheet e tratamento para DF ser único

            print(self.m305.head())

        except:
            self.errorMsg = tk.messagebox.showerror("Erro",
                                                    "Não foi possível carregar a tabela")

    def importM310(self):
        try:
            path = tk.filedialog.askopenfilename(initialdir="/",
                                                 title="Selecione o arquivo do bloco M305",
                                                 filetypes=[("Arquivos Suportados",
                                                            "*.csv *.xls *.xlsx")]) # "*.csv *.xls *.xlsx"

            if "xls" not in path:
                self.m310 = pd.read_csv(path, decimal=',')

            else:
                self.m310 = pd.read_excel(path,
                                          dtype={"Bloco":str,
                                          "Conta":str,
                                          "CC":str,
                                          "Saldo":float,
                                          "DC":str,
                                          "Lalur":str,
                                          "Data":str})
                # TODO: Chamada de janela de seleção de sheet e tratamento para DF ser único

            print(self.m310.head())

        except:
            self.errorMsg = tk.messagebox.showerror("Erro",
                                                    "Não foi possível carregar a tabela")

    def validador(self):
        if len(self.anoEntry.get()) != 4:
            self.errorMsg = tk.messagebox.showerror("Erro Ano",
                                                    "Por favor digite uma data válida(yyyy)")

        if len(self.m300.columns) != 8:
            self.errorMsg = tk.messagebox.showerror("Erro M300",
                                                    "Número de colunas do bloco M300 inválida. Verifique o template")

        if len(self.m305.columns) != 6:
            self.errorMsg = tk.messagebox.showerror("Erro M305",
                                                    "Número de colunas do bloco M305 inválida. Verifique o template")

        if len(self.m310.columns) != 7:
            self.errorMsg = tk.messagebox.showerror("Erro M310",
                                                    "Número de colunas do bloco M310 inválida. Verifique o template")

        self.conversor()

    def conversor(self):

        print("Start: " + datetime.now().strftime("%d-%b-%Y (%H:%M:%S.%f)"))

        # zero to NA and drop
        self.m305 = self.m305.loc[self.m305["Saldo"] != 0].copy()
        self.m310 = self.m310.loc[ self.m310[ "Saldo" ] != 0 ].copy()

        # Standard columns to concatenate
        self.m300.columns = ['Bloco',
                             'field2',
                             'field3',
                             'field4',
                             'field5',
                             'field6',
                             'field7',
                             'field8']

        self.m305.columns = ['Bloco',
                             'field2',
                             'field3',
                             'field4',
                             'field5',
                             'field6']

        self.m310.columns = ['Bloco',
                             'field2',
                             'field3',
                             'field4',
                             'field5',
                             'field6',
                             'field7']

        blocoM = pd.DataFrame(columns=['Bloco',
                                       'field2',
                                       'field3',
                                       'field4',
                                       'field5',
                                       'field6',
                                       'field7',
                                       'field8' ] )

        blocoMes = pd.DataFrame(columns=['Bloco',
                                         'field2',
                                         'field3',
                                         'field4',
                                         'field5',
                                         'field6',
                                         'field7',
                                         'field8'])

        # Iterando meses para o bloco M030
        for month in range(1, 13):

            print("Começando mes: " +str(month))
            if month < 10:
                monthNormal = "0" + str(month)

            m030entry = {'Bloco': "M030",
                         'field2': "0101" + self.anoEntry.get(),
                         'field3': str(monthrange(int(self.anoEntry.get()), month)[1]) + monthNormal + self.anoEntry.get(),
                         'field4': "A" + monthNormal,
                         'field5': "",
                         'field6': "",
                         'field7': "",
                         'field8': ""}

            m030entry = pd.DataFrame(m030entry, index=[0])

            self.m300 = self.m300.astype(str)
            self.m305 = self.m305.astype(str)
            self.m310 = self.m310.astype(str)

            m300mes = self.m300.loc[self.m300['field8'].str.match(m030entry.iloc[0, 2])]

            blocoMes = blocoMes.append(m030entry,
                                       ignore_index=True)

            # Iterando linhas do bloco m300
            for i in range(len(m300mes)):

                m300entry = m300mes.iloc[i, 0:6].copy()
                m300entry = m300entry.astype(str)
                blocoMes = blocoMes.append(m300entry,
                                           ignore_index=True)

                m305entry = self.m305.loc[(self.m305['field6'] == m030entry.iloc[0, 2]) & (self.m305['field5'] == m300entry.iloc[1])].copy()
                m305entry.drop(columns=['field5',
                                        'field6'],
                               inplace=True)

                blocoMes = blocoMes.append(m305entry,
                                           ignore_index=True)

                m310entry = self.m310.loc[(self.m310['field7'] == m030entry.iloc[0, 2]) & (self.m310['field6'] == m300entry.iloc[1])].copy()
                m310entry.drop(columns=['field6',
                                        'field7'],
                               inplace=True)

                blocoMes = blocoMes.append(m310entry,
                                           ignore_index=True)

            blocoM = blocoM.append(blocoMes,
                                   ignore_index=True)

            blocoMes = pd.DataFrame(columns=['Bloco',
                                              'field2',
                                              'field3',
                                              'field4',
                                              'field5',
                                              'field6',
                                              'field7',
                                              'field8'])

        m030entry = {'Bloco': "M030",
                     'field2': "0101" + self.anoEntry.get(),
                     'field3': str(monthrange(int(self.anoEntry.get()), month)[1]) + monthNormal + self.anoEntry.get(),
                     'field4': "A00",
                     'field5': "",
                     'field6': "",
                     'field7': "",
                     'field8': ""}

        m030entry = pd.DataFrame(m030entry, index=[0])

        mAcum = blocoMes.loc[(blocoMes['Bloco'] != "M030")]
        mAcum = m030entry.append(mAcum,
                                 ignore_index=True)
        blocoM = mAcum.append(blocoM,
                              ignore_index=True)

        print("Finish: " + datetime.now().strftime("%d-%b-%Y (%H:%M:%S.%f)"))

        print(blocoM.describe())

        # prepara o dataframe para exportação

        blocoM.insert(0,
                      "field0",
                      "")
        blocoM["lastfield"] = ""


        savepath = tk.filedialog.asksaveasfilename(title="Selecione o diretório para salvar o arquivo",
                                                   filetypes=[("Arquivos Suportados",
                                                               "*.txt")])

        blocoM.to_csv(savepath,
                      sep="|",
                      float_format='%.2f',
                      header=False,
                      index=False,
                      na_rep='',
                      decimal=',')
        # TODO: Não está exportando com , no deciamal nem NA como branco


        savepath = tk.filedialog.asksaveasfilename( title="Selecione o diretório para salvar o arquivo",
                                                        filetypes=[ ("Arquivos Suportados",
                                                                     "*.txt") ] )

        for i in range(len(blocoM)):

            if blocoM.loc[i, 'Bloco'] == "M030" or "M305":

                blocoMEntry = blocoM.loc[i, ['Bloco',
                                             'field2',
                                             'field3',
                                             'field4',
                                             'lastfield']].copy()

            elif blocoM.loc[i, 'Bloco'] == "M300":

                blocoMEntry = blocoM.loc[i, ['Bloco',
                                             'field2',
                                             'field3',
                                             'field4',
                                             'field5',
                                             'field6',
                                             'field7',
                                             'lastfield']].copy()

            else:

                blocoMEntry = blocoM.loc[i, ['Bloco',
                                             'field2',
                                             'field3',
                                             'field4',
                                             'field5',
                                             'lastfield']].copy()

            blocoMEntry = blocoMEntry.to_frame()
            blocoMEntry.to_csv(savepath,
                               sep="|",
                               float_format='%.2f',
                               header=False,
                               index=False,
                               decimal=",",
                               na_rep="",
                               mode='a')

            # TODO: Não está colocando separador. Formato estranho. Escrever um exportador?

root = tk.Tk()
root.iconbitmap(r'C:\Users\luizh\PycharmProjects\blocoM\icons\blocoM.ico')
app = Main(master=root)
app.mainloop()
