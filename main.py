import pandas as pd
import numpy as np
import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
from calendar import monthrange

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



class main(tk.Frame):

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

        self.m300 = tk.Button(self,
                              text="M300",
                              command=self.importM300)
        self.m300.pack()

        self.m305 = tk.Button(self,
                              text="M305",
                              command=self.importM305)
        self.m305.pack()

        self.m310 = tk.Button(self,
                              text="M310",
                              command=self.importM310)
        self.m310.pack()

        self.quit = tk.Button(self,
                              text="Sair",
                              fg="red",
                              command=self.master.destroy)
        self.quit.pack(side="bottom")

    def importM300(self):
        try:
            path = tk.filedialog.askopenfilename(initialdir="/",
                                                 title="Selecione o arquivo do bloco M300",
                                                 filetypes=[("Arquivos Suportados",
                                                            "*.csv")])  # "*.csv *.xls *.xlsx"

            if "xls" not in path:
                self.m300 = pd.read_csv(path)

            else:
                self.m300 = pd.read_excel(path, None)
                # TODO: Chamada de janela de seleção de sheet e tratamento para DF ser único

        except:
            self.errorMsg = tk.messagebox.showerror("Erro","Não foi possível carregar a tabela")

    def importM305(self):
        try:
            path = tk.filedialog.askopenfilename(initialdir="/",
                                                 title="Selecione o arquivo do bloco M305",
                                                 filetypes=[("Arquivos Suportados",
                                                            "*.csv")])  # "*.csv *.xls *.xlsx"

            if "xls" not in path:
                self.m305 = pd.read_csv(path)

            else:
                self.m305 = pd.read_excel(path, None)
                # TODO: Chamada de janela de seleção de sheet e tratamento para DF ser único

        except:
            self.errorMsg = tk.messagebox.showerror("Erro","Não foi possível carregar a tabela")

    def importM310(self):
        try:
            path = tk.filedialog.askopenfilename(initialdir="/",
                                                 title="Selecione o arquivo do bloco M305",
                                                 filetypes=[("Arquivos Suportados",
                                                            "*.csv")]) # "*.csv *.xls *.xlsx"

            if "xls" not in path:
                self.m310 = pd.read_csv(path)

            else:
                self.m310 = pd.read_excel(path, None)
                # TODO: Chamada de janela de seleção de sheet e tratamento para DF ser único

        except:
            self.errorMsg = tk.messagebox.showerror("Erro","Não foi possível carregar a tabela")

    def validador(self):
        if len(self.anoLabel.get()) != 4 :
            self.errorMsg = tk.messagebox.showerror("Erro Ano", "Por favor digite uma data válida(yyyy)")

        if len(self.m300.columns) != 8 :
            self.errorMsg = tk.messagebox.showerror("Erro M300", "Número de colunas do bloco M300 inválida. Verifique o template")

        if len(self.m305.columns) != 6 :
            self.errorMsg = tk.messagebox.showerror("Erro M305", "Número de colunas do bloco M305 inválida. Verifique o template")

        if len(self.m310.columns) != 7 :
            self.errorMsg = tk.messagebox.showerror("Erro M310", "Número de colunas do bloco M310 inválida. Verifique o template")

    def conversor(self):

        # zero to NA and drop
        self.m300.replace(0, np.nan)
        self.m300.dropna(inplace=True)
        self.m305.replace(0, np.nan)
        self.m305.dropna(inplace=True)
        self.m310.replace(0, np.nan)
        self.m310.dropna(inplace=True)

        # Standard columns to concatenate
        self.m300.rename(index={0:'Bloco',
                                1:'field2',
                                2:'field3',
                                3:'field4',
                                4:'field5',
                                5:'field6',
                                6:'field7',
                                7:'field8'})

        self.m305.rename(index={0:'Bloco',
                                1:'field2',
                                2:'field3',
                                3:'field4',
                                4:'field5',
                                5:'field6'})

        self.m310.rename(index={0:'Bloco',
                                1:'field2',
                                2:'field3',
                                3:'field4',
                                4:'field5',
                                5:'field6',
                                6:'field7'})

        blocoM = pd.DataFrame(data={'Bloco',
                                    'field2',
                                    'field3',
                                    'field4',
                                    'field5',
                                    'field6',
                                    'field7',
                                    'field8'})

        blocoMes = pd.DataFrame(data={'Bloco',
                                      'field2',
                                      'field3',
                                      'field4',
                                      'field5',
                                      'field6',
                                      'field7',
                                      'field8'})

        # Iterando meses para o bloco M030
        for month in range(1, 13):

            if len(month) == 1:
                monthNormal = "0" & month

            m030entry = {'Bloco': "M030",
                    'field2': "0101" & self.anoLabel.get(),
                    'field3': monthrange(self.anoLabel.get(),month)[1] & monthNormal & self.anoLabel.get(),
                    'field4': "A" & monthNormal,
                    'field5':"",
                    'field6':"",
                    'field7':"",
                    'field8':""}

            blocoMes.append(m030entry)

            # Iterando linhas do bloco m300
            for i in self.m300.index:
                m300entry = self.m300[['Bloco',
                                       'field2',
                                       'field3',
                                       'field4',
                                       'field5',
                                       'field6',
                                       'field7']][i]

                blocoMes = pd.concat([blocoMes, m300entry], ignore_index=True)

                for j in self.m305.index:
                    m305entry = self.m305.loc[
                        (self.m305['field6'] == m030entry['field3'][0]) &
                        (self.m305['field5']) == m300entry['field2'][0]]
                    m305entry.drop(columns=['field5','field6'], inplace=True)

                    blocoMes = pd.concat([blocoMes, m305entry], ignore_index=True)

                    for k in self.m310.index:
                        m310entry = self.m310.loc[
                            (self.m310['field7'] == m030entry['field3'][0]) &
                            (self.m310['field6']) == m300entry['field2'][0]]
                        m310entry.drop(columns=['field6', 'field7'], inplace=True)

                        blocoMes = pd.concat([blocoMes, m310entry], ignore_index=True)

            blocoM = pd.concat([blocoM,blocoMes], ignore_index=True)

        m030entry = {'Bloco': "M030",
                    'field2': "0101" & self.anoLabel.get(),
                    'field3': monthrange(self.anoLabel.get(),month)[1] & monthNormal & self.anoLabel.get(),
                    'field4': "A00",
                    'field5':"",
                    'field6':"",
                    'field7':"",
                    'field8':""}
        mAcum = blocoMes.loc[(blocoMes['Bloco'] != "M030")]
        mAcum = pd.concat([m030entry,mAcum], ignore_index=True)
        blocoM = pd.concat([mAcum,blocoM], ignore_index=True)

        # prepara o dataframe para exportação

        blocoM.insert(0,"field0", "")
        blocoM["lastfield"] = ""
        blocoM.to_csv("blocoM.txt", sep="|", float_format='%.2f', header=False, index=False)
        # TODO: cada bloco tem um tamanho diferente (ex: bloco M030 só deve ter 4 campos separados por |)



root = tk.Tk()
root.iconbitmap(r'C:\Users\luizh\PycharmProjects\blocoM\icons\blocoM.ico')
app = main(master=root)
app.mainloop()