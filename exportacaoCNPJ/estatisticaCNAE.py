import pandas as pd
import numpy as np


entradaExterna = ["Município", "Cnae Fiscal - Descrição", "Estado"]

# Inicia a criação do BOT

novatabela = {}


class EstatisticaCNAE:
    def __init__(self, entrada):
        self.lista = {}
        self.CampoName = entrada

    def ler(self):
        self.arquivo = pd.read_excel(r'./lista.xlsx')
        novatabela['Tabela Principal'] = self.arquivo
        print(self.arquivo)

        estCNAE.modificar()

    def modificar(self):
        self.arquivo[f"Quantidade de {self.CampoName}'s"] = self.arquivo['Cnae Fiscal']

        y = pd.pivot_table(self.arquivo,
                           index=self.CampoName,
                           values=[f"Quantidade de {self.CampoName}'s"],
                           aggfunc='count', fill_value=0).reset_index(level=-1)
        print(y)
        self.arquivo = y.sort_values(
            by=[f"Quantidade de {self.CampoName}'s"], ascending=False)
        estCNAE.salvar()

    def salvar(self):
        self.df_arquivo = pd.DataFrame(self.arquivo)
        self.df_arquivo.style.set_properties(**{'text-align': 'center'})

        novatabela[self.CampoName] = self.arquivo

    def formatarTabela(self):
        self.writer = pd.ExcelWriter(
            f"./listaTest.xlsx", engine='openpyxl')

    def sheetName(self, Nome):
        novatabela[Nome].to_excel(
            self.writer, sheet_name=Nome, index=False)

        if(Nome == "Cnae Fiscal - Descrição"):
            self.writer.sheets[Nome].column_dimensions['A'].width = 80

        elif(Nome == "Município"):
            self.writer.sheets[Nome].column_dimensions['A'].width = 30

        elif(Nome == "Estado"):
            self.writer.sheets[Nome].column_dimensions['A'].width = 20

        elif(Nome == "Tabela Principal"):
            self.writer.sheets[Nome].column_dimensions['A'].width = 15
            self.writer.sheets[Nome].column_dimensions['C'].width = 25
            self.writer.sheets[Nome].column_dimensions['D'].width = 35
            self.writer.sheets[Nome].column_dimensions['E'].width = 35
            self.writer.sheets[Nome].column_dimensions['F'].width = 20
            self.writer.sheets[Nome].column_dimensions['G'].width = 25
            self.writer.sheets[Nome].column_dimensions['H'].width = 20
            self.writer.sheets[Nome].column_dimensions['I'].width = 20
            self.writer.sheets[Nome].column_dimensions['J'].width = 20
            self.writer.sheets[Nome].column_dimensions['K'].width = 20
            self.writer.sheets[Nome].column_dimensions['L'].width = 35
            self.writer.sheets[Nome].column_dimensions['M'].width = 20
            self.writer.sheets[Nome].column_dimensions['N'].width = 20
            self.writer.sheets[Nome].column_dimensions['O'].width = 20
            self.writer.sheets[Nome].column_dimensions['P'].width = 20
            self.writer.sheets[Nome].column_dimensions['Q'].width = 35
            self.writer.sheets[Nome].column_dimensions['R'].width = 20
            self.writer.sheets[Nome].column_dimensions['S'].width = 20
            self.writer.sheets[Nome].column_dimensions['T'].width = 20
            self.writer.sheets[Nome].column_dimensions['U'].width = 30
            self.writer.sheets[Nome].column_dimensions['V'].width = 20
            self.writer.sheets[Nome].column_dimensions['W'].width = 20
            self.writer.sheets[Nome].column_dimensions['X'].width = 20
            self.writer.sheets[Nome].column_dimensions['Y'].width = 40
            self.writer.sheets[Nome].column_dimensions['Z'].width = 35
            self.writer.sheets[Nome].column_dimensions['AA'].width = 20
            self.writer.sheets[Nome].column_dimensions['AB'].width = 20
            self.writer.sheets[Nome].column_dimensions['AC'].width = 20
            self.writer.sheets[Nome].column_dimensions['AD'].width = 20

        else:
            self.writer.sheets[Nome].column_dimensions['A'].width = 30

        self.writer.sheets[Nome].column_dimensions['B'].width = 20

    def saveF(self):
        self.writer.save()


for entra in entradaExterna:
    estCNAE = EstatisticaCNAE(entra)
    estCNAE.ler()
    print("tabela 1:")
    print(novatabela[entra])

estCNAE.formatarTabela()

for entra in novatabela:
    estCNAE.sheetName(entra)


print("tem algo:")
print(novatabela)

estCNAE.saveF()
