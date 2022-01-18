import csv
from numpy import printoptions
import requests
import xlsxwriter
import pandas as pd
from IPython.display import display
import os
import time
import numpy as np


# Cria um nome aleatório com data / hora / minutos e segundos
timestr = time.strftime("%Y%m%d-%H%M%S")

nomeArquivo = "planilha1.xlsx"
# Verifica existe a pasta "exportado"
if os.path.exists('./exportado') == False:
    # Se não existeir ele cria a pasta
    os.mkdir(f'./exportado')

# Cria um arquivo dentro da pasta "exportado" com nome ateatório
os.mkdir(f'./exportado/{timestr}')


# Inicia a criação do BOT
class BotCsv:

    # define a quantidade inicial da contagem para 0
    quantidadeEmpresa = 0

    def ler_csv(self):
        self.df = pd.read_excel(nomeArquivo,
                                engine='openpyxl', usecols=[0, 1, 2])
        numero_linhas = int(self.df.count().iloc[0])
        for linha in range(0, numero_linhas):

            if 'Distribuidor' not in self.df.columns:
                print('Não existe uma coluna chamada "Distribuidor"')
                quit()
            self.distribuidor = (self.df['Distribuidor'].values[linha])

            if self.distribuidor == False:
                quit()
            self.praca = (self.df['Cidades'].values[linha])
            self.praca = self.praca.split(',')
            self.cnae = (self.df['CNAE'].values[linha])
            self.cnae = self.cnae.split(',')

            self.folder = timestr
            os.mkdir(
                f'./exportado/{timestr}/Distribuidor - {self.distribuidor}')
            for self.pracas in self.praca:
                self.pracas = self.pracas.strip()  # tirar espaço antes e depois da variavel
                print(self.pracas)
                for self.CNAEs in self.cnae:
                    self.CNAEs = self.CNAEs.strip()
                    bot.ler_nome_lojas()

    def ler_nome_lojas(self):
        self.lista = {}
        response = requests.get(
            f'http://localhost:3001/consulta/superconsulta?cnae={self.CNAEs}&municipio={self.pracas}')
        response = response.json()
        quantidade_itens = len(response)
        for c in range(0, quantidade_itens):
            self.lista[f'Empresa: {c}'] = response[c]
            self.contador = c
            print(c)
        bot.criar_arquivo()

    def criar_arquivo(self):
        self.df_arquivo = pd.DataFrame(self.lista)
        self.df_arquivo.replace(to_replace=[None], value="-", inplace=True)
        self.df_arquivo_transposed = self.df_arquivo.T  # or df1.transpose()
        print(self.df_arquivo)
        self.df_arquivo_transposed.rename(columns={
            'cnpj_nofilter': 'CNPJ não filtrado',
            'matriz_filial': 'Matriz ou Filial?',
            'cnpj': 'CNPJ',
            'qualificacao_nofilter': 'Código da qualificação Social',
            'situacao_cadastral_nofilter': 'código da Situação Cadastral',
            'razao_social': 'Razão Social',
            'cnae_fiscal': 'Cnae FF',
            'nome_fantasia': 'Nome Fantasia',
            'situacao_cadastral': 'Situação Cadastral',
            'data_da_situacao': 'Data da Situação',
            'motivo_situacao': 'Motivo da Situação',
            'situacao_especial': 'Situação Especial',
            'data_situacao_especial': 'Data da Situação Especial',
            'natureza_juridica': 'Natureza Juridica',
            'descricao_responsavel': 'Qualificação do Responsável',
            'capital_social': 'Capital Social',
            'porte_empresa': 'Porte da Empresa',
            'porte_empresa_nofilter': 'Código da Qualificação da Empresa',
            'cidade_exterior': 'Cidade Exterior',
            'cnae_fiscal': 'Cnae Fiscal',
            'cnae_fiscal_desc': 'Cnae Fiscal - Descrição',
            'cnae_secundario': 'Cnae Secundário',
            'telefone1': 'Telefone (1)',
            'telefone2': 'Telefone (2)',
            'email': 'Email',
            'municipio_nome': 'Município',
            'endereco': 'Endereço',
            'cep': 'CEP',
            'municipio': 'Município',
            'uf': 'Estado',
            'nome_socio': 'Nome do Sócio',
            'cpf_cnpj_socio': 'CPF/CNPJ Sócio',
            'nome_representante_legal': 'Nome do Representante legal',
            'cpf_representante_legal': 'CPF do Representante legal',
            'data_entrada_sociedade': 'Data de Entrada na Sociedade',
            'opcao_simples': 'Optou pelo Simples',
            'data_opcao_simples': 'Data que optou pelo Simples',
            'data_exclusao': 'Data que saiu do Simples',
            'opcao_mei': 'Optou pelo MEI',
            'data_opcao_mei': 'Data que optou pelo MEI',
            'exclusao_mei': 'Data de exclusão do MEI',
        },
            inplace=True)
        writer = pd.ExcelWriter(
            f"./exportado/{timestr}/Distribuidor - {self.distribuidor}/{self.pracas} - {self.CNAEs}.xlsx", engine='openpyxl', type='3_color_scale')

        self.df_arquivo_transposed.to_excel(writer, sheet_name='Sheet1')

        writer.sheets['Sheet1'].column_dimensions['A'].width = 15
        writer.sheets['Sheet1'].column_dimensions['B'].width = 15
        writer.sheets['Sheet1'].column_dimensions['C'].width = 15
        writer.sheets['Sheet1'].column_dimensions['D'].width = 15
        writer.sheets['Sheet1'].column_dimensions['E'].width = 25
        writer.sheets['Sheet1'].column_dimensions['F'].width = 40
        writer.sheets['Sheet1'].column_dimensions['G'].width = 40
        writer.sheets['Sheet1'].column_dimensions['H'].width = 15
        writer.sheets['Sheet1'].column_dimensions['I'].width = 15
        writer.sheets['Sheet1'].column_dimensions['J'].width = 30
        writer.sheets['Sheet1'].column_dimensions['K'].width = 10
        writer.sheets['Sheet1'].column_dimensions['L'].width = 10
        writer.sheets['Sheet1'].column_dimensions['M'].width = 10
        writer.sheets['Sheet1'].column_dimensions['N'].width = 30
        writer.sheets['Sheet1'].column_dimensions['O'].width = 15
        writer.sheets['Sheet1'].column_dimensions['P'].width = 30
        writer.sheets['Sheet1'].column_dimensions['Q'].width = 20
        writer.sheets['Sheet1'].column_dimensions['R'].width = 15
        writer.sheets['Sheet1'].column_dimensions['S'].width = 40
        writer.sheets['Sheet1'].column_dimensions['T'].width = 15
        writer.sheets['Sheet1'].column_dimensions['U'].width = 15
        writer.sheets['Sheet1'].column_dimensions['V'].width = 35
        writer.sheets['Sheet1'].column_dimensions['W'].width = 25
        writer.sheets['Sheet1'].column_dimensions['X'].width = 20
        writer.sheets['Sheet1'].column_dimensions['Y'].width = 20
        writer.sheets['Sheet1'].column_dimensions['Z'].width = 40
        writer.sheets['Sheet1'].column_dimensions['AA'].width = 15
        writer.sheets['Sheet1'].column_dimensions['AB'].width = 20
        writer.sheets['Sheet1'].column_dimensions['AC'].width = 10
        writer.sheets['Sheet1'].column_dimensions['AD'].width = 40

        writer.sheets['Sheet1'].column_dimensions['AE'].width = 35
        writer.sheets['Sheet1'].column_dimensions['AF'].width = 15
        writer.sheets['Sheet1'].column_dimensions['AG'].width = 30
        writer.sheets['Sheet1'].column_dimensions['AH'].width = 35
        writer.sheets['Sheet1'].column_dimensions['AI'].width = 40
        writer.sheets['Sheet1'].column_dimensions['AJ'].width = 30
        writer.sheets['Sheet1'].column_dimensions['AK'].width = 30
        writer.sheets['Sheet1'].column_dimensions['AL'].width = 20
        writer.save()
        QuantE = self.df_arquivo_transposed.shape[0]
        self.quantidadeEmpresa += QuantE
        print("Quantidade de empresas: ", self.quantidadeEmpresa)


if os.path.exists(nomeArquivo) == True:
    print('existe o arquivo chamado planilha.xlsx')
    bot = BotCsv()
    bot.ler_csv()

else:
    print('Preciso de um arquivo chamado planilha.xlsx')
    quit()
