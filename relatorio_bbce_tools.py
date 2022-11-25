import datetime, pandas as pd, numpy as np, tools as tl, docx, matplotlib.pyplot as plt, os, win32com.client as win32
from datetime import datetime as dt
import mysql.connector
class Relatorio_BBCE:
    def __init__(self):
        while True:
            try:
                self.periodo = input("Informe a data que quer o relatório (Dd/Mm/Aa): ")
                self.mes = int(input("Até qual mês? Ex: 2 "))
                self.ano = int(input("Qual ano? Ex: 2023 "))
                self.novo_periodo = dt.strptime(self.periodo, '%d/%m/%Y').date()
                if not self.novo_periodo.weekday() >= 4:
                    raise ValueError("Dia fora do range permitido")
            except ValueError as e:
                print("Valor inválido:")
            else:
                break
        self.lista_semana = [self.novo_periodo-datetime.timedelta(days=contador) for contador in range(0,5)]
        print(self.lista_semana)
    def query_principal(self, tabela, inicio='2022-12-31'):
        query1 = f'''
        SELECT produto, dia, precos_interpolation.preco, inicio FROM {tabela} JOIN produtos_bbce ON id_produto = id
        WHERE DATEDIFF(fim,inicio) < 32 AND inicio < '2023-04-01' AND inicio > {inicio}
        AND (dia = "{self.lista_semana[4]}"
        OR dia = "{self.lista_semana[3]}"
        OR dia = "{self.lista_semana[2]}"
        OR dia = "{self.lista_semana[1]}"
        OR dia = "{self.lista_semana[0]}")
        AND submercado = "SE"
        AND energia = "CON"
        AND produtos_bbce.preco = "Fixo"
        ORDER BY inicio;
        '''
    def faz_grafico(self):
        db = tl.connection_db('BBCE')
        query1 = self.query_principal(tabela="precos_bbce_geral")
        print()
        query2 = self.query_principal(tabela="precos_interpolation")
        print(query2)
        tabela1 = pd.DataFrame(db.query(query1)) #transforma tabela em dataframe
        tabela2 = pd.DataFrame(db.query(query2)) #transforma tabela em dataframe
        a = pd.concat([tabela1, tabela2]) #junta as duas tabelas
        print(a)
        b = a.sort_values(['inicio', 'dia']) #ordena os valores das colunas inicio e dia
        b.reset_index(inplace=True, drop=True) #inplace=muda na tabela principal
        tabela = b.drop('inicio', axis=1)
        ymin = 10 * round(tabela['preco'].min() / 10)
        ymax = 10 * round(tabela['preco'].max() / 10) + 10
        produtos = list(dict.fromkeys(tabela['produto']))
        plt.figure(figsize=(6, 5))
        for produto in produtos:
            valores = tabela.loc[tabela['produto'] == produto]
            if len(valores['dia']) == 5:
                plt.plot_date([dt.strftime(i, "%d/%m") for i in valores['dia']], valores['preco'], '--o', label='',
                              alpha=0)
                break
        for produto in produtos:
            valores = tabela.loc[tabela['produto'] == produto]
            if len(valores['dia']) >= 3:
                plt.plot_date([dt.strftime(i, "%d/%m") for i in valores['dia']], valores['preco'], '--o', label=produto)
        plt.title('')
        plt.ylabel("Preço em R$")
        plt.yticks(np.arange(ymin, ymax, 5.0))
        plt.grid(linestyle='--')
        plt.savefig(f'./graficos/grafico_semana_{self.lista_semana[5].strftime("%d-%m")}.jpg',
                    bbox_extra_artists=(plt.legend(bbox_to_anchor=(1.58, 0), loc="lower right"),), bbox_inches='tight')
        plt.clf()



mapa = Relatorio_BBCE()
mapa.faz_grafico()
