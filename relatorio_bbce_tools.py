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
    def query_principal(self, tabela, tabela2, inicio='2022-12-31'):
        query1 = f'''
        SELECT produto, dia, {tabela2}, inicio FROM {tabela} JOIN produtos_bbce ON id_produto = id
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
        query1 = self.query_principal(tabela="precos_bbce_geral", tabela2="precos_bbce_geral.preco")
        print()
        query2 = self.query_principal(tabela="precos_interpolation", tabela2="precos_interpolation.preco")
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

   def faz_tabelas(self):
        db = tl.connection_db('BBCE')
        query1 = self.query_principal(tabela2="precos_bbce_geral.preco", tabela="precos_bbce_geral")
        print(query1)
        query2 = self.query_principal(tabela2= "precos_interpolation.preco", tabela="precos_interpolation")
        print(query2)
        tabela1 = pd.DataFrame(db.query(query1))
        tabela2 = pd.DataFrame(db.query(query2))
        a = pd.concat([tabela1, tabela2])
        b = a.sort_values(['inicio', 'dia'])
        b.reset_index(inplace=True, drop=True)
        tabela_preco = b  # .drop('inicio',axis=1)
        produtos = list(dict.fromkeys(tabela_preco['produto']))
        colunas = ['Produto', 'Preço inicial', 'Preço final', 'Variação', 'Qt. Negócios', 'Volume']
        col_pro, col_pri, col_ult, col_var, col_qtn, col_vol = [], [], [], [], [], []
        for produto in produtos:
            valores = tabela_preco.loc[tabela_preco['produto'] == produto]
            if len(valores['dia']) >= 4:
                inicio = valores['inicio'].tolist()[0]
                fim = valores['fim'].tolist()[0]
                primeiro_preco = valores['preco'].tolist()[0]
                ultimo_preco = valores['preco'].tolist()[-1]
                variacao = (ultimo_preco - primeiro_preco) * 100 / primeiro_preco
                query = f'''
                SELECT volume_medio FROM bbce
                WHERE submercado = "SE"
                AND tipo_energia = "CON"
                AND data_inicio = "{inicio}"
                AND data_fim = "{fim}"
                AND tipo_preco = "Fixo"
                AND data_hora > "{self.semana[0]}"
                AND data_hora < "{self.semana[4] + datetime.timedelta(days=1)}"
                '''
                tabela = pd.DataFrame(db.query(query))
                # print(tabela)
                qt_negocios = len(tabela['volume_medio'])
                volume = sum(tabela['volume_medio'])
                col_pro.append(produto)
                col_pri.append("R$ %.2f" % primeiro_preco)
                col_ult.append("R$ %.2f" % ultimo_preco)
                col_var.append("%.2f %%" % variacao)
                col_qtn.append(f'{qt_negocios}')
                col_vol.append("%.2f MwM" % volume)
        tabela1 = pd.DataFrame(
            {colunas[0]: [i[6:15] for i in col_pro], colunas[1]: col_pri, colunas[2]: col_ult, colunas[3]: col_var,
             colunas[4]: col_qtn, colunas[5]: col_vol})
        semana_passada = [dia - datetime.timedelta(days=7) for dia in self.semana]
        query = self.query_principal(tabela2="precos_bbce_geral.preco", tabela="precos_bbce_geral")
        print(query)
        query_i = self.query_principal(tabela2= "precos_interpolation.preco", tabela="precos_interpolation")
        print(query_i)
        tabela_preco_passada = pd.DataFrame(db.query(query))
        tabela_preco_passada_i = pd.DataFrame(db.query(query_i))
        print(tabela_preco_passada_i)
        produtos_passada = list(dict.fromkeys(tabela_preco_passada['produto']))
        colunas = ['Produto', 'Preço passado', 'Preço atual', 'Variação']
        col_pro, col_prp, col_pra, col_var = [], [], [], []
        for produto in produtos_passada:
            valores = tabela_preco_passada.loc[tabela_preco_passada['produto'] == produto]
            if len((tabela_preco.loc[tabela_preco['produto'] == produto])['dia']) >= 3 and len(valores['dia']) >= 3:
                preco_atual = (tabela_preco.loc[tabela_preco['produto'] == produto])['preco'].tolist()[-1]
                preco_passada = valores['preco'].tolist()[-1]
                variacao = (preco_atual - preco_passada) * 100 / preco_passada
                col_pro.append(produto)
                col_prp.append("R$ %.2f" % preco_passada)
                col_pra.append("R$ %.2f" % preco_atual)
                col_var.append("%.2f%%" % variacao)
        tabela2 = pd.DataFrame(
            {colunas[0]: [i[6:15] for i in col_pro], colunas[1]: col_prp, colunas[2]: col_pra, colunas[3]: col_var})

        tabela1.to_excel(f'./tabelas/tabela_semana_{self.semana[0]}.xlsx', sheet_name='sheet1', index=False)
        tabela2.to_excel(f'./tabelas/tabela_comparativa_semana_{self.semana[0]}.xlsx', sheet_name='sheet2', index=False)
        return tabela1, tabela2

mapa = Relatorio_BBCE()
mapa.faz_grafico()
