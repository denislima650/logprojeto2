import datetime, pandas as pd, numpy as np, tools as tl, docx, matplotlib.pyplot as plt, os, win32com.client as win32
from datetime import datetime as dt
class Relatorio_BBCE:
    def __init__(self):
        while True:
            try:
                self.periodo = input("Informe a data que quer o relatório (Dd/Mm/Aa): ")
                self.novo_periodo = dt.strptime(self.periodo, '%d/%m/%Y').date()
                if not self.novo_periodo.weekday() >= 4:
                    raise ValueError("Dia fora do range permitido")
            except ValueError as e:
                print("Valor inválido:")
            else:
                break
        self.lista_semana = [self.novo_periodo-datetime.timedelta(days=contador) for contador in range(0,5)]
    def remove_repetidos(self, lista):
        l = []
        for i in lista:
            if i not in l:
                l.append(i)
        return l
    def query_principal(self, lista, tabela, tabela2, inicio='2022-12-31', tem_fim=''):
        query_padrao = f'''
        SELECT produto, dia, {tabela2}, inicio{tem_fim} FROM {tabela} JOIN produtos_bbce ON id_produto = id
        WHERE DATEDIFF(fim,inicio) < 32 AND inicio < '2023-04-01' AND inicio > '{inicio}'
        AND (dia = "{lista[4]}"
        OR dia = "{lista[3]}"
        OR dia = "{lista[2]}"
        OR dia = "{lista[1]}"
        OR dia = "{lista[0]}")
        AND submercado = "SE"
        AND energia = "CON"
        AND produtos_bbce.preco = "Fixo"
        ORDER BY inicio;
        '''
        return query_padrao
    def faz_grafico(self):
        db = tl.connection_db('BBCE')
        query1 = self.query_principal(lista=self.lista_semana, tabela="precos_bbce_geral", tabela2="precos_bbce_geral.preco", tem_fim='')
        query2 = self.query_principal(lista=self.lista_semana, tabela="precos_interpolation", tabela2="precos_interpolation.preco", tem_fim='')
        tabela1 = pd.DataFrame(db.query(query1))	 #transforma tabela em dataframe
        tabela2 = pd.DataFrame(db.query(query2))
        a = pd.concat([tabela1, tabela2]) 		 #junta as duas tabelas
        b = a.sort_values(['inicio', 'dia'])
        b.reset_index(inplace=True, drop=True) 		 #inplace=muda na tabela principal
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
        plt.savefig(f'./graficos/grafico_semana_{self.lista_semana[4].strftime("%d-%m")}.jpg',
                    bbox_extra_artists=(plt.legend(bbox_to_anchor=(1.58, 0), loc="lower right"),), bbox_inches='tight')
    def faz_tabelas(self):
        db = tl.connection_db('BBCE')
        query3 = self.query_principal(lista=self.lista_semana, tabela2="precos_bbce_geral.preco", tabela="precos_bbce_geral", tem_fim=', fim')
        query4 = self.query_principal(lista=self.lista_semana, tabela2= "precos_interpolation.preco", tabela="precos_interpolation", tem_fim=', fim')
        tabela1 = pd.DataFrame(db.query(query3))
        tabela2 = pd.DataFrame(db.query(query4))
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
                AND data_hora > "{self.lista_semana[4]}"
                AND data_hora < "{self.lista_semana[0] + datetime.timedelta(days=1)}"
                '''
                tabela = pd.DataFrame(db.query(query))
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
        semana_passada = [dia - datetime.timedelta(days=7) for dia in self.lista_semana]
        query5 = self.query_principal(lista=semana_passada, tabela="precos_bbce_geral", tabela2="precos_bbce_geral.preco", tem_fim=', fim')
        query6 = self.query_principal(lista=semana_passada, tabela="precos_interpolation", tabela2="precos_interpolation.preco", tem_fim=', fim')
        tabela_preco_passada = pd.DataFrame(db.query(query5))
        tabela_preco_passada_i = pd.DataFrame(db.query(query6))
        tabela_preco_passada_j = pd.concat([tabela_preco_passada_i, tabela_preco_passada])
        produtos_passada = list(dict.fromkeys(tabela_preco_passada['produto']))
        produtos_passada_i = list(dict.fromkeys(tabela_preco_passada_i['produto']))
        produtos_passada_j = produtos_passada_i+produtos_passada
        print(produtos_passada_j)
        lista = self.remove_repetidos(produtos_passada_j)
        print(lista)
        colunas = ['Produto', 'Preço passado', 'Preço atual', 'Variação']
        col_pro, col_prp, col_pra, col_var = [], [], [], []
        for produto in lista:
            valores = tabela_preco_passada_j.loc[tabela_preco_passada_j['produto'] == produto]
            if len((tabela_preco.loc[tabela_preco['produto'] == produto])['dia']) >= 1 and len(valores['dia']) >= 1:
                preco_atual = (tabela_preco.loc[tabela_preco['produto'] == produto])['preco'].tolist()[-1]
                preco_passada = valores['preco'].tolist()[-1]
                variacao = (preco_atual - preco_passada) * 100 / preco_passada
                col_pro.append(produto)
                col_prp.append("R$ %.2f" % preco_passada)
                col_pra.append("R$ %.2f" % preco_atual)
                col_var.append("%.2f%%" % variacao)
        tabela2 = pd.DataFrame(
            {colunas[0]: [i[6:15] for i in col_pro], colunas[1]: col_prp, colunas[2]: col_pra, colunas[3]: col_var})
        tabela1.to_excel(f'./tabelas/tabela_semana_{self.lista_semana[4]}.xlsx', sheet_name='sheet1', index=False)
        tabela2.to_excel(f'./tabelas/tabela_comparativa_semana_{self.lista_semana[4]}.xlsx', sheet_name='sheet2', index=False)
        return tabela1, tabela2
    def escreve_relatorio(self):
        tabela_info, tabela_comparativa = self.faz_tabelas()
        self.faz_grafico()
        semana_passada = [dia - datetime.timedelta(days=7) for dia in self.lista_semana]
        doc = docx.Document()
        doc.add_heading('Relatório Semanal BBCE', 0)
        doc.add_heading(f"Semana {self.lista_semana[4].strftime('%d/%m')} - {self.lista_semana[0].strftime('%d/%m')}\n", 1)
        doc.add_paragraph("Produtos com alta liquidez: Sudeste; Convencional; Preço fixo\n")
        doc.add_picture(f'./graficos/grafico_semana_{self.lista_semana[4].strftime("%d-%m")}.jpg', width=docx.shared.Cm(15.82))
        table = doc.add_table(rows=1, cols=6)
        row = table.rows[0].cells
        lista_row = ['Produto      ', 'Preço inicial', 'Preço final  ', 'Variação     ', 'Qt. Negócios ', 'Volume total ']
        for linha in range(0, 6):
            row[linha].text = lista_row[linha]
        for linha in tabela_info.itertuples(index=False):
            row = table.add_row().cells
            for i in [0, 1, 2, 3, 4, 5]:
                row[i].text = linha[i]
        table.style = 'Colorful Grid Accent 1'
        lista_row3 = [2.47, 2.72, 2.50, 2.17, 2.72, 3.04]
        for linha in range(0, 6):
            for cell in table.columns[linha].cells:
                cell.width = docx.shared.Cm(lista_row3[linha])
        doc.add_paragraph(
            f"\nVariações em relação ao preço da semana anterior ({semana_passada[0].strftime('%d/%m')}-{semana_passada[4].strftime('%d/%m')}) \n")
        table2 = doc.add_table(rows=1, cols=4)
        row = table2.rows[0].cells
        lista_table2 = ['Produto', 'Preço passado', 'Preço atual', 'Variação']
        for indice in range(0, 4):
            row[indice].text = lista_table2[indice]
        for linha in tabela_comparativa.itertuples(index=False):
            row = table2.add_row().cells
            for i in [0, 1, 2, 3]:
                row[i].text = linha[i]
        table2.style = "Colorful Grid Accent 1"
        lista_tamanhos = [2.46, 3.18, 2.80, 2.10]
        for indice in range(0, 4):
            for cell in table2.columns[indice].cells:
                cell.width=docx.shared.Cm(lista_tamanhos[indice])
        doc.save(f'./relatorios_bbce/relatorio_semana_{self.lista_semana[4].strftime("%d-%m")}.docx')
        try:
            word = win32.Dispatch('Word.Application')
            wdFormatPDF = 17
            in_file = os.path.abspath(f'./relatorios_bbce/relatorio_semana_{self.lista_semana[4].strftime("%d-%m")}.docx')
            out_file = os.path.abspath(f'./relatorios_bbce/relatorio_semana_{self.lista_semana[4].strftime("%d-%m")}.pdf')
            doc = word.Documents.Open(in_file)
            doc.SaveAs(out_file, FileFormat=wdFormatPDF)
            doc.Close()
            word.Quit()
        except Exception:
            print('Arquivo PDF não foi criado')

relatorio = Relatorio_BBCE()
relatorio.escreve_relatorio()
