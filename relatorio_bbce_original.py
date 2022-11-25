import datetime, pandas as pd, numpy as np, tools as tl, docx, matplotlib.pyplot as plt, os, win32com.client as win32
from datetime import datetime as dt
class Relatorio:
    def __init__(self, essa_semana=False):
        # essa_semana = True assume que hoje é sexta e ja temos dados da semana inteira para fazer o relatório
        # essa_semana = False faz um relatório da semana anterior
        if essa_semana:
            if dt.now().date().weekday() < 4:
                print('Hoje não é sexta')
                return
            else:
                hoje = dt.now().date().weekday()
                semana = [dt.now().date()]
                while hoje > 0:
                    semana.append(semana[-1] - datetime.timedelta(days=1))
                    hoje -= 1
        if not essa_semana:
            hoje = (dt.now().date() - datetime.timedelta(days=7)).weekday()
            semana = [dt.now().date() - datetime.timedelta(days=7)]
            while hoje > 0:
                semana.append(semana[-1] - datetime.timedelta(days=1))
                hoje -= 1
            semana = semana[::-1]
            hoje = (dt.now().date() - datetime.timedelta(days=7)).weekday()
            while hoje < 4:
                semana.append(semana[-1] + datetime.timedelta(days=1))
                hoje += 1
            semana = semana[::-1]
        self.semana = semana[::-1]
        # self.semana_passada = [day - datetime]

    def faz_grafico(self):
        db = tl.connection_db('BBCE')
        query1 = f'''
        SELECT produto, dia, precos_bbce_geral.preco, inicio FROM precos_bbce_geral JOIN produtos_bbce ON id_produto = id
        WHERE DATEDIFF(fim,inicio) < 32 AND inicio < '2023-04-01'  AND inicio > '2022-12-31'
        AND (dia = "{self.semana[0]}"
        OR dia = "{self.semana[1]}"
        OR dia = "{self.semana[2]}"
        OR dia = "{self.semana[3]}"
        OR dia = "{self.semana[4]}")
        AND submercado = "SE"
        AND energia = "CON"
        AND produtos_bbce.preco = "Fixo"
        ORDER BY inicio;
        '''
        print(query1)
        query2 = f'''
        SELECT produto, dia, precos_interpolation.preco, inicio FROM precos_interpolation JOIN produtos_bbce ON id_produto = id
        WHERE DATEDIFF(fim,inicio) < 32 AND inicio < '2023-04-01' AND inicio > '2022-12-31'
        AND (dia = "{self.semana[0]}"
        OR dia = "{self.semana[1]}"
        OR dia = "{self.semana[2]}"
        OR dia = "{self.semana[3]}"
        OR dia = "{self.semana[4]}")
        AND submercado = "SE"
        AND energia = "CON"
        AND produtos_bbce.preco = "Fixo"
        ORDER BY inicio;
        '''
        print(query2)
        tabela1 = pd.DataFrame(db.query(query1)) #transforma tabela em dataframe
        # print(tabela1)
        tabela2 = pd.DataFrame(db.query(query2)) #transforma tabela em dataframe
        # print(tabela2)
        a = pd.concat([tabela1, tabela2]) #junta as duas tabelas
        b = a.sort_values(['inicio', 'dia']) #ordena os valores das colunas inicio e dia
        b.reset_index(inplace=True, drop=True) #inplace=muda na tabela principal
        # print(b)
        # b.at[10,'preco'] = 72.1936
        # print(b)
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
        plt.savefig(f'./graficos/grafico_semana_{self.semana[0].strftime("%d-%m")}.jpg',
                    bbox_extra_artists=(plt.legend(bbox_to_anchor=(1.58, 0), loc="lower right"),), bbox_inches='tight')
        plt.clf()

    def faz_tabelas(self):
        db = tl.connection_db('BBCE')
        query1 = f'''
        SELECT produto, dia, precos_bbce_geral.preco, inicio, fim FROM precos_bbce_geral JOIN produtos_bbce ON id_produto = id
        WHERE DATEDIFF(fim,inicio) < 32 AND inicio < '2023-04-01'
        AND (dia = "{self.semana[0]}"
        OR dia = "{self.semana[1]}"
        OR dia = "{self.semana[2]}"
        OR dia = "{self.semana[3]}"
        OR dia = "{self.semana[4]}")
        AND submercado = "SE"
        AND energia = "CON"
        AND produtos_bbce.preco = "Fixo"
        ORDER BY inicio;
        '''
        print(query1)
        query2 = f'''
        SELECT produto, dia, precos_interpolation.preco, inicio, fim FROM precos_interpolation JOIN produtos_bbce ON id_produto = id
        WHERE DATEDIFF(fim,inicio) < 32 AND inicio < '2023-04-01'
        AND (dia = "{self.semana[0]}"
        OR dia = "{self.semana[1]}"
        OR dia = "{self.semana[2]}"
        OR dia = "{self.semana[3]}"
        OR dia = "{self.semana[4]}")
        AND submercado = "SE"
        AND energia = "CON"
        AND produtos_bbce.preco = "Fixo"
        ORDER BY inicio;
        '''
        print(query2)
        tabela1 = pd.DataFrame(db.query(query1))
        # print(tabela1)
        tabela2 = pd.DataFrame(db.query(query2))
        # print(tabela2)
        a = pd.concat([tabela1, tabela2])
        b = a.sort_values(['inicio', 'dia'])
        b.reset_index(inplace=True, drop=True)
        # print(b)
        # b.at[10,'preco'] = 72.1936
        tabela_preco = b  # .drop('inicio',axis=1)
        produtos = list(dict.fromkeys(tabela_preco['produto']))
        colunas = ['Produto', 'Preço inicial', 'Preço final', 'Variação', 'Qt. Negócios', 'Volume']
        col_pro, col_pri, col_ult, col_var, col_qtn, col_vol = [], [], [], [], [], []
        for produto in produtos:
            valores = tabela_preco.loc[tabela_preco['produto'] == produto]
            if len(valores['dia']) >= 4:
                # print(produto)
                # print('aaa')
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

        query = f'''
        SELECT produto, dia, precos_bbce_geral.preco, inicio, fim FROM precos_bbce_geral JOIN produtos_bbce ON id_produto = id
        WHERE DATEDIFF(fim, inicio) < 32 AND inicio < '2023-04-01'
        AND (dia = "{semana_passada[0]}"
        OR dia = "{semana_passada[1]}"
        OR dia = "{semana_passada[2]}"
        OR dia = "{semana_passada[3]}"
        OR dia = "{semana_passada[4]}")
        AND submercado = "SE"
        AND energia = "CON"
        AND produtos_bbce.preco = "Fixo"
        ORDER BY inicio,dia;
        '''
        print(query)
        query_i = f'''
        SELECT produto, dia, precos_interpolation.preco, inicio, fim FROM precos_interpolation JOIN produtos_bbce ON id_produto = id
        WHERE DATEDIFF(fim, inicio) < 32 AND inicio < '2023-04-01'
        AND (dia = "{semana_passada[0]}"
        OR dia = "{semana_passada[1]}"
        OR dia = "{semana_passada[2]}"
        OR dia = "{semana_passada[3]}"
        OR dia = "{semana_passada[4]}")
        AND submercado = "SE"
        AND energia = "CON"
        AND produtos_bbce.preco = "Fixo"
        ORDER BY inicio,dia;
        '''
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
    def escreve_relatorio(self):
        tabela_info, tabela_comparativa = self.faz_tabelas()
        self.faz_grafico()
        semana_passada = [dia - datetime.timedelta(days=7) for dia in self.semana]
        doc = docx.Document()
        doc.add_heading('Relatório Semanal BBCE', 0)
        doc.add_heading(f"Semana {self.semana[0].strftime('%d/%m')} - {self.semana[4].strftime('%d/%m')}\n", 1)
        doc.add_paragraph("Produtos com alta liquidez: Sudeste; Convencional; Preço fixo\n")
        doc.add_picture(f'./graficos/grafico_semana_{self.semana[0].strftime("%d-%m")}.jpg',
                        width=docx.shared.Cm(15.82))
        table = doc.add_table(rows=1, cols=6)
        row = table.rows[0].cells
        row[0].text = 'Produto      '
        row[1].text = 'Preço inicial'
        row[2].text = 'Preço final  '
        row[3].text = 'Variação     '
        row[4].text = 'Qt. Negócios '
        row[5].text = 'Volume total '
        for linha in tabela_info.itertuples(index=False):
            row = table.add_row().cells
            row[0].text = linha[0]
            row[1].text = linha[1]
            row[2].text = linha[2]
            row[3].text = linha[3]
            row[4].text = linha[4]
            row[5].text = linha[5]
        table.style = 'Colorful Grid Accent 1'
        for cell in table.columns[0].cells:
            cell.width = docx.shared.Cm(2.47)
        for cell in table.columns[1].cells:
            cell.width = docx.shared.Cm(2.72)
        for cell in table.columns[2].cells:
            cell.width = docx.shared.Cm(2.50)
        for cell in table.columns[3].cells:
            cell.width = docx.shared.Cm(2.17)
        for cell in table.columns[4].cells:
            cell.width = docx.shared.Cm(2.72)
        for cell in table.columns[5].cells:
            cell.width = docx.shared.Cm(3.04)
        doc.add_paragraph(
            f"\nVariações em relação ao preço da semana anterior ({semana_passada[0].strftime('%d/%m')}-{semana_passada[4].strftime('%d/%m')}) \n")
        table2 = doc.add_table(rows=1, cols=4)
        row = table2.rows[0].cells
        row[0].text = 'Produto'
        row[1].text = 'Preço passado'
        row[2].text = 'Preço atual'
        row[3].text = 'Variação'

        for linha in tabela_comparativa.itertuples(index=False):
            row = table2.add_row().cells
            row[0].text = linha[0]
            row[1].text = linha[1]
            row[2].text = linha[2]
            row[3].text = linha[3]
        table2.style = "Colorful Grid Accent 1"
        for cell in table2.columns[0].cells:
            cell.width = docx.shared.Cm(2.46)
        for cell in table2.columns[1].cells:
            cell.width = docx.shared.Cm(3.18)
        for cell in table2.columns[2].cells:
            cell.width = docx.shared.Cm(2.80)
        for cell in table2.columns[3].cells:
            cell.width = docx.shared.Cm(2.10)
        doc.save(f'./relatorios_bbce/relatorio_semana_{self.semana[0].strftime("%d-%m")}.docx')
        try:
            word = win32.Dispatch('Word.Application')
            wdFormatPDF = 17
            in_file = os.path.abspath(f'./relatorios_bbce/relatorio_semana_{self.semana[0].strftime("%d-%m")}.docx')
            out_file = os.path.abspath(f'./relatorios_bbce/relatorio_semana_{self.semana[0].strftime("%d-%m")}.pdf')
            doc = word.Documents.Open(in_file)
            doc.SaveAs(out_file, FileFormat=wdFormatPDF)
            doc.Close()
            word.Quit()
        except Exception:
            print('Arquivo PDF não foi criado')

relatorio = Relatorio(essa_semana=True)
# relatorio.faz_tabelas()
relatorio.escreve_relatorio()
