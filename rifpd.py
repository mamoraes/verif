# -- coding: utf-8 --'
import pandas as pd
import sqlite3

import six
import pyanx
import textwrap

pasta='../rif/'

class Pyanx_macros(pyanx.Pyanx):
    def __init__(self):
        pyanx.Pyanx.__init__(self)

    def createStream(self, layout='spring_layout', pretty=True, iterations=50):
        self.layout(layout, iterations)

        chart = pyanx.anx.Chart(IdReferenceLinking=False)
        self._Pyanx__add_entity_types(chart)
        self._Pyanx__add_link_types(chart)
        self._Pyanx__add_entities(chart)
        self._Pyanx__add_links(chart)
        output_stream = six.StringIO()
        chart.export(output_stream, 0, pretty_print=pretty, namespacedef_=None)
        return output_stream


def gerar_planilha(arquivo, df, nome, indice=False):
    df.style.bar(color='#d65f5f')
    df.to_excel(arquivo, sheet_name=nome, index=indice)
       
    formato_cabecalho = arquivo.book.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 1})
    formato_cabecalho2 = arquivo.book.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': 'yellow',
    'border': 1})
        
# Write the column headers with the defined format.
    #print(df.index.names)
    if not indice:
        for col_num, value in enumerate(df.columns.values):
            arquivo.sheets[nome].write(0, col_num, value, formato_cabecalho)
    else:
        for col_num, value in enumerate(df.index.names):
            arquivo.sheets[nome].write(0, col_num, value, formato_cabecalho2)
        for col_num, value in enumerate(df.columns.values):
            arquivo.sheets[nome].write(0, col_num + len(df.index.names), value, formato_cabecalho)

def gerar_planilhaXLS(arquivo, df, nome, indice=False):
    df.style.bar(color='#d65f5f')
    df.to_excel(arquivo, sheet_name=nome, index=indice)

def tipoi2F(umou2=1,linha=None,carJuncao='\r '):
    descricao = linha[1 if umou2==1 else 3]
#     if descricao == '': #telefone ou endereco
#         descricao = carJuncao.join(node[4:].split('__'))
#     else:
#         if self.GNX.node[node]['tipo'] !='TEL':   
#             descricao = Obj.parseCPFouCNPJ(node) + carJuncao + carJuncao.join(textwrap.wrap(descricao,30))
    
    dicTipo = {'TEL':u'Telefone', 'END':u'Local', 'PF':u'PF', 'PJ':u'PJ', 'PE':u'Edifício', 'ES':u'Edifício', 'CC':u'Conta','INF':u'Armário' }
    tipo = linha[7 if umou2==1 else 8]
    tipoi2 = dicTipo[tipo]
    if tipo in('TEL','END','CC'):
        descricao = ''
    else:
        descricao = carJuncao.join(textwrap.wrap(descricao,30))
    sexo = 1

    if tipo=='PF':
        #if self.GNX.node[node]['sexo']==1:
        if not sexo or sexo == 1:
            tipoi2 = u'Profissional (masculino)'
        elif sexo == 2:
            tipoi2 = u'Profissional (feminino)'     
    elif tipo=='PJ':
        #if node[8:12]!='0001':
        #if sexo != 1: #1=matriz
        if sexo % 2==0: #1=matriz
            tipoi2 = u'Apartamento' #filial de empresa
        else:
            tipoi2 = u'Escritório' 
    elif tipo=='PE':
        tipoi2 = u'Oficina'

    corSituacao = linha[9 if umou2 ==1 else 10]  
    if linha[4 if umou2 ==1 else 5]==0:
        corSituacao='Vermelho'
    return (tipoi2, descricao, corSituacao)
       
def to_i2(df, arquivo = None):
    dicTiposIngles = {u'Profissional (masculino)':u'Person',
                   u'Profissional (feminino)':u'Woman',
                   u'Escritório':u'Office',
                   u'Apartamento':u'Workshop',
                   u'Governo':u'House',
                   u'Casa':u'House',
                   u'Loja':u'Office',
                   u'Oficina':u'Office',
                   u'Telefone':u'Phone',
                   u'Local':u'Place',
                   u'Conta':u'Account',
                   u'Armário':u'Cabinet',
                   u'Edifício':u'Office'
                   } 
    chart = Pyanx_macros()
    noi2origem= {}
    noi2destino={}
    
    idc=0
    for campos in df.iterrows():
        tipo, descricao, corSituacao = tipoi2F(linha=campos, umou2=1, carJuncao=' ')
        noi2origem[idc] = chart.add_node(entity_type=dicTiposIngles.get(tipo,''), label=(campos['cpfcnpj1']) + u'-' +(descricao))
        tipo, descricao, corSituacao = tipoi2F(linha=campos, umou2=2, carJuncao=' ')
        noi2destino[idc] = chart.add_node(entity_type=dicTiposIngles.get(tipo,''), label=(campos['cpfcnpj1']) + u'-' +(descricao))

        nomeLigacao = campos['descrição']
        chart.add_edge(noi2origem[idc], noi2destino[idc], removeAcentos(nomeLigacao))
        idc += 1

    fstream = chart.createStream(layout='spring_layout',iterations=0) #não calcula posição

    retorno = fstream.getvalue()
    fstream.close()
    if arquivo is not None:
      f = open(arquivo, 'w')
      f.write(retorno)
      f.close()
    return retorno


df_com = pd.read_excel(pasta+'Comunicacoes.xlsx',options={'strings_to_numbers': False}, converters={'Indexador':str})
df_com['Data_da_operacao'] = pd.to_datetime(df_com['Data_da_operacao']) 
df_env = pd.read_excel(pasta+'Envolvidos.xlsx',options={'strings_to_numbers': False}, converters={'Indexador':str})
df_oco = pd.read_excel(pasta+'Ocorrencias.xlsx',options={'strings_to_numbers': False})
df_ped = pd.read_excel(pasta+'Pedido.xlsx',options={'strings_to_numbers': False})
df_consolida = pd.merge(df_com,df_env, how='left', on='Indexador')
df_consolida = pd.merge(df_consolida,df_ped, how='left', on='cpfCnpjEnvolvido')
df_consolida.Justificativa.fillna("-?-",inplace=True) # CPFCNPJ que não constam do pedido

consolidado = pd.ExcelWriter(pasta+'pdConsolidados.xlsx', engine='xlsxwriter', options={'strings_to_numbers': False}, datetime_format='dd/mm/yyyy', date_format='dd/mm/yyyy')

table = pd.pivot_table(df_consolida,index=['Indexador','Data_da_operacao','cpfCnpjEnvolvido','nomeEnvolvido','informacoesAdicionais', 'Justificativa'],columns=['tipoEnvolvido'],margins=False)
df_pivot=table.stack()
#print(df_pivot)

gerar_planilha(consolidado,df_pivot,'INDEXADOR',indice=True)

table = pd.pivot_table(df_consolida,index=['cpfCnpjEnvolvido','nomeEnvolvido','Data_da_operacao', 'Justificativa','informacoesAdicionais','Indexador'],columns=['tipoEnvolvido'],margins=False)
df_pivot=table.stack()

gerar_planilha(consolidado,df_pivot,'CPFCNPJ',indice=True)

gerar_planilha(consolidado,df_consolida,'ComunicXEnvolvidos')
gerar_planilha(consolidado,df_com,'Comunicacoes')
gerar_planilha(consolidado,df_env,'Envolvidos')
gerar_planilha(consolidado,df_oco,'Ocorrencias')
gerar_planilha(consolidado,df_ped,'Pedido')
consolidado.save()


rede_rel = pd.ExcelWriter(pasta+'rede_de_rel.xls', engine = 'openpyxl', options={'strings_to_numbers': False}, datetime_format='dd/mm/yyyy', date_format='dd/mm/yyyy')

conx = sqlite3.connect(":memory:")  # ou use :memory: para botÃ¡-lo na memÃ³ria RAM
curs = conx.cursor()            

"""
r= curs.execute(sql)
r= curs.fetchall()"""


df_ped.to_sql('Pedido',conx, index=False,if_exists="replace")
df_env.to_sql('Envolvidos',conx, index=False,if_exists="replace")

sql  = 'select '
sql += " REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,'.',''),'-',''),'/','') as cpfcnpj,"
sql += ' case length(REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,".",""),"-",""),"/","")) when 14 then "PJ" when 11 then "PF" else "-" end as tipo,'
sql += ' upper(Empresa_nome) as nome, 0 as camada, "" as [componente/grupo],'   
sql += ' 0 as situacao, 0 as [sexo(PF)/Matriz-Filial(PJ)], 0 as [servidor(PF)/nat.jur(PJ)],'
sql += ' 0 as [salario<2min], 0 as OB, 0 as pad, 0 as [PF candidato], 0 as [CEIS/CEPIM], 0 as doadorTSE, 0 as CadUnico, 0 as Falecido'
sql += ' from Pedido '
sql += ' union select distinct'
sql += " REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,'.',''),'-',''),'/','') as cpfcnpj,"
sql += ' case length(REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,".",""),"-",""),"/","")) when 14 then "PJ" when 11 then "PF" else "-" end as tipo,'
sql += ' upper(nomeEnvolvido) as nome, 1 as camada, "" as [componente/grupo],'   
sql += ' 0 as situacao, 0 as [sexo(PF)/Matriz-Filial(PJ)], 0 as [servidor(PF)/nat.jur(PJ)],'
sql += ' 0 as [salario<2min], 0 as OB, 0 as pad, 0 as [PF candidato], 0 as [CEIS/CEPIM], 0 as doadorTSE, 0 as CadUnico, 0 as Falecido'
sql += ' from Envolvidos where cpfCnpjEnvolvido not in (select cpfCnpjEnvolvido from Pedido) '
sql += ' union select distinct "" as cpfcnpj, "CC" as tipo,'
sql += ' agenciaEnvolvido || "-"|| contaEnvolvido as nome, 1 as camada, "" as [componente/grupo],'
sql += ' 0 as situacao, 0 as [sexo(PF)/Matriz-Filial(PJ)], 0 as [servidor(PF)/nat.jur(PJ)],'
sql += ' 0 as [salario<2min], 0 as OB, 0 as pad, 0 as [PF candidato], 0 as [CEIS/CEPIM], 0 as doadorTSE, 0 as CadUnico, 0 as Falecido'
sql += ' from Envolvidos where agenciaEnvolvido not in ("-","0")'

df_rr = pd.read_sql(sql,conx)
gerar_planilhaXLS(rede_rel, df_rr,'cpfcnpj')

sql = 'select distinct indexador, '
sql += ' (agenciaEnvolvido || "-"|| contaEnvolvido) as CC'
sql += ' from Envolvidos where agenciaEnvolvido not in ("-","0") '
df_cc = pd.read_sql(sql,conx)
df_cc.to_sql('Contas',conx, index=False,if_exists="replace")

sql  = 'select distinct'
sql += " REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,'.',''),'-',''),'/','') as cpfcnpj1,"
sql += 'upper(nomeEnvolvido) as nome1, contas.CC as cpfcnpj2, "" as nome2, 1 as camada, "CC" as [descrição]'
sql += ' from Envolvidos inner join contas'
sql += ' on Envolvidos.indexador = contas.indexador'
df_rr = pd.read_sql(sql,conx)
gerar_planilhaXLS(rede_rel, df_rr,'ligacoes')

sql  = 'select distinct '
sql += " cpfCnpjEnvolvido as cpfcnpj1,"
sql += ' upper(nomeEnvolvido) as nome1, contas.CC as cpfcnpj2, "" as nome2, 1 as camada1, 1 as camada2,"CC" as [descrição], '
sql += ' case length(REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,".",""),"-",""),"/","")) when 14 then "Escritório" when 11 then "Profissional (masculino)" else "-" end as tipo1,'
sql += ' "Conta" as tipo2, "Nenhum" as cor_situacao1, "Nenhum" as cor_situacao2'
sql += ' from Envolvidos inner join contas'
sql += ' on Envolvidos.indexador = contas.indexador'
df_rr = pd.read_sql(sql,conx)
gerar_planilhaXLS(rede_rel, df_rr,'ligacoes')

df_rr.to_csv("I2.csv",sep=';', header=True, encoding= 'utf-8', decimal=',', index=False )
#to_i2(df=df_rr,arquivo='rede_rel.anx')

gerar_planilhaXLS(rede_rel, df_rr,'I2')

#df_rr = pd.DataFrame(r, columns=[i[0] for i in curs.description])

rede_rel.save()
print('ok')

