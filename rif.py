# -- coding: utf-8 --'

import pandas as pd
import os
import textwrap
import string
import unicodedata
import sys
import sqlite3


class estrutura:
    def __init__(self,nome='',estr=[], pasta='./'):
        self.nome = nome
        self.estr=estr
        self.pasta=pasta
    def mudar_pasta(self,pasta):
        self.pasta =pasta        
    def xlsx(self):
        return(self.nome+'.xlsx')
    def estr_upper(self):
        result = []
        for elem in self.estr:
            result.append(elem.upper())
        return result
    def nomearq(self):
        return os.path.join(self.pasta,self.xlsx())
    def arquivo_existe(self):
        return (os.path.isfile(self.nomearq()))
    def estr_completa(self, outra_estr=[]):
        return (all(elem.upper() in self.estr_upper()  for elem in outra_estr))
class log:
    def __init__(self):
        self.logs=u''
    def gravalog(self,linha):
        print(linha)
        self.logs += linha +'/n'
    def lelog(self):
        return self.logs

lg=log()
        
ped = estrutura('Pedido',['cpfCnpjEnvolvido','Empresa_Nome','Justificativa'])        
com = estrutura('Comunicacoes',["Indexador","Data_do_Recebimento","Data_da_operacao","DataFimFato","cpfCnpjComunicante","nomeComunicante","CidadeAgencia","UFAgencia","NomeAgencia","NumeroAgencia","informacoesAdicionais","CampoA","CampoB","CampoC","CampoD","CampoE"])
env = estrutura('Envolvidos',["Indexador","cpfCnpjEnvolvido","nomeEnvolvido","tipoEnvolvido","agenciaEnvolvido","contaEnvolvido","DataAberturaConta","DataAtualizacaoConta","bitPepCitado","bitPessoaObrigadaCitado","intServidorCitado"])
oco = estrutura('Ocorrencias',["Indexador","Ocorrencia"])

estruturas = [com,env,oco,ped]

def removeAcentos(data):
  if data is None:
    return u''
#  if isinstance(data,str):
#    data = unicode(data,'latin-1','ignore')
  return ''.join(x for x in unicodedata.normalize('NFKD', data) if x in string.printable)

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
    print('linha= ',linha)
    descricao = linha[1 if umou2==1 else 3]
#     if descricao == '': #telefone ou endereco
#         descricao = carJuncao.join(node[4:].split('__'))
#     else:
#         if self.GNX.node[node]['tipo'] !='TEL':   
#             descricao = Obj.parseCPFouCNPJ(node) + carJuncao + carJuncao.join(textwrap.wrap(descricao,30))
    
    #dicTipo = {'TEL':u'Telefone', 'END':u'Local', 'PF':u'PF', 'PJ':u'PJ', 'PE':u'Edifício', 'ES':u'Edifício', 'CC':u'Conta','INF':u'Armário' }

    tipo = linha[7 if umou2==1 else 8]
    # tipoi2 = dicTipo[tipo]
    tipoi2= u'Escritório' 
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
    #chart = Pyanx_macros()
    noi2origem= {}
    noi2destino={}
    
    
    for idc, campos in df.iterrows():
      #  print('campos= ',campos)
               
        tipo, descricao, corSituacao = tipoi2F(linha=campos, umou2=1, carJuncao=' ')
        noi2origem[idc] = chart.add_node(entity_type=dicTiposIngles.get(tipo,''), label=(campos['cpfcnpj1']) + u'-' +(descricao))
        tipo, descricao, corSituacao = tipoi2F(linha=campos, umou2=2, carJuncao=' ')
        noi2destino[idc] = chart.add_node(entity_type=dicTiposIngles.get(tipo,''), label=(campos['cpfcnpj1']) + u'-' +(descricao))

        nomeLigacao = campos['descrição']
        chart.add_edge(noi2origem[idc], noi2destino[idc], removeAcentos(nomeLigacao))
       # idc += 1
        
    fstream = chart.createStream(layout='spring_layout',iterations=0) #não calcula posição

    retorno = fstream.getvalue()
    fstream.close()
    if arquivo is not None:
      f = open(arquivo, 'w')
      f.write(retorno)
      f.close()
    return retorno

def consolidar_pd(pasta):
    """Processa as planilhas comunicacoes, envolvidos, ocorrencias e pedido em planilhas com agrupamento """
    arq = com.nomearq() # Comunicacoes
    try:
        df_com = pd.read_excel(arq,options={'strings_to_numbers': False}, converters={'Indexador':str})
        df_com['Data_da_operacao'] = pd.to_datetime(df_com['Data_da_operacao']) 
        if not com.estr_completa(df_com.columns):
            
            mostra_erro('O arquivo '+arq+ ' não contém todas as colunas necessárias: ')
            raise ('Estrutura incompleta')
        lg.gravalog('Arquivo '+arq+' lido.')
    except Exception as exc:
        print('Erro ao ler o arquivo '+arq+'\n'+str(type(exc)))
        
    arq = env.nomearq() # Envolvidos
    try:    
        df_env = pd.read_excel(arq,options={'strings_to_numbers': False}, converters={'Indexador':str})
        if not env.estr_completa(df_env.columns):
            print (env.estr_upper())
            mostra_erro('O arquivo '+arq+ ' não contém todas as colunas necessárias: ')
            raise ('Estrutura incompleta')
        lg.gravalog('Arquivo '+arq+' lido.')
    except Exception as exc:
        lg.gravalog('Erro ao ler o arquivo '+arq+'\n'+str(type(exc)))
    
    arq = oco.nomearq() # Ocorrencias
    try:
        df_oco = pd.read_excel(arq,options={'strings_to_numbers': False})
        if not oco.estr_completa(df_oco.columns):
            print (oco.estr_upper())
            mostra_erro('O arquivo '+arq+ ' não contém todas as colunas necessárias: ')
            raise ('Estrutura incompleta')
    except Exception as exc:
        lg.gravalog('Erro ao ler o arquivo '+arq+'\n'+str(type(exc)))
        
    arq = ped.nomearq()  # Pedido
    if not os.path.isfile(arq): # criar arquivo vazio
        consolidado = pd.ExcelWriter(arq, engine='xlsxwriter', options={'strings_to_numbers': False}, datetime_format='dd/mm/yyyy', date_format='dd/mm/yyyy')
        gerar_planilha(consolidado, pd.DataFrame(columns=ped.estr), ped.nome, indice=False)
        consolidado.save()
        lg.gravalog('O arquivo '+arq+' não foi encontrado. Um novo foi criado com as colunas ', ped.estr)
    try:
        df_ped = pd.read_excel(arq,options={'strings_to_numbers': False})
        if not ped.estr_completa(df_ped.columns):
            print (ped.estr_upper())
            mostra_erro('O arquivo '+arq+ ' não contém todas as colunas necessárias: ')
            raise ('Estrutura incompleta')        
        lg.gravalog('Arquivo '+arq+' lido.')
    except Exception as exc:
        lg.gravalog('Erro ao ler o arquivo '+arq+'\n'+str(type(exc)))
        
    lg.gravalog('Consolidando')
    arq = os.path.join(pasta,'RIF_consolidados.xlsx')    
    try:
        df_consolida = pd.merge(df_com,df_env, how='left', on='Indexador')
        df_consolida = pd.merge(df_consolida,df_ped, how='left', on='cpfCnpjEnvolvido')
        df_consolida.Justificativa.fillna("-?-",inplace=True) # CPFCNPJ que não constam do pedido

        consolidado = pd.ExcelWriter(arq, engine='xlsxwriter', options={'strings_to_numbers': False}, datetime_format='dd/mm/yyyy', date_format='dd/mm/yyyy')

        table = pd.pivot_table(df_consolida,index=['Indexador','Data_da_operacao','cpfCnpjEnvolvido','nomeEnvolvido','informacoesAdicionais', 'Justificativa'],columns=['tipoEnvolvido'],margins=False)
        df_pivot=table.stack()
    except Exception as exc:
         lg.gravalog('Erro ao consolidar planilhas no arquivo '+arq+'\n'+str(type(exc)))
       
    try:
        gerar_planilha(consolidado,df_pivot,'INDEXADOR',indice=True)

        table = pd.pivot_table(df_consolida,index=['cpfCnpjEnvolvido','nomeEnvolvido','Data_da_operacao', 'Justificativa','informacoesAdicionais','Indexador'],columns=['tipoEnvolvido'],margins=False)
        df_pivot=table.stack()

        gerar_planilha(consolidado,df_pivot,'CPFCNPJ',indice=True)

        gerar_planilha(consolidado,df_consolida,'ComunicXEnvolvidos')
        gerar_planilha(consolidado,df_com,'Comunicacoes')
        gerar_planilha(consolidado,df_env,'Envolvidos')
        gerar_planilha(consolidado,df_oco,'Ocorrencias')
        gerar_planilha(consolidado,df_ped,'Pedido')
    except Exception as exc:
        lg.gravalog('Erro ao gerar planilhas para o arquivo '+arq+'\n'+str(type(exc)))
    
    try:
        consolidado.save()
    except Exception as exc:
        lg.gravalog('Erro ao gravar o arquivo '+arq+'\n'+str(type(exc)))
    lg.gravalog("Planilhas geradas")
    return df_ped, df_env

def exportar_rede_rel(pasta,dfped, dfenv):
# criando as tabelas no SQLITE
    try:
        conx = sqlite3.connect(":memory:")  # ou use :memory: para botÃ¡-lo na memÃ³ria RAM
        curs = conx.cursor()            
    except Exception as exc:
        lg.gravalog('Erro criar conexão com SQLITE memory\n'+str(type(exc)))
        
    try:
        dfped.to_sql('Pedido',conx, index=False,if_exists="replace")
    except Exception as exc:
        lg.gravalog('Erro carregar pedidos no SQLITE memory\n'+str(type(exc)))
    try:
        dfenv.to_sql('Envolvidos',conx, index=False,if_exists="replace")
    except Exception as exc:
        lg.gravalog('Erro carregar envolvidos no SQLITE memory\n'+str(type(exc)))
   

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

    try:
        arq = os.path.join(pasta,'RIF_rede_de_rel.xls')    
        rede_rel = pd.ExcelWriter(arq, engine = 'openpyxl', options={'strings_to_numbers': False}, datetime_format='dd/mm/yyyy', date_format='dd/mm/yyyy')
    except Exception as exc:
        lg.gravalog('Erro abrir planilha de rede de relacionamento\n'+str(type(exc)))

    try:
        df_rr = pd.read_sql(sql,conx)
        gerar_planilhaXLS(rede_rel, df_rr,'cpfcnpj')
    except Exception as exc:
        lg.gravalog('Erro gerar planilha de rede de relacionamento CPFCNPJ\n'+str(type(exc)))
        
    try:
        sql = 'select distinct indexador, '
        sql += ' (agenciaEnvolvido || "-"|| contaEnvolvido) as CC'
        sql += ' from Envolvidos where agenciaEnvolvido not in ("-","0") '
        df_cc = pd.read_sql(sql,conx)
        df_cc.to_sql('Contas',conx, index=False,if_exists="replace")
    except Exception as exc:
        lg.gravalog('Erro gerar tabela auxiliar de contas\n'+str(type(exc)))
        
    try:    
        sql = 'select distinct'
        sql += " REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,'.',''),'-',''),'/','') as cpfcnpj1,"
        sql += 'upper(nomeEnvolvido) as nome1, contas.CC as cpfcnpj2, "" as nome2, 1 as camada, "CC" as [descrição]'
        sql += ' from Envolvidos inner join contas'
        sql += ' on Envolvidos.indexador = contas.indexador'
        df_rr = pd.read_sql(sql,conx)
        gerar_planilhaXLS(rede_rel, df_rr,'ligacoes')
    except Exception as exc:
        lg.gravalog('Erro gerar planilha de rede de relacionamento LIGACOES\n'+str(type(exc)))
    
    try:
        sql  = 'select distinct '
        sql += " cpfCnpjEnvolvido as cpfcnpj1,"
        sql += ' upper(nomeEnvolvido) as nome1, contas.CC as cpfcnpj2, "" as nome2, 1 as camada1, 1 as camada2,"CC" as [descrição], '
        sql += ' case length(REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,".",""),"-",""),"/","")) when 14 then "Escritório" when 11 then "Profissional (masculino)" else "-" end as tipo1,'
        sql += ' "Conta" as tipo2, "Nenhum" as cor_situacao1, "Nenhum" as cor_situacao2'
        sql += ' from Envolvidos inner join contas'
        sql += ' on Envolvidos.indexador = contas.indexador'
        df_rr = pd.read_sql(sql,conx)
        gerar_planilhaXLS(rede_rel, df_rr,'ligacoes')
    except Exception as exc:
        lg.gravalog('Erro gerar planilha de rede de relacionamento LIGACOES - complementares\n'+str(type(exc)))
        
    try:
        arqcsv = os.path.join(pasta,'I2.csv')
        df_rr.to_csv(arqcsv,sep=';', header=True, encoding= 'utf-8', decimal=',', index=False )
    except Exception as exc:
        lg.gravalog('Erro gerar csv de rede de relacionamento i2\n'+str(type(exc)))


    #to_i2(df=df_rr,arquivo='rede_rel.anx')
    try:
        gerar_planilhaXLS(rede_rel, df_rr,'I2')
    except Exception as exc:
        lg.gravalog('Erro gerar planilha de rede de relacionamento i2\n'+str(type(exc)))

#df_rr = pd.DataFrame(r, columns=[i[0] for i in curs.description])

    try:
        rede_rel.save()
    except Exception as exc:
        lg.gravalog('Erro ao gravar o arquivo '+arq+'\n'+str(type(exc)))
    lg.gravalog("Rede de relacionamento gerada:\n"+arq)


def validar_pasta(pasta, planilhas):
    for p in planilhas:
        p.mudar_pasta(pasta)
        if not p.arquivo_existe(): return(p.nomearq())
    return ''

def mostra_erro(msg):
    print(msg)
    return
    """    
    import wx
    app=wx.App()
    dlg=wx.MessageDialog(None,msg,'Error', wx.ICON_ERROR)
    dlg.ShowModal()
    dlg.Destroy()"""

def parse_args():
    """ Use ArgParser to build up the arguments we will use in our script
    Save the arguments in a default json file so that we can retrieve them
    every time we run the script.
    """
    import json
    import argparse

    stored_args = {}
    # get the script name without the extension & use it to build up
    # the json filename
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    args_file = "{}-args.json".format(script_name)
    # Read in the prior arguments as a dictionary
    if os.path.isfile(args_file):
        with open(args_file) as data_file:
            stored_args = json.load(data_file)
    
    desc = "Processa planilhas de RIF (comunicações, envolvidos, ocorrencias e pedido), \ngerando planilhas com agrupamento e com dados para i2"
    file_help_msg = "Nome da pasta onde estão as planilhas"

    my_cool_parser = argparse.ArgumentParser(description=desc)
    my_cool_parser.add_argument(
            '--pasta', 
            action = 'store',
            dest= 'pasta',
            help=file_help_msg, 
            default=stored_args.get('Pasta'),
            required = False)

    
    args = my_cool_parser.parse_args()
    if not args.pasta:
        args.pasta = os.getcwd()
    if not os.path.isdir(args.pasta):
        print(args.pasta + ' nao e pasta')
    
    # Store the values of the arguments so we have them next time we run
    with open(args_file, 'w') as data_file:
        # Using vars(args) returns the data as a dictionary
        json.dump(vars(args), data_file)
    
    return args
  #  display_message()
      
def executar(pasta,logs=''):
    if not pasta:
        args=parse_args()
        pasta=args.pasta
    print(pasta)
    pl = validar_pasta(pasta,estruturas)
    if pl =='':
        dped, denv = consolidar_pd(pasta)
        exportar_rede_rel(pasta, dped, denv)
        logs=lg.lelog()
    else:
        mostra_erro('Arquivo '+ pl +' não encontrado nesta pasta:\n'+pasta)
        sys.exit(1)

def pasta_valida(pasta):
    pl = validar_pasta(pasta,estruturas)
    if pl =='':
        return True
    else:
        return False


if __name__ == '__main__':
    executar('')