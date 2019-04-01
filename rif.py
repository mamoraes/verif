# -- coding: utf-8 --'

import pandas as pd
import os
import textwrap
import string
import unicodedata
import sys
import sqlite3
import easygui
import re
import copy
import json


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
        if self.nome.upper()=='grupos'.upper() or self.nome.upper()=='vinculos'.upper(): # um novo é criado vazio, uma vez que não vem do COAF
            return(True)
        else:
            return (os.path.isfile(self.nomearq()))
    def estr_completa(self, outra_estr=[]):
        return (all(elem.upper() in self.estr_upper()  for elem in outra_estr))
    def exibir(self):
        strestr =','.join(self.estr)
        return (self.nome + ': '+strestr)
def help_estruturas(estruturas):
    print('Estruturas esperadas das planilhas:')
    for e in estruturas:
        print('  '+e.exibir())

class log:
    def __init__(self):
        self.logs=u''
    def gravalog(self,linha):
        print(linha)
        self.logs += linha +'\n'
    def lelog(self):
        return self.logs

lg=log()
        
com = estrutura('Comunicacoes',["Indexador","Data_do_Recebimento","Data_da_operacao","DataFimFato","cpfCnpjComunicante","nomeComunicante","CidadeAgencia","UFAgencia","NomeAgencia","NumeroAgencia","informacoesAdicionais","CampoA","CampoB","CampoC","CampoD","CampoE"])
env = estrutura('Envolvidos',["Indexador","cpfCnpjEnvolvido","nomeEnvolvido","tipoEnvolvido","agenciaEnvolvido","contaEnvolvido","DataAberturaConta","DataAtualizacaoConta","bitPepCitado","bitPessoaObrigadaCitado","intServidorCitado"])
oco = estrutura('Ocorrencias',["Indexador","Ocorrencia"])
# opcionais
gru = estrutura('Grupos',['cpfCnpjEnvolvido','Grupo','Detalhe'])        
vin = estrutura('Vinculos',['cpfCnpjEnvolvido','cpfCnpjVinculado','Descricao'])        


estruturas = [com,env,oco,gru,vin]
#help_estruturas(estruturas)

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


def soDigitos(texto):
    return  re.sub("[^0-9]", "",texto)

def estimarFluxoDoDinheiro(tInformacoesAdicionais):
    #normalmente aparece algo como R$ 20,8 Mil enviada para Jardim Indústria e Comércio -  CNPJ 606769xxx
    #inicialmente quebramos o texto por R$ e verifica quais são seguidos por CPF ou CNPJ
    #pega o texto da coluna InformacoesAdicionais do arquivo Comunicacoes.csv e tenta estimar o valor para cada cpf/cnpj
    #normalmente aparece algo como R$ 20,8 Mil enviada para Indústria e Comércio -  CNPJ 6067xxxxxx
    #inicialmente quebramos o texto por R$ e verifica quais são seguidos por CPF ou CNPJ
    #retorna dicionário
    #como {'26106949xx': 'R$420 MIL RECEBIDOS, R$131 MIL POR', '68360088xxx': 'R$22 MIL, RECEBIDAS'}    
    #lista = re.sub(' +', ' ',tInformacoesAdicionais).upper().split('R$')
    t = re.sub(' +', ' ',tInformacoesAdicionais).upper()
    lista = t.split('R$') 
    listaComTermoCPFCNPJ = []
    for item in lista:
        if 'CPF' in item or 'CNPJ' in item:
            listaComTermoCPFCNPJ.append(item.strip())

    listaValores = []
    valoresDict = {}
    for item in listaComTermoCPFCNPJ:
        valorPara = ''
        cpn = ''
        le = item.split(' ')
        valor = 'R$' + le[0] #+ ' ' + le[1] # + ' ' + le[2]
        if le[1].upper().rstrip(',').rstrip('S').rstrip(',') in ('MIL','MI','RECEBIDO','RECEBIDA','ENVIADA','RETIRADO','DEPOSITADO','CHEQUE'):
            valor += ' ' + le[1]   
        if le[2].upper().rstrip(',').rstrip('S') in ('MIL','MI','RECEBIDO','RECEBIDA','ENVIADA','RETIRADO','DEPOSITADO','CHEQUE'):
            valor += ' ' + le[2]
        if 'CPF' in item:
            aux1 = item.split('CPF ')
            try:
                aux2  = aux1[1].split(' ') 
                cpn = soDigitos(aux2[0])
            except: 
                pass
        elif 'CNPJ' in item:
            aux1 = item.split('CNPJ ')
            try:
                aux2  = aux1[1].split(' ') 
                cpn = soDigitos(aux2[0])
            except:
                pass
        if cpn:
            listaValores.append(valorPara)
            if cpn in valoresDict:
                v = valoresDict[cpn]
                v.add(valor)
                valoresDict[cpn] = v
            else:
                valoresDict[cpn] = set([valor,])
    d = {}
    for k,v in valoresDict.items():
        d[k] = ', '.join(v)
    return d
#.def estimaFluxoDoDinheiro(t):



def consolidar_pd(pasta):
    """Processa as planilhas comunicacoes, envolvidos, ocorrencias e grupo em planilhas com agrupamento """
    arq = com.nomearq() # Comunicacoes
    try:
        df_com = pd.read_excel(arq,options={'strings_to_numbers': False}, converters={'Indexador':str})
        df_com['Indexador'] = pd.to_numeric(df_com['Indexador'], errors='coerce')
        df_com['Data_da_operacao'] = pd.to_datetime(df_com['Data_da_operacao']) 
        if not com.estr_completa(df_com.columns):
            print (com.estr_upper())
            mostra_erro('O arquivo '+arq+ ' não contém todas as colunas necessárias: ')
            raise ('Estrutura incompleta')
        lg.gravalog('Arquivo '+arq+' lido.')
    except Exception as exc:
        print('Erro ao ler o arquivo '+arq+'\n'+str(type(exc)))
        
    arq = env.nomearq() # Envolvidos
    try:    
        df_env = pd.read_excel(arq,options={'strings_to_numbers': False}, converters={'Indexador':str})
        df_env['Indexador'] = pd.to_numeric(df_env['Indexador'], errors='coerce')
        df_env = df_env[pd.notnull(df_env['Indexador'])]
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
        df_oco['Indexador'] = pd.to_numeric(df_oco['Indexador'], errors='coerce')
        df_oco = df_oco[pd.notnull(df_oco['Indexador'])]
        dictOco ={}
        dictOco2 ={}
        for r in df_oco.itertuples(index=False):
            if r.Indexador in dictOco:
                s=dictOco[r.Indexador]
                s+= "; "+r.Ocorrencia
                dictOco[r.Indexador]= s
            else:
                dictOco[r.Indexador]= r.Ocorrencia
        dictOco2['Indexador']=[]
        dictOco2['Ocorrencia']=[]
        for k,v in dictOco.items():
            dictOco2['Indexador'].append(k)
            dictOco2['Ocorrencia'].append(v)

        df_oco2=pd.DataFrame.from_dict(dictOco2)            
        
        if not oco.estr_completa(df_oco.columns):
            print (oco.estr_upper())
            mostra_erro('O arquivo '+arq+ ' não contém todas as colunas necessárias: ')
            raise ('Estrutura incompleta')
        lg.gravalog('Arquivo '+arq+' lido.')
    except Exception as exc:
        lg.gravalog('Erro ao ler o arquivo '+arq+'\n'+str(type(exc)))
        
    arq = gru.nomearq()  # Pedido
    if not os.path.isfile(arq): # criar arquivo vazio
        consolidado = pd.ExcelWriter(arq, engine='xlsxwriter', options={'strings_to_numbers': False}, datetime_format='dd/mm/yyyy', date_format='dd/mm/yyyy')
        gerar_planilha(consolidado, pd.DataFrame(columns=gru.estr), gru.nome, indice=False)
        consolidado.save()
        lg.gravalog('O arquivo '+arq+' não foi encontrado. Um novo foi criado com as colunas '+ gru.exibir())
    try:
        df_gru = pd.read_excel(arq,options={'strings_to_numbers': False})
        if not gru.estr_completa(df_gru.columns):
            print (gru.estr_upper())
            mostra_erro('O arquivo '+arq+ ' não contém todas as colunas necessárias: ')
            raise ('Estrutura incompleta')        
        lg.gravalog('Arquivo '+arq+' lido.')
    except Exception as exc:
        lg.gravalog('Erro ao ler o arquivo '+arq+'\n'+str(type(exc)))

    arq = vin.nomearq()  # Vinculos
    if not os.path.isfile(arq): # criar arquivo vazio
        consolidado = pd.ExcelWriter(arq, engine='xlsxwriter', options={'strings_to_numbers': False}, datetime_format='dd/mm/yyyy', date_format='dd/mm/yyyy')
        gerar_planilha(consolidado, pd.DataFrame(columns=vin.estr), vin.nome, indice=False)
        consolidado.save()
        lg.gravalog('O arquivo '+arq+' não foi encontrado. Um novo foi criado com as colunas '+ vin.exibir())
    try:
        df_vin = pd.read_excel(arq,options={'strings_to_numbers': False})
        if not vin.estr_completa(df_vin.columns):
            print (vin.estr_upper())
            mostra_erro('O arquivo '+arq+ ' não contém todas as colunas necessárias: ')
            raise ('Estrutura incompleta')        
        lg.gravalog('Arquivo '+arq+' lido.')
    except Exception as exc:
        lg.gravalog('Erro ao ler o arquivo '+arq+'\n'+str(type(exc)))            
        
    print('Consolidando')
    arq = os.path.join(pasta,'RIF_consolidados.xlsx') 
    porGrupo = len(df_gru['Grupo'].unique()) > 1  
    try:
        df_consolida = pd.merge(df_com,df_env, how='left', on='Indexador')
        df_consolida = pd.merge(df_consolida,df_oco2, how='left', on='Indexador')
        df_consolida = pd.merge(df_consolida,df_gru, how='left', on='cpfCnpjEnvolvido')
        df_consolida.Detalhe.fillna("-?-",inplace=True) # CPFCNPJ que não constam do grupo

        consolidado = pd.ExcelWriter(arq, engine='xlsxwriter', options={'strings_to_numbers': False}, datetime_format='dd/mm/yyyy', date_format='dd/mm/yyyy')
        if porGrupo:  # tem agrupamentos
            table = pd.pivot_table(df_consolida,index=['Grupo','Indexador','Data_da_operacao','cpfCnpjEnvolvido','nomeEnvolvido','informacoesAdicionais', 'Detalhe','Ocorrencia'],columns=['tipoEnvolvido'],margins=False)
        else:
            table = pd.pivot_table(df_consolida,index=['Indexador','Data_da_operacao','cpfCnpjEnvolvido','nomeEnvolvido','informacoesAdicionais', 'Detalhe','Ocorrencia'],columns=['tipoEnvolvido'],margins=False)
        df_pivot=table.stack()
    except Exception as exc:
         lg.gravalog('Erro ao consolidar planilhas no arquivo '+arq+'\n'+str(type(exc)))

    dicAdic = {'Indexador':[],'cpfCnpjEnvolvido':[],'valor':[]}
    valoresPorIndexador = {}
    for row in df_com.itertuples(index=False):
        if row.informacoesAdicionais:
            valoresPorIndexador[row.Indexador] = estimarFluxoDoDinheiro(str(row.informacoesAdicionais))
    for k, v in valoresPorIndexador.items():
        if v != {}:
            for kk, vv in v.items(): 
                dicAdic['Indexador'].append(k)
                dicAdic['cpfCnpjEnvolvido'].append(kk)
                dicAdic['valor'].append(vv)
    df_Adic = pd.DataFrame.from_dict(dicAdic)

    try:
        gerar_planilha(consolidado,df_pivot,'INDEXADOR',indice=True)

        if porGrupo:  # tem agrupamentos
            table = pd.pivot_table(df_consolida,index=['Grupo','cpfCnpjEnvolvido','nomeEnvolvido','Data_da_operacao', 'Detalhe','informacoesAdicionais','Indexador','Ocorrencia'],columns=['tipoEnvolvido'],margins=False)
        else:
            table = pd.pivot_table(df_consolida,index=['cpfCnpjEnvolvido','nomeEnvolvido','Data_da_operacao', 'Detalhe','informacoesAdicionais','Indexador','Ocorrencia'],columns=['tipoEnvolvido'],margins=False)
        df_pivot=table.stack()

        gerar_planilha(consolidado,df_pivot,'CPFCNPJ',indice=True)

        gerar_planilha(consolidado,df_consolida,'ComunicXEnvolvidos')
        gerar_planilha(consolidado,df_com,'Comunicacoes')
        gerar_planilha(consolidado,df_env,'Envolvidos')
        gerar_planilha(consolidado,df_oco2,'Ocorrencias')
        gerar_planilha(consolidado,df_gru,'Grupos')
        gerar_planilha(consolidado,df_vin,'Vinculos')
        gerar_planilha(consolidado,df_Adic,'InfoAdicionais')
    except Exception as exc:
        lg.gravalog('Erro ao gerar planilhas para o arquivo '+arq+'\n'+str(type(exc)))
    
    try:
        consolidado.save()
    except Exception as exc:
        lg.gravalog('Erro ao gravar o arquivo '+arq+'\n'+str(type(exc)))
    lg.gravalog("Planilhas consolidadas: "+arq)
    return df_gru, df_env, df_com, df_oco2, df_vin

def exportar_rede_rel(pasta,dfgru, dfenv):
# criando as tabelas no SQLITE
    try:
        conx = sqlite3.connect(":memory:")  # ou use :memory: para botÃ¡-lo na memÃ³ria RAM
        curs = conx.cursor()            
    except Exception as exc:
        lg.gravalog('Erro criar conexão com SQLITE memory\n'+str(type(exc)))
        
    try:
        dfgru.to_sql('Pedido',conx, index=False,if_exists="replace")
    except Exception as exc:
        lg.gravalog('Erro carregar grupos no SQLITE memory\n'+str(type(exc)))
    try:
        dfenv.to_sql('Envolvidos',conx, index=False,if_exists="replace")
    except Exception as exc:
        lg.gravalog('Erro carregar envolvidos no SQLITE memory\n'+str(type(exc)))
   

    sql  = 'select '
    sql += " REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,'.',''),'-',''),'/','') as cpfcnpj,"
    sql += ' case length(REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,".",""),"-",""),"/","")) when 14 then "PJ" when 11 then "PF" else "-" end as tipo,'
    sql += ' upper(Grupo) as nome, 0 as camada, "" as [componente/grupo],'   
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
    #dicRede = {'cpfcnpj'=[],'tipo'=[],'nome'=[],'camada'=[],}
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
    lg.gravalog("Rede de relacionamento gerada: "+arq)


def validar_pasta(pasta, planilhas):
    for p in planilhas:
        p.mudar_pasta(pasta)
        if not p.arquivo_existe(): return(p.nomearq())
    return ''

def mostra_erro(msg):
    print(msg)
    easygui.msgbox(msg)
    return

def parse_args():
    """ Obtém do usuário a definição da pasta onde estão as planilhas a serem processadas
    e persiste num arquivo de configuração json
    """
    import json
    import argparse

    #futuro
    procGrupos = True
    procVinculos = True
    procContas = True

    proc_grupos_msg = "Processar grupos"
    proc_vinculos_msg = "Processar outros vinculos"
    proc_contas_msg = "Processar contas"

    stored_args = {}
    # usar o nome do script sem a extensão para formar o nome do arquivo json
    script_name = os.path.splitext(os.path.basename(__file__))[0]

    args_file = "{}-args.json".format(script_name)
    pasta=''
    # ler os parâmetros persistidos, gravados no arquivo json
    if os.path.isfile(args_file):
        with open(args_file) as data_file:
            stored_args = json.load(data_file)
            procContas = stored_args.get('procContas')
            procVinculos = stored_args.get('procVinculos')
            procGrupos = stored_args.get('procGrupos')

    # se o programa for chamado sem especificar a pasta de origem, abrir GUI para obtê-la do usuário
    if len(sys.argv)<=1:
        pasta=stored_args.get('Pasta')
        pasta= gui_pasta(pasta)
        

    # definir o processamento de parâmetros passados pela linha de comando
    desc = "Processa planilhas de RIF (comunicações, envolvidos, ocorrencias e grupo), \ngerando planilhas com agrupamento e com dados para i2"
    file_help_msg = "Nome da pasta onde estão as planilhas"


    linha_comando = argparse.ArgumentParser(description=desc)
    linha_comando.add_argument(
            '--pasta', 
            action = 'store',
            dest= 'pasta',
            help=file_help_msg, 
            default=pasta,
            required = False)
    linha_comando.add_argument(
            '-p', 
            action = 'store',
            dest= 'pasta',
            help=file_help_msg, 
            default=pasta,
            required = False)
    linha_comando.add_argument(
            '-e', 
            action = help_estruturas(estruturas),
            help='Exibe as estruturas esperadas das planilhas de entrada', 
            required = False)
    linha_comando.add_argument(
            '-g', 
            action = 'store',
            dest= 'procGrupos',
            help=proc_grupos_msg, 
            default=procGrupos,
            required = False)
    linha_comando.add_argument(
            '-c', 
            action = 'store',
            dest= 'procContas',
            help=proc_contas_msg, 
            default=procContas,
            required = False)

    
    args = linha_comando.parse_args()
    if not args.pasta:
        args.pasta = os.getcwd() # pasta atual do script como default
    if not os.path.isdir(args.pasta): # ver sé pasta
        print(args.pasta + ' nao é pasta')
        exit(1)
    
    args.procContas = gui_sn(procContas, 'Gerar nós de contas correntes?')
    args.procGrupos = gui_sn(procGrupos,'Gerar nós de grupo?')
    args.procVinculos = gui_sn(procVinculos,'Gerar vínculos adicionais a partir da planilha?')
    
    # persistir os parâmetros
    with open(args_file, 'w') as data_file:
        # Using vars(args) returns the data as a dictionary
        json.dump(vars(args), data_file)
    
    return args
  #  display_message()
      
def executar(pasta):
    # Selecionar a pasta de origem/destino, validá-la e executar a consolidação e a geração das planilhas
    if not pasta:
        args=parse_args()
        pasta=args.pasta
    procContas = args.procContas
    procGrupos = args.procGrupos
    procVinculos = args.procVinculos
    print('Pasta selecionada: '+pasta)
    print('Gerar nós de conta: '+'sim' if procContas else 'não')
    print('Gerar nós de grupo: '+'sim' if procGrupos else 'não')
    print('Gerar vínculos adicionais a partir de planilha: '+'sim' if procVinculos else 'não')
    pl = validar_pasta(pasta,estruturas)
    if pl =='':
        dfGrupos, dfEnvolvidos, dfComunicacoes, dfOcorrencias, dfVinculos = consolidar_pd(pasta) # gera planilhas, retorna dataframes
        exportar_rede_rel(pasta, dfGrupos, dfEnvolvidos) # gera planilha para i2
        criarArquivoMacrosGrafo(pasta, procContas, procGrupos, procVinculos, dfGrupos, dfEnvolvidos, dfComunicacoes, dfOcorrencias, dfVinculos) # gera json para o Macros
        logs=lg.lelog()
        easygui.msgbox(logs)
    else:
        mostra_erro('Arquivo '+ pl +' não encontrado nesta pasta:\n'+pasta)
        sys.exit(1)
    
    return logs  # mensagens sobre a execução

def gui_pasta(pasta):
    # GUI para obter a pastas escolhida pelo usuário
    nomepasta= easygui.diropenbox(default=pasta,msg='Selecione a pasta onde estão as planilhas Comunicações, Envolvidos e Ocorrências',title='Gera planilhas a partir de dados de RIF')
    if not nomepasta:
        exit(1)
    return(nomepasta)
def gui_sn(sn,texto:str):
    resp= easygui.choicebox(msg=texto,choices=['Sim','Não'],preselect= 0 if sn else 1)
    return(resp=='Sim')

def pasta_valida(pasta): 
    # verifica se a pasta escolida é válida
    pl = validar_pasta(pasta,estruturas)
    if pl =='':
        return True
    else:
        return False

def criarArquivoMacrosGrafo(pasta, processaConta, processaGrupos, processaVinculos, dfGrupos, dfEnvolvidos, dfComunicacoes, dfOcorrencias, dfVinculos):
    '''gera o arquivo do Macros a partir dos dataframes'''
    
    # cria tabela de nós
    nos=[]

    #procura tipoEnvolvido=Titular, para por no tooltip do nó 
    titularOcorrencia = {}
    conta ={}
    for row in dfEnvolvidos.itertuples(index=False):
        if row.tipoEnvolvido.upper()=='TITULAR':
            titularOcorrencia[row.Indexador] =  'Titular: ' + row.nomeEnvolvido.strip() + '(' + row.cpfCnpjEnvolvido  +')' 

        if processaConta:    
            if row.contaEnvolvido !='-' and row.contaEnvolvido !='0':
                nroConta = 'CC: '+str(row.agenciaEnvolvido)+'/'+str(row.contaEnvolvido)
                campoA = (dfComunicacoes['CampoA'].where(dfComunicacoes['Indexador']==row.Indexador).sum())
                if nroConta not in conta:
                    conta[nroConta] = campoA
                else:
                    conta[nroConta] += campoA
    if processaConta:            
        for key, value in conta.items():
            item = {"id":key, "tipo":"ENT", "cor":"Green","label":"R$ " + str(value),"camada":0,"situacao":0}
            item['texto_tooltip'] = ''
            nos.append(copy.deepcopy(item))

    for row in dfComunicacoes.itertuples(index=False):
        if (str(row.Indexador)=='nan'): continue
        campos = str(row)[len('Pandas'):]
        item = {"id":"ENT_" + str(row.Indexador),"tipo":"ENT","sexo":0,"label":"COAF R$" + str(row.CampoA),"camada":0,"situacao":0,"m1":0,"m2":0,"m3":0,"m4":0,"m5":0,"m6":0,"m7":0,"m8":0,"m9":0,"m10":0,"m11":0} #,"posicao":{"x":0,"y":0}}
        d = dfOcorrencias[dfOcorrencias.Indexador==row.Indexador]['Ocorrencia']
        campos = titularOcorrencia.get(row.Indexador,'') + ' ' + campos + ' . Ocorrência(s): ' + '; '.join(d)
        #campos = ' '.join(campos.split(r'\t')) #funciona
        #campos = re.sub('\t',' ',campos) #não funciona
        campos = campos.replace(r'\t',' ')
        campos = campos.replace(r'\x96', ' ').replace(r'\x92', ' ')
        item['texto_tooltip'] = str(re.sub(' +', ' ',campos)) #remove espaços a mais
        nos.append(copy.deepcopy(item))
    
    valoresPorIndexador = {}
    for row in dfComunicacoes.itertuples(index=False):
        if row.informacoesAdicionais:
            valoresPorIndexador[row.Indexador] = estimarFluxoDoDinheiro(str(row.informacoesAdicionais))
            
    cpfcnpjset = set()
    for row in dfEnvolvidos.itertuples(index=False):
        cpfcnpj = row.cpfCnpjEnvolvido.replace('.','').replace('-','').replace('/','')
        cpfcnpjset.add(cpfcnpj)
    
    cs = set()
    for row in dfEnvolvidos.itertuples(index=False):
        cpfcnpj = row.cpfCnpjEnvolvido.replace('.','').replace('-','').replace('/','')
        if cpfcnpj not in cs:
            cs.add(cpfcnpj)
            #item = {"id":cpfcnpj,"tipo":"","sexo":0,"label":row.nomeEnvolvido,"camada":1,"situacao":0,"m1":0,"m2":0,"m3":0,"m4":0,"m5":0,"m6":0,"m7":0,"m8":0,"m9":0,"m10":0,"m11":0,"posicao":{"x":0,"y":0}}
            item = {"id":cpfcnpj} #passando somente o id, na hora de importar o navegador vai pedir os dados dos nós para o servidor
            if len(cpfcnpj)==11:
                item['tipo'] = 'PF'
            else:
                item['tipo'] = 'PJ'
                item['sexo'] = 1
            item['texto_tooltip'] = ''
            nos.append(copy.deepcopy(item))
        
    #cria tabela de ligacoes
    ligacoes = []
    ligadict = {}
    ligaContadict = {}

    det = {} # detalhes
    grp = {} # grupos
    idDet=0
    idGru=0
    dmatch = pd.merge(dfEnvolvidos, dfGrupos, how='inner', on='cpfCnpjEnvolvido')
    
    if processaGrupos:
        for gr in dmatch['Grupo'].unique():
            idGru += 1
            grp[gr] = idGru
            texto = gr[:80]
            item = {"id":'#'+str(idGru), "tipo":"ENT", "cor":"Blue","label":texto,"camada":0,"situacao":0}
            nos.append(copy.deepcopy(item))

        for row in dmatch.itertuples(index=False):
            gr = row.Grupo
            obDetalhe = row.Detalhe[:80]
            cpfcnpj = row.cpfCnpjEnvolvido.replace('.','').replace('-','').replace('/','')
            item = {"origem": cpfcnpj, "destino":'#'+str(grp[gr]),"cor":"Blue","camada":1,"tipoDescricao":{"0":obDetalhe}}
            ligacoes.append(copy.deepcopy(item))
    #dfEnvolvidos.sort_values(by=['Indexador', 'cpfCnpjEnvolvido'])
    for row in dfEnvolvidos.itertuples(index=False):
        if (str(row.Indexador)=='nan'): continue
        cpfcnpj = row.cpfCnpjEnvolvido.replace('.','').replace('-','').replace('/','')
        #origemDestino = "ENT_"+row.Indexador+"-"+cpfcnpj
        origemDestino = str(row.Indexador)+"-"+cpfcnpj
        if origemDestino not in ligadict:
            ligadict[origemDestino] = set([row.tipoEnvolvido,])
        else:
            ligadict[origemDestino].add(row.tipoEnvolvido)
        #conta
        if row.contaEnvolvido !='-' and row.contaEnvolvido !='0':
            nroConta = 'CC: '+str(row.agenciaEnvolvido)+'/'+str(row.contaEnvolvido)
            origemDestino = nroConta+"-"+cpfcnpj
            if origemDestino not in ligadict:
                ligaContadict[origemDestino] = set([row.tipoEnvolvido,])
            else:
                ligaContadict[origemDestino].add(row.tipoEnvolvido)
            
    for k,v in ligadict.items():
        origem = k.split('-')[0]
        destino = k.split('-')[1]
        valor = valoresPorIndexador.get(origem,{}).get(destino,'')
        if valor: valor = '-' + valor
        #descricao = ','.join(sorted([k for k in v if k!='Outros'])) + valor
        descricao = ','.join(sorted(v)) + valor
        item = {"origem":'ENT_' + origem, "destino":destino,"cor":"Yellow","camada":1,"tipoDescricao":{"0":descricao}}
        ligacoes.append(copy.deepcopy(item))
        
    for k,v in ligaContadict.items():
        origem = k.split('-')[0]
        destino = k.split('-')[1]
        valor = valoresPorIndexador.get(origem,{}).get(destino,'')
        if valor: valor = '-' + valor
        #descricao = ','.join(sorted([k for k in v if k!='Outros'])) + valor
        descricao = ','.join(sorted(v)) + valor
        item = {"origem": origem, "destino":destino,"cor":"Blue","camada":1,"tipoDescricao":{"0":descricao}}
        ligacoes.append(copy.deepcopy(item))

    if processaVinculos:
        for row in dfVinculos.itertuples(index=False):
            cpfcnpj = row.cpfCnpjEnvolvido.replace('.','').replace('-','').replace('/','')
            if cpfcnpj not in cs:
                cs.add(cpfcnpj)
                item = {"id":cpfcnpj} #passando somente o id, na hora de importar o navegador vai pedir os dados dos nós para o servidor
                if len(cpfcnpj)==11:
                    item['tipo'] = 'PF'
                else:
                    item['tipo'] = 'PJ'
                    item['sexo'] = 1
                item['texto_tooltip'] = ''
                nos.append(copy.deepcopy(item))
        for row in dfVinculos.itertuples(index=False):
            cpfcnpj = row.cpfCnpjEnvolvido.replace('.','').replace('-','').replace('/','')
            cpfcnpjvinc = row.cpfCnpjVinculado.replace('.','').replace('-','').replace('/','')
            descricao = row.Descricao[:80]
            item = {"origem": cpfcnpj, "destino":cpfcnpjvinc,"cor":"Blue","camada":1,"tipoDescricao":{"0":descricao}}
            ligacoes.append(copy.deepcopy(item))

            if cpfcnpj not in cs: # incluir CPFCNPJ
                cs.add(cpfcnpj)
                item = {"id":cpfcnpj} #passando somente o id, na hora de importar o navegador vai pedir os dados dos nós para o servidor
                if len(cpfcnpj)==11:
                    item['tipo'] = 'PF'
                else:
                    item['tipo'] = 'PJ'
                    item['sexo'] = 1
                item['texto_tooltip'] = ''
                nos.append(copy.deepcopy(item))
            cpfcnpj = cpfcnpjvinc
            if cpfcnpj not in cs: # incluir CPFCNPJ
                cs.add(cpfcnpj)
                item = {"id":cpfcnpj} #passando somente o id, na hora de importar o navegador vai pedir os dados dos nós para o servidor
                if len(cpfcnpj)==11:
                    item['tipo'] = 'PF'
                else:
                    item['tipo'] = 'PJ'
                    item['sexo'] = 1
                item['texto_tooltip'] = ''
                nos.append(copy.deepcopy(item))

    textoJson=json.dumps({'no': nos, 'ligacao':ligacoes}) #, ensure_ascii=False)
    #print(textoJson)
    #textoJson = textoJson.replace(r'\t',' ').replace(r'\x96',' ').replace(r'\x92', ' ') #isso dá erro
    arq = os.path.join(pasta, 'macros_grafo.json')
    with open(arq, 'wt', encoding='latin1') as out:
        out.write(textoJson)
    lg.gravalog("Grafo gerado para o sistema Macros: "+arq)

if __name__ == '__main__':
    # caso o script seja utilizado como um package por outro, basta chamar a função "executar" com um nome de pasta
    executar('')