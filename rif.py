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
import xlsxwriter

# import pyanx

MAX_TAM_LABEL = 100  # nro  máximo de caracteres nos labels


class estrutura:  # especificações das planilhas
    def __init__(self, nome="", estr=[], pasta="./"):
        self.nome = nome
        self.estr = estr
        self.pasta = pasta

    def mudar_pasta(self, pasta):
        self.pasta = pasta

    def xlsx(self):
        return self.nome + ".xlsx"

    def estr_upper(self):
        result = []
        for elem in self.estr:
            result.append(elem.upper())
        return result

    def nomearq(self):
        return os.path.join(self.pasta, self.xlsx())

    def arquivo_existe(self):
        if (
            self.nome.upper() == "grupos".upper()
            or self.nome.upper() == "vinculos".upper()
        ):  # um novo é criado vazio, uma vez que não vem do COAF
            return True
        else:
            return os.path.isfile(self.nomearq())

    def estr_compativel(self, outra_estr=[]):
        return all(elem.upper() in self.estr_upper() for elem in outra_estr)

    def exibir(self):
        strestr = ",".join(self.estr)
        return self.nome + ": " + strestr


def help_estruturas(estruturas):
    print("Estruturas esperadas das planilhas:")
    for e in estruturas:
        print("  " + e.exibir())


class log:
    def __init__(self):
        self.logs = u""

    def gravalog(self, linha):
        print(linha)
        self.logs += linha + "\n"

    def lelog(self):
        return self.logs


class nodo:
    def __init__(self, id, label, tipo="ENT", tooltip="", fonte="RIF"):
        self.id = id
        self.tipo = tipo
        self.label = label
        self.cor = "Silver"
        self.sexo = 0
        self.m1 = 0
        self.m2 = 0
        self.situacao = ""
        self.dataOperacao = ""
        self.texto_tooltip = tooltip
        self.fonte = fonte
        self.camada = 0 if self.fonte == "RIF" else 1

    def todict(self):
        return {
            "id": self.id,
            "tipo": self.tipo,
            "sexo": self.sexo,
            "label": self.label,
            "camada": self.camada,
            "situacao": self.situacao,
            "cor": self.cor,
            "texto_tooltip": self.texto_tooltip,
            "m1": self.m1,
            "m2": self.m2,
            "m3": 0,
            "m4": 0,
            "m5": 0,
            "m6": 0,
            "m7": 0,
            "m8": 0,
            "m9": 0,
            "m10": 0,
            "m11": 0,
            "dataoperacao": self.dataOperacao,
        }


class noPF(nodo):
    def __init__(self, id, label="", cor="Silver", sexo=0, fonte="RIF"):
        nodo.__init__(self, id, label, "PF")
        self.sexo = sexo

    def todict(self):
        return nodo.todict(self)


class noPJ(nodo):
    def __init__(self, id, label="", cor="Silver", fonte="RIF"):
        nodo.__init__(self, id, label, "PJ")
        self.cor = cor
        self.sexo = 1


class noConta(nodo):
    def __init__(self, id, label="CONTA", cor="Green"):
        nodo.__init__(self, id, label, "CCR")
        self.cor = cor


class noGrupo(nodo):
    def __init__(self, id, label="GRUPO", cor="Blue"):
        nodo.__init__(self, id, label, "GR")
        self.cor = cor
        self.fonte = "grupos"


class noComunicacao(nodo):
    def __init__(self, id, label="COMUNICACAO", cor="Red", dataOperacao=None):
        nodo.__init__(self, id, label, "COM")
        self.cor = cor

    # self.dataOperacao=dataOperacao


class aresta:
    def __init__(self, origem, destino, descricao="", cor="Silver", fonte="RIF"):
        self.origem = origem
        self.destino = destino
        self.descricao = descricao
        self.cor = cor
        self.fonte = fonte
        self.camada = 0 if self.fonte == "RIF" else 1

    def todict(self):
        return {
            "origem": self.origem,
            "destino": self.destino,
            "cor": self.cor,
            "camada": self.camada,
            "tipoDescricao": {"0": self.descricao},
        }


lg = log()

com = estrutura(
    "Comunicacoes",
    [
        "Indexador",
        "Data_do_Recebimento",
        "Data_da_operacao",
        "DataFimFato",
        "cpfCnpjComunicante",
        "nomeComunicante",
        "CidadeAgencia",
        "UFAgencia",
        "NomeAgencia",
        "NumeroAgencia",
        "informacoesAdicionais",
        "CampoA",
        "CampoB",
        "CampoC",
        "CampoD",
        "CampoE",
    ],
)
env = estrutura(
    "Envolvidos",
    [
        "Indexador",
        "cpfCnpjEnvolvido",
        "nomeEnvolvido",
        "tipoEnvolvido",
        "agenciaEnvolvido",
        "contaEnvolvido",
        "DataAberturaConta",
        "DataAtualizacaoConta",
        "bitPepCitado",
        "bitPessoaObrigadaCitado",
        "intServidorCitado",
    ],
)
oco = estrutura("Ocorrencias", ["Indexador", "Ocorrencia"])
# opcionais
gru = estrutura("Grupos", ["cpfCnpjEnvolvido", "nome_Envolvido", "Grupo", "Detalhe"])
vin = estrutura(
    "Vinculos",
    [
        "cpfCnpjEnvolvido",
        "nome_Envolvido",
        "cpfCnpjVinculado",
        "nome_Vinculado",
        "Descricao",
    ],
)


estruturas = [com, env, oco, gru, vin]
# help_estruturas(estruturas)


def removeAcentos(data):
    if data is None:
        return u""
    #  if isinstance(data,str):
    #    data = unicode(data,'latin-1','ignore')
    return "".join(
        x for x in unicodedata.normalize("NFKD", data) if x in string.printable
    )


def gerar_planilha(arquivo, df, nome, indice=False):
    def formatar_cabecalho(cor):
        return arquivo.book.add_format(
            {
                "bold": True,
                "text_wrap": True,
                "valign": "top",
                "fg_color": cor,
                "border": 1,
            }
        )

    # Palette URL: http://paletton.com/#uid=43K0I0kw0w0jyC+oRxVy4oIDfjr
    PALETA = [
        "#5778C0",
        "#a4b3b6",
        "#FF8D63",
        "#FFE700",
        "#FFA900",
        "#000000",
    ]  # azul, cinza, verm, amarelo, lara, preto
    COR_PRINCIPAL = PALETA[0]
    COR_NEUTRA_CLARA = PALETA[1]
    COR_SECUNDARIA = PALETA[2]
    COR_TERCIARIA = PALETA[4]
    COR_NEUTRA_ESCURA = PALETA[5]

    df.style.bar(color=COR_PRINCIPAL)
    print("antes " + nome)
    df.to_excel(arquivo, sheet_name=nome, index=indice)
    print("depois " + nome)
    # Write the column headers with the defined format.
    # print(df.index.names)
    if len(arquivo.sheets) > 6:
        cor_basica = COR_SECUNDARIA
    elif len(arquivo.sheets) < 3:
        cor_basica = COR_PRINCIPAL
    else:
        cor_basica = COR_NEUTRA_CLARA

    if not indice:
        for col_num, value in enumerate(df.columns.values):
            arquivo.sheets[nome].write(
                0, col_num, value, formatar_cabecalho(cor_basica)
            )
        arquivo.sheets[nome].set_tab_color(cor_basica)
    else:
        for col_num, value in enumerate(df.index.names):
            arquivo.sheets[nome].write(
                0, col_num, value, formatar_cabecalho(cor_basica)
            )
        for col_num, value in enumerate(df.columns.values):
            arquivo.sheets[nome].write(
                0,
                col_num + len(df.index.names),
                value,
                formatar_cabecalho(COR_NEUTRA_CLARA),
            )
        arquivo.sheets[nome].set_tab_color(cor_basica)


def gerar_planilhaXLS(arquivo, df, nome, indice=False):
    df.style.bar(color="#99ccff")
    df.to_excel(arquivo, sheet_name=nome, index=indice)


def tipoi2F(umou2=1, linha=None, carJuncao="\r "):
    print("linha= ", linha)
    descricao = linha[1 if umou2 == 1 else 3]
    #     if descricao == '': #telefone ou endereco
    #         descricao = carJuncao.join(node[4:].split('__'))
    #     else:
    #         if self.GNX.node[node]['tipo'] !='TEL':
    #             descricao = Obj.parseCPFouCNPJ(node) + carJuncao + carJuncao.join(textwrap.wrap(descricao,30))

    # dicTipo = {'TEL':u'Telefone', 'END':u'Local', 'PF':u'PF', 'PJ':u'PJ', 'PE':u'Edifício', 'ES':u'Edifício', 'CC':u'Conta','INF':u'Armário' }

    tipo = linha[7 if umou2 == 1 else 8]
    # tipoi2 = dicTipo[tipo]
    tipoi2 = u"Escritório"
    if tipo in ("TEL", "END", "CC"):
        descricao = ""
    else:
        descricao = carJuncao.join(textwrap.wrap(descricao, 30))
    sexo = 1

    if tipo == "PF":
        # if self.GNX.node[node]['sexo']==1:
        if not sexo or sexo == 1:
            tipoi2 = u"Profissional (masculino)"
        elif sexo == 2:
            tipoi2 = u"Profissional (feminino)"
    elif tipo == "PJ":
        # if node[8:12]!='0001':
        # if sexo != 1: #1=matriz
        if sexo % 2 == 0:  # 1=matriz
            tipoi2 = u"Apartamento"  # filial de empresa
        else:
            tipoi2 = u"Escritório"
    elif tipo == "PE":
        tipoi2 = u"Oficina"

    corSituacao = linha[9 if umou2 == 1 else 10]
    if linha[4 if umou2 == 1 else 5] == 0:
        corSituacao = "Vermelho"
    return (tipoi2, descricao, corSituacao)


def to_i2(df, arquivo=None):
    dicTiposIngles = {
        u"Profissional (masculino)": u"Person",
        u"Profissional (feminino)": u"Woman",
        u"Escritório": u"Office",
        u"Apartamento": u"Workshop",
        u"Governo": u"House",
        u"Casa": u"House",
        u"Loja": u"Office",
        u"Oficina": u"Office",
        u"Telefone": u"Phone",
        u"Local": u"Place",
        u"Conta": u"Account",
        u"Armário": u"Cabinet",
        u"Edifício": u"Office",
    }
    # chart = Pyanx_macros()
    noi2origem = {}
    noi2destino = {}

    for idc, campos in df.iterrows():
        #  print('campos= ',campos)

        tipo, descricao, corSituacao = tipoi2F(linha=campos, umou2=1, carJuncao=" ")
        noi2origem[idc] = chart.add_node(
            entity_type=dicTiposIngles.get(tipo, ""),
            label=(campos["cpfcnpj1"]) + u"-" + (descricao),
        )
        tipo, descricao, corSituacao = tipoi2F(linha=campos, umou2=2, carJuncao=" ")
        noi2destino[idc] = chart.add_node(
            entity_type=dicTiposIngles.get(tipo, ""),
            label=(campos["cpfcnpj1"]) + u"-" + (descricao),
        )

        nomeLigacao = campos["descrição"]
        chart.add_edge(noi2origem[idc], noi2destino[idc], removeAcentos(nomeLigacao))
    # idc += 1

    fstream = chart.createStream(
        layout="spring_layout", iterations=0
    )  # não calcula posição

    retorno = fstream.getvalue()
    fstream.close()
    if arquivo is not None:
        f = open(arquivo, "w")
        f.write(retorno)
        f.close()
    return retorno


def soDigitos(texto):
    return re.sub("[^0-9]", "", texto)


def estimarFluxoDoDinheiro(tInformacoesAdicionais):
    # normalmente aparece algo como R$ 20,8 Mil enviada para Jardim Indústria e Comércio -  CNPJ 606769xxx
    # inicialmente quebramos o texto por R$ e verifica quais são seguidos por CPF ou CNPJ
    # pega o texto da coluna InformacoesAdicionais do arquivo Comunicacoes.csv e tenta estimar o valor para cada cpf/cnpj
    # normalmente aparece algo como R$ 20,8 Mil enviada para Indústria e Comércio -  CNPJ 6067xxxxxx
    # inicialmente quebramos o texto por R$ e verifica quais são seguidos por CPF ou CNPJ
    # retorna dicionário
    # como {'26106949xx': 'R$420 MIL RECEBIDOS, R$131 MIL POR', '68360088xxx': 'R$22 MIL, RECEBIDAS'}
    # lista = re.sub(' +', ' ',tInformacoesAdicionais).upper().split('R$')
    t = re.sub(" +", " ", tInformacoesAdicionais).upper()
    lista = t.split("R$")
    listaComTermoCPFCNPJ = []
    for item in lista:
        if "CPF" in item or "CNPJ" in item:
            listaComTermoCPFCNPJ.append(item.strip())

    listaValores = []
    valoresDict = {}
    for item in listaComTermoCPFCNPJ:
        valorPara = ""
        cpn = ""
        le = item.split(" ")
        valor = "R$" + le[0]  # + ' ' + le[1] # + ' ' + le[2]
        if le[1].upper().rstrip(",").rstrip("S").rstrip(",") in (
            "MIL",
            "MI",
            "RECEBIDO",
            "RECEBIDA",
            "ENVIADA",
            "RETIRADO",
            "DEPOSITADO",
            "CHEQUE",
        ):
            valor += " " + le[1]
        if le[2].upper().rstrip(",").rstrip("S") in (
            "MIL",
            "MI",
            "RECEBIDO",
            "RECEBIDA",
            "ENVIADA",
            "RETIRADO",
            "DEPOSITADO",
            "CHEQUE",
        ):
            valor += " " + le[2]
        if "CPF" in item:
            aux1 = item.split("CPF ")
            try:
                aux2 = aux1[1].split(" ")
                cpn = soDigitos(aux2[0])
            except:
                pass
        elif "CNPJ" in item:
            aux1 = item.split("CNPJ ")
            try:
                aux2 = aux1[1].split(" ")
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
                valoresDict[cpn] = set([valor])
    d = {}
    for k, v in valoresDict.items():
        d[k] = ", ".join(v)
    return d


# .def estimaFluxoDoDinheiro(t):


def consolidar_pd(pasta):
    """Processa as planilhas comunicacoes, envolvidos, ocorrencias e grupo em planilhas com agrupamento """
    arq = com.nomearq()  # Comunicacoes
    try:
        df_com = pd.read_excel(
            arq, options={"strings_to_numbers": False}, converters={"Indexador": str}
        )
        df_com["Indexador"] = pd.to_numeric(df_com["Indexador"], errors="coerce")
        df_com["Data_da_operacao"] = pd.to_datetime(df_com["Data_da_operacao"])
        if not com.estr_compativel(df_com.columns):
            print(com.estr_upper())
            mostra_erro("O arquivo " + arq + " contém colunas incompatíveis: ")
            raise ("Estrutura incompatível")
        lg.gravalog("Arquivo " + arq + " lido.")
    except Exception as exc:
        print("Erro ao ler o arquivo " + arq + "\n" + str(type(exc)))

    arq = env.nomearq()  # Envolvidos
    try:
        df_env = pd.read_excel(
            arq, options={"strings_to_numbers": False}, converters={"Indexador": str}
        )
        df_env["Indexador"] = pd.to_numeric(df_env["Indexador"], errors="coerce")
        df_env = df_env[pd.notnull(df_env["Indexador"])]
        if not env.estr_compativel(df_env.columns):
            print(env.estr_upper())
            mostra_erro("O arquivo " + arq + " contém colunas incompatíveis: ")
            raise ("Estrutura incompatível")
        lg.gravalog("Arquivo " + arq + " lido.")
    except Exception as exc:
        lg.gravalog("Erro ao ler o arquivo " + arq + "\n" + str(type(exc)))

    arq = oco.nomearq()  # Ocorrencias
    try:
        df_oco = pd.read_excel(arq, options={"strings_to_numbers": False})
        df_oco["Indexador"] = pd.to_numeric(df_oco["Indexador"], errors="coerce")
        df_oco = df_oco[pd.notnull(df_oco["Indexador"])]
        dictOco = {}
        dictOco2 = {}
        for r in df_oco.itertuples(index=False):
            if r.Indexador in dictOco:
                s = dictOco[r.Indexador]
                s += "; " + r.Ocorrencia
                dictOco[r.Indexador] = s
            else:
                dictOco[r.Indexador] = r.Ocorrencia
        dictOco2["Indexador"] = []
        dictOco2["Ocorrencia"] = []
        for k, v in dictOco.items():
            dictOco2["Indexador"].append(k)
            dictOco2["Ocorrencia"].append(v)

        df_oco2 = pd.DataFrame.from_dict(dictOco2)

        if not oco.estr_compativel(df_oco.columns):
            print(oco.estr_upper())
            mostra_erro("O arquivo " + arq + " contém colunas incompatíveis: ")
            raise ("Estrutura incompatível")
        lg.gravalog("Arquivo " + arq + " lido.")
    except Exception as exc:
        lg.gravalog("Erro ao ler o arquivo " + arq + "\n" + str(type(exc)))

    arq = gru.nomearq()  # Grupos/detalhes
    if not os.path.isfile(arq):  # criar arquivo vazio
        consolidado = pd.ExcelWriter(
            arq,
            engine="xlsxwriter",
            options={"strings_to_numbers": False},
            datetime_format="dd/mm/yyyy",
            date_format="dd/mm/yyyy",
        )
        gerar_planilha(
            consolidado, pd.DataFrame(columns=gru.estr), gru.nome, indice=False
        )
        consolidado.save()
        lg.gravalog(
            "O arquivo "
            + arq
            + " não foi encontrado. Um novo foi criado com as colunas "
            + gru.exibir()
        )
    try:
        df_gru = pd.read_excel(arq, options={"strings_to_numbers": False})
        df_gru = df_gru.fillna("-")
        if not gru.estr_compativel(df_gru.columns):
            print(gru.estr_upper())
            mostra_erro("O arquivo " + arq + " contém colunas incompatíveis: ")
            raise ("Estrutura incompatível")
        lg.gravalog("Arquivo " + arq + " lido.")
    except Exception as exc:
        lg.gravalog("Erro ao ler o arquivo " + arq + "\n" + str(type(exc)))

    arq = vin.nomearq()  # Vinculos
    if not os.path.isfile(arq):  # criar arquivo vazio
        consolidado = pd.ExcelWriter(
            arq,
            engine="xlsxwriter",
            options={"strings_to_numbers": False},
            datetime_format="dd/mm/yyyy",
            date_format="dd/mm/yyyy",
        )
        gerar_planilha(
            consolidado, pd.DataFrame(columns=vin.estr), vin.nome, indice=False
        )
        consolidado.save()
        lg.gravalog(
            "O arquivo "
            + arq
            + " não foi encontrado. Um novo foi criado com as colunas "
            + vin.exibir()
        )
    try:
        df_vin = pd.read_excel(arq, options={"strings_to_numbers": False})
        if not vin.estr_compativel(df_vin.columns):
            print(vin.estr_upper())
            mostra_erro("O arquivo " + arq + " contém colunas incompatíveis: ")
            raise ("Estrutura incompatível")
        lg.gravalog("Arquivo " + arq + " lido.")
    except Exception as exc:
        lg.gravalog("Erro ao ler o arquivo " + arq + "\n" + str(type(exc)))

    nenhumgrupo = len(df_gru["Grupo"].unique())==0
    if nenhumgrupo:
        grupos_selecionados = None
    else:
        grupos_selecionados = gui_grupos(df_gru["Grupo"].unique()) # selecao
        if grupos_selecionados == None : 
            grupos_selecionados = df_gru["Grupo"].unique() # nenhum = todos

    print("Consolidando")
    arq = os.path.join(pasta, "RIF_consolidados.xlsx")
    porGrupo = len(df_gru["Grupo"].unique()) > 1
    try:
        print("antes merge")
        df_consolida = pd.merge(df_com, df_env, how="left", on="Indexador")
        df_consolida = pd.merge(df_consolida, df_oco2, how="left", on="Indexador")

        df_consolida = pd.merge(df_consolida, df_gru, how="left", on="cpfCnpjEnvolvido")
        print("depois merge")
        df_consolida.Detalhe.fillna(
            "-?-", inplace=True
        )  # CPFCNPJ que não constam do grupo

        indexadores_selecionados = df_consolida["Indexador"].values

        if porGrupo:
            indexadores_grupo = df_consolida["Indexador"].loc[
            df_consolida["Grupo"].isin(grupos_selecionados)
        ]
            indexadores_selecionados = indexadores_grupo.values
            df_congrupo = df_consolida.loc[
                df_consolida["Indexador"].isin(indexadores_grupo.values)
            ]
            df_consolida = df_congrupo

        consolidado = pd.ExcelWriter(
            arq,
            engine="xlsxwriter",
            options={"strings_to_numbers": False},
            datetime_format="dd/mm/yyyy",
            date_format="dd/mm/yyyy",
        )
        if porGrupo and not nenhumgrupo:  # tem agrupamentos
            table = pd.pivot_table(
                df_consolida,
                index=[
                    "Grupo",
                    "Indexador",
                    "Data_da_operacao",
                    "cpfCnpjEnvolvido",
                    "nomeEnvolvido",
                    "informacoesAdicionais",
                    "Detalhe",
                    "Ocorrencia",
                ],
                columns=["tipoEnvolvido"],
                margins=False,
            )
        else:
            table = pd.pivot_table(
                df_consolida,
                index=[
                    "Indexador",
                    "Data_da_operacao",
                    "cpfCnpjEnvolvido",
                    "nomeEnvolvido",
                    "informacoesAdicionais",
                    "Detalhe",
                    "Ocorrencia",
                ],
                columns=["tipoEnvolvido"],
                margins=False,
            )
        df_pivot = table.stack()
    except Exception as exc:
        lg.gravalog(
            "Erro ao consolidar planilhas no arquivo " + arq + "\n" + str(type(exc))
        )
        print("depois grupo")

    dicAdic = {"Indexador": [], "cpfCnpjEnvolvido": [], "valor": []}
    valoresPorIndexador = {}
    for row in df_com.itertuples(index=False):
        if row.informacoesAdicionais:
            valoresPorIndexador[row.Indexador] = estimarFluxoDoDinheiro(
                str(row.informacoesAdicionais)
            )
    for k, v in valoresPorIndexador.items():
        if v != {}:
            for kk, vv in v.items():
                dicAdic["Indexador"].append(k)
                dicAdic["cpfCnpjEnvolvido"].append(kk)
                dicAdic["valor"].append(vv)
    df_Adic = pd.DataFrame.from_dict(dicAdic)

    try:
        gerar_planilha(consolidado, df_pivot, "INDEXADOR", indice=True)

        if porGrupo:  # tem agrupamentos
            table = pd.pivot_table(
                df_consolida,
                index=[
                    "Grupo",
                    "cpfCnpjEnvolvido",
                    "nomeEnvolvido",
                    "Data_da_operacao",
                    "Detalhe",
                    "informacoesAdicionais",
                    "Indexador",
                    "Ocorrencia",
                ],
                columns=["tipoEnvolvido"],
                margins=False,
            )
        else:
            table = pd.pivot_table(
                df_consolida,
                index=[
                    "cpfCnpjEnvolvido",
                    "nomeEnvolvido",
                    "Data_da_operacao",
                    "Detalhe",
                    "informacoesAdicionais",
                    "Indexador",
                    "Ocorrencia",
                ],
                columns=["tipoEnvolvido"],
                margins=False,
            )
        df_pivot = table.stack()

        gerar_planilha(consolidado, df_pivot, "CPFCNPJ", indice=True)
        print("df_pivot")

        gerar_planilha(consolidado, df_consolida, "ComunicXEnvolvidos")
        print("df_consolida")
        gerar_planilha(consolidado, df_com, "Comunicacoes")
        print("df_com")
        gerar_planilha(consolidado, df_env, "Envolvidos")
        print("df_env")
        gerar_planilha(consolidado, df_oco2, "Ocorrencias")
        print("df_oco2")
        gerar_planilha(consolidado, df_gru, "Grupos")
        print("df_gru")
        gerar_planilha(consolidado, df_vin, "Vinculos")
        print("df_vin")
        gerar_planilha(consolidado, df_Adic, "InfoAdicionais")
        print("df_Adic")

        df_consolida.to_csv(os.path.join(pasta, 'consolidado.csv'))

    except Exception as exc:
        lg.gravalog(
            "Erro ao gerar planilhas para o arquivo " + arq + "\n" + str(type(exc))
        )

    try:
        consolidado.save()
    except Exception as exc:
        lg.gravalog("Erro ao gravar o arquivo " + arq + "\n" + str(type(exc)))
    lg.gravalog("Planilhas consolidadas: " + arq)
    return df_gru, df_env, df_com, df_oco2, df_vin, indexadores_selecionados


def exportar_rede_rel(pasta, dfgru, dfenv):
    # criando as tabelas no SQLITE
    try:
        conx = sqlite3.connect(
            ":memory:"
        )  # ou use :memory: para botÃ¡-lo na memÃ³ria RAM
        curs = conx.cursor()
    except Exception as exc:
        lg.gravalog("Erro criar conexão com SQLITE memory\n" + str(type(exc)))

    try:
        dfgru.to_sql("Pedido", conx, index=False, if_exists="replace")
    except Exception as exc:
        lg.gravalog("Erro carregar grupos no SQLITE memory\n" + str(type(exc)))
    try:
        dfenv.to_sql("Envolvidos", conx, index=False, if_exists="replace")
    except Exception as exc:
        lg.gravalog("Erro carregar envolvidos no SQLITE memory\n" + str(type(exc)))

    sql = "select "
    sql += (
        " REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,'.',''),'-',''),'/','') as cpfcnpj,"
    )
    sql += ' case length(REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,".",""),"-",""),"/","")) when 14 then "PJ" when 11 then "PF" else "-" end as tipo,'
    sql += ' upper(Grupo) as nome, 0 as camada, "" as [componente/grupo],'
    sql += " 0 as situacao, 0 as [sexo(PF)/Matriz-Filial(PJ)], 0 as [servidor(PF)/nat.jur(PJ)],"
    sql += " 0 as [salario<2min], 0 as OB, 0 as pad, 0 as [PF candidato], 0 as [CEIS/CEPIM], 0 as doadorTSE, 0 as CadUnico, 0 as Falecido"
    sql += " from Pedido "
    sql += " union select distinct"
    sql += (
        " REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,'.',''),'-',''),'/','') as cpfcnpj,"
    )
    sql += ' case length(REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,".",""),"-",""),"/","")) when 14 then "PJ" when 11 then "PF" else "-" end as tipo,'
    sql += ' upper(nomeEnvolvido) as nome, 1 as camada, "" as [componente/grupo],'
    sql += " 0 as situacao, 0 as [sexo(PF)/Matriz-Filial(PJ)], 0 as [servidor(PF)/nat.jur(PJ)],"
    sql += " 0 as [salario<2min], 0 as OB, 0 as pad, 0 as [PF candidato], 0 as [CEIS/CEPIM], 0 as doadorTSE, 0 as CadUnico, 0 as Falecido"
    sql += " from Envolvidos where cpfCnpjEnvolvido not in (select cpfCnpjEnvolvido from Pedido) "
    sql += ' union select distinct "" as cpfcnpj, "CC" as tipo,'
    sql += ' agenciaEnvolvido || "-"|| contaEnvolvido as nome, 1 as camada, "" as [componente/grupo],'
    sql += " 0 as situacao, 0 as [sexo(PF)/Matriz-Filial(PJ)], 0 as [servidor(PF)/nat.jur(PJ)],"
    sql += " 0 as [salario<2min], 0 as OB, 0 as pad, 0 as [PF candidato], 0 as [CEIS/CEPIM], 0 as doadorTSE, 0 as CadUnico, 0 as Falecido"
    sql += ' from Envolvidos where agenciaEnvolvido not in ("-","0")'
    # dicRede = {'cpfcnpj'=[],'tipo'=[],'nome'=[],'camada'=[],}
    try:
        arq = os.path.join(pasta, "RIF_rede_de_rel.xls")
        rede_rel = pd.ExcelWriter(
            arq,
            engine="openpyxl",
            options={"strings_to_numbers": False},
            datetime_format="dd/mm/yyyy",
            date_format="dd/mm/yyyy",
        )
    except Exception as exc:
        lg.gravalog("Erro abrir planilha de rede de relacionamento\n" + str(type(exc)))

    try:
        df_rr = pd.read_sql(sql, conx)
        gerar_planilhaXLS(rede_rel, df_rr, "cpfcnpj")
    except Exception as exc:
        lg.gravalog(
            "Erro gerar planilha de rede de relacionamento CPFCNPJ\n" + str(type(exc))
        )

    try:
        sql = "select distinct indexador, "
        sql += ' (agenciaEnvolvido || "-"|| contaEnvolvido) as CC'
        sql += ' from Envolvidos where agenciaEnvolvido not in ("-","0") '
        df_cc = pd.read_sql(sql, conx)
        df_cc.to_sql("Contas", conx, index=False, if_exists="replace")
    except Exception as exc:
        lg.gravalog("Erro gerar tabela auxiliar de contas\n" + str(type(exc)))

    try:
        sql = "select distinct"
        sql += " REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,'.',''),'-',''),'/','') as cpfcnpj1,"
        sql += 'upper(nomeEnvolvido) as nome1, contas.CC as cpfcnpj2, "" as nome2, 1 as camada, "CC" as [descrição]'
        sql += " from Envolvidos inner join contas"
        sql += " on Envolvidos.indexador = contas.indexador"
        df_rr = pd.read_sql(sql, conx)
        gerar_planilhaXLS(rede_rel, df_rr, "ligacoes")
    except Exception as exc:
        lg.gravalog(
            "Erro gerar planilha de rede de relacionamento LIGACOES\n" + str(type(exc))
        )

    try:
        sql = "select distinct "
        sql += " cpfCnpjEnvolvido as cpfcnpj1,"
        sql += ' upper(nomeEnvolvido) as nome1, contas.CC as cpfcnpj2, "" as nome2, 1 as camada1, 1 as camada2,"CC" as [descrição], '
        sql += ' case length(REPLACE(REPLACE(REPLACE(cpfCnpjEnvolvido,".",""),"-",""),"/","")) when 14 then "Escritório" when 11 then "Profissional (masculino)" else "-" end as tipo1,'
        sql += ' "Conta" as tipo2, "Nenhum" as cor_situacao1, "Nenhum" as cor_situacao2'
        sql += " from Envolvidos inner join contas"
        sql += " on Envolvidos.indexador = contas.indexador"
        df_rr = pd.read_sql(sql, conx)
        gerar_planilhaXLS(rede_rel, df_rr, "ligacoes")
    except Exception as exc:
        lg.gravalog(
            "Erro gerar planilha de rede de relacionamento LIGACOES - complementares\n"
            + str(type(exc))
        )

    try:
        arqcsv = os.path.join(pasta, "I2.csv")
        df_rr.to_csv(
            arqcsv, sep=";", header=True, encoding="utf-8", decimal=",", index=False
        )
    except Exception as exc:
        lg.gravalog("Erro gerar csv de rede de relacionamento i2\n" + str(type(exc)))

    # to_i2(df=df_rr,arquivo='rede_rel.anx')
    try:
        gerar_planilhaXLS(rede_rel, df_rr, "I2")
    except Exception as exc:
        lg.gravalog(
            "Erro gerar planilha de rede de relacionamento i2\n" + str(type(exc))
        )

    # df_rr = pd.DataFrame(r, columns=[i[0] for i in curs.description])

    try:
        rede_rel.save()
    except Exception as exc:
        lg.gravalog("Erro ao gravar o arquivo " + arq + "\n" + str(type(exc)))
    lg.gravalog("Rede de relacionamento gerada: " + arq)


def validar_pasta(pasta, planilhas):
    for p in planilhas:
        p.mudar_pasta(pasta)
        if not p.arquivo_existe():
            return p.nomearq()
    return ""


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

    # futuro
    procGrupos = True
    procVinculos = True
    procContas = True
    procAusentes = True

    proc_grupos_msg = "Processar grupos/detalhes de CPF/CNPJ"
    proc_vinculos_msg = "Incluir vínculos complementares ao RIF"
    proc_contas_msg = "Agrupar contas bancárias no grafo"
    proc_ausentes_msg = "Incluir CPF/CNPF ausentes do RIF no grafo"

    stored_args = {}
    # usar o nome do script sem a extensão para formar o nome do arquivo json
    script_name = os.path.splitext(os.path.basename(__file__))[0]

    args_file = "{}-args.json".format(script_name)
    pasta = ""
    # ler os parâmetros persistidos, gravados no arquivo json
    if os.path.isfile(args_file):
        with open(args_file) as data_file:
            stored_args = json.load(data_file)
            procContas = stored_args.get("procContas")
            procVinculos = stored_args.get("procVinculos")
            procGrupos = stored_args.get("procGrupos")
            procAusentes = stored_args.get("procAusentes")

    # se o programa for chamado sem especificar a pasta de origem, abrir GUI para obtê-la do usuário
    if len(sys.argv) <= 1:
        pasta = stored_args.get("pasta")
        pasta = gui_pasta(pasta)

    # definir o processamento de parâmetros passados pela linha de comando
    desc = "Processa planilhas de RIF (comunicações, envolvidos, ocorrencias e grupo), \ngerando planilhas com agrupamento e com dados para i2"
    file_help_msg = "Nome da pasta onde estão as planilhas"

    linha_comando = argparse.ArgumentParser(description=desc)
    linha_comando.add_argument(
        "--pasta",
        action="store",
        dest="pasta",
        help=file_help_msg,
        default=pasta,
        required=False,
    )
    linha_comando.add_argument(
        "-p",
        action="store",
        dest="pasta",
        help=file_help_msg,
        default=pasta,
        required=False,
    )
    linha_comando.add_argument(
        "-e",
        action=help_estruturas(estruturas),
        help="Exibe as estruturas esperadas das planilhas de entrada",
        required=False,
    )
    linha_comando.add_argument(
        "-g",
        action="store",
        dest="procGrupos",
        help=proc_grupos_msg,
        default=procGrupos,
        required=False,
    )
    linha_comando.add_argument(
        "-c",
        action="store",
        dest="procContas",
        help=proc_contas_msg,
        default=procContas,
        required=False,
    )
    linha_comando.add_argument(
        "-a",
        action="store",
        dest="procAusentes",
        help=proc_ausentes_msg,
        default=procAusentes,
        required=False,
    )

    args = linha_comando.parse_args()
    if not args.pasta:
        args.pasta = os.getcwd()  # pasta atual do script como default
    if not os.path.isdir(args.pasta):  # ver sé pasta
        print(args.pasta + " nao é pasta")
        exit(1)

    args.procContas = gui_sn(procContas, proc_contas_msg + "?")
    args.procGrupos = gui_sn(procGrupos, proc_grupos_msg + "?")
    args.procVinculos = gui_sn(procVinculos, proc_vinculos_msg + "?")
    args.procAusentes = gui_sn(procAusentes, proc_ausentes_msg + "?")

    # persistir os parâmetros
    with open(args_file, "w") as data_file:
        # Using vars(args) returns the data as a dictionary
        json.dump(vars(args), data_file)

    return args


#  display_message()


def executar(pasta):
    # Selecionar a pasta de origem/destino, validá-la e executar a consolidação e a geração das planilhas
    if not pasta:
        args = parse_args()
        pasta = args.pasta
    procContas = args.procContas
    procGrupos = args.procGrupos
    procVinculos = args.procVinculos
    procAusentes = args.procAusentes

    proc_grupos_msg = "Processar grupos/detalhes de CPF/CNPJ"
    proc_vinculos_msg = "Incluir vínculos complementares ao RIF"
    proc_contas_msg = "Agrupar contas bancárias no grafo"
    proc_ausentes_msg = "Incluir CPF/CNPF ausentes do RIF no grafo"

    print("Pasta selecionada: " + pasta)
    print(proc_contas_msg + ": " + "sim" if procContas else "não")
    print(proc_grupos_msg + ": " + "sim" if procGrupos else "não")
    print(proc_vinculos_msg + ": " + "sim" if procVinculos else "não")
    print(proc_ausentes_msg + ": " + "sim" if procAusentes else "não")

    pl = validar_pasta(pasta, estruturas)
    if pl == "":
        dfGrupos, dfEnvolvidos, dfComunicacoes, dfOcorrencias, dfVinculos, indexadores_selecionados = consolidar_pd(
            pasta
        )  # gera planilhas, retorna dataframes
        #   exportar_rede_rel(pasta, dfGrupos, dfEnvolvidos)  # gera planilha para i2
        nos, ligacoes = criarArquivoMacrosGrafo(
            pasta,
            procContas,
            procGrupos,
            procVinculos,
            procAusentes,
            dfGrupos,
            dfEnvolvidos,
            dfComunicacoes,
            dfOcorrencias,
            dfVinculos,
            indexadores_selecionados,
        )  # gera json para o Macros
        #   toAnx(pasta,nos,ligacoes)
        logs = lg.lelog()
        easygui.msgbox(logs)
    else:
        mostra_erro("Arquivo " + pl + " não encontrado nesta pasta:\n" + pasta)
        sys.exit(1)

    return logs  # mensagens sobre a execução


def gui_pasta(pasta):
    # GUI para obter a pastas escolhida pelo usuário
    nomepasta = easygui.diropenbox(
        default=pasta,
        msg="Selecione a pasta onde estão as planilhas Comunicações, Envolvidos e Ocorrências",
        title="Gera planilhas a partir de dados de RIF",
    )
    if not nomepasta:
        exit(1)
    return nomepasta


def gui_sn(sn, texto: str):
    resp = easygui.choicebox(
        msg=texto, choices=["Sim", "Não"], preselect=0 if sn else 1
    )
    return resp == "Sim"


def gui_grupos(grupos):
    if len(grupos) == 0:
        return None
    resp = easygui.multchoicebox(
        msg="Selecione os grupos a serem tratados:", title="Grupos", choices=grupos
    )
    return resp


def pasta_valida(pasta):
    # verifica se a pasta escolida é válida
    pl = validar_pasta(pasta, estruturas)
    if pl == "":
        return True
    else:
        return False


def criarArquivoMacrosGrafo(
    pasta,
    processaConta,
    processaGrupos,
    processaVinculos,
    processaAusentes,
    dfGrupos,
    dfEnvolvidos,
    dfComunicacoes,
    dfOcorrencias,
    dfVinculos,
    indexadores_selecionados,
):
    """gera o arquivo do Macros a partir dos dataframes"""

    # cria tabela de nós
    nos = []

    # procura tipoEnvolvido=Titular, para por no tooltip do nó
    if len(indexadores_selecionados):
        dfEnvolvidos = dfEnvolvidos.loc[
            dfEnvolvidos["Indexador"].isin(indexadores_selecionados)
        ]
        dfComunicacoes = dfComunicacoes.loc[
            dfComunicacoes["Indexador"].isin(indexadores_selecionados)
        ]
        dfOcorrencias = dfOcorrencias.loc[
            dfOcorrencias["Indexador"].isin(indexadores_selecionados)
        ]
        # dfGrupos = dfGrupos.loc[dfOcorrencias['Indexador'].isin(indexadores_selecionados)]

    titularOcorrencia = {}
    conta = {}
    for row in dfEnvolvidos.itertuples(index=False):
        if row.tipoEnvolvido.upper() == "TITULAR":
            titularOcorrencia[row.Indexador] = (
                "Titular: "
                + row.nomeEnvolvido.strip()
                + "("
                + row.cpfCnpjEnvolvido
                + ")"
            )

        if processaConta:
            if row.contaEnvolvido != "-" and str(row.contaEnvolvido) != "0":
                nroConta = (
                    "CC: " + str(row.agenciaEnvolvido) + "/" + str(row.contaEnvolvido)
                )
                campoA = (
                    dfComunicacoes["CampoA"]
                    .where(dfComunicacoes["Indexador"] == row.Indexador)
                    .sum()
                )
                if nroConta not in conta:
                    conta[nroConta] = campoA
                else:
                    conta[nroConta] += campoA
    if processaConta:
        for key, value in conta.items():
            label = (
                "R$ {:,}".format((value))
                .replace(",", "X")
                .replace(".", ",")
                .replace("X", ".")
            )
            nos.append(copy.deepcopy(noConta(key, label).todict()))

    for row in dfComunicacoes.itertuples(index=False):
        if str(row.Indexador) == "nan" or str(row.Indexador) == "0":
            continue
        campos = str(row)[len("Pandas") :]
        label = (
            "COAF R$ {:,}".format((row.CampoA))
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )
        noCOM = noComunicacao(id="COM_" + str(int(row.Indexador)), label=label)

        d = dfOcorrencias[dfOcorrencias.Indexador == row.Indexador]["Ocorrencia"]
        campos = (
            titularOcorrencia.get(row.Indexador, "")
            + " "
            + campos
            + " . Ocorrência(s): "
            + "; ".join(d)
        )

        campos = campos.replace(r"\t", " ")
        campos = campos.replace(r"\x96", " ").replace(r"\x92", " ")
        noCOM.texto_tooltip = str(re.sub(" +", " ", campos))
        nos.append(copy.deepcopy(noCOM.todict()))

    valoresPorIndexador = {}
    for row in dfComunicacoes.itertuples(index=False):
        if row.informacoesAdicionais:
            valoresPorIndexador[row.Indexador] = estimarFluxoDoDinheiro(
                str(row.informacoesAdicionais)
            )

    cpfcnpjset = set()
    for row in dfEnvolvidos.itertuples(index=False):
        cpfcnpj = (
            row.cpfCnpjEnvolvido.replace(".", "").replace("-", "").replace("/", "")
        )
        cpfcnpjset.add(cpfcnpj)

    cs = set()
    for row in dfEnvolvidos.itertuples(index=False):
        cpfcnpj = (
            row.cpfCnpjEnvolvido.replace(".", "").replace("-", "").replace("/", "")
        )
        nome = str(
            row.nomeEnvolvido
        )  # para exportar para i2 (nao busca o nome no Macros)
        if cpfcnpj not in cs:
            cs.add(cpfcnpj)
            if len(cpfcnpj) == 11:
                nos.append(copy.deepcopy(noPF(id=cpfcnpj, label=nome).todict()))
            else:
                nos.append(copy.deepcopy(noPJ(id=cpfcnpj, label=nome).todict()))

    # cria tabela de ligacoes
    ligacoes = []
    ligadict = {}
    ligaContadict = {}

    det = {}  # detalhes
    grp = {}  # grupos
    idDet = 0
    idGru = 0

    if processaGrupos:
        if processaAusentes:  # inclui todos os grupos
            df = dfGrupos
        else:  # inclui os grupos com incidencia no RIF
            df = pd.merge(dfEnvolvidos, dfGrupos, how="inner", on="cpfCnpjEnvolvido")
        csgrupos = df["Grupo"].unique()
        for gr in csgrupos:
            idGru += 1
            grp[gr] = idGru
            texto = gr[:MAX_TAM_LABEL]
            nos.append(
                copy.deepcopy(noGrupo(id="#" + str(idGru), label=texto).todict())
            )
        for row in df.itertuples(index=False):
            cpfcnpj = (
                row.cpfCnpjEnvolvido.replace(".", "").replace("-", "").replace("/", "")
            )
            nome = str(row.nome_Envolvido)
            if cpfcnpj not in cs:
                cs.add(cpfcnpj)
                if len(cpfcnpj) == 11:
                    nos.append(copy.deepcopy(noPF(id=cpfcnpj, label=nome).todict()))
                else:
                    nos.append(copy.deepcopy(noPJ(id=cpfcnpj, label=nome).todict()))

        for row in dfGrupos.itertuples(index=False):
            gr = row.Grupo
            if not (gr in grp):
                continue
            obDetalhe = str(row.Detalhe)[:MAX_TAM_LABEL]
            cpfcnpj = (
                row.cpfCnpjEnvolvido.replace(".", "").replace("-", "").replace("/", "")
            )
            lig = aresta(
                origem=cpfcnpj, destino="#" + str(grp[gr]), descricao=obDetalhe
            ).todict()
            ligacoes.append(copy.deepcopy(lig))
    # dfEnvolvidos.sort_values(by=['Indexador', 'cpfCnpjEnvolvido'])
    for row in dfEnvolvidos.itertuples(index=False):
        if str(row.Indexador) == "nan":
            continue
        cpfcnpj = (
            row.cpfCnpjEnvolvido.replace(".", "").replace("-", "").replace("/", "")
        )
        # origemDestino = "COM_"+row.Indexador+"-"+cpfcnpj
        origemDestino = str(int(row.Indexador)) + "-" + cpfcnpj
        if origemDestino not in ligadict:
            ligadict[origemDestino] = set([row.tipoEnvolvido])
        else:
            ligadict[origemDestino].add(row.tipoEnvolvido)
        # conta
        if (
            row.contaEnvolvido != "-"
            and str(row.contaEnvolvido) != "0"
            and str(row.contaEnvolvido) != "nan"
        ):
            nroConta = (
                "CC: " + str(row.agenciaEnvolvido) + "/" + str(row.contaEnvolvido)
            )
            origemDestino = nroConta + "-" + cpfcnpj
            if origemDestino not in ligadict:
                ligaContadict[origemDestino] = set([row.tipoEnvolvido])
            else:
                ligaContadict[origemDestino].add(row.tipoEnvolvido)

    for k, v in ligadict.items():
        origem = k.split("-")[0]
        destino = k.split("-")[1]
        valor = valoresPorIndexador.get(origem, {}).get(destino, "")
        if valor:
            valor = "-" + valor

        descricao = ",".join(sorted(v)) + valor
        ligacoes.append(
            copy.deepcopy(
                aresta(
                    origem="COM_" + origem, destino=destino, descricao=descricao
                ).todict()
            )
        )

    for k, v in ligaContadict.items():
        origem = k.split("-")[0]
        destino = k.split("-")[1]
        valor = valoresPorIndexador.get(origem, {}).get(destino, "")
        if valor:
            valor = "-" + valor
        # descricao = ','.join(sorted([k for k in v if k!='Outros'])) + valor
        descricao = ",".join(sorted(v)) + valor
        ligacoes.append(
            copy.deepcopy(
                aresta(origem=origem, destino=destino, descricao=descricao).todict()
            )
        )

    if processaVinculos:
        for row in dfVinculos.itertuples(index=False):
            cpfcnpj = (
                row.cpfCnpjEnvolvido.replace(".", "").replace("-", "").replace("/", "")
            )
            nome = str(row.nome_Envolvido)
            if cpfcnpj not in cs:
                cs.add(cpfcnpj)
                if len(cpfcnpj) == 11:
                    nos.append(
                        copy.deepcopy(
                            noPF(id=cpfcnpj, label=nome, fonte="VIN").todict()
                        )
                    )
                else:
                    nos.append(
                        copy.deepcopy(
                            noPJ(id=cpfcnpj, label=nome, fonte="VIN").todict()
                        )
                    )

        for row in dfVinculos.itertuples(index=False):
            cpfcnpj = (
                row.cpfCnpjEnvolvido.replace(".", "").replace("-", "").replace("/", "")
            )
            cpfcnpjvinc = (
                row.cpfCnpjVinculado.replace(".", "").replace("-", "").replace("/", "")
            )
            descricao = str(row.Descricao)[:MAX_TAM_LABEL]
            ligacoes.append(
                copy.deepcopy(
                    aresta(
                        origem=cpfcnpj,
                        destino=cpfcnpjvinc,
                        descricao=descricao,
                        fonte="VIN",
                    ).todict()
                )
            )
            nome = str(row.nome_Envolvido)
            if processaAusentes and cpfcnpj not in cs:  # incluir CPFCNPJ
                cs.add(cpfcnpj)
                if len(cpfcnpj) == 11:
                    nos.append(
                        copy.deepcopy(
                            noPF(id=cpfcnpj, label=nome, fonte="VIN").todict()
                        )
                    )
                else:
                    nos.append(
                        copy.deepcopy(
                            noPJ(id=cpfcnpj, label=nome, fonte="VIN").todict()
                        )
                    )

            cpfcnpj = cpfcnpjvinc
            nome = str(row.nome_Vinculado)
            if processaAusentes and cpfcnpj not in cs:  # incluir CPFCNPJ
                cs.add(cpfcnpj)
                if len(cpfcnpj) == 11:
                    nos.append(
                        copy.deepcopy(
                            noPF(id=cpfcnpj, label=nome, fonte="VIN").todict()
                        )
                    )
                else:
                    nos.append(
                        copy.deepcopy(
                            noPJ(id=cpfcnpj, label=nome, fonte="VIN").todict()
                        )
                    )

    textoJson = json.dumps({"no": nos, "ligacao": ligacoes})  # , ensure_ascii=False)
    # print(textoJson)
    # textoJson = textoJson.replace(r'\t',' ').replace(r'\x96',' ').replace(r'\x92', ' ') #isso dá erro
    arq = os.path.join(pasta, "macros_grafo.json")
    with open(arq, "wt", encoding="latin1") as out:
        out.write(textoJson)
    lg.gravalog("Grafo gerado para o sistema Macros: " + arq)
    return (nos, ligacoes)


def exportar_rede_rel_xlsx(pasta, nos, ligacoes):
    workbook = xlsxwriter.Workbook(os.path.join(pasta, "RIF_rede.xlsx"))
    worksheet = workbook.add_worksheet()
    headings = [
        "cpfcnpj1",
        "nome1",
        "cpfcnpj2",
        "nome2",
        "camada1",
        "camada2",
        "descrição",
        "tipo1",
        "tipo2",
        "cor_situacao1",
        "cor_situacao2",
    ]

    row = 0
    col = 0
    for h in headings:
        worksheet.write(row, col, h)
        col += 1

    for l in ligacoes:
        col = 0
        worksheet.write(row, col, l.origem)
        col += 1
        worksheet.write(row, col, l.label)
        col += 1
        worksheet.write(row, col, l.destino)
        """
        self.origem=origem
        self.destino=destino
        self.descricao=descricao
        self.cor=cor
        self.fonte=fonte
        self.camada=0 if self.fonte == 'RIF' else 1
        """
    return


def toAnx(pasta, nos, ligacoes):
    def tipoNo(no):
        if no[tipo] == "PF":
            tipo = "Person"
        elif no[tipo] == "PJ":
            tipo = "Office"
        elif no[tipo] == "CC":
            tipo = "Account"
        elif no[tipo] == "ENT":
            tipo = "Cabinet"
        else:
            tipo = "Cabinet"
        return tipo

    chart = pyanx.Pyanx()

    for no in nos:
        noAnx = chart.add_node(entity_type=tipoNo, label=removeAcentos(no["label"]))
    """        'label': label,
      'color': color,
      'style': style,
      'description': description,
      'datetime': _datetime,
      'datetime_description': datestr_description,
      'timezone': timezone
    """
    for lig in ligacoes:
        ligAnx = chart.add_edge(
            removeAcentos(lig["origem"]),
            removeAcentos(lig["destino"]),
            removeAcentos(lig["tipoDescricao"]["0"]),
        )
    arqAnx = os.path.join(pasta, "RIF_Grafo.anx")

    chart.create(arqAnx)
    return


if __name__ == "__main__":
    # caso o script seja utilizado como um package por outro, basta chamar a função "executar" com um nome de pasta
    executar("")
