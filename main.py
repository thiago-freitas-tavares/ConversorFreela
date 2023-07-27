import xmltodict                                                #01 biblioteca que faz a leitura de um xml e transforma em um dicionário python.
import os                                                       #02 biblioteca que permite manusear arquivos.
import json                                                     #10 biblioteca que permite manuserar dados de arquivos.
import pandas as pd                                             #22 biblioteca que permite criar tabelas e transformar em arquivos excel.


def pegar_infos(nome_arquivo, dados_linhas):                    #06 definição da função pegar_infos, recebendo nome_arquivos como parâmetro.
#   print(f"Pegou as informações da {nome_arquivo}")
    with open(f'nfs/{nome_arquivo}', "rb") as arquivo_xml:      #07 abre arquivos ("rb" leitura retorna bytes) com os nomes passados (incluir pasta) e salva na variável arquivo_xml.
        dic_arquivo = xmltodict.parse(arquivo_xml)              #08 método parse transforma os dados do arquivo.xml em dicionário python.
#       print(json.dumps(dic_arquivo, indent=4))                #11 método dumps permite formatar o dicionário, nesse caso deixar 4 espaços para infos dentro de outras.
#       try:                                                    #14 deu erro para alguma NF, inserimos o try para tratar o erro.
        if "NFe" in dic_arquivo:                                #18 if/else para tratar o erro 'NFe'.
            infos_nf = dic_arquivo["NFe"]["infNFe"]             #12 na NFe1 todos os dados necessários estão dentro de NFe e infNFe. infos_nf ajuda a reduzir o código.
        else:                                                   #19 else trata o erro 'NFe', porém ocorre outro erro 'vol' do peso.
            infos_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
        numero_nota = infos_nf ["@Id"]                          #13 com isso, extraímos as informações necessárias e rodamos o app sem o break do for para testar todas NFs.
        empresa_emissora = infos_nf ["emit"]["xNome"]
        nome_cliente = infos_nf ["dest"]["xNome"]
        endereço = infos_nf ["dest"]["enderDest"]
        if "vol" in infos_nf ["transp"]:                        #20 if/else para tratar o erro 'vol'.
            peso = infos_nf ["transp"]["vol"]["pesoB"]
        else:                                                   #21 else trata o erro 'vol' e não ocorre mais erro, com isso, podemos retirar o try/except.
            peso = "Não informado"
#       print(numero_nota, empresa_emissora, nome_cliente, endereço, peso, sep="\n")
        dados_linhas.append([numero_nota, empresa_emissora, nome_cliente, endereço, peso])
#                                                               #25 incluímos os dados de cada NF na lista dados_linhas, sendo que cada item é uma lista com todos os dados da NF.
#       except Exception as e:                                  #15 salvamos o erro na variável e.
#           print(e)                                            #16 printamos o erro e o arquivo que deu erro com o método dumps.
#           print(json.dumps(dic_arquivo, indent=4))            #17 observamos que as NFs que ocorrem erro possuem um item antes de NFe, que é o nfeProc.


lista_arquivos = os.listdir("nfs")                              #03 método listdir retorna uma lista com o nome dos arquivos dentro da pasta nfs, passada como argumento.
#                                                               #23 criarmos uma lista com os nomes das colunas da tabela que queremos criar.
colunas = ["numero_nota", "empresa_emissora", "nome_cliente", "endereço", "peso"]
linhas = []                                                     #24 criamos a lista linhas vazia e a incluímos no argumento da chamada e definição da função pegar_infos.
for arquivo in lista_arquivos:                                  #04 for cada arquivo na lista_arquivos.
    pegar_infos(arquivo, linhas)                                #05 chama a função pegar_infos passando o nome de cada arquivo como argumento.
#   break                                                       #09 break colocado para construirmos o app usando apenas 1 arquivo como referência. Depois adaptamos para 1+ arquivos.

tabela = pd.DataFrame(columns=colunas, data=linhas)             #26 método DataFrame de pandas monta a tabela com as listas passadas como argumento.
tabela.to_excel("NotasFiscais.xlsx", index=False)               #27 método to_excel cria um excel na pasta do app com a tabela dos dados coletados.