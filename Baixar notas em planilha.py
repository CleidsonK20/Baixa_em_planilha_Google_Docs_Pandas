from __future__ import print_function
import os.path
import time
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import pandas as pd
import datetime
import win32com.client  
import string
import re

# If modifying these scopes, delete the file token.json.vids
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
# Planilhha de qualidade
SAMPLE_SPREADSHEET_ID1 = '1xyITceuvLlv71iCfSEwCXeL2M4me0KIo4gWcHIvmNtY'
SAMPLE_SPREADSHEET_ID = '1MJBXn2KhX9G3POtGmkmz4UJ6npVs1-P8nxdxR1HIbpw'
SAMPLE_SPREADSHEET_ID2 = '13IsLbdd7bJ5iPbsw5k_zRgAkVSgBFvI2-nwCgmev9hU'
SAMPLE_SPREADSHEET_ID3 = '122wH0OlFL7JBgA4j5UsVXsiugvykD6z7y7lI9iZ39Yc'

creds = None

# If there are no (valid) credentials available, let the user log in.
creds = None
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            r'\\10.20.21.55\supply\Matheus\Arlete\client_secret.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

service = build('sheets', 'v4', credentials=creds)
# Call the Sheets API
sheet = service.spreadsheets()




#print(lista_cods)

#planilha falta com sobra
intervalo_plan_qua = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID1,
                            range='Falta & Sobra!A2:AE').execute()
Planilha_FeS = intervalo_plan_qua.get('values', [])


#Cria o dataframe
dados_plan_FeS = pd.DataFrame(Planilha_FeS, columns = [ 'STATUS',
                                                        'CONTROLE',
                                                        'DATA DE COLETA',
                                                        'COLABORADOR DEVOLUÇÃO',
                                                        'DATA NOTIFICAÇÃO',
                                                        'DATA DA TRIAGEM',
                                                        'DATA AUTORIZAÇÃO',
                                                        'JIRA',
                                                        'COLABORADOR BK',
                                                        'NFO',
                                                        'Fornecedor',
                                                        'MATERIAL',
                                                        'OCORRÊNCIA',
                                                        'QTD',
                                                        'CNPJ Transp',
                                                        'Nome Transportadora',
                                                        'QTDE SOBRA',
                                                        'Nº MIGO',
                                                        'SM',
                                                        'MIRO',
                                                        'DADOS ADICIONAIS',
                                                        'DATA NFD',
                                                        'NFD',
                                                        'RESPONSÁVEL',
                                                        'CÓD.INTERNO',
                                                        '[]',
                                                        'Vol',
                                                        'POSIÇÕES',
                                                        'AC',
                                                        'REF JIRA',
                                                        'ref cod'])


intervalo_plan_qua1 = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                            range='Qualidade Extrema!D2:Y').execute()
Planilha_qualidade = intervalo_plan_qua1.get('values', [])

#data frame qualidade
Planilha_qualidade = pd.DataFrame(Planilha_qualidade, columns= ['Data da coleta',
                                                                'Resp. Notificação', 
                                                                'Data da Notificação',
                                                                'Data de inserção',
                                                                'Data autorização',
                                                                'NFO',
                                                                'NFO - CD',
                                                                'Código Interno',
                                                                'Material',
                                                                'CNPJ',
                                                                'Fornecedor',
                                                                'Marca',
                                                                'Descrição',
                                                                'Referência',
                                                                'Cor',
                                                                'Quantidade',
                                                                'Ocorrência',
                                                                'MIGO',
                                                                'SM',
                                                                'MIRO',
                                                                'DATA DE EMISSÃO',
                                                                'NFD'])


#Data frame site
intervalo_plan_site = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID3,
                            range='Devolução - Site!B3:AF').execute()
Planilha_site = intervalo_plan_site.get('values', [])


#Cria o dataframe site
dados_plan_site = pd.DataFrame(Planilha_site,columns = ['Status',
                                                        'Estado',
                                                        'Tipo de Tratativa',
                                                        'Protocolo',
                                                        'Data de Triagem',
                                                        'Responsável por Triagem',
                                                        'Fornecedor',
                                                        'NFO',
                                                        'Referência',
                                                        'Material',
                                                        'Cód. interno',
                                                        'Posição',
                                                        'Ocorrência - Qualidade',
                                                        'Descrição',
                                                        'Cor do Fornecedor',
                                                        'Quantidade',
                                                        'UC',
                                                        'Colaborador Devolução',
                                                        'Data de Notificação',
                                                        'MIGO',
                                                        'SM',
                                                        'Volumetria',
                                                        'MIRO',
                                                        'Data de Autorização da NFD',
                                                        'Data de emissão',
                                                        'NFD',
                                                        'Responsável NFD',
                                                        'Observação',
                                                        'Motivo de Finalização',
                                                        'Data de Finalização',
                                                        'Qualidade Físico Recebido ?'])


#Data frame RETIRA DE ESTOQUE
intervalo_plan_retira = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID3,
                            range='Devolução - Retira de Estoque!A3:AC').execute()
Planilha_retira = intervalo_plan_retira.get('values', [])


#Cria o dataframe RETIRA DE ESTOQUE
dados_plan_retira = pd.DataFrame(Planilha_retira,columns = ['Status',
                                                        'Estado',
                                                        'Tipo de Tratativa',
                                                        'Data de Autorização da NFD',
                                                        'Data de Triagem',
                                                        'Resposável por Triagem',
                                                        'COD',
                                                        'NFO',
                                                        'Fornecedor',
                                                        'Código Interno / JIRA / Cód. de Rastreio',
                                                        'Posição',
                                                        'Material',
                                                        'Quantidade',
                                                        'Resposável por Devolução',
                                                        'Data de Notificação',
                                                        'MIGO',
                                                        'SM',
                                                        'Volumetria',
                                                        'MIRO',
                                                        'Dados adicionais NFD',
                                                        'Data de emissão',
                                                        'NFD',
                                                        'Responsável NFD',
                                                        'Motivo Devolução',
                                                        'Observação - Motivo Devolução',
                                                        'Fórum Sobre Divergência',
                                                        'Observação',
                                                        'Data de Finalização',
                                                        'Data de Atualização'])

#Data frame Contorle de coletas
controle = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID2,
                range="Controle virtual!B2:C").execute()
controle = controle.get('values', [])


#Cria o dataframe Controle de coletas
controle = pd.DataFrame(controle,columns = ['NFD',
                                            'DATA_DA_COLETA'])

#data = "{:%d.%m.%Y}".format(datetime.date.today())

data_da_coleta = input("Por favor me diga qual a data da coleta:  \n")

time.sleep(1)
#Busca a lista de NFDs da planilha de controle com base na Data

controle1 = controle[controle.DATA_DA_COLETA == data_da_coleta]
controle1 = controle1.reset_index()
controle1 = controle1.NFD.tolist()

#Lista de NFDs da planilha F&S e Qualidade

nfd_qualidade = Planilha_qualidade.NFD.tolist()
nfd_fes = dados_plan_FeS.NFD.tolist()
nfd_site = dados_plan_site.NFD.tolist()
nfd_retira = dados_plan_retira.NFD.tolist()

contagem = 0


for lista_nfds in controle1:
    
    nfd_controle = lista_nfds

    if nfd_controle in nfd_qualidade:

        nfd_qualidade_list = Planilha_qualidade[Planilha_qualidade.NFD == nfd_controle]
        nfds_qualidade_reset = nfd_qualidade_list.reset_index()
        nfds_qualidade1 = nfds_qualidade_reset["index"].tolist()
        nfds_linha_cod_qualidade = (int(nfds_qualidade1[0])) + 2
        
        print("Linha da nota e NFD (Planilha Qualidade): ",nfds_linha_cod_qualidade," - ",nfd_controle)
        
        valor_novo = [["ENVIADO",data_da_coleta]]
        
        try:
            result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range= "Qualidade Extrema!C{}".format(nfds_linha_cod_qualidade), valueInputOption = "USER_ENTERED",body= {"values" : valor_novo}).execute()
            print("\n")
        except:                           
            print("Nota preenchida !\n")
        
    elif nfd_controle in nfd_fes:
        
        nfd_fes = dados_plan_FeS[dados_plan_FeS.NFD == nfd_controle]
        nfds_fes_reset = nfd_fes.reset_index()
        num_linha_codfes = nfds_fes_reset["index"].tolist()
        nfds_linha_cod_fes = (int(num_linha_codfes[0])) + 2
        print("Linha da nota e NFD (Planilha F&S): ",nfds_linha_cod_fes," - ",nfd_controle)
        
        Data_fes_1 = [data_da_coleta]
        Data_fes = []
            
        cont = 0
        while cont < len(nfd_fes):
            Data_fes.append(Data_fes_1)
            cont = cont + 1
            
        try:
            result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID1,
                                    range= "Falta & Sobra!C{}".format(nfds_linha_cod_fes), valueInputOption = "USER_ENTERED",body= {"values" : Data_fes}).execute()
            print("\n")                    
        except:
            print("Nota preenchida !\n")
    elif nfd_controle in nfd_site:
            nfd_site_list = dados_plan_site[dados_plan_site.NFD == nfd_controle]
            nfds_site_reset = nfd_site_list.reset_index()
            num_linha_cod_site = nfds_site_reset["index"].tolist()
            num_linha_codsite = (int(num_linha_cod_site[0])) + 3
            
            print("Linha da nota e NFD (Planilha SITE): ",num_linha_cod_site," - ",nfd_controle)
            
            
            dados1 = ["Finalizado","Coletado"]
            dados2 = ["Devolvido para o fornecedor",data_da_coleta]
            Data_site1 = []
            Data_site2 = []
            
            contador = 0
            while contador < len(nfd_site_list):
                Data_site1.append(dados1)
                Data_site2.append(dados2)
                contador = contador + 1       
                
            try:
                result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID3,
                                    range= "Devolução - Site!B{}".format(num_linha_codsite), valueInputOption = "USER_ENTERED",body= {"values" : Data_site1}).execute()
                result1 = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID3,
                                    range= "Devolução - Site!AD{}".format(num_linha_codsite), valueInputOption = "USER_ENTERED",body= {"values" : Data_site2}).execute()
                print("\n")
                
            except:
                print("Nota preenchida !\n")
    elif nfd_controle in nfd_retira:
            nfd_retira_list = dados_plan_retira[dados_plan_retira.NFD == nfd_controle]
            nfds_retira_reset = nfd_retira_list.reset_index()
            nfd_retira_cod = nfds_retira_reset['index'].tolist()
            retira_linha_cod = (int(nfd_retira_cod[0])) + 3
        
            print("Linha da nota e NFD (Planilha RETIRA): ",retira_linha_cod," - ",nfd_controle)
        
            dados_1 = ["Finalizado","Coletado"]
            dados_2 = [data_da_coleta]
            Dados_retira1 = []
            Dados_retira2 = []

            i = 0
            while i < len(nfd_retira_list):
                Dados_retira1.append(dados_1)
                Dados_retira2.append(dados_2)
                i = i + 1
        
            try:
                result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID3,
                                    range= "Devolução - Retira de Estoque!A{}".format(retira_linha_cod), valueInputOption = "USER_ENTERED",body= {"values" : Dados_retira1}).execute()
                                    
                result1 = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID3,
                                    range= "Devolução - Retira de Estoque!AB{}".format(retira_linha_cod), valueInputOption = "USER_ENTERED",body= {"values" : Dados_retira2}).execute()
                print("\n")
            except:
                print("Nota preenchida !\n")
    else:
        print("\nNota ",nfd_controle," não encontrada!\n\n")

contagem = contagem + 1

