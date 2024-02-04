import os.path
import math
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Defina os escopos necessários para o acesso à planilha
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ID e intervalo da planilha de exemplo
spreadsheet_id = "1hvsc9B8fNhsoSiSflyDts_qlEXcGZmeMtxHErKvEqzU"
range_name = "Página1!C4:G27"


def main():
    # Inicializa as credenciais como None
    creds = None

    # Verifica se o arquivo de token existe
    if os.path.exists("token.json"):
        # Carrega as credenciais do arquivo de token
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # Se não houver credenciais válidas, solicita ao usuário que faça login
    if not creds or not creds.valid:
        # Se as credenciais estiverem expiradas e houver um token de atualização, atualiza as credenciais
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Cria um fluxo para obter as credenciais
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            # Executa o fluxo para autenticação do usuário
            creds = flow.run_local_server(port=0)

        # Salva as credenciais para o próximo uso
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        # Cria o serviço Sheets com as credenciais autenticadas
        service = build("sheets", "v4", credentials=creds)

        # Executa a solicitação para obter os valores da planilha
        result = (
            service.spreadsheets()
            .values()
            .get(spreadsheetId=spreadsheet_id, range=range_name)
            .execute()
        )

        # Extrai os valores da resposta
        valores = result['values']

        # Exibe os valores na saída
        print(valores)

        # Inicializa uma lista vazia para armazenar as situações dos alunos
        Addsituacao = []

        for linhas in valores:  # Itera sobre cada linha (aluno) na lista de valores
            faltas = linhas[0]  # Atribui a quantidade de faltas do aluno
            p1 = linhas[1]  # Atribui a nota da primeira prova do aluno
            p2 = linhas[2]  # Atribui a nota da segunda prova do aluno
            p3 = linhas[3]  # Atribui a nota da terceira prova do aluno

            # Calcula a soma das notas das três provas
            soma = int(p1) + int(p2) + int(p3)
            # Calcula a média das notas das três provas
            media = soma / 3
            # Verifica a situação do aluno com base nas notas e faltas
            if media < 50:
                situacao = "Reprovado por Nota"
            elif media >= 50 and media < 70:
                situacao = "Exame Final"
            elif int(faltas) > 15:
                situacao = "Reprovado por falta"
            else:
                situacao = "Aprovado"

            # Adiciona a situação do aluno à lista Addsituacao
            Addsituacao.append([situacao])

        # Atualiza as situações dos alunos na planilha do Google Sheets
        result = (
            service.spreadsheets()
            .values()
            .update(spreadsheetId=spreadsheet_id, range="G4:G27",
                    valueInputOption="USER_ENTERED", body={'values': Addsituacao})
            .execute()
        )

        Addnotaprovacao = []

        for linha in valores:
            p1 = linha[1]  # Atribui a nota da primeira prova do aluno
            p2 = linha[2]  # Atribui a nota da segunda prova do aluno
            p3 = linha[3]  # Atribui a nota da terceira prova do aluno
            soma = int(p1) + int(p2) + int(p3)  # Calcula a soma das notas das três provas
            media = soma / 3  # Calcula a média das notas das três provas
            situacoes = linha[4]
            if situacoes == 'Aprovado':
                notaaprov = 0
            elif situacoes == 'Exame Final':
                notaaprov = math.ceil(2 * 50 - media)

            Addnotaprovacao.append([notaaprov])

            # Atualiza as situações dos alunos na planilha do Google Sheets
            result = (
                service.spreadsheets()
                .values()
                .update(spreadsheetId=spreadsheet_id, range="H4:H27",
                        valueInputOption="USER_ENTERED", body={'values': Addnotaprovacao})
                .execute()
            )

    except HttpError as err:
        print(err)


if __name__ == "__main__":
    main()
