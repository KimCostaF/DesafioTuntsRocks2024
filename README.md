# DesafioTuntsRocks2024
API que faz integração com Google Sheets e realiza operações na planilha

Este script Python foi desenvolvido para interagir com planilhas do Google Sheets, calcular a média dos alunos, determinar suas situações acadêmicas e calcular as notas necessárias para aprovação final.

## Requisitos

Para executar este script, é necessário ter as seguintes bibliotecas Python instaladas:

- `google-auth`
- `google-auth-oauthlib`
- `google-auth-httplib2`
- `google-api-python-client`

Além disso, é preciso configurar as credenciais de autenticação do Google Sheets conforme descrito na documentação oficial.

## Utilização

1. Clone o repositório do GitHub ou baixe o script Python.
2. Instale as bibliotecas necessárias via pip:
   `pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client`
3. Configure as credenciais de autenticação do Google Sheets conforme descrito na documentação oficial.
4. Execute o script Python `main.py`:


## Funcionalidades

- **Cálculo de Média dos Alunos:** O script lê os dados de uma planilha do Google Sheets, calcula a média das notas dos alunos e determina suas situações acadêmicas com base nas notas e na quantidade de faltas.
- **Cálculo de Notas Necessárias para Aprovação Final:** Além disso, o script calcula as notas necessárias para a aprovação final dos alunos que estão em situação de "Exame Final".

## Estrutura do Código

O código do script está dividido em diferentes seções:

- Importação de bibliotecas necessárias.
- Definição de escopos e intervalo da planilha.
- Função principal `main()` que realiza a interação com o Google Sheets.
- Carregamento das credenciais de autenticação do Google Sheets.
- Extração dos dados da planilha e cálculo das médias dos alunos.
- Determinação das situações acadêmicas dos alunos.
- Cálculo das notas necessárias para aprovação final.
- Atualização das informações na planilha do Google Sheets.
