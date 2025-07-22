**README.md**

# Projeto de Planejamento de Sprint

Este projeto tem como objetivo automatizar o planejamento de sprints, analisando a disponibilidade de horários em uma planilha do Google Sheets, criando eventos no Google Calendar e enviando mensagens via WhatsApp. Foi criado para facilitar a coordenação da equipe de tecnologia da ONG Escrevendo na Quebrada, marcando automaticamente as quatro sprints do mês, sem intervenção do usuário. Isso evita esquecimentos e simplifica o processo de identificar os horários com maior disponibilidade da equipe, agendando as reuniões nessas datas e horários automaticamente.

Esse projeto é uma continuação do seguinte projeto: https://github.com/taina-coder/envio-programado-de-mensagens

## Tecnologias Utilizadas

- Python
- gspread
- Google API Client
- Selenium
- datetime
- re

## Funcionalidades

- **Análise de Disponibilidade**: Lê uma planilha do Google Sheets para verificar a disponibilidade de horários.
- **Geração de Datas**: Calcula as datas para os próximos 4 encontros com base na disponibilidade.
- **Criação de Eventos no Google Calendar**: Cria eventos no Google Calendar com os horários selecionados.
- **Envio de Mensagens via WhatsApp**: Envia uma mensagem para um grupo no WhatsApp com os detalhes dos encontros.

## Pré-requisitos

- Python 3.x
- Bibliotecas Python:
  - `gspread`
  - `google-auth`
  - `google-api-python-client`
  - `selenium`
  
Instale as bibliotecas necessárias com o seguinte comando:

```bash
pip install gspread google-auth google-api-python-client selenium
```

## Configuração

1. **Credenciais do Google**:
   - Crie um projeto no [Google Cloud Console](https://console.cloud.google.com/).
   - Ative as APIs do Google Sheets e Google Calendar.
   - Crie uma conta de serviço e baixe o arquivo `credentials.json`.
   - Compartilhe a planilha do Google Sheets com o e-mail da conta de serviço.

2. **Configuração do Selenium**:
   - Certifique-se de ter o ChromeDriver instalado e que ele seja compatível com a versão do seu navegador Chrome.

## Uso

1. Altere a variável `url_planilha` no código para o link da sua planilha do Google Sheets.
2. Execute o script:

```bash
python seu_script.py
```

## Estrutura do Código

- `analisar_disponibilidade(sheet_url)`: Analisa a planilha e retorna os horários disponíveis.
- `gerar_datas_para_4_semanas(selecionados)`: Gera as datas para os próximos 4 encontros.
- `criar_eventos_google_calendar(eventos)`: Cria eventos no Google Calendar.
- `enviar_mensagem_whatsapp(mensagem, grupo)`: Envia uma mensagem para um grupo no WhatsApp.
- `montar_mensagem(links)`: Monta a mensagem a ser enviada no WhatsApp.

## Contribuição

Sinta-se à vontade para contribuir com melhorias ou correções. Faça um fork do repositório e envie um pull request.

---

**Nota**: Certifique-se de que todas as dependências estão instaladas e que as credenciais estão configuradas corretamente antes de executar o script.
