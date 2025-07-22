import re
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
from googleapiclient.discovery import build
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import time


def analisar_disponibilidade(sheet_url):
    print("Iniciando análise da planilha...")
    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
        client = gspread.authorize(creds)

        match = re.search(r"/d/([a-zA-Z0-9-_]+)", sheet_url)
        if not match:
            raise ValueError("URL da planilha inválida.")
        sheet_id = match.group(1)

        sheet = client.open_by_key(sheet_id)
        aba = sheet.sheet1
        dados = aba.get_all_values()

        inicio_tabela = 2
        fim_tabela = inicio_tabela + 14
        dias_semana = dados[1][1:7]

        contagem = []

        for indice_linha, linha in enumerate(dados[inicio_tabela:fim_tabela]):
            horario = linha[0]
            for j, valor in enumerate(linha[1:7]):
                try:
                    qtd = int(valor)
                    dia = dias_semana[j]
                    contagem.append(((dia, horario), qtd))
                except ValueError:
                    pass

        # Ordena pela quantidade desc, depois horário asc
        contagem.sort(key=lambda x: (-x[1], x[0][1]))

        selecionados_local = []
        dias_usados = set()

        for (dia, horario), qtd in contagem:
            # Permite 2 dias diferentes (dias_usados < 2) ou mesmo dia repetido se já estiver na lista
            if dia in dias_usados or len(dias_usados) < 2:
                selecionados_local.append((dia, horario))
                dias_usados.add(dia)
            if len(selecionados_local) == 2:
                break

        # Se não achou dois horários, mantém o que tiver
        if len(selecionados_local) < 2:
            selecionados_local = selecionados_local[:2]

        # Não retorna os 4 horários repetidos aqui, só os 2 escolhidos
        print(f"Análise da planilha concluída. Selecionados locais: {selecionados_local}")
        return selecionados_local

    except Exception as e:
        print(f"Erro ao analisar a planilha: {e}")
        return None


# Mapeia dia da semana em português para índice do Python (segunda=0)
dias_semana_map = {
    'segunda-feira': 0,
    'terça-feira': 1,
    'quarta-feira': 2,
    'quinta-feira': 3,
    'sexta-feira': 4,
    'sábado': 5,
    'domingo': 6
}


def gerar_datas_para_4_semanas(selecionados):
    hoje = datetime.now()
    resultados = []

    n = len(selecionados)
    dia_semana = selecionados[0][0].lower()
    dia_idx = dias_semana_map.get(dia_semana)
    if dia_idx is None:
        print(f"Dia da semana desconhecido: {dia_semana}")
        return []

    # Calcula quantos dias até o próximo dia da semana desejado (pula hoje)
    dias_ate_dia = (dia_idx - hoje.weekday() + 7) % 7
    if dias_ate_dia == 0:
        dias_ate_dia = 7
    primeira_data_base = hoje + timedelta(days=dias_ate_dia)
    primeira_data_base = primeira_data_base.replace(second=0, microsecond=0)

    for i in range(4):
        horario = selecionados[i % n][1]
        hora_dt = datetime.strptime(horario, "%H:%M").time()

        data_evento = primeira_data_base + timedelta(weeks=i)
        data_evento = data_evento.replace(hour=hora_dt.hour, minute=hora_dt.minute)

        resultados.append(data_evento.strftime("%d/%m/%Y %H:%M"))

    return resultados


def criar_eventos_google_calendar(eventos):
    print("Iniciando criação dos eventos no Google Calendar...")
    try:
        scopes = ['https://www.googleapis.com/auth/calendar']
        creds = Credentials.from_service_account_file('credentials.json', scopes=scopes)

        print("Email da conta de serviço:", creds.service_account_email)

        service = build('calendar', 'v3', credentials=creds)
        links_criados = []

        for data_hora in eventos:
            # data_hora é string no formato "dd/mm/yyyy HH:MM"
            evento_data = datetime.strptime(data_hora, "%d/%m/%Y %H:%M")
            fim = evento_data + timedelta(hours=1)

            evento = {
                'summary': 'Planejamento de Sprint',
                'description': 'Reunião semanal de planejamento da sprint.',
                'start': {'dateTime': evento_data.isoformat(), 'timeZone': 'America/Sao_Paulo'},
                'end': {'dateTime': fim.isoformat(), 'timeZone': 'America/Sao_Paulo'},
            }

            evento_criado = service.events().insert(calendarId='ID DO SEU GOOGLE CALENDAR', body=evento).execute()
            links_criados.append((evento_data.strftime('%d/%m %H:%M'), evento_criado.get('hangoutLink', '')))

        print("Eventos criados com sucesso no Google Calendar.")
        return links_criados

    except Exception as e:
        print(f"Erro ao criar eventos no Google Calendar: {e}")
        return None


def enviar_mensagem_whatsapp(mensagem, grupo):
    def log(m):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {m}")

    log("Iniciando envio de mensagem pelo WhatsApp...")

    options = Options()
    options.add_argument(r"CAMINHO DO SEU PERFIL")

    driver = None

    try:
        driver = webdriver.Chrome(options=options)
        driver.get("https://web.whatsapp.com/")

        log("Aguardando login e carregamento do WhatsApp Web...")

        WebDriverWait(driver, 10).until(
            ec.any_of(
                ec.presence_of_element_located((By.XPATH, "//canvas[@aria-label='Scan me!']")),
                ec.presence_of_element_located((By.XPATH, "//div[@role='button' or @role='gridcell']"))
            )
        )
        log("WhatsApp Web carregado ou aguardando login manual.")

        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((By.XPATH, "//div[@role='button' or @role='gridcell']"))
        )
        log("Interface do WhatsApp Web pronta.")

        try:
            log("Tentando encontrar campo de busca pelo XPath com title...")
            search_box = WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((By.XPATH, '//div[@title="Caixa de texto de pesquisa"]'))
            )
            log("Campo de busca encontrado pelo XPath com title.")
        except Exception as e1:
            log(f"Falha ao encontrar campo de busca pelo XPath com title: {e1}")
            log("Tentando encontrar campo de busca pelo XPath com data-tab='3'...")
            search_box = WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((By.XPATH, "//div[@contenteditable='true' and @data-tab='3']"))
            )
            log("Campo de busca encontrado pelo XPath com data-tab='3'.")

        search_box.click()
        log("Clicando no campo de busca...")
        search_box.clear()
        search_box.send_keys(grupo)
        time.sleep(2)
        search_box.send_keys(Keys.ENTER)
        log(f"Grupo pesquisado: {grupo}")

        grupo_element = WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((By.XPATH, f"//span[@title='{grupo}']"))
        )
        grupo_element.click()
        log(f"Entrou no grupo {grupo}.")

        log("Aguardando caixa de mensagem...")
        msg_box = WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((By.XPATH, "//div[@contenteditable='true' and @data-tab='10']"))
        )
        log("Caixa de mensagem encontrada.")

        msg_box.click()
        msg_box.send_keys(mensagem)
        msg_box.send_keys(Keys.ENTER)
        log("Mensagem enviada.")

    except Exception as e:
        log(f"Erro ao enviar mensagem: {e}")

    finally:
        time.sleep(5)
        if driver is not None:
            driver.quit()
            log("Navegador encerrado.")


def montar_mensagem(links):
    mes_atual = datetime.now().strftime("%B").upper()

    mensagem = f"Olá, pessoal! Aqui estão os encontros definidos para o planejamento de sprint para o mês de {mes_atual}:\n\n"

    for data, _ in links:
        mensagem += f"{data}\n"

    return mensagem


if __name__ == "__main__":
    url_planilha = "LINK DA SUA PLANILHA AQUI"

    selecionados = analisar_disponibilidade(url_planilha)

    if selecionados is None:
        print("Erro ao analisar a planilha. Encerrando.")
        exit()

    datas_completas = gerar_datas_para_4_semanas(selecionados)
    print(f"Datas completas para criação dos eventos: {datas_completas}")

    links = criar_eventos_google_calendar(datas_completas)

    if links is None:
        print("Não foi possível criar eventos no Google Calendar. Encerrando.")
        exit()

    mensagem = montar_mensagem(links)

    enviar_mensagem_whatsapp(mensagem, grupo="teste")
