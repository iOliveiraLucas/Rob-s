''' app.py '''

# Libraries
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime as dt
from time import sleep
import schedule

# init.py
from init import execute_query_connection, conn_db

# util.py
from util import remove_special_chars, read_excel, get_date, convert_dict
from util import rename_file, remove_old, convert_date, normalization, is_loaded

# dados de acesso
data = {
        'url':'https://hubdigital.stoque.com.br/PortalImobiliario/Portal/',
        'Login': '',
        'Senha': ''
        }

# Configurações WebDriver
options = Options()
options.add_argument('--headless') #enabled in productions
options.add_argument('--no-sandbox') #enabled in productions
options.add_argument("--disable-setuid-sandbox") #enabled in productions
options.add_argument("window-size=1920x1080") #enabled in productions
options.add_argument("--disable-dev-shm-usage") #enabled in productions
options.add_argument("start-maximized") #enabled in productions
options.add_argument("--log-level=3") #enabled in productions

chrome = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


def login():
    logged = False

    chrome.get(url=data['url'])
    chrome.find_element(By.ID,"txtUsuario").send_keys(data["Login"])
    chrome.find_element(By.ID,"txtSenha").send_keys(data["Senha"])
    chrome.find_element(By.ID,"btnEntrar").click()
    loaded = is_loaded(chrome, By)
    if loaded:
        try:
            log_in = chrome.find_element(By.ID,"logoff-action")
            if log_in:
                logged = True
        except:
            pass

    return logged


def filter(date_start, date_end):
    filtro = chrome.find_element(By.ID,"btnFiltros")
    if filtro:
        filtro.click()

    data_inicio = chrome.find_element(By.ID,"txtDataInicio")
    if data_inicio:
        chrome.execute_script(f"arguments[0].value='{date_start}';", data_inicio)

    data_fim = chrome.find_element(By.ID,"txtDataFim")
    if data_fim:
        chrome.execute_script(f"arguments[0].value='{date_end}';", data_fim)

    carregar = chrome.find_element(By.ID,"btnCarregar")
    if carregar:
        carregar.click()

    loaded = is_loaded(chrome, By)
    if loaded:

        relatorio = chrome.find_element(By.XPATH,'//*[@id="btnRelatorio"]')
        if relatorio:
            chrome.execute_script("arguments[0].click();", relatorio)
            sleep(5)
            rename_file()


def insert_db(value):
    conn = conn_db()
    if conn['status']:
        conn=conn['content']
        cur = conn.cursor()

        update = update_db(value)
        if update != []:
            dict_update = convert_dict(update)
            for item in dict_update:

                numero_proposta= str(item['numero_proposta'])
                status= str(item['status'])
                primeiro_comprador= str(item['primeiro_comprador'])
                cpf= str(item['cpf'])
                corban= str(item['corban'])
                pdv= str(item['pdv'])
                valor_financiamento= remove_special_chars(item['valor_financiamento'])
                taxa_anual= str(item['taxa_anual'])
                data_envio= convert_date(item['data_envio'])
                data_emissao= convert_date(item['data_emissao'])
                data_assinatura= convert_date(item['data_assinatura'])
                proposta_enviada_por= str(item['proposta_enviada_por'])
                produto= str(item['produto'])
                created_at = str(dt.now()).split('.', maxsplit=1)[0]

                content= "INSERT INTO projects " \
                        "(numero_proposta, status, primeiro_comprador, cpf, corban, pdv, valor_financiamento, taxa_anual, data_envio, data_emissao, data_assinatura, proposta_enviada_por, produto, created_at) " \
                        "VALUES ("+numero_proposta+",'"+status+"','"+primeiro_comprador+"','"+cpf+"','"+corban+"',"+pdv+","+valor_financiamento+","+taxa_anual+","+data_envio+","+data_emissao+","+data_assinatura+",'"+proposta_enviada_por+"','"+produto+"','"+created_at+"')"
                query1 = execute_query_connection(cur=cur,content=content)
                if query1['status']:
                    conn.commit()
                else:
                    print(query1['content'])
        else:
            print("Sem alteração no db...")

        cur.close()
        conn.close()
        remove_old()


def update_db(list_excel):
    update_list = []

    conn = conn_db()
    if conn['status']:
        conn=conn['content']
        cur = conn.cursor()

        for item_excel in list_excel:
            content = "SELECT numero_proposta, status, primeiro_comprador, cpf, corban, pdv, valor_financiamento, taxa_anual, data_envio, data_emissao, data_assinatura, proposta_enviada_por, produto " \
                "FROM projects WHERE numero_proposta ="+str(item_excel[0])+" ORDER BY created_at DESC LIMIT 1"
            query1 = execute_query_connection(cur=cur,content=content)

            if query1['content'] != []:
                alter = False
                data_db = query1['content'][0]

                for i, item_db in enumerate(data_db):
                    item_norm = normalization(item_excel[i])
                    if str(item_db) != str(item_norm):
                        alter = True
                if alter:
                    update_list.append(item_excel)
            else:
                update_list.append(item_excel)

        cur.close()
        conn.close()

    return update_list


def lgpd(): # temporario
    try:
        btn_lgpd = chrome.find_element(By.XPATH,'//div[@id="ModalLGPD"]//button[@class="btn btn-default bg-itau"]')
        if btn_lgpd:
            btn_lgpd.click()
    except:
        pass

def main():
    logged= login()
    if logged:
        lgpd() # temporario
        DATES_LIST = [(150,121),(120,91),(90,61),(60,31),(30,0)] # Intervalo de tempo em dias passado no filtro do site (intervalo maximo= 30 dias)

        for dates in DATES_LIST:
            filter(get_date(dates[0]), get_date(dates[1]))
            list_excel = read_excel()
            insert_db(list_excel)

        logoff = chrome.find_element(By.ID,"logoff-action")
        if logoff:
            logoff.click()
      
    else:
        print('Falha no login...')


# schedule.every().day.at("19:00").do(main)
schedule.every(1).hour.do(main)


if __name__ == '__main__':
    main()
    while 1:
        schedule.run_pending()
        sleep(1)