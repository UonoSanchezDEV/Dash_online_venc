import shutil
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from time import sleep
from datetime import date, timedelta, datetime
from os import listdir
from os.path import isfile, join, basename
from pathlib import Path
from time import sleep
from dotenv import load_dotenv
import time
from glob import glob

load_dotenv()

login_usuario_bradesco = os.getenv('LOGIN_USUARIO_BRADESCO')
login_senha_bradesco = os.getenv('LOGIN_SENHA_BRADESCO')

login_usuario_uono_relatorio = os.getenv('LOGIN_USUARIO_EMAIL_UONO')
login_senha_uono_relatorio = os.getenv('LOGIN_SENHA_EMAIL_UONO')

login_usuario_inspectos = os.getenv('LOGIN_USUARIO_INSPECTOS')
login_senha_inspectos = os.getenv('LOGIN_SENHA_INSPECTOS')

login_usuario_uono = os.getenv('LOGIN_USUARIO_EMAIL_UONO_UONO')
login_senha_uono = os.getenv('LOGIN_SENHA_EMAIL_UONO_UONO')

login_usuario_cetip = os.getenv('LOGIN_USUARIO_EMAIL_CETIP')
login_senha_cetip = os.getenv('LOGIN_SENHA_EMAIL_CETIP')

login_usuario_viva = os.getenv('LOGIN_USUARIO_EMAIL_VIVA')
login_senha_viva = os.getenv('LOGIN_SENHA_EMAIL_VIVA')



#Função para dar tempo de espera no splinter
def wait_for_element_xpath(browser, xpath, timeout=10):
    start_time = time.time()
    while time.time() - start_time < timeout:
        if browser.is_element_present_by_xpath(xpath, wait_time=1):
            return True
        time.sleep(0.5)
    return False


#Funções para ajustar vencimentos
def calcular_vencimento(data_envio, horas):
        """
        Função para calcular a data de vencimento com base em dias úteis e horas remanescentes.

        Args:
            data_envio: Data e hora de envio da solicitação (objeto datetime).
            horas: Número de horas a serem adicionadas (float).

        Returns:
            Data de vencimento (objeto datetime).
        """
        try:
            # Criando objetos timedelta
            hora_por_dia = pd.Timedelta(hours=10)
            hora_total = pd.Timedelta(hours=horas)

            # Calculando dias úteis e horas remanescentes
            dias_uteis_necessarios = hora_total // hora_por_dia
            hora_remanescente = hora_total % hora_por_dia

            # Ajustando data de vencimento com base em dias úteis
            data_vencimento = data_envio + dias_uteis_necessarios * pd.offsets.BDay()

            # Considerando horas remanescentes no último dia útil
            while data_vencimento.weekday() in [5, 6]:  # Verificando se cai em fim de semana
                data_vencimento += pd.offsets.BDay(1)  # Adiantando para o próximo dia útil

            # Adicionando horas remanescentes à data de vencimento final
            data_vencimento += hora_remanescente

            if data_vencimento.hour >= 18:
                data_vencimento += pd.offsets.BDay(1)
                data_vencimento = data_vencimento.replace(hour=8 + (data_vencimento.hour - 18))
            elif data_vencimento.hour < 8:
                data_vencimento = data_vencimento.replace(hour=8, minute=0)

            return data_vencimento
        
        except Exception as e:
            print(f'Ocorreu o seguinte erro {e}')


def somar_horas_uteis(data_envio, horas):
    """
    Função para somar horas considerando apenas dias úteis.

    Args:
        data_envio: Data e hora de envio da solicitação (objeto datetime).
        horas: Número de horas a serem adicionadas (float).

    Returns:
        Data de vencimento (objeto datetime).
    """
    try:
        horas_por_dia = 10
        dias = int(horas // horas_por_dia)
        horas_restantes = horas % horas_por_dia
        
        # Adiciona dias úteis
        data_vencimento = data_envio
        for _ in range(dias):
            data_vencimento += pd.offsets.BDay(1)
        
        # Adiciona horas restantes
        data_vencimento += timedelta(hours=horas_restantes)
        
        # Ajusta caso caia fora do horário comercial (8h-18h)
        if data_vencimento.hour >= 18:
            data_vencimento += pd.offsets.BDay(1)
            data_vencimento = data_vencimento.replace(hour=8 + (data_vencimento.hour - 18))
        elif data_vencimento.hour < 8:
            data_vencimento = data_vencimento.replace(hour=8, minute=0)
    
        return data_vencimento
    
    except Exception as e:
        print(f'Ocorreu um erro {e}')


def data_vistoria_pra_vencimento(data_vistoria, data_criação):

    try:
        
        # Adiciona dias úteis
        if data_criação == data_vistoria:
            data_vencimento = data_vistoria
            data_vencimento += pd.offsets.BDay(2)
        else:
            data_vencimento = data_vistoria
            data_vencimento += pd.offsets.BDay(1)
        
        if data_vencimento.weekday in [5,6]:
            data_vencimento += pd.offsets.BDay(1)
        
        return data_vencimento
    
    except Exception as e:
        print(f'Ocorreu um erro {e}')


#Funções de Busca de Laudos
def busca_bradesco():
    """_Buscador de arquivos no site do Bradesco_

    Sem necessidade de argumentos.

    Entra no site, faz todos os acessos sozinho e conclui baixando os arquivos de Excel disponivel
    """
    try:
        driver = webdriver.Chrome()
        driver.get('https://avaliacaobra.com.br/')
        driver.minimize_window()
        
        elemento_botao = driver.find_element(By.ID, 'btnFornec')
        elemento_botao.click()

        elemento_input_login = driver.find_element(By.ID, 'txtUsuario')
        elemento_input_login.send_keys(login_usuario_bradesco)

        elemento_input_senha = driver.find_element(By.ID, 'txtSenha')
        elemento_input_senha.send_keys(login_senha_bradesco)

        elemento_botao_enviar = driver.find_element(By.ID, 'btnEnviar')
        elemento_botao_enviar.click()

        janelas = driver.window_handles
        driver.switch_to.window(janelas[0])

        elemento_hamb = driver.find_element(By.XPATH, '/html/body/nav/div[2]/div[1]/ul/li[3][@title ="Relatórios"]')
        elemento_hamb.click()

        sleep(1)

        elemento_laudo = driver.find_element(By.XPATH, '/html/body/nav/div[2]/div[2]/ul/li[1][@title ="Em Andamento"]')
        elemento_laudo.click()

        sleep(6)

        elemento_hamb_concluidos = driver.find_element(By.XPATH, '//*[@id="my-icon"]/span')
        elemento_hamb_concluidos.click()

        sleep(1)

        elemento_laudo = driver.find_element(By.XPATH, '//*[@id="id_3,02"]')
        elemento_laudo.click()

        sleep(5)

        iframe = driver.find_element(By.XPATH, '//*[@id="Cont"]')
        driver.switch_to.frame(iframe)
        element = driver.find_element(By.XPATH, '//*[@id="TextBox4"]')
        element.send_keys(Keys.CONTROL, 'a')
        element.send_keys(Keys.ENTER)

        sleep(5)
    
    except:
        print("Não foi possivel baixar os arquivos do BRADESCO")


def busca_inspectos():
    def acessa_email():
        """_Captura o codigo da B3 no email 'webmail'_

        Args:
            driver (_str_): _Driver do navegador_

        Returns:
            _num_: _retorna o capturado do email_
        """
        try:
            driver = webdriver.Chrome()
            driver.get('https://webmail.uonosanchez.com.br/?_task=mail&_mbox=INBOX')
            wait = WebDriverWait(driver, 100)

            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="rcmloginuser"]'))).send_keys(login_usuario_uono_relatorio)
            driver.find_element(By.XPATH, '//*[@id="rcmloginpwd"]').send_keys(login_senha_uono_relatorio)
            driver.find_element(By.XPATH, '//*[@id="rcmloginsubmit"]').click()

            sleep(13)
            # Atualiza a lista de emails
            driver.find_element(By.XPATH, '//*[@id="rcmbtn108"]').click()
            driver.refresh()

            # Acessa o primeiro email
            wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[4]/table[2]/tbody/tr[1]'))).click()
            sleep(2.5)

            # Switch para o iframe e captura o código
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="messagecontframe"]')))
            elements = driver.find_elements(By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/div[2]/div/div/center/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td/b')
            if elements:
                codigo = elements[0].text
            else:
                codigo = None

            driver.switch_to.default_content()
            driver.quit()

            return codigo
        
        except:
            print(f"Não foi possivel baixar os Laudos INSPECTOS")
    """_Buscador de arquivos no site do Santander_

    Sem necessidade de argumentos.

    Entra no site, faz todos os acessos sozinho e conclui baixando os arquivos de Excel disponivel
    """
    try:
        driver = webdriver.Chrome()
        wait = WebDriverWait(driver, 10) 
        driver.get('https://inspectos.com/sistema/index.html#/home')

        elemento_input_login = wait.until(EC.visibility_of_element_located((By.NAME, 'email'))).send_keys(login_usuario_inspectos)
        elemento_input_senha = wait.until(EC.visibility_of_element_located((By.NAME, 'senha'))).send_keys(login_senha_inspectos)
        elemento_botao_enviar = driver.find_element(By.ID, 'enter').click()
        enviar_codigo = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/div/div/div[4]/div[2]/div/div/div/div[2]/div/a[1]/div[1]/span/i[2]'))).click()
        button_avanc = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/div/div/div[4]/div[2]/div/div/div/div[2]/button'))).click()
        codigo = acessa_email()
        print(codigo)
        codigos_ = {
            '//*[@id="first"]': codigo[0],
            '//*[@id="second"]': codigo[1],
            '//*[@id="third"]': codigo[2],
            '//*[@id="fourth"]': codigo[3]
        }

        for xpath, value in codigos_.items():
            try:
                elemento_input_codigo = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
                elemento_input_codigo.clear()  # Limpa o campo antes de inserir o valor
                elemento_input_codigo.send_keys(value)
            except Exception as e:
                print(f"Erro ao processar o elemento {xpath}: {e}")
        enviar = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/div/div/div[4]/div[2]/div/div/div/div[2]/form/button')))
        enviar.click()
        elemento_hamb = wait.until(EC.visibility_of_element_located((By.XPATH, "//*[@id='menu-principal']/button"))).click()
        elemento_botao = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'accordion-toggle'))).click()
        elemento_botao_prox = wait.until(EC.visibility_of_element_located((By.XPATH, '//div[@ng-click="go(item.recurso.rota)"]'))).click()
        elemento_botao_analitico = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/uib-accordion/div/div/div[2]/div/div/div/div[2]/div/div/div[1][@is-open="item.open"]'))).click()
        data_hoje = date.today()

        data_anterior = data_hoje - timedelta(days=25)
        data_mes_anterior = data_hoje - timedelta(days=20)

        data_formatada = data_anterior.strftime('%d/%m/%Y')
        elemento_input_data_inicio = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div[2]/div/div/div[2]/div[2]/div[2]/div[3]/datepicker[1]/input[@type="text"]')))

        sleep(3.5)
        elemento_input_data_inicio.send_keys(Keys.CONTROL, 'a')
        sleep(0.5)
        elemento_input_data_inicio.send_keys(data_formatada)
        sleep(5)

        elemento_input_botao_filtro1 = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div[2]/div/div/div[2]/div[1]/div[2]/div[11]/div[2]/div/div[1]/input').click()
        sleep(0.5)

        elemento_input_botao_data = driver.find_element(By.XPATH, '//*[@id="idFiltroDataAgendamento"]/datepicker[1]/input')
        elemento_input_botao_data.click()
        elemento_input_botao_data.send_keys(data_mes_anterior.strftime('%d/%m/%Y'))
        
        elemento_input_botao_data_hj = driver.find_element(By.XPATH, '//*[@id="idFiltroDataAgendamento"]/datepicker[2]/input')
        elemento_input_botao_data_hj.click()
        elemento_input_botao_data_hj.send_keys(data_hoje.strftime('%d/%m/%Y'))
        
        elemento_input_botao = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div[2]/div/div/div[2]/div[2]/div[2]/div[5]/button[2][@type="button"]').click()
        sleep(5.5)

        elemento_input_botao_excel = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div[2]/div/div/div[2]/div[2]/div[2]/div[5]/button[1][@type="button"]').click()
        sleep(7)

    except Exception as e:
        print(f"Ocorreu um erro {e}")

    finally:
        driver.quit()


def busca_cetip():
    def acessa_email():
        """Captura o código da B3 no email 'webmail'
        Returns:
            num: retorna o capturado do email
        """
        try:
            driver = webdriver.Chrome()
            driver.get('https://webmail.uonosanchez.com.br/?_task=mail&_mbox=INBOX')
            wait = WebDriverWait(driver, 100)

            # Login no email
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="rcmloginuser"]'))).send_keys(login_usuario_uono)
            driver.find_element(By.XPATH, '//*[@id="rcmloginpwd"]').send_keys(login_senha_uono)
            driver.find_element(By.XPATH, '//*[@id="rcmloginsubmit"]').click()

            sleep(11)
            # Atualiza a lista de emails
            driver.find_element(By.XPATH, '//*[@id="rcmbtn108"]').click()
            driver.refresh()

            # Acessa o primeiro email
            wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[4]/table[2]/tbody/tr[1]'))).click()
            sleep(2.5)

            # Switch para o iframe e captura o código
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="messagecontframe"]')))
            elements = driver.find_elements(By.XPATH, '//*[@id="message-htmlpart1"]/div/table[2]/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr/td/div/p[3]/b')

            if elements:
                codigo = elements[0].text
            else:
                codigo = None

            driver.switch_to.default_content()
            driver.quit()

            return codigo

        except Exception as e:
            print(f"Ocorreu um erro: {e}")
            return None

    # Acessando o site da B3
    options = webdriver.ChromeOptions()
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3")
    driver = webdriver.Chrome(options=options)
    driver.set_window_size(width=945, height= 1020)
    wait = WebDriverWait(driver, 100)

    try:
        driver.get('https://auth4b3-pf.b3.com.br/as/authorization.oauth2?client_id=CADU_UIF_CLIENT&response_type=code&redirect_url=https://cadu2.b3.com.br/auth/home&spa=CADUAUTHMFARISK')
        # Preenchendo login e senha
        wait.until(EC.visibility_of_element_located((By.NAME, 'pf.username'))).send_keys(login_usuario_cetip)
        driver.find_element(By.NAME, 'pf.pass').send_keys(login_senha_cetip)
        driver.find_element(By.ID, 'Btn_CONTINUE').click()

        # Envia código de verificação para o email
        if wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div[2]/div[2]'))):
            driver.find_element(By.XPATH, '//*[@id="app"]/div[2]/div[2]').click()

        # Captura o código do email
        codigo = acessa_email()
        if codigo:
            driver.find_element(By.NAME, 'otp').send_keys(codigo)
            driver.find_element(By.ID, 'sign-on').click()

        # Acessa a página de produtos
        if wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="products"]/div[1]/div[1]/a'))):
            driver.find_element(By.XPATH, '//*[@id="products"]/div[1]/div[1]/a').click()

        driver.find_element(By.XPATH, '//*[@id="AssignedToTeam0"]/td[1]/span').click()

        #  sleep(12)

        # Preenchendo filtros de busca
        # pyautogui.press('tab', presses=20)
        # pyautogui.press('enter')

        # pyautogui.press('tab', presses=16)
        # pyautogui.press('down')
        # pyautogui.press('tab')

        # dia_hoje = datetime.today()
        # data_anterior = dia_hoje - timedelta(days=210)
        # pyautogui.write(data_anterior.strftime('%d/%m/%Y'))

        # pyautogui.press('tab')
        # pyautogui.write(dia_hoje.strftime('%d/%m/%Y'))

        # pyautogui.press('tab', presses=42)
        # pyautogui.press('enter')

        # sleep(12)

        # # Selecionando laudos fechados
        # pyautogui.press('tab', presses=7)
        # pyautogui.press('down')
        # pyautogui.press('tab')
        # pyautogui.press('enter')
        # pyautogui.press('tab')
        # pyautogui.press('enter')

        # # Selecionando vencimentos do dia
        # pyautogui.keyDown('shift')
        # pyautogui.press('tab', presses=5)
        # pyautogui.keyUp('shift')
        # pyautogui.press('enter')

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
    finally:
        sleep(5)
        driver.quit()


def busca_vivaintra_bradesco():
    """_Buscador de Laudos no site 'VIVAINTRA'
    Busca os 3 tipos de Laudos, e busca cada parte individualmente
    """
    try:
        driver = webdriver.Chrome()

        data_hoje = date.today()
        data_inicio_mes = data_hoje.replace(day=1)
        data_anterior = data_inicio_mes - timedelta(days=10)

        wait = WebDriverWait(driver, 140) 
        driver.get("https://uonosanchez.vivaintra.com/admin-blog")
        
        # Entra na conta
        elem_email_vivaintra = wait.until(EC.visibility_of_element_located((By.NAME, "username")))
        elem_email_vivaintra.send_keys(login_usuario_viva)
        elem_email_vivaintra.send_keys(Keys.ENTER)
        
        elem_senha = driver.find_element(By.NAME, "password")
        elem_senha.send_keys(login_senha_viva)
        elem_senha.send_keys(Keys.ENTER)

        #Entrando na pagina ADM
        elem_seleciona = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '/html/body/div[3]/div/div[2]/div/form/input'
        )))
        elem_seleciona.send_keys(login_senha_viva)
        elem_seleciona.send_keys(Keys.ENTER)

        elem_seleciona_produtividade = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '/html/body/nav/div[1]/div[1]/a[2]'
        ))).click()

        elem_seleciona_requisicoes = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[2]/div/div[2]/div/div[10]/a/div'
        ))).click()

        elem_seleciona_toda_as_requisicoes = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[1]/div[2]/ul/li[1]'
        ))).click()

        seleciona_bradesco = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[2]/div/select/option[10]'
        ))).click()

        seleciona_bradesco_data = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[11]/div[1]/input'
        ))).send_keys(data_anterior.strftime('%d/%m/%Y'))

        seleciona_bradesco_data_segundo_input = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[11]/div[3]/div/input'
        ))).send_keys(data_hoje.strftime('%d/%m/%Y'))
        
        seleciona_bradesco_data_segundo_input_botao = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[1]/div[4]/div/div[2]/form/div[13]/div/button'
        ))).send_keys(Keys.ENTER)

        seleciona_bradesco_data_segundo_input_botao_cria_excel = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '/html/body/nav/div[3]/a'
        ))).click()

        seleciona_bradesco_data_segundo_input_botao_baixa_excel = wait.until(EC.visibility_of_element_located(
        (By.XPATH,
            '//*[@id="body-content"]/div[2]/div[2]/div[1]/a' 
        ))).click()
        
        caminho = encontrar_arquivo_mais_recente('Exportacao','C:\\Users\\thiago.oliveira\\Downloads','csv')

        converte_em_excel(caminho,'bradesco_viva')


        sleep(10)

    except Exception as e:
        print(f'Ocorreu o seguinte erro {e}')


    driver.quit()


#Funções de manipulação de arquivo
def move(path_origem, path_destino):
    """Move o arquivo da pasta de origem (path_origem) para a pasta de destino (path_destino)

    Args:
        path_origem (_str_): _Pasta onde o arquivo está localizado_
        path_destino (_str_): _Pasta onde o arquivo será encaminhado_
    """
    try:
        for item in [join(path_origem, f) for f in listdir(path_origem) if isfile(join(path_origem, f))]:
            #print(item)
            shutil.move(item, join(path_destino, basename(item)))
            print(f'Arquivo(s) movido de "{item}" para --> "{join(path_destino, basename(item))}"')
    except Exception as e:
        print(f'Ocorreu o seguinte erro ao tentar mover o arquivo {e}')


def converte_em_excel(caminho, nome_saida=None):
    """
    Converte um arquivo CSV ou Excel (presumindo XLSX para Excel) para o formato Excel.
    
    Argumentos:
        caminho (str ou Path): Caminho para o arquivo de entrada.
        
    Retorno:
        Nenhum
    """
    try:
        caminho = str(caminho)  # Converte para string se for um Path
        if caminho.endswith('.csv'):
            # Se for um arquivo CSV
            read_file = pd.read_csv(caminho, sep=';')  # Ignora linhas com erros
        elif caminho.endswith('.xls') or caminho.endswith('.xlsx'):
            # Se for um arquivo Excel (XLS ou XLSX)
            read_file = pd.read_excel(caminho, engine='xlrd' if caminho.endswith('.xls') else 'openpyxl')
        else:
            print(f'Tipo de arquivo não suportado: {caminho}')
            return
    except Exception as e:
        print(f'Erro ao ler o arquivo {caminho}: {e}')
        return

    if nome_saida:
        caminho_saida = os.path.join(os.path.dirname(caminho), nome_saida + ".xlsx")
    else:
        if caminho.endswith('.csv'):
            caminho_saida = caminho.replace(".csv", ".xlsx")
        else:
            caminho_saida = caminho.replace(".xls", ".xlsx")

    try:
        read_file.to_excel(caminho_saida, index=None, header=True)
        print(f'Arquivo convertido com sucesso: {caminho_saida}')
    except Exception as e:
        print(f'Erro ao converter o arquivo {caminho} para Excel: {e}')



def encontrar_arquivo_mais_recente(nome_inicio_arquivo, pasta, extensao):
    """_Na pasta (pasta) buscara o arquivo mais recente_

    Args:
        nome_inicio_arquivo (_str_): _Nome do arquivo_
        pasta (_str_): _Local onde o arquivo está alocado_
        extensao (_str_): _Extensão do arquivo, Ex: .csv, .xlsx, .pdf, etc_

    Returns:
        _str_: _Retorna o caminho do arquivo mais recente_
    """
    try:
        lista_arquivos = Path(pasta).glob(f"{nome_inicio_arquivo}*.{extensao}")
        arquivo_mais_recente = max(lista_arquivos, key=os.path.getmtime, default=None)
        return arquivo_mais_recente
    except Exception as e:
        print(f"Ocorreu o seguinte erro {e}")


#Ajusta todos as planilhas para deixar no formato correto
def ajustar_inspectos():
    """_Ajusta a planilha da inspectos_
    """

    #Encotra o arquivo da inspectos
    try:
        user_dir = os.path.expanduser('~')
        # Constrói o caminho do arquivo de forma independente do usuário
        arquivos = glob(os.path.join(user_dir, 'Downloads', 'InspectosRelAnaliticoInspecoes-*.xls'))
        for arquivo in arquivos:
        #Ajusta a aba de Credito imobiliario
            db_inspectos = pd.read_excel(f"{arquivo}", sheet_name='Crédito imobiliário')
        lista = db_inspectos.columns.to_list()
        colunas_principais = ['Identificador','Data Limite', 'Data Entrega Laudo (B)','Nro. Proposta','Tipo Inspeção', 'Tipo Imovel', 'Município', 'Status']
        for coluna in lista:
            if coluna not in colunas_principais:
                db_inspectos = db_inspectos.drop(columns=[coluna])
                db_inspectos_ajustada_credito = db_inspectos
        
        #Ajusta a aba de Renegociação
        db_inspectos = pd.read_excel(f"{arquivo}", sheet_name='Renegociação')
        lista = db_inspectos.columns.to_list()
        colunas_principais = ['Identificador','Data Limite', 'Data Entrega Laudo (B)','Nro. Proposta','Tipo Inspeção', 'Tipo Imovel', 'Município', 'Status']
        for coluna in lista:
            if coluna not in colunas_principais:
                db_inspectos = db_inspectos.drop(columns=[coluna])
                db_inspectos_ajustada_rene = db_inspectos
            
        #Salva as abas em arquivos excel diferentes, depois move para a pasta de 'Controle de Laudos' e exclui o arquivo antigo da inspectos
        dataframe_sheet_credito = pd.DataFrame(db_inspectos_ajustada_credito)
        dataframe_sheet_rene = pd.DataFrame(db_inspectos_ajustada_rene)
        dataframe_sheet_rene.to_excel('ajustes\\inspectos_rene.xlsx', index=False)
        dataframe_sheet_credito.to_excel('ajustes\\inspectos_credito.xlsx', index=False)
        try:
            os.unlink(arquivo)
        except Exception as e:
            print(f'Ocorreu um erro {e}')
        dados = [dataframe_sheet_credito, dataframe_sheet_rene]
        pasta_base = os.listdir('ajustes')
        print(pasta_base)
        #for arquivo_dados in pasta_base:
            # print(arquivo_dados)
            #df = pd.read_excel(arquivo_dados)
            # dados.append(df)
            #print(dados)
        #Concatena os dois Excel's criados e junta em um unico 
        df_concatenado = pd.concat(dados, ignore_index=True)
        df_concatenado_excel = pd.DataFrame(df_concatenado)
        df_concatenado_excel.to_excel('ajustes\\inspectos_ajustada.xlsx', index=False)
        try:
            os.unlink('M:\\Thiago\\Laudos busca automatico\\ajustes\\inspectos_rene.xlsx')
            os.unlink('M:\\Thiago\\Laudos busca automatico\\ajustes\\inspectos_credito.xlsx')
        except Exception as e:
            print(f'Ocorreu um erro ao excluir os arquivos: {e}')
        move('ajustes', 'm:\\Thiago\\Controle de Laudos')
    except:
        print("Arquivo Inspectos não encontrado")


def ajustar_bradesco():
    """_Cria a coluna de vencimentos_
    """
    try:
        user_dir = os.path.expanduser('~')

        # Constrói o caminho do arquivo de forma independente do usuário
        file_path = os.path.join(user_dir, 'Downloads', 'EmAndamento.xlsx')

        db_base = pd.read_excel(file_path, skiprows=1)

        # Convertendo a coluna "Data Envio Solicitação" para datetime
        db_base["Data Envio Solicitação"] = pd.to_datetime(db_base["Data Envio Solicitação"], format='mixed', dayfirst=True)

        # Calculando datas de vencimento para cada linha do dataframe
        db_base["Vencimentos"] = db_base.apply(lambda row: somar_horas_uteis(row["Data Envio Solicitação"], 40), axis=1)

        # Exibindo as primeiras linhas do dataframe para verificar os resultados
        db_base.head()

        # Salvando o resultado no arquivo Excel
        db_base_excel = pd.DataFrame(db_base)
        db_base_excel.to_excel('ajustes\\EmAndamento_Atualizado.xlsx', index=False, engine='openpyxl')
        move('ajustes', 'm:\\Thiago\\Controle de Laudos')
    except:
        print('Arquivo Bradesco não encontrado')


def ajustar_cetip():
    """_Transforma o arquivo em um .xlsx_
    """
    try:
        user_dir = os.path.expanduser('~')

        # Constrói o caminho do arquivo de forma independente do usuário
        file_path = os.path.join(user_dir, 'Downloads', 'search_results.csv')
        db = pd.read_csv(file_path)

        # Salvando o resultado no arquivo Excel
        db_excel = pd.DataFrame(db)
        db_excel.to_excel('ajustes\\cetip.xlsx', index=False, engine='openpyxl')
        move('ajustes', 'm:\\Thiago\\Controle de Laudos')
    except FileNotFoundError:
        print("Arquivo CETIP não encontrado.")


def ajustar_bradesco_producao():
    """_Cria a coluna de vencimentos_

    Returns:
        _text_: _retorna o texto se concluido ou erro ocorrido._
    """
    try:
        user_dir = os.path.expanduser('~')

        # Constrói o caminho do arquivo de forma independente do usuário
        file_path = os.path.join(user_dir, 'Downloads', 'EmAndamento.xlsx')

        db_base = pd.read_excel(file_path, skiprows=1)

        # Convertendo a coluna "Data Envio Solicitação" para datetime
        db_base["Data Envio Solicitação"] = pd.to_datetime(db_base["Data Envio Solicitação"], format='mixed', dayfirst=True)

        # Calculando datas de vencimento para cada linha do dataframe
        db_base["Vencimentos"] = db_base.apply(lambda row: somar_horas_uteis(row["Data Envio Solicitação"], 40), axis=1)

        # Exibindo as primeiras linhas do dataframe para verificar os resultados
        db_base.head()

        # Salvando o resultado no arquivo Excel
        output_dir = os.path.join(user_dir, 'Downloads')
        output_path = os.path.join(output_dir, 'EmAndamento_Atualizado.xlsx')
        db_base.to_excel(output_path, index=False, engine='openpyxl')

        return f'Arquivo salvo em: {output_path}'
    except FileNotFoundError:
        return 'Arquivo Bradesco não encontrado'
    except Exception as e:
        return f'Ocorreu um erro: {e}'

busca_vivaintra_bradesco()