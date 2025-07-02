"""
Módulo de automação para extração de relatórios do sistema SSW.
Este script automatiza o processo de download de relatórios 455 do sistema SSW,
realizando login, preenchimento de formulários e download dos arquivos.

Requer:
- Arquivo credenciais.env com as credenciais do SSW
- Selenium WebDriver para Microsoft Edge
- Acesso ao sistema SSW
"""

from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import calendar
import pandas as pd
import os
from dotenv import load_dotenv
import time
import locale

# Configuração de localidade para datas em português
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')

# Configuração do diretório de download e carregamento das credenciais
download_folder = os.path.expanduser('I:\\.shortcut-targets-by-id\\1BbEijfOOPBwgJuz8LJhqn9OtOIAaEdeO\\Logdi\\Relatório e Dashboards\\DB_COMUM\\DB_455')
load_dotenv("credenciais.env")

# Configuração das opções do navegador Edge
edge_options = Options()
edge_options.add_experimental_option('prefs', {
    "download.default_directory": download_folder,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safeBrowse.enabled": True
})

def realizar_login(driver, stop_event):
    """
    Realiza o login no sistema SSW usando as credenciais do arquivo .env
    
    Args:
        driver: Instância do WebDriver
        stop_event: Evento para controle de parada da automação
    """
    if stop_event and stop_event.is_set(): return
    driver.get("https://sistema.ssw.inf.br/bin/ssw0422")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f1")))
    driver.find_element(By.NAME, "f1").send_keys(os.getenv("SSW_EMPRESA"))
    driver.find_element(By.NAME, "f2").send_keys(os.getenv("SSW_CNPJ"))
    driver.find_element(By.NAME, "f3").send_keys(os.getenv("SSW_USUARIO"))
    driver.find_element(By.NAME, "f4").send_keys(os.getenv("SSW_SENHA"))
    login_button = driver.find_element(By.ID, "5")
    driver.execute_script("arguments[0].click();", login_button)
    time.sleep(5)

def preencher_formulario(driver, data_inicio, data_fim, stop_event):
    """
    Preenche o formulário de pesquisa do relatório 455
    
    Args:
        driver: Instância do WebDriver
        data_inicio: Data inicial no formato DDMMYY
        data_fim: Data final no formato DDMMYY
        stop_event: Evento para controle de parada da automação
    """
    if stop_event and stop_event.is_set(): return
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f2")))
    driver.find_element(By.NAME, "f3").send_keys("455")
    time.sleep(1)
    abas = driver.window_handles
    driver.switch_to.window(abas[-1])
    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.ID, "11")))
    time.sleep(1)
    driver.find_element(By.ID, "11").clear()
    time.sleep(0.3)
    driver.find_element(By.ID, "11").send_keys(data_inicio)
    time.sleep(3)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "12")))
    campo_data_fim = driver.find_element(By.ID, "12")
    driver.execute_script("arguments[0].value = '';", campo_data_fim)
    time.sleep(0.3)
    campo_data_fim.send_keys(data_fim)
    # ... (o resto da função continua igual)
    time.sleep(1)
    driver.find_element(By.NAME, "f21").clear()
    time.sleep(0.3)
    driver.find_element(By.NAME, "f21").send_keys("t")
    time.sleep(0.3)
    driver.find_element(By.NAME, "f35").clear()
    time.sleep(0.3)
    driver.find_element(By.NAME, "f35").send_keys("e")
    time.sleep(0.3)
    input_element = driver.find_element(By.NAME, "f37")
    driver.execute_script("arguments[0].value = '';", input_element)
    driver.find_element(By.NAME, "f37").send_keys("b")
    time.sleep(0.3)
    driver.find_element(By.NAME, "f38").send_keys("g")
    time.sleep(0.3)
    driver.find_element(By.NAME, "f39").send_keys("h")
    login_button = driver.find_element(By.ID, "40")
    driver.execute_script("arguments[0].click();", login_button)
    time.sleep(0.8)
    actions = ActionChains(driver)
    actions.send_keys("1").perform()
    time.sleep(5)
    abas = driver.window_handles
    driver.switch_to.window(abas[-1])
    time.sleep(2)


def capturar_seq(driver, stop_event):
    """
    Captura o número de sequência da requisição gerada
    
    Args:
        driver: Instância do WebDriver
        stop_event: Evento para controle de parada da automação
    
    Returns:
        str: Número da sequência ou None se não encontrado
    """
    if stop_event and stop_event.is_set(): return None
    try:
        tabela = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "tblsr"))
        )
        linhas = tabela.find_elements(By.TAG_NAME, "tr")
        if len(linhas) > 1:
            seq_da_requisicao = linhas[1].find_element(By.TAG_NAME, "td").text
            print(f"Seq da requisição: {seq_da_requisicao}")
            return seq_da_requisicao
        else:
            print("Não há linhas suficientes na tabela para capturar o seq.")
            return None
    except Exception as e:
        print(f"Erro ao capturar o seq: {e}")
        return None

def atualizar_relatorio(driver, seq_da_requisicao, stop_event):
    """
    Atualiza o relatório e inicia o download quando disponível
    
    Args:
        driver: Instância do WebDriver
        seq_da_requisicao: Número da sequência do relatório
        stop_event: Evento para controle de parada da automação
    
    Returns:
        bool: True se o download foi iniciado com sucesso, False caso contrário
    """
    if stop_event and stop_event.is_set(): return False
    try:
        # Loop para esperar de forma cancelável
        for _ in range(150):
            if stop_event and stop_event.is_set(): return False
            time.sleep(1)
            
        update_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "2"))
        )
        driver.execute_script("arguments[0].click();", update_button)
        time.sleep(5)
    except Exception as e:
        print(f"Botão de atualização não encontrado ou não clicável: {e}")
        return False
    
    # ... (o resto da função continua igual)
    relatorios_atualizados = driver.find_elements(By.CSS_SELECTOR, "table#tblsr tr")
    relatorio_encontrado = None
    for relatorio in relatorios_atualizados[1:]:
        seq_atual = relatorio.find_element(By.TAG_NAME, "td").text
        if seq_atual == seq_da_requisicao:
            relatorio_encontrado = relatorio
            break
    if relatorio_encontrado:
        try:
            link = relatorio_encontrado.find_element(By.TAG_NAME, "u")
            driver.execute_script("arguments[0].click();", link)
            print("Clicou no link da requisição correspondente para fazer o download.")
            return True
        except Exception as e:
            print(f"Não foi possível encontrar ou clicar no link de download: {e}")
            return False
    else:
        print("Nenhum relatório correspondente ao seq encontrado após a atualização.")
        return False

def renomear_ultimo_arquivo_baixado(pasta_download, nome_base_novo):
    """
    Renomeia o último arquivo baixado na pasta de download
    
    Args:
        pasta_download: Caminho da pasta de download
        nome_base_novo: Novo nome base para o arquivo (sem extensão)
    """
    try:
        arquivos = [
            os.path.join(pasta_download, f)
            for f in os.listdir(pasta_download)
            if f.lower() != "desktop.ini"
        ]
        if not arquivos:
            print("Nenhum arquivo encontrado na pasta de download.")
            return
        arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
        print(f"Arquivo mais recente encontrado: {os.path.basename(arquivo_mais_recente)}")
        _, extensao = os.path.splitext(arquivo_mais_recente)
        novo_nome_completo = os.path.join(pasta_download, nome_base_novo + extensao)
        if os.path.exists(novo_nome_completo):
            print(f"Arquivo '{os.path.basename(novo_nome_completo)}' já existe. Removendo o antigo.")
            os.remove(novo_nome_completo)
        print(f"Renomeando '{os.path.basename(arquivo_mais_recente)}' para '{os.path.basename(novo_nome_completo)}'")
        os.rename(arquivo_mais_recente, novo_nome_completo)
        print("Arquivo renomeado com sucesso.")
    except Exception as e:
        print(f"Ocorreu um erro ao gerenciar o arquivo: {e}")


def main(stop_event=None):
    """
    Função principal que coordena todo o processo de extração
    Extrai relatórios dos últimos 3 meses
    
    Args:
        stop_event: Evento opcional para controle de parada da automação
    """
    hoje = datetime.now()
    for i in range(3):
        if stop_event and stop_event.is_set():
            print("Sinal de parada recebido. Interrompendo a extração.")
            return

        driver = None
        try:
            primeiro_dia_mes_atual = hoje.replace(day=1)
            ano = primeiro_dia_mes_atual.year
            mes = primeiro_dia_mes_atual.month - i
            if mes < 1:
                mes = 12 + mes
                ano -= 1
            primeiro_dia = datetime(ano, mes, 1)
            ultimo_dia_numero = calendar.monthrange(ano, mes)[1]
            ultimo_dia = datetime(ano, mes, ultimo_dia_numero)
            data_inicio_str = primeiro_dia.strftime('%d%m%y')
            data_fim_str = ultimo_dia.strftime('%d%m%y')
            nome_arquivo_mes = primeiro_dia.strftime('%b').upper() + str(ano)
            print(f"\n--- Iniciando extração para o período: {data_inicio_str} a {data_fim_str} ---")
            
            driver = webdriver.Edge(options=edge_options)
            realizar_login(driver, stop_event)
            if stop_event and stop_event.is_set(): raise InterruptedError
            
            preencher_formulario(driver, data_inicio_str, data_fim_str, stop_event)
            if stop_event and stop_event.is_set(): raise InterruptedError

            seq = capturar_seq(driver, stop_event)
            if stop_event and stop_event.is_set(): raise InterruptedError

            if seq:
                if atualizar_relatorio(driver, seq, stop_event):
                    print("Aguardando 20 segundos para a conclusão do download...")
                    time.sleep(20) 
                    renomear_ultimo_arquivo_baixado(download_folder, nome_arquivo_mes)
            print(f"--- Finalizada extração para o período: {data_inicio_str} a {data_fim_str} ---")
        
        except InterruptedError:
            print("Execução interrompida pelo usuário.")
            break # Sai do loop for
        except Exception as e:
            print(f"Ocorreu um erro geral na automação para o mês {mes}/{ano}: {e}")
        finally:
            if driver:
                print("Encerrando a sessão do navegador para esta iteração.")
                driver.quit()

if __name__ == "__main__":
    main()