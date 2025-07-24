from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import UnexpectedAlertPresentException, NoAlertPresentException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from datetime import datetime
from shutil import copyfile
import pyautogui
import re
import os
import pandas as pd
import gcpj_utils
import time
 
def main_bacenjud(): 

    index_path = rf"\\192.168.100.10\dados\Pesquisa e Desenvolvimento\Development\CODIGOS_VM\GCPJ-SELENIUM\GCPJ_ADM\index-bacenjud.txt"
    with open(index_path, "r") as index:
        index_read = int(index.read().strip() or 0)

    print(f"Iniciando do index: {index_read}")

    ie_options = webdriver.IeOptions()
    ie_options.attach_to_edge_chrome = True
    ie_options.ignore_zoom_level = True
    ie_options.headless = True
    driver = webdriver.Ie(options=ie_options)
    driver.get("https://juridico8.bradesco.com.br/gcpj/")
    
    gcpj_utils.gcpj_access(driver)
    
    df = pd.read_excel(
        r'\\192.168.100.10\dados\Pesquisa e Desenvolvimento\Development\BASES\ADM\requerimento-bacenjud\INPUT\AUDITORIA - DR.KARINA - ULTIMA RC.xlsx',
        usecols=['GCPJ'],
        engine='openpyxl'
    )
    df_acordo_parcelado = df
    resultados = []
    
    def mover_renomear_pdf(numero_gcpj, destino_final):
        downloads_padrao = os.path.join(os.environ['USERPROFILE'], 'Downloads')
        time.sleep(5)
    
        arquivos_pdf = [f for f in os.listdir(downloads_padrao) if f.lower().endswith('.pdf')]
        arquivos_pdf.sort(key=lambda f: os.path.getmtime(os.path.join(downloads_padrao, f)), reverse=True)
    
        if not arquivos_pdf:
            print("Nenhum PDF encontrado na pasta de downloads.")
            return None
    
        caminho_original = os.path.join(downloads_padrao, arquivos_pdf[0])
        novo_nome = f"{numero_gcpj}.pdf"
        caminho_renomeado = os.path.join(destino_final, novo_nome)
    
        os.makedirs(destino_final, exist_ok=True)
    
        if os.path.exists(caminho_renomeado):
            try:
                os.remove(caminho_renomeado)
                print(f"Arquivo existente removido: {caminho_renomeado}")
            except Exception as e:
                print(f"Erro ao remover arquivo existente: {e}")
    
        try:
            copyfile(caminho_original, caminho_renomeado)
            print(f"PDF movido e renomeado para: {caminho_renomeado}")
            return caminho_renomeado
        except Exception as e:
            print(f"Erro ao mover PDF: {e}")
            return None
    
    def verificar_evidencias_em_paginas(driver, numero_gcpj):
        evidencias_encontradas = False
        ultima_rc, data_inclusao = None, None
    
        WebDriverWait(driver, 3).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, '[alt="última página"]'))
        )
        driver.find_element(By.CSS_SELECTOR, '[alt="última página"]').click()
        time.sleep(1)
    
        for tentativa in range(3):
            try:
                rows = driver.find_elements(By.XPATH, "//table[@id='oTable']//tr")
                for row in rows:
                    columns = row.find_elements(By.TAG_NAME, "td")
                    if len(columns) < 7:
                        continue
    
                    descricao = columns[3].text.strip().upper()
                    data_texto = columns[6].text.strip()
    
                    if descricao == "RC-BACENJUD REQ.":
                        data_dt = datetime.strptime(data_texto, "%d/%m/%Y")
                        hoje = datetime.today()
                        ultima_rc = descricao
                        data_inclusao = data_texto
    
                        if data_dt.month == hoje.month and data_dt.year == hoje.year:
                            print(f"RC-BACENJUD REQ. deste mês encontrado na tentativa {tentativa + 1}")
                            evidencias_encontradas = True
    
                            checkbox = row.find_element(By.XPATH, ".//input[@name='selecionados' and @type='checkbox']")
                            checkbox.click()
                            print(f"Checkbox da linha clicado com sucesso")
    
                            btn_vizualizar = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, "//button[@id='visualizar' and @type='button']"))
                            )
                            btn_vizualizar.click()
                            time.sleep(3)
                            pyautogui.hotkey("alt", "s")
                            time.sleep(3)
    
                            downloads_dir = r"\\192.168.100.10\dados\Pesquisa e Desenvolvimento\Development\BASES\BACENJUD\PDF"
                            caminho_renomeado = mover_renomear_pdf(numero_gcpj, downloads_dir)
                        
                            if caminho_renomeado is None:
                                print("Erro no processo de salvar.")
                                return evidencias_encontradas, ultima_rc, data_inclusao
    
                            dropdown = driver.find_element(By.NAME, "cmdAnexos")
                            gcpj_utils.select_in_dropdown_by_javascript(driver, dropdown, "195")
                            time.sleep(2)
    
                            campo = driver.find_element(By.XPATH, "//input[@type='file' and @name='formFile']")
                            driver.execute_script("arguments[0].style.display = 'block';", campo)
                            escrever_gcpj = driver.find_element(By.XPATH, "//input[@type='text' and @name='nmAnexoProcesso']")
                            driver.execute_script("arguments[0].style.display = 'block';", escrever_gcpj)
                            if escrever_gcpj.is_enabled() and escrever_gcpj.get_attribute("readonly") is None:
                                escrever_gcpj.send_keys('BACENJUD -'f"{numero_gcpj}")
                            if campo.is_enabled() and campo.get_attribute("readonly") is None:
                                campo.send_keys(caminho_renomeado)
                                print("Upload realizado com campo forçado visível.")
                            else:
                                print(f"Campo de upload desabilitado ou readonly para GCPJ {numero_gcpj}. Pulando...")
                                return evidencias_encontradas, ultima_rc, data_inclusao
                        
                            botao_incluir = WebDriverWait(driver, 10).until(
                                EC.presence_of_all_elements_located((By.XPATH, "//input[@type='button' and @name='btoIncluir']"))
                            )
                            driver.execute_script("arguments[0].click();", botao_incluir[0])
                        
                            btn_voltar = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and @name='btoVoltar']"))
                            )
                            btn_voltar.click()
    
                if evidencias_encontradas:
                    break    
                
                try:
                    driver.find_element(By.CSS_SELECTOR, '[alt="página anterior"]').click()
                    print(f"Tentativa {tentativa + 1}: RC-BACENJUD REQ. não encontrado. Voltando página...")
                    time.sleep(1)
                except:
                    print(f"Tentativa {tentativa + 1}: Elemento para voltar página não encontrado. Tentando próximo GCPJ...")
                    return evidencias_encontradas, ultima_rc, data_inclusao
            
            except Exception as e:
                print("Erro ao verificar páginas:", e)
            break
    
        return evidencias_encontradas, ultima_rc, data_inclusao
    
    
    def busca_processos(nmr_gcpj):
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "(//a[@class='lnk1'])[3]"))
        ).click()
        print("RYAN DIAS ")
        print(f"Buscando GCPJ: {nmr_gcpj}")
    
        campo_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "cdNumeroProcessoBradesco"))
        )
        btn_pesquisar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.NAME, "btoPesquisar"))
        )
        campo_input.clear()
        campo_input.send_keys(str(nmr_gcpj))
        btn_pesquisar.click()
    
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
    
        janela_original = driver.current_window_handle
        montar_tela = driver.find_element(By.CLASS_NAME, "lnk1")
        montar_tela.click()
    
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
    
        for handle_janela in driver.window_handles:
            if handle_janela != janela_original:
                driver.switch_to.window(handle_janela)
                break
        driver.refresh()
        elemento_resumo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@value='resumo' and @type='button']"))
        )
        elemento_resumo.click()
        
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "here"))
        )

        link_geral = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@class='lnk1' and contains(text(), 'Geral')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", link_geral)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", link_geral)
    
        time.sleep(2)
        trecho = ""
        time.sleep(5)
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        texto = soup.get_text(separator='\n', strip=True)
        time.sleep(3)
        padrao = r"Referência:\s*RC-BACENJUD.*?REQUERIMENTO\s*(.*?)\s*Data:"
        matches = re.findall(padrao, texto, re.DOTALL | re.IGNORECASE)
        time.sleep(5)
        
        trecho = matches[-1].strip() if matches else "Trecho abaixo da referência não encontrado"
        print("Trecho encontrado:")
        print(trecho)        
    
        voltar_btn = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//input[@type='button' and @value='voltar']"))
        )
        driver.execute_script("arguments[0].click();", voltar_btn[0])
    
        driver.switch_to.default_content()
        driver.close()
        driver.switch_to.window(janela_original)
        driver.switch_to.frame("frameCorpoINET")
    
        dropdown_salvo = driver.find_element(By.NAME, "cdReferenciaAndamentoProcesso")
        gcpj_utils.select_in_dropdown_by_javascript(driver, dropdown_salvo, "00363")
        time.sleep(2)
    
        andamento_processual = driver.find_element(By.NAME, "dsAndamentoProcessoEscritorio")
        andamento_processual.send_keys(f"{trecho}")
    
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@value='anexos' and @type='button']"))
        ).click()
    
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "here"))
        )
    
        evidencias, ultima_rc, data_inclusao = verificar_evidencias_em_paginas(driver, nmr_gcpj)
    
        resultados.append({
            'GCPJ': nmr_gcpj,
            'Trecho': trecho,
            'Última RC': ultima_rc,
            'Data de inclusão': data_inclusao if evidencias else "Não é deste mês"
        })
    
        driver.switch_to.default_content()
        driver.switch_to.frame("frameCorpoINET")
    
        if evidencias:
            print("Resultado salvo.")
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '[X]')]"))
            ).click()
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.NAME, "btoVoltar"))
            ).click()
            print("teria salvo !!!!")
        else:
            print("Nenhuma EVIDÊNCIA deste mês encontrada.")
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '[X]')]"))
            ).click()
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.NAME, "btoVoltar"))
            ).click()
            time.sleep(1)
            print("não teria salvo!!!")
    
        driver.switch_to.default_content()
        driver.switch_to.frame(1)
    
        menu_principal = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Menu Principal')]"))
        )
        driver.execute_script("arguments[0].click();", menu_principal)
    
        driver.switch_to.default_content()
        driver.switch_to.frame("frameCorpoINET")
    
        return resultados, busca_processos
    
    
    
    with open(index_path, "w+") as r:
        try:
            index_content = index_read
            for index, row in df_acordo_parcelado.iterrows():
                numero = row['GCPJ']
                sucesso = busca_processos(numero)
                if not sucesso:
                    index_content += 1
                    print(f" Falhou ao buscar o processo {numero}")
                else:
                    index_content += 1
                    print(f" Processo {numero} processado com sucesso.\n")
    
            driver.quit(
            )
        
        finally:
            r.seek(0)
            r.truncate(0)
            r.write(f"{index_content}")

            df_resultados = pd.DataFrame(resultados)
            df_resultados.to_excel(
                r'\\192.168.100.10\dados\Pesquisa e Desenvolvimento\Development\BASES\BACENJUD\OUTPUT\resultados_acordo_parcelado_teste_25-1.07.xlsx',
                index=False,
                engine='openpyxl'
            )
            print(" Arquivo salvo com sucesso.")
            
main_bacenjud()        