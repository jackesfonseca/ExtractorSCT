import time
from tkinter.simpledialog import askstring
import tkinter.messagebox
import datetime as dt
from selenium import webdriver
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
import win32com.client
import win32gui
import keyboard
from pynput.mouse import Controller
import threading

def blockinput():
    global block_input_flag
    block_input_flag = 1
    t1 = threading.Thread(target=blockinput_start)
    t1.start()
    print("[SUCCESS] Input blocked!")

def unblockinput():
    blockinput_stop()
    print("[SUCCESS] Input unblocked!")

def blockinput_start():
    mouse = Controller()
    global block_input_flag
    for i in range(300):
        keyboard.block_key(i)
    while block_input_flag == 1:
        mouse.position = (0, 0)
    time.sleep(10)
    
def blockinput_stop():
    global block_input_flag
    for i in range(150):
        keyboard.unblock_key(i)
    block_input_flag = 0
def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))

# seta configuração inicial
def setup():
    # setar driver e opções do browser
    driver = webdriver.Firefox(service=FirefoxService('../driver/geckodriver.exe'))
    driver.get('https://sistemas.aneel.gov.br/concessionarios/administracao/')

    driver.maximize_window()

    return driver

# obtém os dados a serem baixados
def get_cycle(root):
    date = dt.date.today()
    year = date.year

    # obtém usuário, senha e dados mensais
    while True:
        user = askstring('Dados', 'Usuário ANEEL', parent=root)
        if user == None:
            return
        elif user == '':
            tkinter.messagebox.showwarning('Atenção!', 'Informe o usuário para continuar!')
        else:
            break
    
    while True:
        password = askstring('Dados', 'Senha ANEEL', show='*', parent=root)
        if password == None:
            return
        elif password == '':
            tkinter.messagebox.showwarning('Atenção!', 'Informe a senha para continuar!')
        else:
            break

    while True:
        start_date = askstring('SAMP', 'Início (MM/AAAA)', parent=root)
        if start_date == None:
            return
        elif len(start_date) == 7 and start_date[2] == '/':
            if len(start_date.split('/')[0]) == 2 and len(start_date.split('/')[1]) == 4 and start_date[2] == '/' and int(start_date.split('/')[0]) <= 12 and int(start_date.split('/')[0]) >= 1 and int(start_date.split('/')[1]) <= year:
                break
        tkinter.messagebox.showwarning('Atenção!', 'Data inválida')

    while True:
        end_date = askstring('SAMP', 'Fim (MM/AAAA)', parent=root)
        if end_date == None:
            return
        elif len(end_date) == 7 and end_date[2] == '/':
            if len(end_date.split('/')[0]) == 2 and len(end_date.split('/')[1]) == 4 and end_date[2] == '/' and int(end_date.split('/')[0]) <= 12 and int(end_date.split('/')[0]) >= 1 and int(end_date.split('/')[1]) <= year:
                break
        tkinter.messagebox.showwarning('Atenção!', 'Data inválida')

    start_month = start_date.split('/')[0]
    start_month_int = int(start_month)
    start_year = start_date.split('/')[1]
    start_year_int = int(start_year)

    end_month = end_date.split('/')[0]
    end_month_int = int(end_month)
    end_year = end_date.split('/')[1]
    end_year_int = int(end_year)

    if start_month_int != end_month_int:
        one_file_ans = tkinter.messagebox.askquestion('Opção de Download', 'Deseja baixar o SAMP em um único arquivo?', icon='warning')
    else:
        if start_year_int != end_year_int:
            one_file_ans = tkinter.messagebox.askquestion('Opção de Download', 'Deseja baixar o SAMP em um único arquivo?', icon='warning')
        else:
            one_file_ans = 'no'

    tkinter.messagebox.showwarning('Atenção!', 'Não utilize o mouse e o teclado até que o primeiro arquivo tenha sido baixado!')

    return user, password, start_month, start_month_int, start_year, start_year_int, end_month, end_month_int, end_year, end_year_int, one_file_ans

# chama função de download
def download(root):
    user, password, start_month, start_month_int, start_year, start_year_int, end_month, end_month_int, end_year, end_year_int, one_file_ans = get_cycle(root)

    driver = setup()

    firstButtonWait = WebDriverWait(driver, 60)
    firstButtonWait.until(EC.element_to_be_clickable(('xpath', '//*[@id="myModal"]/div/div/div[1]/button')))
    driver.find_element('xpath', '//*[@id="myModal"]/div/div/div[1]/button').click()

    acessoSistemasWait = WebDriverWait(driver, 60)
    acessoSistemasWait.until(EC.element_to_be_clickable(('xpath', '//*[@id="accordion"]/div[2]/div[1]/h4/a/span')))
    element = driver.find_element('xpath', '//*[@id="accordion"]/div[2]/div[1]/h4/a/span')
    driver.execute_script("arguments[0].click();", element)

    sampButtonWait = WebDriverWait(driver, 60)
    sampButtonWait.until(EC.element_to_be_clickable(('xpath', '//*[@id="collapseFour"]/div/ul/li[1]/a')))
    driver.find_element('xpath', '//*[@id="collapseFour"]/div/ul/li[1]/a').click()

    receita(driver, user, password, start_month, start_month_int, start_year, start_year_int, end_month, end_month_int, end_year, end_year_int, one_file_ans)

# realiza download da Receita Uso 
def receita(driver, user, password, start_month, start_month_int, start_year, start_year_int, end_month, end_month_int, end_year, end_year_int, one_file_ans):

    securityUser = 'usrRelatorio'
    securityPass = '12345678'

    parent_handle = driver.current_window_handle
    all_handles = driver.window_handles

    finish = False

    for handle in all_handles:
        if handle != parent_handle:
            driver.switch_to.window(handle)

            # selecionar distribuidora
            selectDisWait = WebDriverWait(driver, 60)
            selectDisWait.until(EC.visibility_of_element_located(('xpath', '//*[@id="form1"]/table/tbody/tr[4]/td[2]/table/tbody/tr[1]/td[2]/select')))
            selectDis = driver.find_element('xpath', '//*[@id="form1"]/table/tbody/tr[4]/td[2]/table/tbody/tr[1]/td[2]/select')
            dropdown = Select(selectDis)
            dropdown.select_by_value('169')

            # adicionar usuário
            userInput = driver.find_element('xpath', '//*[@id="form1"]/table/tbody/tr[4]/td[2]/table/tbody/tr[2]/td[2]/input')
            userInput.send_keys(user)

            # adicionar senha
            userInput = driver.find_element('xpath', '//*[@id="form1"]/table/tbody/tr[4]/td[2]/table/tbody/tr[3]/td[2]/input')
            userInput.send_keys(password)

            # clica em avançar
            driver.find_element('xpath', '//*[@id="form1"]/table/tbody/tr[4]/td[2]/table/tbody/tr[4]/td/input').click()

            # seleciona SAMP
            sampWait = WebDriverWait(driver, 60)
            sampWait.until(EC.visibility_of_all_elements_located(('xpath', '//*[@id="form1"]/table/tbody/tr[2]/td[2]/p/a[3]')))
            driver.find_element('xpath', '//*[@id="form1"]/table/tbody/tr[2]/td[2]/p/a[3]').click()

            # seleciona consulta
            consultaWait = WebDriverWait(driver, 60)
            consultaWait.until(EC.visibility_of_all_elements_located(('xpath', '/html/body/table/tbody/tr[1]/td[3]/p[2]/a[1]')))
            driver.find_element('xpath', '/html/body/table/tbody/tr[1]/td[3]/p[2]/a[1]').click()

            # seleciona relatórios consolidados
            selectDisWait = WebDriverWait(driver, 60)
            selectDisWait.until(EC.visibility_of_element_located(('xpath', '//*[@id="lstConsulta"]')))
            selectDis = driver.find_element('xpath', '//*[@id="lstConsulta"]')
            dropdown = Select(selectDis)
            dropdown.select_by_value('../Rel_Consolidados/irt_modalidades.asp?id_empresa=5160')

            # seleciona avançar
            driver.find_element('xpath', '/html/body/form/table/tbody/tr[3]/td/input[1]').click()

            # caso deseja baixar em um único arquivo
            if one_file_ans == 'yes':

                # seleciona Receita de Uso no Transporte de Energia e insere periodo
                recUsoEnergWait = WebDriverWait(driver, 60)
                recUsoEnergWait.until(EC.visibility_of_element_located(('xpath', '//*[@id="lstConsulta"]')))
                recUsoEnerg = driver.find_element('xpath', '//*[@id="lstConsulta"]')
                dropdown = Select(recUsoEnerg)
                dropdown.select_by_value('Receita_selecao.asp?empresa=5160')

                # mês inicial
                userInput = driver.find_element('id', 'mesIni')
                userInput.send_keys(start_month)

                # ano incial
                userInput = driver.find_element('id', 'anoIni')
                userInput.send_keys(start_year)

                # mês final
                userInput = driver.find_element('id', 'mesFim')
                userInput.send_keys(end_month)

                # ano final
                userInput = driver.find_element('id', 'anoFim')
                userInput.send_keys(end_year)

                # selecionar avançar
                driver.find_element('xpath','/html/body/form/table/tbody/tr[4]/td/input[1]').click()

                # selecionar Receita de Uso Relatório Completo
                recUsoWait = WebDriverWait(driver, 60)
                recUsoWait.until(EC.visibility_of_element_located(('xpath', '//*[@id="modalidades"]')))
                recUso = driver.find_element('xpath', '//*[@id="modalidades"]')
                dropdown = Select(recUso)
                dropdown.select_by_value('receita_selecao.asp?flag=2&empresa=5160')

                # selecionar avançar
                driver.find_element('xpath', '//*[@id="Table1"]/tbody/tr[4]/td/input[1]').click()

                # blockinput()

                # insere credenciais windows security (alterar para apenas durante primeiro acesso e modo background)
                samp_handle = driver.current_window_handle

                # joga a tela para frente
                results = []
                top_windows = []
                win32gui.EnumWindows(windowEnumerationHandler, top_windows)

                time.sleep(0.5)
                
                # for i in top_windows:
                #     if "Mozilla Firefox" in i[1]:
                #         win32gui.ShowWindow(i[0],5)
                #         win32gui.SetForegroundWindow(i[0])
                #         time.sleep(0.01)
                #         break

                shell = win32com.client.Dispatch("WScript.Shell")

                time.sleep(0.3)
                shell.Sendkeys(securityUser)
                time.sleep(0.1)
                shell.Sendkeys('{TAB}')
                time.sleep(0.1)
                shell.Sendkeys(securityPass)
                time.sleep(0.1)
                shell.Sendkeys('{ENTER}')

                # unblockinput()
                
                time.sleep(2)

                for h in driver.window_handles:
                    if h != parent_handle and h != samp_handle:
                        driver.switch_to.window(h)

                time.sleep(1.5)

                # clica em exportar
                exportWait = WebDriverWait(driver, 180)
                exportWait.until(EC.visibility_of_all_elements_located(('xpath', '//*[@id="ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink"]')))
                driver.find_element('xpath', '//*[@id="ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink"]').click()

                # seleciona a opção excel
                excelWait = WebDriverWait(driver, 60)
                excelWait.until(EC.visibility_of_all_elements_located(('xpath', '/html/body/form/span[1]/div/table/tbody/tr[4]/td/div/div/div[3]/table/tbody/tr/td/div[2]/div[5]/a')))
                driver.find_element('xpath', '/html/body/form/span[1]/div/table/tbody/tr[4]/td/div/div/div[3]/table/tbody/tr/td/div[2]/div[5]/a').click()
        
                finish = True
            # caso deseja baixar em vários arquivos
            else:
                insert_cred_first = True
                while not finish:
                    # seleciona Receita de Uso no Transporte de Energia e insere periodo
                    recUsoEnergWait = WebDriverWait(driver, 60)
                    recUsoEnergWait.until(EC.visibility_of_element_located(('xpath', '//*[@id="lstConsulta"]')))
                    recUsoEnerg = driver.find_element('xpath', '//*[@id="lstConsulta"]')
                    dropdown = Select(recUsoEnerg)
                    dropdown.select_by_value('Receita_selecao.asp?empresa=5160')

                    # mês inicial
                    userInput = driver.find_element('id', 'mesIni')
                    userInput.send_keys(start_month)

                    # ano incial
                    userInput = driver.find_element('id', 'anoIni')
                    userInput.send_keys(start_year)

                    # mês final
                    userInput = driver.find_element('id', 'mesFim')
                    userInput.send_keys(start_month)

                    # ano final
                    userInput = driver.find_element('id', 'anoFim')
                    userInput.send_keys(start_year)

                    # selecionar avançar
                    driver.find_element('xpath','/html/body/form/table/tbody/tr[4]/td/input[1]').click()

                    # selecionar Receita de Uso Relatório Completo
                    recUsoWait = WebDriverWait(driver, 60)
                    recUsoWait.until(EC.visibility_of_element_located(('xpath', '//*[@id="modalidades"]')))
                    recUso = driver.find_element('xpath', '//*[@id="modalidades"]')
                    dropdown = Select(recUso)
                    dropdown.select_by_value('receita_selecao.asp?flag=2&empresa=5160')

                    # selecionar avançar
                    driver.find_element('xpath', '//*[@id="Table1"]/tbody/tr[4]/td/input[1]').click()

                    # insere credenciais win sec apenas uma vez
                    if insert_cred_first:

                        samp_handle = driver.current_window_handle

                        # blockinput()

                        # joga a tela para frente
                        results = []
                        top_windows = []
                        win32gui.EnumWindows(windowEnumerationHandler, top_windows)

                        time.sleep(0.5)

                        for i in top_windows:
                            if "Mozilla Firefox" in i[1]:
                                win32gui.ShowWindow(i[0],5)
                                win32gui.SetForegroundWindow(i[0])
                                time.sleep(0.01)
                                break

                        # insere credenciais windows security (alterar para apenas durante primeiro acesso e modo background)
                        shell = win32com.client.Dispatch("WScript.Shell")

                        time.sleep(0.3)
                        shell.Sendkeys(securityUser)
                        time.sleep(0.1)
                        shell.Sendkeys('{TAB}')
                        time.sleep(0.1)
                        shell.Sendkeys(securityPass)
                        time.sleep(0.1)
                        shell.Sendkeys('{ENTER}')

                        # unblockinput()

                        insert_cred_first = False

                    time.sleep(2)

                    for h in driver.window_handles:
                        if h != parent_handle and h != samp_handle:
                            driver.switch_to.window(h)

                    time.sleep(1.5)

                    # clica em exportar 
                    exportWait = WebDriverWait(driver, 180)
                    exportWait.until(EC.visibility_of_all_elements_located(('xpath', '//*[@id="ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink"]')))
                    driver.find_element('xpath', '//*[@id="ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink"]').click()

                    # seleciona a opção excel
                    excelWait = WebDriverWait(driver, 60)
                    excelWait.until(EC.visibility_of_all_elements_located(('xpath', '/html/body/form/span[1]/div/table/tbody/tr[4]/td/div/div/div[3]/table/tbody/tr/td/div[2]/div[5]/a')))
                    driver.find_element('xpath', '/html/body/form/span[1]/div/table/tbody/tr[4]/td/div/div/div[3]/table/tbody/tr/td/div[2]/div[5]/a').click()

                    # calcula próximo mês
                    if start_year_int < end_year_int:
                        if start_month_int != end_month_int and start_year_int != end_year_int:
                            if start_month_int < 12:
                                if start_month_int < 10:
                                    start_month_int += 1
                                    # start_month = '0' + str(start_month_int)
                                    start_month = str(start_month_int)
                                else:
                                    start_month_int += 1
                                    start_month = str(start_month_int)
                            else:
                                start_month_int = 1
                                start_year_int += 1
                                # start_month = '0' + str(start_month_int)
                                start_month = str(start_month_int)
                                start_year = str(start_year_int)
                        else:
                            finish = True
                    else:
                        if start_month_int < end_month_int:
                            if start_month_int < 10:
                                start_month_int += 1
                                # start_month = '0' + str(start_month_int)
                                start_month = str(start_month_int)
                            else:
                                start_month_int += 1
                                start_month = str(start_month_int)
                        else:
                            finish = True

                    # altera para tela ***SAMP*** e clica em voltar
                    driver.switch_to.window(samp_handle)

                    excelWait = WebDriverWait(driver, 60)
                    excelWait.until(EC.visibility_of_all_elements_located(('xpath', '/html/body/form/table/tbody/tr[4]/td/input[2]'))) 
                    driver.find_element('xpath', '/html/body/form/table/tbody/tr[4]/td/input[2]').click()

            if finish:
                tkinter.messagebox.showinfo(title="Atenção", message="Arquivos baixados com sucesso!")
                driver.quit()