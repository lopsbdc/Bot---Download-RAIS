from playwright.sync_api import sync_playwright

#Pandas - módulo para mexer com BD em excel
import pandas as pd

#módulo de geração de log
import logging

#módulo de controle de teclado e mouse
import pyautogui as pyg

#módulo de tempo de espera
import time

import cred
#**************************** Inicio ****************************

#iniciando o navegador
with sync_playwright() as p:
    navegador = p.chromium.launch(headless=False, slow_mo=100)  #Headless = não fazer no modo invisvel, para reconhecimento de imagens.

    #lendo o banco de dados em excel
    tabela = pd.read_excel('CNPJ.xlsx') # planilha com dados de CNPJ


    #configs finais do navegador, e entrar no site do governo
    pagina = navegador.new_page(accept_downloads=True, ignore_https_errors=True) #ignorar erro de certificado do governo
    pyg.write('https://rais.gov.br/sitio/private/lista_declaracoes.jsf')
    pyg.press('enter')

    procurar = "sim"

    #entrar em loop para achar onde clicar
    while procurar == "sim":

        try:
            #primeiro o botao de ok
            img = pyg.locateCenterOnScreen('ok.png', confidence=0.8)
            pyg.click(img.x, img.y)
            procurar = "não"

        except:
            time.sleep(1)
            print("Não localizado")

    procurar1 = "sim"

    while procurar1 == "sim":
        try:
            #agora o botao do certificado
            img = pyg.locateCenterOnScreen('2.png', confidence=0.9)
            pyg.click(img.x, img.y)
            procurar1 = "não"

        except:
            time.sleep(1)
            print("Não localizado")

    procurar2 = "sim"

    while procurar2 == "sim":
        try:
            # agora o botao da senha
            img = pyg.locateCenterOnScreen('Senha.png', confidence=0.8)
            pyg.click(img.x, img.y)
            pyg.write(cred.senha)
            pyg.press('enter')
            procurar2 = "não"

        except:
            time.sleep(1)
            print("Não localizado")

    #tempo de espera para o site do governo sumir o erro
    pyg.click()
    time.sleep(5)
    pyg.press('f5')
    pyg.press('f5')
    pyg.press('f5')
    time.sleep(2)
    pyg.press('f5')
    pyg.press('f5')
    pyg.press('f5')

    #log config básico
    logging.basicConfig(filename='RAIS.log', filemode='a', format='%(levelname)s - %(message)s')

    #ao entrar, temos que inciar a consulta, e forçar uma primeira busca/erro (senão o site apresenta bug)
    pagina.locator('xpath=//*[@id="form:Nova"]').click()
    pagina.locator('xpath=//*[@id="form:tipo-identificacao:0"]').click()
    pagina.locator('xpath=//*[@id="form:localizar"]').click()
    time.sleep(2)
    pagina.locator('xpath=//*[@id="form:localizar"]').click()
    pagina.locator('xpath=//*[@id="form:Nova"]').click()

    #iniciar a repetição, de acordo com o tamanho da tabela (usando enumerate)
    for i, filial in enumerate(tabela['Filial']):

        try:

            #identificando onde ta a informação no excel, e o i, indica a linha. Conversão em String para evitar erro
            filial = str(tabela.loc[i, 'Filial'])
            cnpj = str(tabela.loc[i, 'Final'])

            #O pandas não importa os zeros a esquerda do excel, nem mesmo por texto. Então, zfill preenche com zeros
            cnpj1 = cnpj.zfill(6)

            #Etapa de preencher linhas, e indicando onde preencher na página

            #CNPJ
            pagina.locator('xpath=//*[@id="form:cnpj3"]').click()
            pagina.locator('xpath=//*[@id="form:cnpj3"]').fill(cnpj1)

            vazio = ''
            #zerar o CREA pra evitar erro
            pagina.locator('xpath=//*[@id="form:crea"]').click()
            time.sleep(1)
            pagina.locator('xpath=//*[@id="form:crea"]').fill(vazio)

            #pesquisar
            pagina.locator('xpath=//*[@id="form:localizar"]').click()

            #seleção do checkbox e download
            pagina.locator('xpath=//*[@id="form:items:0:selecionar"]').click()

            #Aguardar e renomear downloads
            with pagina.expect_download() as download_info:
                pagina.locator('xpath=//*[@id="form:imprimir"]').click()
            download = download_info.value
            # selecionando o caminho R é de Raw, pra ele entender que é um caminho. \\ em todos tambem funciona
            download.save_as(r'\PDFs\ ' + filial + ".pdf")

            #log de print screen
            screen = pyg.screenshot()
            screen.save(r'Prints\'' + filial + 'filial.png')

            #iniciando nova consulta
            pagina.locator('xpath=//*[@id="form:Nova"]').click()

            #salvando no log em texto, de acordo com o numero da filial na planilha (usado como identificador)
            logging.warning('RAIS da filial ' + filial + " foi feito com sucesso!")

        except:
            #salvando erro no log, de acordo com o numero da filial na planilha (usado como identificador)
            logging.warning('RAIS da filial ' + filial + " não foi feito, isto é um erro!")


    #salvando no log,informando que finalizou o bot
    logging.warning('Bot finalizado')
