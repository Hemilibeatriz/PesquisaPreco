'''
    Instalado pelo Pycharm:
        install package selenium
        install package pandas
    Instalado pelo Terminal:
        pip install openpyxl
        pip install pywin32
'''
#Bibliotecas
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
from selenium.webdriver.common.by import By
import time
import win32com.client as win32

#criar navegador
navegador = webdriver.Chrome('chromedriver')

#importar a base de dados que contém os produtos que se quer fazer as buscas
tabela_produtos = pd.read_excel("buscas.xlsx")

#Funções
def to_lower(palavra):
    palavra = palavra.lower()
    return palavra

def to_split(lista_palavras):
    lista_palavras = lista_palavras.split(' ')
    return lista_palavras

def to_float(numero):
    numero = float(numero)
    return numero

def busca_google_shopping(navegador, produto, termos_banidos, preco_minimo, preco_maximo):
    #entrar no google
    navegador.get("https://www.google.com/")

    produto = to_lower(produto)
    lista_produto = to_split(produto)
    termos_banidos = to_lower(termos_banidos)
    lista_termos_banidos = to_split(termos_banidos)
    preco_minimo = to_float(preco_minimo)
    preco_maximo = to_float(preco_maximo)

    #pesquisar o nome de cada produto que está na tabela
    navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(produto)
    navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

    #clicar na aba shopping
    lista_elementos = navegador.find_elements(By.CLASS_NAME, 'hdtb-mitem')
    for item in lista_elementos:
        if "Shopping" in item.text:
            item.click()
            break

    #lista o resultado dos produtos
    lista_resultados = navegador.find_elements(By.CLASS_NAME, 'sh-dgr__grid-result')

    lista_ofertas = []
    for resultado in lista_resultados:

        nome = resultado.find_element(By.CLASS_NAME, 'Xjkr3b').text
        nome = nome.lower()

        #verifica nome
        tem_termos_banidos = False
        for palavra in lista_termos_banidos:
            if palavra in nome:
                tem_termos_banidos =  True

        tem_todos_termos_produto = True
        for palavra in lista_produto:
            if palavra not in nome:
                tem_todos_termos_produto = False

        #se passa das regras dos nomes, começa o tratamento dos preços
        if not tem_termos_banidos and tem_todos_termos_produto:
            try:
                preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
                preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                preco = float(preco)

                #verifica se o preço está na faixa de preços
                if preco >= preco_minimo and preco <= preco_maximo:
                #se passa dessas duas verificações, o link do produto nos interessa
                    elemento_link = resultado.find_element(By.CLASS_NAME, 'aULzUe')
                    elemento_pai = elemento_link.find_element(By.XPATH, '..')
                    link = elemento_pai.get_attribute('href')
                    lista_ofertas.append((nome, preco, link))
            except:
                continue

    return lista_ofertas

def busca_buscape(navegador, produto, termos_banidos, preco_minimo, preco_maximo):
    # entrar no google
    navegador.get("https://www.buscape.com.br/")

    produto = to_lower(produto)
    lista_produto = to_split(produto)
    termos_banidos = to_lower(termos_banidos)
    lista_termos_banidos = to_split(termos_banidos)
    preco_minimo = to_float(preco_minimo)
    preco_maximo = to_float(preco_maximo)

    #pesquisar o nome de cada produto que está na tabela
    navegador.find_element(By.CLASS_NAME, 'search-bar__text-box').send_keys(produto, Keys.ENTER)

    time.sleep(5)
    lista_resultados = navegador.find_elements(By.CLASS_NAME, 'Cell_Content__1630r')

    lista_ofertas = []
    #para cada resultado
    for resultado in lista_resultados:
        try:
            preco = resultado.find_element(By.CLASS_NAME, 'CellPrice_MainValue__3s0iP').text
            nome = resultado.get_attribute('title')
            nome = nome.lower()
            link = resultado.get_attribute('href')

            # verifica nome
            tem_termos_banidos = False
            for palavra in lista_termos_banidos:
                if palavra in nome:
                    tem_termos_banidos = True

            tem_todos_termos_produto = True
            for palavra in lista_produto:
                if palavra not in nome:
                    tem_todos_termos_produto = False

            # se passa das regras dos nomes, começa o tratamento dos preços
            if not tem_termos_banidos and tem_todos_termos_produto:
                preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                preco = float(preco)
                # verifica se o preço está na faixa de preços
                if preco >= preco_minimo and preco <= preco_maximo:
                    # se passa dessas duas verificações, o link do produto nos interessa
                    lista_ofertas.append((nome, preco, link))
        except:
            pass
    return lista_ofertas

def busca_zoom(navegador, produto, termos_banidos, preco_minimo, preco_maximo):
    # entrar no google
    navegador.get("https://www.zoom.com.br/")

    produto = to_lower(produto)
    lista_produto = to_split(produto)
    termos_banidos = to_lower(termos_banidos)
    lista_termos_banidos = to_split(termos_banidos)
    preco_minimo = to_float(preco_minimo)
    preco_maximo = to_float(preco_maximo)

    #pesquisar o nome de cada produto que está na tabela
    navegador.find_element(By.CLASS_NAME, 'search-bar__text-box').send_keys(produto, Keys.ENTER)

    time.sleep(5)
    lista_resultados = navegador.find_elements(By.CLASS_NAME, 'Cell_Content__1630r')

    lista_ofertas = []
    #para cada resultado
    for resultado in lista_resultados:
        try:
            preco = resultado.find_element(By.CLASS_NAME, 'CellPrice_MainValue__3s0iP').text
            nome = resultado.get_attribute('title')
            nome = nome.lower()
            link = resultado.get_attribute('href')

            # verifica nome
            tem_termos_banidos = False
            for palavra in lista_termos_banidos:
                if palavra in nome:
                    tem_termos_banidos = True

            tem_todos_termos_produto = True
            for palavra in lista_produto:
                if palavra not in nome:
                    tem_todos_termos_produto = False

            # se passa das regras dos nomes, começa o tratamento dos preços
            if not tem_termos_banidos and tem_todos_termos_produto:
                preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                preco = float(preco)
                # verifica se o preço está na faixa de preços
                if preco >= preco_minimo and preco <= preco_maximo:
                    # se passa dessas duas verificações, o link do produto nos interessa
                    lista_ofertas.append((nome, preco, link))
        except:
            pass
    return lista_ofertas

tabela_ofertas = pd.DataFrame()
for linha in tabela_produtos.index:
    produto = tabela_produtos.loc[linha, "Produto"]
    termos_banidos = tabela_produtos.loc[linha, "Termos Banidos"]
    preco_minimo = tabela_produtos.loc[linha, "Preço Mínimo"]
    preco_maximo = tabela_produtos.loc[linha, "Preço Máximo"]
    buscagoogle = busca_google_shopping(navegador, produto, termos_banidos, preco_minimo, preco_maximo)
    if buscagoogle:
        tabela_buscagoogle = pd.DataFrame(buscagoogle, columns=["Produto", "Preço", "Link"])
        tabela_ofertas = tabela_ofertas.append(tabela_buscagoogle)
    else:
        tabela_buscagoogle = None
    buscabuscape = busca_buscape(navegador, produto, termos_banidos, preco_minimo, preco_maximo)
    if buscabuscape:
        tabela_buscabuscape = pd.DataFrame(buscabuscape, columns=["Produto", "Preço", "Link"])
        tabela_ofertas = tabela_ofertas.append(tabela_buscabuscape)
    else:
        tabela_buscabuscape = None
    buscazoom = busca_zoom(navegador, produto, termos_banidos, preco_minimo, preco_maximo)
    if buscazoom:
        tabela_buscazoom = pd.DataFrame(buscazoom, columns=["Produto", "Preço", "Link"])
        tabela_ofertas = tabela_ofertas.append(tabela_buscazoom)
    else:
        tabela_buscabuscape = None

#Exportar para uma base de dados Excel, os resultados satisfatórios
tabela_ofertas = tabela_ofertas.reset_index(drop=True)
tabela_ofertas.to_excel("Ofertas.xlsx", index=False)

#Enviar e-mail se existir ofertas
if len(tabela_ofertas.index) > 0:
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'hemilibeatriz@gmail.com'
    mail.Subject = 'Produtos Encontrados na Faixa de Preços Desejada'
    mail.HTMLBody = f'''
    <p>Queridona Prepara o Cartão que tem produto na Promoção (UHUUL)</p>
    <p>Segue a tabela com as informações certinhas!</p>
    {tabela_ofertas.to_html(index=False)}
    <p>Beijos da Programadora,</p>
    <p>Mi</p>
    '''
    mail.Send()

navegador.quit()



