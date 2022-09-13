from tkinter import *
from bibliotecas import *

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

def acessarQB(email, senha):
    #Abrindo o site
    navegador.get('######')
    #inserindo o email 
    navegador.find_element('xpath', '######').send_keys(email)
    navegador.find_element('xpath', '######').click()
    #inserindo a senha
    navegador.find_element('xpath', '######').send_keys(senha)
    navegador.find_element('xpath', '######').click()
def acessoEmpresa():
    #acessando todas empresas
    navegador.find_element('xpath', '######').click()
    #acessando sempre verde irrigação
    navegador.find_element('xpath', "######").click()
data_inicial_fatura = "14/06/2022"
data_final_fatura = data_inicial_fatura #para teste
def acessarVendas():
    #Acessando vendas
    navegador.find_element('######').click()
    #abre o filtro
    try:
        navegador.find_element('xpath', '######').click()
    except:
        navegador.find_element('xpath', '######').click()
        print("errorrrrr")
    #limpa a data inicial
    navegador.find_element('xpath', '######').send_keys(Keys.BACKSPACE)
    navegador.find_element('xpath', '######').send_keys(Keys.BACKSPACE)
    navegador.find_element('xpath', '######').send_keys(Keys.BACKSPACE)
    navegador.find_element('xpath', '######').send_keys(Keys.BACKSPACE)
    navegador.find_element('xpath', '######').send_keys(Keys.BACKSPACE)
    navegador.find_element('xpath', '######').send_keys(Keys.BACKSPACE)
    navegador.find_element('xpath', '######').send_keys(Keys.BACKSPACE)
    navegador.find_element('xpath', '######').send_keys(Keys.BACKSPACE)
    navegador.find_element('xpath', '######').send_keys(Keys.BACKSPACE)
    navegador.find_element('xpath', '######').send_keys(Keys.BACKSPACE)
    navegador.find_element('xpath', '######').send_keys(data_inicial_fatura)

    #insere a data final
    navegador.find_element('xpath', '######').click()
    navegador.find_element('xpath', '######').send_keys(data_final_fatura)

    #botao buscar
    navegador.find_element('xpath', '######').click()
nomeDoCliente= "######"
def encontrarCliente():
    tag_td = navegador.find_elements(By.TAG_NAME, 'td') #busca pela tag demonstrou melhor resultado
    contador = 0
    nome_cliente = nomeDoCliente
    for t in tag_td:
        if tag_td[contador].text == nome_cliente: #nome do cliente sera uma variavel
            tag_td[contador].click()
        contador = contador + 1
numFatura = "1069"
def numeroDaFatura():
    tag_a = navegador.find_elements(By.TAG_NAME, 'a') #busca pela tag demonstrou melhor resultado
    contador = 0
    numero_fatura = numFatura
    for t in tag_a:
        if tag_a[contador].text == f'Fatura # {numero_fatura}':
            tag_a[contador].click()
            break
        contador = contador + 1
def fazendoDownload():
    
    try:                                 #imprimir ou visualizar                      
        navegador.find_element('xpath', '######').click()
    except:
        try:
            navegador.find_element('xpath', '######').click()
        except:
            try:
                botdownload = navegador.find_elements(By.CLASS_NAME, "bottomCenterButton") #busca plea classe demonstrou melhor resultadp
                if botdownload.text == "Imprimir ou Visualizar":
                    botdownload.click()
            except:
                print("Erro no botão 'imprimir ou visualizar")
                None
    try:
        botdown = navegador.find_element('xpath', '######')
        if botdown.text == "Download":
            botdown.click()
        else:
            print("botão download não encontrado")
    except:
        navegador.find_element('xpath', '######').click()

def encontrandoPDF():
    usuario = getpass.getuser() #lozalizando o nome do pc para o caminho do arquivo não dar erro
    caminho = "C:\\Users\\" + usuario +"\\Downloads" 
    testando_caminho = os.path.exists(caminho)
    arquivos_pasta = os.listdir(caminho)
    numero_fatura = numFatura
    for arquivos in arquivos_pasta:
        if arquivos == f'Fatura {numero_fatura}.pdf':
            caminho_pdf = caminho + "\\" + f'Fatura {numero_fatura}.pdf'
    return caminho_pdf
def modificandoPlanilha(caminho_pdf): #organizando a planilha 
    lista_pdf = tabula.read_pdf(caminho_pdf, pages="all")
    tabela_vendas = pd.DataFrame(lista_pdf[0])
    tabela_vendas.to_excel(f"Fatura {numFatura}.xlsx")
    planilha = load_workbook(f"Fatura {numFatura}.xlsx")
    aba = planilha.active
    aba.delete_rows(1)
    aba["B1"]= "Produtos"
    planilha.save(f"Fatura {numFatura} modificada.xlsx") #salvando a planilha


 
def main():
    try:
        acessarQB("######", "######")
    except:
        print("Erro: Acesso QB primeira tentativa")
        try:
            acessarQB("######", "######")
        except:
            print("Erro: Acesso QB na segunda tentativa")
            return
    sleep(10)
    try:
        acessoEmpresa()
    except:
        print("Erro: Localizar empresa")
        try:
            acessoEmpresa()
        except:
            print("Erro: Localizar empresa na segunda tentativa")
            return
    try:
        acessarVendas()
    except:
        print("Erro: Acessar vendas")
        try:
            acessarVendas()
        except:
            print("Erro: Acessar vendas na segunda tentativa")
            return
    try:
        encontrarCliente()
    except:
        print("Erro: Encontrar cliente")
        try:
            encontrarCliente()
        except:
            print("Erro: Encontrar cliente na segunda tentativa")
            return
    try:
        numeroDaFatura()
    except:
        print("Erro: Encontrar fatura")
        try:
            numeroDaFatura()
        except:
            print("Erro: Encontrar fatura na segunda tentativa")
            return
    try:
        fazendoDownload()
    except:
        print("Erro: Fazer o download da fatura")
        try:
            fazendoDownload()
        except:
            print("Erro: Fazer o download da fatura na segunda tentativa")
            return
    try:    
        caminhoPDF = encontrandoPDF()
    except:
        print("Erro: Localizar PDF")
        try:   
            caminhoPDF = encontrandoPDF()
        except:
            print("Erro: Localizar PDF na segunda tentativa")
            return
    try:
        modificandoPlanilha(caminhoPDF)
    except:
        print("Erro: Alterando dados da planilha")
        try:
            modificandoPlanilha(caminhoPDF)
        except:
            print("Erro: Alterando dados da planilha na segunda tentativa")
