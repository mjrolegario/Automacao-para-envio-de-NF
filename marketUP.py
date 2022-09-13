from select import select
from traceback import print_tb
from bibliotecas import *
from QB import *

#Entrar no marketup
def acessandoMP():
    navegador.get('##########')
#Fazer login 
def fazerLoginMP(email, senha):
    navegador.find_element('xpath', '######').click()
    navegador.find_element('xpath', '######').send_keys(email)
    navegador.find_element('xpath', '######').click()
    navegador.find_element('xpath', '######').send_keys(senha)
    navegador.find_element('xpath', '######').click()
#vendas - pedidos
def acessandoVendas():
    navegador.find_element('xpath', '######').click()
    navegador.find_element('xpath', '######').click()
#adicionar novo pedido
def botNovoPedido():
    navegador.find_element('xpath', '######').click()      
def nome_cliente_func():
    navegador.find_element('xpath', '######').send_keys(nomeDoCliente)
    navegador.find_element('xpath', '######').send_keys(Keys.ENTER)
def adicionandoProduto(produto):
    navegador.find_element('xpath', '######').click()
    #produto
    navegador.find_element('xpath', '######').send_keys(produto)
    navegador.find_element('xpath', '######').send_keys(Keys.ENTER)
#quantidade a adicionar
def quantidadeProduto():    
    navegador.find_element('xpath', '######').click()
#retirar
def retirarProduto():
    navegador.find_element('xpath', '######').click()
#botão incluir
def incluirProduto():
    navegador.find_element('xpath', '######').click()
def lendoPlanilhaAddProdutos():
    fatura = pd.read_excel(f'Fatura {numFatura} modificada.xlsx')
    contador = 0
    while contador < 10:
        produto_tabela = fatura.loc[contador, 'Produtos']
        quantidade_tabela = fatura.loc[contador, 'QTDE.']
        valor_tabela = fatura.loc[contador, 'V UNITARIO']
        contador += 1
        adicionandoProduto(produto_tabela)
        contador2 = 0
        while contador2 < quantidade_tabela:
            quantidadeProduto()
            contador2 = contador2 + 1
        retirarProduto()
        incluirProduto()
        if produto_tabela and quantidade_tabela and valor_tabela == "":
            break
try:
    navegador.execute_script('window.scrollBy(0, 800)')
except:
    None

def formaDePagamento(forma_de_pagamento):
    #dinheiro
    if forma_de_pagamento == "dinheiro":
        navegador.find_element('xpath', '######').click("")
    #vale
    elif forma_de_pagamento == "vale":
        navegador.find_element('xpath', '######').click()
    #cartão de crédito
    elif forma_de_pagamento == "cartão de crédito":
        navegador.find_element('xpath', '######').click()
    #cheque
    elif forma_de_pagamento == "cheque":
        navegador.find_element('xpath', '######').click()
    #cartão de débito
    elif forma_de_pagamento == "cartão de débito":
        navegador.find_element('xpath', '######').click()
    #cartão benefício
    elif forma_de_pagamento == "cartão benefício":
        navegador.find_element('xpath', '######').click()
    #boleto
    elif forma_de_pagamento == "boleto":
        navegador.find_element('xpath', '######').click()
    #depósito
    elif forma_de_pagamento == "depósito":
        navegador.find_element('xpath', '######').click()
    #pagamento digital
    elif forma_de_pagamento == "pagamento digital":
        navegador.find_element('xpath', '######').click()
    #TEF
    elif forma_de_pagamento == "tef":
        navegador.find_element('xpath', '######').click()
    #Transferência
    elif forma_de_pagamento == "transferência":
        navegador.find_element('xpath', '######').click()
    else:
        print("Erro na forma de pagamento")
        return
#gerar contas
def gerar_contas():
    navegador.find_element('xpath', '######').click()
#concluir venda
def concluir_venda():
    navegador.find_element('xpath', '######').click()


def mainMP(email, senha,  Forma_de_pagamento):
    try:
        acessandoMP()
    except:
        print("Erro: Acesso MarketUP primeira tentativa")
        try:
            acessandoMP()
        except:
            print("Erro: Acesso MarketUP na segunda tentativa")
            return
    try:
        fazerLoginMP(email, senha)
    except:
        print("Erro: Fazer login primeira tentativa")
        try:
            fazerLoginMP(email, senha)
        except:
            print("Erro: Fazer login na segunda tentativa")
            try:
                navegador.find_element('xpath', '######').click() #caso a pagina para carregar aguardar 1 segundo e tentar novamente
                fazerLoginMP(email, senha)
            except:
                print("Erro: Fazer login na terceita tentativa")
                return
    try:
        acessandoVendas()
    except:
        print("Erro: Acesso vendas primeira tentativa")
        try:
            acessandoVendas()
        except:
            print("Erro: Acesso vendas na segunda tentativa")
            try:
                navegador.find_element('xpath', '######').click() #caso a pagina demore carregar
                acessandoVendas()
            except:
                print("Erro: Acesso vendas na terceira tentativa")
                return
    try:
        botNovoPedido()
    except:
        print("Erro: Acesso novo pedido primeira tentativa")
        try:
            botNovoPedido()
        except:
            print("Erro: Acesso novo pedido na segunda tentativa")
            try:
                navegador.find_element('xpath', '######').click() #caso a pagina demore carregar
            except:   
                print("Erro: Acesso novo pedido na terceira tentativa")         
                return
    try:
        nome_cliente_func()
    except:
        print("Erro: Cliente não cadastrado!")
        try:
            nome_cliente_func()
        except:
            print("Erro ao localizar cliente segunda tentativa")
            return
    try:
        lendoPlanilhaAddProdutos()
    except:
        print("Erro: Adicionar produto")
        try:                              
            navegador.find_element('xpath', '######').click() #caso a pagina demore carregar
            lendoPlanilhaAddProdutos()
        except:
            print("Erro: Adicionar produto segunda tentativa")
            try: #em caso de erro no xpath faz a busca oela tag 
                tag_u = navegador.find_elements(By.TAG_NAME, 'button')
                contador = 0
                for tag in tag_u:
                    if tag_u[contador].text==(" para informar o endereço depois"):
                        tag_u[contador].click()
                    contador = contador+1
                lendoPlanilhaAddProdutos()
            except:
                print("Produtos não adicionados")
                return
    forma_de_pagamento = Forma_de_pagamento
    try:
        formaDePagamento(forma_de_pagamento)
    except:
        print("Erro: Forma de pagamento primeira tentativa")
        try:
            formaDePagamento(forma_de_pagamento)
        except:
            print("Erro: Forma de pagamento na segunda tentativa")
            return
    try:
        gerar_contas()
    except:
        print("Erro: Gerar contas primeira tentativa")
        try:
            gerar_contas()
        except:
            print("Erro: Gerar contas na segunda tentativa")
            return
    try:
        concluir_venda()
    except:
        print("Erro: Concluir venda primeira tentativa")
        try:
            concluir_venda()
        except:
            print("Erro: concluir venda na segunda tentativa")
            return
