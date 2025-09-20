import os
import json
import xlwings as xl

from selenium.webdriver import Edge
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located



class Excel():
    def __init__(self, nome="Precos.xlsm"):
        xl.Book(nome).set_mock_caller()  # Define a planilha a ser executada
        self.book = xl.Book.caller()  # Define este, como o documento a ser executado
        self.precos = self.book.sheets[0]  # Primeira planinlha, onde a mágica acontece        
        
        # Dicionário com as celulas
        self.weg = {"codigo": "F3", "quantidade": "F6", "status": "J3", "nome": "J5", "valor": "M8", "total": "K16",
        "user": "J16", "ipi": "M9", "frete": "M10", "icms": "M11", "faturamento": "L13", "entrega": "N13"}

        self.calculo = {"valor": "C2", "custo": "C5", "lucro": "C6",
        "ipi": "C9", "frete": "C10", "icms": "C11", "preco_venda": "C7"}        


    def status(self, mensagem, cor=True):
        """ Envia uma mensagem de status """
        celula = self.precos[self.weg["status"]]
        if cor:  # Se não ouver argumento, então é amarelo
            celula.font.color = (255, 255, 0)
        else:  # Caso contrário é alerta, portando é vermelho
            celula.font.color = (255, 55, 55)
      
        celula.value = mensagem
        

    def preencher_weg(self, nome=None, valor=None, ipi=None, frete=None, icms=None, faturamento=None, entrega=None):
        self.precos[self.weg["valor"]].value = formatar(valor)
        self.precos[self.weg["nome"]].value = nome
        self.precos[self.weg["ipi"]].value = ipi
        self.precos[self.weg["icms"]].value = icms
        self.precos[self.weg["frete"]].value = frete
        self.precos[self.weg["entrega"]].value = entrega
        self.precos[self.weg["faturamento"]].value = faturamento


    def preencher_valores(self, valor, ipi, frete, icms):
        self.precos[self.calculo["ipi"]].value = ipi
        self.precos[self.calculo["frete"]].value = frete        
        self.precos[self.calculo["icms"]].value = f"{17 - formatar(icms)}%"
        self.precos[self.calculo["valor"]].value = formatar(valor)


    def coletar(self):
        codigo = self.precos[self.weg['codigo']].value
        quantidade = self.precos[self.weg['quantidade']].value
        return (int(codigo), int(quantidade))



class WEG():
    def __init__(self, login, senha, documento: Excel):
        self.login = login
        self.senha = senha
        self.documento = documento  # Objeto fundamental, ele quem irá passar para o excel as informações!
        self.pages = {"login": "https://www.weg.net/catalog/weg/BR/pt/login",
                      "account": "https://www.weg.net/catalog/weg/BR/pt/weg-account",
                      "pesquisar": "https://www.weg.net/catalog/weg/BR/pt/research/delivery-availability"}
        
        self.options = Options()
        self.setup(self.options)        
        self.driver = Edge(options=self.options)
        self.entrar()


    def setup(self, opcoes):
        """ Pra driblar a segurança da WEG, é preciso definir um agente de usuário! Aqui fazemos automaticamente """
        try:  # Tentamos localizar um arquivo de configuração
            with open(f"{caminho}\\config.json", 'r') as configs:
                arguments = json.load(configs)
                self.documento.status("Carregando configurações, por favor aguarde.", False)
                for key in arguments:
                    opcoes.add_argument(arguments[key])
        except:  # Se der erro, criamos um
            self.documento.status("Arquivo de configuração não localizado. Criando um novo.", False)
            temp_driver = Edge()  # Driver temporário
            temp_driver.get(self.pages['login'])  # Navega até a WEG
            agent = temp_driver.execute_script('return navigator.userAgent')  # Coleta o agente!
            with open(f"{caminho}\\config.json", "w") as configs:  # Salva em um arquivo json
                arguments = {"agente": f"user-agent={agent}", "mode": "headless"}  #, "mode": "headless"            
                json.dump(arguments, configs)
                for key in arguments:
                    opcoes.add_argument(arguments[key])
                self.documento.status("Configurações estabelecidas com sucesso!")
                temp_driver.quit()  # Fecha o driver temporário
      

    def entrar(self):
        self.driver.get(self.pages["login"])
        self.driver.delete_all_cookies()
        try:  # Tenta carregar os cookies (pode ser que não exista o arquivo JSON)
            with open(f"{caminho}\\cookies.json", 'r') as dados:
                cookies = json.load(dados)
                self.documento.status("Carregando Cookies da sessão antiga. Por favor aguarde.")
                for cookie in cookies:
                    self.driver.add_cookie(cookie)
        except:  # Se por acaso tiverem expirado, ou o arquivo JSON esteja limpo, a gente faz o login!
            self.documento.status("Cookies não encontrados, criando novo arquivo.", False)
        
        self.documento.status("Tentando conexão com servidor WEG.")
        self.driver.get(self.pages["account"])  # Entra na página da conta!
        
        try:  # Verifica se os cookies eram válidos, caso seja, um elemento com o nome de usuário vai ser detectado
            usuario = self.procurar("//a[@class='navbar-text pull-right xtt-show-popover']/span")         
        except:  # Se os cookies estiverem expirados, então iniciamos uma nova sessão
            self.documento.status(f"Cookies inválidos. Estabelecendo nova sessão.", False)
            self.logar()
            usuario = self.procurar("//a[@class='navbar-text pull-right xtt-show-popover']/span")
            
        self.documento.status(f"Conexão estabelecida! Seja bem vindo {usuario.text}.")
        self.aceitar()  # No final das contas a gente simplesmente aceita o pop-up de conscientização dos cookies


    def pesquisar(self, codigo, quantidade):
        """ Método PRINCIPAL! Procura itens pelo código WEG e quantidade! """
        self.documento.status("Conectado ao host de pesquisa.")
        self.driver.get(self.pages["pesquisar"])  # Primeiro conecta na página de pesquisa (usuário deve estar logad)
        # Identificação dos campos
        barra_codigo = self.procurar("//input[@id='productCode']")
        barra_quantidade = self.procurar("//input[@id='quantity']")
        # Limpeza (caso haja cookies de preenchimento prévio)
        barra_codigo.clear()
        barra_quantidade.clear()
        # Inserção dos dados
        barra_codigo.send_keys(codigo)
        barra_quantidade.send_keys(quantidade)
        # Validação do formulário!
        barra_codigo.send_keys(Keys.ENTER)
        self.documento.status(f"Efetuando pesquisa do item {codigo}.")
    
        try:  # Confere se a página foi carregada corretamente! Se for, então um elemento com o código pesquisado vai corresponder!
            conferir = WebDriverWait(self.driver, 6).until(presence_of_element_located((By.XPATH, f"//dd[@id='clientName']/span[text()='{codigo}']")))

            # Coleta dos dados retornados! Talvez usar um parser? Ou simplesmente uma lista ou dicionário? APRIMORAR!
            nome = self.procurar("//dt[text()='Descrição do Produto']/../dd").text
            valor = self.procurar("//th[text()='Preço Unitário']/../td").text
            faturamento = self.procurar("//th[text()='Entrega Planejada']/../../../tbody/tr/td[2]").text
            entrega = self.procurar("//th[text()='Entrega Planejada']/../../../tbody/tr/td[3]").text
            icms = self.procurar("//tr/td[text()='% ICMS (incluso)']/..//td[2]").text
            ipi = self.procurar("//tr/td[text()='% IPI (não incluso)']/..//td[2]").text            

            try:  # Procuramos se há um frete no produto. Maioria não tem, então precisa verificar
                frete = self.procurar("//tr/td[text()='% Frete']/..//td[2]").text
            except:
                frete = 0

                        # Se encontrar tudo, então preenchemos no monitor
            self.documento.preencher_weg(nome, valor, ipi, frete, icms, faturamento, entrega)
            # E também na área de cáulco de impostos
            self.documento.preencher_valores(valor, ipi, frete, icms)
            self.documento.status(f"Pesquisa do item {codigo} realizada com sucesso.")

        except:
            erro = WebDriverWait(self.driver, 5).until(presence_of_element_located((By.XPATH, "//div[@class='alert alert-danger alert-dismissible xtt-alert']/p")))        
            self.documento.status(erro.text, "vermelho")


    def logar(self):
        self.documento.status("Efetuando login. Por favor aguarde.")
        self.driver.get(self.pages["login"])    
        self.driver.delete_all_cookies()

        login = self.procurar("//input[@id='j_username']")
        senha = self.procurar("//input[@id='j_password']")
        login.send_keys(self.login)
        senha.send_keys(self.senha)
        login.send_keys(Keys.ENTER)
        self.documento.status("Login efetuado com sucesso! Salvando Cookies.")
        self.salvar_cookies()


    def salvar_cookies(self):
        cookies = self.driver.get_cookies()
        with open(f"{caminho}\\cookies.json", 'w') as dados:
            json.dump(cookies, dados, indent=4)
        self.documento.status("Cookies salvos com sucesso.")

    def aceitar(self):
        try:
            popup = WebDriverWait(self.driver, 5).until(presence_of_element_located((By.XPATH, "//div[@class='dp-bar-button dp-bar-dismiss cc-foreground-btn-456']")))        
            popup.click()
        except:
            pass

    def procurar(self, path):
        return self.driver.find_element(By.XPATH, path)

    def fechar(self):
        self.driver.quit()



@xl.func
def formatar(numero):
        """ Formata um valor do tipo 'R$ 12.345,00' para 12345.00 ou remove o % """
        try:  # Bloco try pra evitar problemas. Uma maneira de limpar os campos
            if ("%" in numero):
                return float(numero.replace("%", ""))
            else:            
                valor = numero.strip("R$ ")
                valor = valor.replace(".", "")
                valor = valor.replace(",", ".")
                return float(valor)
        except TypeError:
            return None



@xl.func
def calcular_total(x, y, total, qntd):
    if x == y:
        return float(total) * qntd
    else:
        return None


def main():    
    documento = Excel()  # Abre e configura o documento excel
    scrapper = WEG("eletronvale@eletronvale.com.br", "98@Eletronvale28", documento)  # Novo coletor de dados
    codigo, quantidade = documento.coletar()  # Coleta o código que o usuário definiu
    scrapper.pesquisar(codigo, quantidade)  # Efetua a pesquisa
    scrapper.fechar()



# Como xlwings não funciona com caminho relativo, é preciso pegar o caminho absoluto do arquivo
caminho = os.path.dirname(os.path.abspath(__file__))
if __name__ == "__main__":
    main()