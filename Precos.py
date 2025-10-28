import os
import xlwings as xl
from playwright.sync_api import sync_playwright, expect
PATH = os.path.dirname(os.path.abspath(__file__))


@xl.func
def calcular_total(x, y, total, qntd):
    """ Apenas confere se as celulas estão de acordo e multiplica os valores """
    if x == y:
        return float(total) * qntd
    else:
        return None



class Excel():
    def __init__(self, nome="Precos.xlsm"):
        """ Handler que cuida da coleta, execução e inserção no excel """
        xl.Book(nome).set_mock_caller()  
        self.book = xl.Book.caller()  
        self.sheet = self.book.sheets[0]      
        
        # Dicionário com as celulas
        self.weg = {
            "codigo": "F3",
            "quantidade": "F6",
            "status": "J3",
            "nome": "J5",
            "valor": "M8",
            "total": "K16",
            "user": "J16",
            "ipi": "M9",
            "frete": "M10",
            "icms": "M11",
            "faturamento": "L13",
            "entrega": "N13",
        }

        self.calculo = {
            "valor": "C2",
            "custo": "C5",
            "lucro": "C6",
            "ipi": "C9",
            "frete": "C10",
            "icms": "C11",
            "preco_venda": "C7",
        }   


    def status(self, mensagem: str, cor=True):
        """ Envia uma mensagem de status """
        celula = self.sheet[self.weg["status"]]
        if cor:  # Se não ouver argumento, então é amarelo
            celula.font.color = (255, 255, 0)
        else:  # Caso contrário é alerta, portando é vermelho
            celula.font.color = (255, 55, 55)
      
        celula.value = mensagem


    def preencher_weg(self, nome, valor, ipi, frete, icms, faturamento, entrega):
        """ Coloca os dados coletados no monitor WEG (a caixa azul) """
        self.sheet[self.weg["valor"]].value = self.formatar(valor)
        self.sheet[self.weg["nome"]].value = nome
        self.sheet[self.weg["ipi"]].value = ipi
        self.sheet[self.weg["icms"]].value = icms
        self.sheet[self.weg["frete"]].value = frete
        self.sheet[self.weg["entrega"]].value = entrega
        self.sheet[self.weg["faturamento"]].value = faturamento


    def preencher_calculo(self, valor, ipi, frete, icms):
        """ Coloca os dados nas células para o cálculo do valor """
        self.sheet[self.calculo["ipi"]].value = ipi
        self.sheet[self.calculo["frete"]].value = frete        
        self.sheet[self.calculo["icms"]].value = f"{17 - self.formatar(icms)}%"
        self.sheet[self.calculo["valor"]].value = self.formatar(valor)


    def coletar(self):
        """ Retorna o código e a quantidade definida pelo usuário """
        codigo = self.sheet[self.weg['codigo']].value
        quantidade = self.sheet[self.weg['quantidade']].value
        return (int(codigo), int(quantidade))


    def formatar(self, numero):
        """ Formata um valor do tipo 'R$ 12.345,00' para 12345.00 ou remove o % """
        try: 
            if ("%" in numero):
                return float(numero.replace("%", ""))
            else:            
                valor = numero.strip("R$ ")
                valor = valor.replace(".", "")
                valor = valor.replace(",", ".")
                return float(valor)
        except TypeError:
            return None



class Scrapper:
    def __init__(self, documento: Excel):
        """ Criação do objeto que vai fazer a invasão ao website da weg! Requer um documento xlwings"""
        self.documento = documento
        self.documento.status("Conectando-se ao servidor")

        self.playwright = sync_playwright().start()
        self.device = self.playwright.devices['Desktop Edge']
        self.browser = self.playwright.chromium.launch(headless=True, channel="chromium")

        self.user = {"name": "eletronvale@eletronvale.com.br", "pass": "98@Eletronvale28"}
        self.pages = {
            "login": "https://www.weg.net/catalog/weg/BR/pt/login",
            "search": "https://www.weg.net/catalog/weg/BR/pt/research/delivery-availability"
        }     


        try:
            # Tenta carregar uma sessão anterior
            self.context = self.browser.new_context(**self.device, storage_state=f"{PATH}\\cookies.json")
            self.page = self.context.new_page()
            self.page.goto(self.pages['search'])
            expect(self.page.locator('//h2[@class="page-header" and text()="Disponibilidade e Preço"]')).to_be_visible(timeout=500)
        except FileNotFoundError:
            # Não existe arquivo de sessão
            self.documento.status("Primeira sessão. Gerando cookies")
            self.logar(criar=True)
        except AssertionError:
            # Sessão expirou
            self.documento.status("Sessão expirada. Realizando novo login")
            self.logar()



    def logar(self, criar=False):
        """ Efetua login no servidor. Se não houver arquivo de sessão, cria um novo """        
        if criar:
            self.context = self.browser.new_context(**self.device)
            self.page = self.context.new_page()

        url_logar = self.pages['login']
        if not (self.page.url == url_logar):
            self.page.goto(url_logar)

        self.page.locator('//*[@id="j_username"]').fill(self.user['name'])
        self.page.locator('//*[@id="j_password"]').fill(self.user['pass'])
        self.page.locator('//*[@id="loginForm"]/div[3]/button').click()

        self.context.storage_state(path=f"{PATH}\\cookies.json")        
        self.documento.status("Login efetuado com sucesso")
        


    def pesquisar(self, codigo, quantidade=1):
        """ Pesquisa um determinado código WEG """
        url_pesquisa = self.pages['search']
        if not (self.page.url == url_pesquisa):
            self.page.goto(url_pesquisa)
        
        self.page.locator('//*[@id="productCode"]').fill(str(codigo))
        self.page.locator('//*[@id="quantity"]').fill(str(quantidade))
        self.page.locator('//*[@id="wegDeliveryAvailabilityFormId"]/fieldset/div[2]/div/a').click()
        
        try:
            # Coleta os dados
            self.documento.status(f"Realizando pesquisa do item {codigo}")
            check = "//dt[text()='Descrição do Produto']/../dd"
            expect(self.page.locator(check)).to_be_visible(timeout=500)
            
            nome = self.page.locator(check).inner_text().strip()
            valor = self.page.locator("//th[text()='Preço Unitário']/../td").inner_text().strip()
            faturamento = self.page.locator("//th[text()='Entrega Planejada']/../../../tbody/tr/td[2]").inner_text().strip()
            entrega = self.page.locator("//th[text()='Entrega Planejada']/../../../tbody/tr/td[3]").inner_text().strip()
            icms = self.page.locator("//tr/td[text()='% ICMS (incluso)']/..//td[2]").inner_text().strip()
            ipi = self.page.locator("//tr/td[text()='% IPI (não incluso)']/..//td[2]").inner_text().strip()      

            try:  # Procuramos se há um frete no produto
                local = "//tr/td[text()='% Frete']/..//td[2]"
                frete = expect(self.page.locator(local)).to_be_visible(timeout=500)
                frete = self.page.locator(local).inner_text()            
            except AssertionError:
                frete = 0

            self.documento.preencher_calculo(valor, ipi, frete, icms)
            self.documento.preencher_weg(nome, valor, ipi, frete, icms, faturamento, entrega)
            self.documento.status(f"Pesquisa do item {codigo} realizada com sucesso")
        
        except AssertionError:
            erro = self.page.locator('//div[@class="alert alert-danger alert-dismissible xtt-alert"]/p').inner_text().strip()
            self.documento.status(erro)



    def close(self):
        """ Fecha objeto """
        self.page.close()
        self.context.close()
        self.browser.close()
        self.playwright.stop()


    def __enter__(self):
        return self


    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()



def main():
    Documento = Excel()
    with Scrapper(Documento) as Weg:
        codigo, quantidade = Documento.coletar()
        Weg.pesquisar(codigo, quantidade)



if __name__ == '__main__':
    main()

