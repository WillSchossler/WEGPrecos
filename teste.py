import os
import xlwings as xl
from playwright.sync_api import sync_playwright, expect

PATH = os.path.dirname(os.path.abspath(__file__))



class Scrapper:
    def __init__(self, documento=None):
        """ Criação do objeto que vai fazer a invasão ao website da weg! Requer um documento xlwings"""
        self.documento = documento
        self.playwright = sync_playwright().start()
        self.device = self.playwright.devices['Desktop Edge']
        self.browser = self.playwright.chromium.launch(headless=False, channel="chromium")

        self.user = {
            "gabriel": {"name": "gabrielschossler.principal@gmail.com", "pass": "Copoloco140988."},
            "eletronvale": {"name": "eletronvale@eletronvale.com.br", "pass": "98@Eletronvale28"}
        }
        self.pages = {
            "login": "https://www.weg.net/catalog/weg/BR/pt/login",
            "search": "https://www.weg.net/catalog/weg/BR/pt/research/delivery-availability"
        }     


        try:
            # Tenta carregar uma sessão anterior
            self.context = self.browser.new_context(**self.device, storage_state=f"{PATH}\\cookies.json")
            self.page = self.context.new_page()
            self.page.goto(self.pages['search'])
            expect(self.page.locator('//h2[@class="page-header" and text()="Disponibilidade e Preço"]'), "Não foi possivel logar").to_be_visible(timeout=2000)
        except FileNotFoundError:
            # Não existe arquivo de sessão
            self.logar(criar=True)
        except AssertionError:
            # Sessão expirou
            self.logar()


    def pesquisar(self, codigo, quantidade="1"):
        """ Pesquisa um determinado código WEG """
        url_pesquisa = self.pages['search']

        if not (self.page.url == url_pesquisa):
            self.page.goto(url_pesquisa)
        
        self.page.locator('//*[@id="productCode"]').fill(str(codigo))
        self.page.locator('//*[@id="quantity"]').fill(str(quantidade))
        self.page.locator('//*[@id="wegDeliveryAvailabilityFormId"]/fieldset/div[2]/div/a').click()
        
        # Coleta os dados
        nome = self.page.locator("//dt[text()='Descrição do Produto']/../dd").inner_text().strip()
        valor = self.page.locator("//th[text()='Preço Unitário']/../td").inner_text().strip()
        faturamento = self.page.locator("//th[text()='Entrega Planejada']/../../../tbody/tr/td[2]").inner_text().strip()
        entrega = self.page.locator("//th[text()='Entrega Planejada']/../../../tbody/tr/td[3]").inner_text().strip()
        icms = self.page.locator("//tr/td[text()='% ICMS (incluso)']/..//td[2]").inner_text().strip()
        ipi = self.page.locator("//tr/td[text()='% IPI (não incluso)']/..//td[2]").inner_text().strip()      

        try:  # Procuramos se há um frete no produto
            local = "//tr/td[text()='% Frete']/..//td[2]"
            frete = expect(self.page.locator(local)).to_be_visible(timeout=100)
            frete = self.page.locator(local).inner_text()            
        except AssertionError:
            frete = 0

        for value in [nome, valor, faturamento, entrega, icms, ipi, frete]:
            print(value)





    def logar(self, criar=False):
        """ Efetua login no servidor. Se não houver arquivo de sessão, cria um novo """
        url_logar = self.pages['login']
        if criar:
            self.context = self.browser.new_context(**self.device)
            self.page = self.context.new_page()

        if not (self.page.url == url_logar):
            self.page.goto(url_logar)


        self.page.locator('//*[@id="j_username"]').fill(self.user['eletronvale']['name'])
        self.page.locator('//*[@id="j_password"]').fill(self.user['eletronvale']['pass'])
        self.page.locator('//*[@id="loginForm"]/div[3]/button').click()
        self.context.storage_state(path=f"{PATH}\\cookies.json")
        self.page.goto(self.pages['search'])
        


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
    with Scrapper() as Weg:
        Weg.pesquisar(14226362, 1)
        input()



if __name__ == '__main__':
    main()

