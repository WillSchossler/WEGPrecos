import os
import xlwings as xl
from playwright.sync_api import sync_playwright, expect, TimeoutError



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



class Scrapper:
    def __init__(self, usuario: int, senha: int, documento: Excel, codigo: str, quantidade: str):
        self.codigo = codigo
        self.quantidade = quantidade
        self.documento = documento
        self.user_name = usuario #"gabrielschossler.principal@gmail.com" #"eletronvale@eletronvale.com.br"
        self.user_pass = senha #"Copoloco140988." #"98@Eletronvale28"
        self.page_login = "https://www.weg.net/catalog/weg/BR/pt/login"
        self.page_search = "https://www.weg.net/catalog/weg/BR/pt/research/delivery-availability"
        

        with sync_playwright() as playwright:
            self.browser = playwright.chromium.launch()

            try:
                self.documento.status("Conectanto com o servidor")
                self.context = self.browser.new_context(storage_state=f"{caminho}\\state.json")
                self.page = self.context.new_page()
                self.page.goto(self.page_search)
                expect(self.page.locator('//*[@id="wegDeliveryAvailabilityFormId"]/fieldset/div[2]/div/a')).to_be_visible()
                self.pesquisar(self.codigo, self.quantidade, goto=False)
            except FileNotFoundError:
                """ Não tem arquivo de sessão """
                self.documento.status("Cookies não encontrados. Iniciando nova sessão", False)
                self.logar(create=True)
                self.pesquisar(self.codigo, self.quantidade)
            except AssertionError:
                self.documento.status("Cookies expirados. Iniciando nova sessão", False)
                self.logar()
                self.pesquisar(self.codigo, self.quantidade)
    
    
            

    def pesquisar(self, codigo, quantidade, goto=True):        
        if goto:
            self.page.goto(self.page_search)
        expect(self.page.locator('//*[@id="productCode"]')).to_be_visible()
        self.page.locator('//*[@id="productCode"]').fill(str(codigo))
        self.page.locator('//*[@id="quantity"]').fill(str(quantidade))
        self.page.locator('//*[@id="wegDeliveryAvailabilityFormId"]/fieldset/div[2]/div/a').click()

        self.documento.status(f"Efetuando pesquisa de {quantidade} unidades do item {codigo}")

        # Coleta dos dados retornados! Talvez usar um parser? Ou simplesmente uma lista ou dicionário? APRIMORAR!
        nome = self.page.locator("//dt[text()='Descrição do Produto']/../dd").inner_text()
        valor = self.page.locator("//th[text()='Preço Unitário']/../td").inner_text()
        faturamento = self.page.locator("//th[text()='Entrega Planejada']/../../../tbody/tr/td[2]").inner_text()
        entrega = self.page.locator("//th[text()='Entrega Planejada']/../../../tbody/tr/td[3]").inner_text()
        icms = self.page.locator("//tr/td[text()='% ICMS (incluso)']/..//td[2]").inner_text()
        ipi = self.page.locator("//tr/td[text()='% IPI (não incluso)']/..//td[2]").inner_text()            

        try:  # Procuramos se há um frete no produto. Maioria não tem, então precisa verificar
            path = "//tr/td[text()='% Frete']/..//td[2]"
            frete = expect(self.page.locator(path)).to_be_visible(timeout=2000)
            frete = self.page.locator(path).inner_text()
            
        except AssertionError:
            frete = 0

                    # Se encontrar tudo, então preenchemos no monitor
        self.documento.preencher_weg(nome, valor, ipi, frete, icms, faturamento, entrega)
        # E também na área de cáulco de impostos
        self.documento.preencher_valores(valor, ipi, frete, icms)
        self.documento.status(f"Pesquisa realizada com sucesso")





    def logar(self, create=False):
        if create:
            self.context = self.browser.new_context(user_agent="")
            self.page = self.context.new_page()        
        self.page.goto(self.page_login)
        expect(self.page.locator('//*[@id="j_username"]'))
        self.page.locator('//*[@id="j_username"]').fill(self.user_name)
        self.page.locator('//*[@id="j_password"]').fill(self.user_pass)
        self.page.locator('//*[@id="loginForm"]/div[3]/button').click()
        self.context.storage_state(path=f"{caminho}\\state.json")





caminho = os.path.dirname(os.path.abspath(__file__))
def main():
    Documento = Excel()  # Abre e configura o documento excel
    codigo, quantidade = Documento.coletar()  # Coleta o código que o usuário definiu
    Weg = Scrapper("eletronvale@eletronvale.com.br", "98@Eletronvale28", Documento, codigo, quantidade)  # Novo coletor de dados



if __name__ == '__main__':
    main()