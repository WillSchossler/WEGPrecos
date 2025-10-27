import os
from playwright.sync_api import sync_playwright, expect

PATH = os.path.dirname(os.path.abspath(__file__))



class Scrapper:
    def __init__(self):
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(headless=False, channel="chromium")
        self.device = self.playwright.devices['Desktop Edge']   

        self.user = {
            "gabriel": {"name": "gabrielschossler.principal@gmail.com", "pass": "Copoloco140988."},
            "eletronvale": {"name": "eletronvale@eletronvale.com.br", "pass": "98@Eletronvale28"}
        }
        self.pages = {
            "login": "https://www.weg.net/catalog/weg/BR/pt/login",
            "search": "https://www.weg.net/catalog/weg/BR/pt/research/delivery-availability"
        }     

        try:
            self.context = self.browser.new_context(**self.device, storage_state=f"{PATH}\\cookies.json")
            print("entrou")
        except FileNotFoundError:
            self.login()


    def login(self):
        self.context = self.browser.new_context(**self.device)
        self.page = self.context.new_page()
        self.page.goto(self.pages["login"])
        self.page.locator('//*[@id="j_username"]').fill(self.user['gabriel']['name'])
        self.page.locator('//*[@id="j_password"]').fill(self.user['gabriel']['pass'])
        self.page.locator('//*[@id="loginForm"]/div[3]/button').click()
        self.context.storage_state(path=f"{PATH}\\cookies.json")
        self.page.goto(self.pages['search'])
        expect(self.page.locator('//h2[@class="page-header" and text()="Disponibilidade e Preço"]'))
        input()




    def close(self):
        self.browser.close()
        self.playwright.stop()



"""with sync_playwright() as playwright:
    device = playwright.devices['Desktop Edge']
    browser = playwright.chromium.launch(headless=True, channel="chromium")
    context = browser.new_context(**device)
    page = context.new_page()

    page.goto(weg)
    page.screenshot(path="debug.png", full_page=True)"""


def main():
    Weg = Scrapper()
    
    Weg.close()



if __name__ == '__main__':
    main()