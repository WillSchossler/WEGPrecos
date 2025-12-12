import subprocess

print("""Bem vindo ao instalador!
O script a seguir irá instalar todas as dependências necessárias, porém pode demorar um pouco.
Antes de mais nada, é necessário configurar um usuário WEG que tenha acesso a pesquisas.
    """)

username = input("Por favor, insira o e-mail: ")
password = input("\nPor favor, digite a senha: ")

def rodar(cmd):
    comando = ' '.join(cmd)
    while True:
        try:
            print(f"\nExecutando comando: {comando}. Por favor aguarde!")
            subprocess.run(args=cmd, check=True)
            print("\nComando executado com sucesso!")
            break

        except:
            user = input("\nOcorreu um erro. Gostaria de tentar novamente? Escolha 's': ")
            if user.lower() != "s":
                break

comandos = {
    "xlwings": ["pip", "install", "xlwings"],
}
rodar(["pip", "install", "xlwings"])  # Instala xlwings
rodar(["xlwings", "addin", "install"])  # Installa addin do excel

rodar(["pip", "install", "playwright"])  # Instala o playwright
rodar(["playwright", "install", "chromium"])  # Instala o browser chromium

rodar(["setx", "weguser", username])  # Define a variável de ambiente
rodar(["setx", "wegpass", password])  # E a senha

print("""Tarefas executadas
    """)