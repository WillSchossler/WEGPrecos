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
            print(f"\nExecutando {comando}. Por favor aguarde!")
            subprocess.run(args=cmd, check=True)
            print(f"\n{comando} executado com sucesso!")
            return True

        except:
            user = input(f"\nOcorreu um erro ao executar {comando}. Insira 's' para tentar novamente, ou apenas 'enter' para continuar a instalação: ")
            if user.lower() != "s":
                return False


comandos = {
    "Instalação do xlwings": ["pip", "install", "xlwings"],
    "Atualização do addin xlwings pro excel": ["xlwings", "addin", "install"],
    "Instalação do playwright": ["pip", "install", "playwright"],
    "Instalação do browser chromium": ["playwright", "install", "chromium"],
    }


erros = []
for indice, (comando, cmd) in enumerate(comandos.items()):
    resultado = rodar(cmd)
    
    if not resultado:
        erros.append(comando)


if not erros:
    input("\nParabéns! O instalador executou com sucesso todos os comandos! Pressione 'enter' para sair.")
else:
    input(f"""\nInfelizmente os seguintes comandos não puderam ser executados corretamente:
{"\n".join(erros)}

Por favor, resolva os erros e tente novamente.
Pressine 'enter' para encerrar o instalador.""")


