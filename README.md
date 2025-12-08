# Projeto destinado à coleta de dados a partir de um determinado código WEG

# Antes de começar é necessário sincronizar as dependencias e configurar o xlwings. Para isso, certifique-se de ter instalado o UV (pip install uv) e com o terminal aberto no root do repositório digite "uv sync".
## Para configurar o xlwings, digite no terminal "xlwings install addin".
### Na primeira execução da planilha é necessário conceder permissão. O próprio excel vai solicitar!

# Também é necessário configurar um usuário WEG válido, que tenha acesso à consulta de preços. Para isso crie duas variáveis de ambiente, contendo o usuário "weguser" e a senha "wegpass".
## Para criar as variáveis, no windows usar no CMD o comando "setx"
### Por exemplo: setx weguser "teste123@gmail.com" / setx wegpass "123"

# Após configurado basata abrir a planilha, inserir o código WEG e a quantidade. Para executar, na tab addons, selecione "xlwings" e clique em Run. O programa fará o resto sozinho.