### Web Scraper do site WEG. Coleta os dados e atualiza os preços a partir do código WEG.
#### Execute a planilha, insira o código, e em xlwings, execute "Run main".

### Se estiver usando windows, basta usar o "Instalador.py" para a automatizar a instalação.
#### Caso contrário, terás que instalar manualmente:

OBS: É necessário ter instalado o python e configurado sua localização na variável de ambiente PATH:
```
pip install xlwings; playwright
```
Após instalar as dependências, instale os recursos
```
xlwings addin install
```
```
playwright install chromium
```
Na primeira execução da planilha o excel solicitará permissão. Conceda-a!

### Após instaladas as dependências, é necessário configurar um usuário válido, com acesso à busca da WEG! Para isso, crie uma variável de ambiente para o usuário "weguser" e uma para a senha "wegpass".
No windows basta executar (como administrador) os seguintes comandos no CMD
```
setx weguser "usuario@email.com"
```
```
setx wegpass "senha123"
```

#### Recomenda-se instalar globalmente! Caso queira usar em um ambiente virtual, no addin xlwings, especifique a localização do interpretador virtual!