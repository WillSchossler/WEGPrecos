## Automação da coleta de dados do website da WEG a partir de um código.

### Basta executar a planilha, inserir o código e quantidade, e na aba addons, selecionar xlwings e clicar em "Run".

## Antes de começar é necessário instalar as dependencias. Certifique-se de ter o python instalado e o terminal no diretório com os arquivos.

```
uv sync
```

```
xlwings install addin
```

## É necessário ter um usuário WEG válido com permissão para procurar dados. Armazene o login e senha em variáveis de ambiente (weguser e wegpass)

No windows (CMD executado como administrador):
```
setx weguser "123@teste.com"
```
```
setx wegpas "123"
```

### Na primeira execução é necessário conceder permissão. O próprio excel solicitará!
