# GoogleSheets v1.0.0 beta 1

Módulo criado com o objetivo de facilitar a edição de planilhas do Google.

### Python (com [PIP](https://www.treinaweb.com.br/blog/gerenciando-pacotes-em-projetos-python-com-o-pip))
```py
python ^= 3.10
```
Caso queira instalar o python utilizando [**Anaconda**](https://www.anaconda.com/)
```py
conda create -n googlesheets python=3.10
conda activate googlesheets
```

### Dependências
Envie o seguinte comando para instalar as [dependências Google](https://developers.google.com/sheets/api/quickstart/python)
```
python -m pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

### Como utilizar
1. Crie um arquivo chamado `client_secret.json` no diretório da aplicação
2. Este arquivo deve conter as informações de acesso de sua aplicação Google, para que permita o login e autorização de acesso às planilhas de sua conta Google
    - [📄 Tutorial Criação Aplicação Google](TUTORIAL.md)
3. Modifique o código `main` no final do arquivo `google_sheets.py` (Apenas adicionando o ID de sua planilha em branco o teste irá funcionar)
```py
if __name__ == '__main__':
    RANGE = "Página1!B2:C"
    gs = GoogleSheets('ADICIONE O ID DE SUA PLANILHA')
    
    data = [
        ["NOME", "IDADE"],
        ["Teste", "134"],
        ["Vitor", "27"],
        ["Novo", "Teste"],
        ["Mais", "Um Teste"]
    ]

    print(gs.update(data, RANGE))
    print(gs.sheetpage_id_by_name(RANGE))
    print(gs.add_border(f"{RANGE}{len(data)+1}"))
    print(gs.select(RANGE))
```
4. Rode o seguinte comando para executar suas modificações
```py
python google_sheets.py
```
5. Ele irá solicitar que você logue em sua conta Google
6. Após logar, será gerado um arquivo `token.json` ⚠️ **NÃO COMPARTILHE ELE!** ⚠️ Ele que lhe manterá logado e permite acessar o conteúdo de sua conta
7. Observe que as modificações foram realizadas em sua Planilha
