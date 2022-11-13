# Tutorial Criação Aplicação para Google Sheets

**1.**  Acesse [console.cloud.google.com/projectcreat](https://console.cloud.google.com/projectcreate) e crie uma aplicação Google, caso ainda não tenha

![criação de projeto](https://media.discordapp.net/attachments/979061713171251243/1041442927056007208/image.png)

**2.**  Acesse [console.cloud.google.com/flows/enableapi?apiid=sheets.googleapis.com](https://console.cloud.google.com/flows/enableapi?apiid=sheets.googleapis.com) para permitir o uso da API de Planilhas do Google em sua aplicação

![permitindo aplicação](https://cdn.discordapp.com/attachments/979061713171251243/1041443548916101261/image.png)

**3.** Acesse [console.cloud.google.com/apis/credentials](https://console.cloud.google.com/apis/credentials) e crie uma nova credencial `ID do cliente OAuth` para gerar o arquivo `client_secret.json`

![criando credencial](https://media.discordapp.net/attachments/979061713171251243/1041444445305974824/image.png)

**4.** Após criar, clique no botão `FAZER DOWNLOAD DO JSON`, renomeie para `client_secret.json` e salve no diretório da aplicação

![criando credencial](https://cdn.discordapp.com/attachments/979061713171251243/1041445687709151322/image.png)
![salvando credencial](https://cdn.discordapp.com/attachments/979061713171251243/1041446187967975454/image.png)

**5.** Após isso, configure a tela de permissão de sua aplicação, ela que intermediará a autenticação em seu `google_sheets` com o Google

![tela de permissão](https://cdn.discordapp.com/attachments/979061713171251243/1041448071860588636/image.png)

**6.** Execute seu `google_sheets` através do comando:
```
python google_sheets.py
```

**7.** Caso seja o primeiro acesso, ele solicitará permissão através do Navegador

![tela de permissão](https://cdn.discordapp.com/attachments/979061713171251243/1041449297104867348/image.png)

Se você tiver colocado sua aplicação no modo `Produção` talvez apareça a seguinte mensagem:

![alerta google](https://media.discordapp.net/attachments/979061713171251243/1041450000158306356/image.png)

Basta clicar em `avançado` e prosseguir clicando em `Acessar NOME_APLICACAO (não seguro)` 
Após isso basta permitir que as planilhas google possam ser acessadas e clique em `Continuar`

![autorizando acesso](https://cdn.discordapp.com/attachments/979061713171251243/1041450990496395274/image.png)

**8.** Gerá gerado um arquivo `token.json` que lhe manterá logado enquanto o arquivo estiver lá ⚠️ **NÃO COMPARTILHE ELE!** ⚠️

**9.** Caso tenha configurado corretamente o código `main` do `google_sheets`, ele executará corretamente e modificará a planilha, exemplo:

![exemplo codigo](https://cdn.discordapp.com/attachments/979061713171251243/1041452682013057074/image.png)
![rodando no terminal](https://media.discordapp.net/attachments/979061713171251243/1041451909652938832/image.png)
![resultado](https://cdn.discordapp.com/attachments/979061713171251243/1041453375646085181/image.png)

**10.** Pronto, agora a planílha foi editada, caso tenha dúvida de como identificar o `ID` da planilha para passar no `GoogleSheets` basta copiar da URL dela:

![id planilha](https://cdn.discordapp.com/attachments/979061713171251243/1041454067697864754/image.png)
