# Webcrawler de empresas do LinkedIn

Olá! Este projeto foi desenvolvido para a Distrito pela [Poli Júnior](https://polijunior.com.br), empresa júnior de engenharia da Poli-USP.

O objetivo foi criar uma ferramenta, usando a biblioteca  de Python Scrapy, para conseguir dados de empresas e perfis do LinkedIn que podem ser úteis para a Distrito conseguir entender as redes de conexões feitas nessa rede social.

Em caso de qualquer dúvida, sintam-se à vontade para contatar a Poli Júnior ou a equipe do projeto:

- [Rodrigo Magaldi](mailto:rodrigo.magaldi@polijunior.com.br)
- [Henrique Falconer](mailto:henrique.falconer@polijunior.com.br)

Muito obrigado!

---

## 1. Primeira vez usando o crawler - passo a passo:

### Parte 1 - Crie shortcuts dos executáveis:
  
  1. Entre na pasta **src** deste diretório.
  2. Selecione os executáveis **webcrawler_empresas.exe** e **webcrawler_perfis.exe** e, com o botão direito, crie shortcuts para ambos.
  3. Arraste esses shortcuts para a pasta mãe.

### Parte 2 - Abra os shortcuts:

  1. Dê um duplo clique em ambos **webcrawler_empresas.exe - Shortcut** e **webcrawler_perfis.exe - Shortcut**.
  2. Caso esteja utilizando **Windows 10**, você notará que aparecerá uma janela do Windows Defender com a seguinte mensagem: **"Microsoft Defender SmartScreen prevented an unrecognized app from starting. Running this app might put your PC at risk."** Isso ocorre pois o executável criado não possui um certificado associado a ele e, portanto, o Windows desconfia que possa ser software malicioso. Para continuar, basta apertar **More info** e, depois, **Run anyway**, caso confie que os executáveis não contêm vírus.
  3. Logo em seguida, quando cada um dos executáveis começar a rodar, aparecerá o texto **"config.json carregado!"** em verde. Vendo isso, simplesmente feche cada uma das janelas.

### Parte 3 - Configure os cookies do webcrawler:

  1. Acesse o [link](https://chrome.google.com/webstore/detail/cookiestxt/njabckikapfpffapmjgojcnbfjonfjfg) para usar no Chrome uma extensão capaz de pegar os cookies do navegador. 
  2. Aperte o botão **Usar no Chrome** e, em seguida, em **Adicionar extensão**.
  3. Após aparecer uma janela de confirmação, aperte o botão com formato de quebra-cabeça, na parte direita do menu do topo do Chrome. Em seguida, aperte o botão de fixar, do lado da extensão **cookies.txt**. Você verá que aparecerá um botão da extensão do lado do com formato de quebra-cabeça.



Pronto! as configurações iniciais estão prontas. Agora é só seguir com os passos abaixo.

---

## 2. Passos para usar o crawler:

### Parte 1 - Copie os cookies do Chrome:

  1. Entre e faça login em [LinkedIn](https://www.linkedin.com) com os mesmos usuário e senha do arquivo **config.json**.
  2. Após navegar para alguma página do LinkedIn, aperte o botão da extensão do **cookies.txt**, na parte direita do menu do topo do Chrome. Deverá aparecer uma janela com texto.
  3. Copie todo o texto desta nova janela e insira-o no arquivo **cookies.txt** deste diretório, substituindo o texto anterior caso necessário.

> Obs.: Cookies deveriam tornar-se inválidos aproximadamente a cada 3 meses, porém, esse nem sempre é o caso. Quando você manualmente faz logout ou quando o LinkedIn o faz por você (especificamente em casos de suspeita do usuário ser um bot), os seus cookies do LinkedIn são automaticamente invalidados. Para impedir que isso aconteça, reduza o  limite de páginas acessadas por dia, no arquivo "config.json" deste diretório.

### Parte 2 - Rode o webcrawler

  1. Abra o arquivo **input.xlsx** deste diretório.
  2. Na parte inferior do Excel, você encontrará duas abas: **Empresas** e **Perfis**. Escolha a primeira caso queira obter dados de funcionários de empresas, ou escolha a segunda caso queira obter dados de perfis específicos. Siga as instruções da tabela que escolheu para adicionar novos links.
  2. Caso não queira que seus dados da última vez que você rodou o webcrawler se misturem com o desta sessão, entre em **output_empresas.json** ou **output_perfis.json** (dependendo de qual webcrawler você planeja rodar) e apague todo seu conteúdo.
  3. Rode **webcrawler_empresas.exe - Shortcut** deste diretório para carregar dados de funcionários das empresas cujos links você inseriu na aba **Empresas** de **input.xlsx**, ou **webcrawler_perfis.exe - Shortcut** para carregar dados de perfis específicos cujos links você inseriu na aba **Perfis** de **input.xlsx**.

---

## 3. Informações importantes sobre o crawler:

  - A recomendação é que esta ferramenta de crawling acesse, no máximo, **80** páginas por dia usando uma conta de categoria abaixo de **Premium Business**. Porém, caso a conta utilizada seja **Premium Business** ou superior, esse limite torna-se **150**.
  - No caso do webcrawler de empresas, como é necessário acessar perfis de pessoas através de menus do próprio site do LinkedIn, o uso de uma conta sem muitas conexões não permite que muitos dados de funcionários sejam obtidos, já que o LinkedIn não compartilha as URLs de perfis fora de sua rede de conexões. Por isso, tenha em mente que **quanto mais conexões a conta utilizada possuir, maior a proporção de dados de usuários obtidos de uma empresa**.
  - Na mesma lógica do tópico anterior, o LinkedIn também limita o acesso a conexões de perfis que passam pelo crawler em questão da rede da sua conta, de modo que nem sempre conexões de tais perfis possam ser obtidas. Novamente, a regra é que, **quanto mais conexões a conta utilizada possuir, mais dados serão obtidos**. Eis o que foi observado:
    * Se um perfil possui **conexão de 1° grau** com a conta utilizada no crawler, o acesso é ilimitado às conexões desse.
    * Se um perfil possui **conexão de 2º grau** com a conta utilizada, pode-se acessar apenas as conexões dele que também são conexões de 1º grau da própria conta utilizada.
    * E se um perfil possui **conexão de 3º ou maior grau** com a conta, os resultados que o Linkedin envia não parecem fazer sentido. Aparecem resultados, porém claramente não são os verdadeiros do perfil em questão. Por conta disso, foi implementada a opção de acessar apenas as páginas de conexões de, no máximo, 2° grau de conexão da conta utilizada, que pode ser desativada no arquivo **config.json**.