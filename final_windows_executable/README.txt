Webcrawler de empresas do Linkedin

~~~~~~ PASSOS NA PRIMEIRA VEZ USANDO O WEBCRAWLER: ~~~~~~

Parte 1 - Crie shortcuts dos executáveis:
  1. Entre na pasta "src" deste diretório.
  2. Selecione os executáveis "webcrawler_empresas.exe" e "webcrawler_perfis.exe" e, com o botão direito, crie 
shortcuts para ambos.
  3. Arraste esses shortcuts para a pasta mãe.

Parte 2 - Configure os cookies do webcrawler:
  1. Acesse o link https://chrome.google.com/webstore/detail/cookiestxt/njabckikapfpffapmjgojcnbfjonfjfg
  2. Aperte o botão "Usar no Chrome" e, em seguida, em "Adicionar extensão".
  3. Após aparecer uma janela de confirmação, aperte o botão com formato de quebra-cabeça, na parte direita 
do menu do topo do Chrome. Em seguida, aperte o botão de fixar, do lado da extensão "cookies.txt". Você verá
que aparecerá um botão do próprio "cookies.txt" do lado do botão com formato de quebra-cabeça.

Com isso, as configurações iniciais estão prontas. Agora é só seguir com os passos abaixo.

~~~~~~~~~~~~ PASSOS PARA USAR O WEBCRAWLER: ~~~~~~~~~~~~

Parte 1 - Copie os cookies do Chrome:
  1. Entre e faça login em www.linkedin.com com os mesmos usuário e senha do arquivo "config.json".
  2. Após navegar para alguma página do Linkedin, aperte o botão da extensão do "cookies.txt", na parte direita 
do menu do topo do Chrome. Deverá aparecer uma janela com texto.
  3. Copie todo o texto desta nova janela e insira-o no arquivo "cookies.txt" deste diretório, substituindo 
o texto anterior caso necessário.

Obs. 1: Cookies deveriam expirar aproximadamente uma vez a cada 3 meses, porém, esse nem sempre é o caso. Quando 
você manualmente faz logout ou quando o Linkedin o faz por você (especificamente em casos de suspeita do usuário
ser um bot), os seus cookies do Linkedin são automaticamente expirados. Para impedir que isso aconteça, reduza o 
limite de páginas acessadas por dia, no arquivo "config.json" deste diretório.

Parte 2 - Rode o webcrawler
  1. Abra o arquivo "input.xlsx" deste diretório.
  2. Na parte inferior do Excel, você encontrará duas abas: "Empresas" e "Perfis". Escolha a primeira caso queira
obter dados de funcionários de empresas específicas, ou escolha a segunda caso queira obter dados de perfis 
específicos. Siga as instruções da tabela que escolheu para adicionar novos links.
  2. Caso não queira que seus dados da última vez que você rodou o webcrawler se misturem com o desta sessão,
entre em "output_empresas.json" ou "output_perfis.json" (dependendo de qual webcrawler você planeja rodar) e 
apague todo seu conteúdo.
  3. Rode "webcrawler_empresas.exe - Shortcut" deste diretório para carregar dados de funcionários das empresas
cujos links você inseriu na aba "Empresas" de "input.xlsx", ou "webcrawler_perfis.exe - Shortcut" para carregar 
dados de perfis específicos cujos links você inseriu na aba "Perfis" de "input.xlsx".

Obs. 2: A recomendação é que esta ferramenta de crawling acesse, no máximo, 80 páginas por dia usando uma conta
não Premium Business. Porém, caso a conta utilizada seja Premium Business, esse limite torna-se 150.