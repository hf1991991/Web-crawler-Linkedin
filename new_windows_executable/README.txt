PASSOS NA PRIMEIRA VEZ USANDO O WEBCRAWLER:

  1. Pesquise "cookies.txt extension" no google, usando o Chrome, e baixe a extensão.
  2. Entre e faça login em www.linkedin.com com os mesmos usuário e senha do arquivo "config.json".
  3. Apertando o botão de extensões na parte direita do menu do topo do Chrome, selecione a opção cookies.txt. 
Deverá aparecer uma janela com texto.
  4. Copie todo o texto desta nova janela e insira-o no arquivo "cookies.txt" deste diretório, substituindo 
o texto anterior caso necessários.
  5. As configurações iniciais estão prontas. Agora é só seguir com os passos abaixo.


PASSOS PARA RODAR O WEBCRAWLER:

  1. Atualize o conteúdo do arquivo "cookies.txt" deste diretório com os cookies da sua sessão atual do Linkedin 
no Chrome (Nota: toda vez que você fizer um novo login no Linkedin, você deverá atualizar este arquivo!).
  2. Entre em "Empresas.xlsx" e siga as instruções da tabela para adicionar links de empresas para passarem
pelo crawler.
  3. Caso não queira que seus dados da última vez que você rodou o webcrawler se misturem com o desta sessão,
entre em "output.json" neste diretório e apague todo seu conteúdo.
  4. Rode "webcrawler_empresas.exe - Shortcut" deste diretório para carregar dados e funcionários das empresas
cujos links você inseriu em "Empresas.xlsx".
  5. Rode webcrawler_funcionarios.exe - Shortcut" deste diretório para carregar os dados dos funcionários das empresas
carregadas pelo processo anterior.

Obs.: A recomendação é que ambas as ferramentas de crawling (de empresas e funcionários) acessem, no máximo e ao total, 
80 páginas por dia. Porém, caso a conta utilizada seja Premium Business, esse limite torna-se 150.