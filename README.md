# Web scraper Linkedin

Autor: Henrique Falconer

## Instalação de módulos para Python

Para rodar o programa em Python, siga os passos dos seguintes links:

- https://docs.scrapy.org/en/latest/intro/install.html
- https://openpyxl.readthedocs.io/en/stable/#installation
- https://pypi.org/project/colorama/
- https://pypi.org/project/termcolor/

## Como transformar os arquivos de Python em executável (Windows)

Primeiro, instale o seguinte módulo. Ele é uma versão GUI do pyinstaller.

```pip install auto-py-to-exe```

Após a instalação, execute o seguinte código no terminal. Isso deve fazer com que seja aberta uma nova janela.

```auto-py-to-exe```

Com isso, selecione o arquivo .py que você queira transformar em um executável, assim como todos os arquivos secundários do seu programa.

Por fim, aperte o botão "Convert .py to .exe" et voilà!

## Como transformar os arquivos de Python em executável (MacOS)

Rode o seguinte comando:

```
pyinstaller --noconfirm --onefile --console --name "webcrawler" --add-data "/Users/henriquefalconer/Desktop/Poli Júnior/NTec /Projetos/Webcrawler Linkedin/python_files/webcrawler:webcrawler/" --hidden-import "pkg_resources.py2_warn" --hidden-import "scrapy.spiderloader" --hidden-import "scrapy.statscollectors" --hidden-import "scrapy.logformatter" --hidden-import "scrapy.extensions" --hidden-import "scrapy.extensions.logstats" --hidden-import "scrapy.extensions.corestats" --hidden-import "scrapy.extensions.memusage" --hidden-import "scrapy.extensions.feedexport" --hidden-import "scrapy.extensions.memdebug" --hidden-import "scrapy.extensions.closespider" --hidden-import "scrapy.extensions.throttle" --hidden-import "scrapy.extensions.telnet" --hidden-import "scrapy.extensions.spiderstate" --hidden-import "scrapy.core.scheduler" --hidden-import "scrapy.core.downloader" --hidden-import "scrapy.downloadermiddlewares" --hidden-import "scrapy.downloadermiddlewares.robotstxt" --hidden-import "scrapy.downloadermiddlewares.httpauth" --hidden-import "scrapy.downloadermiddlewares.downloadtimeout" --hidden-import "scrapy.downloadermiddlewares.defaultheaders" --hidden-import "scrapy.downloadermiddlewares.useragent" --hidden-import "scrapy.downloadermiddlewares.retry" --hidden-import "scrapy.core.downloader.handlers.http" --hidden-import "scrapy.core.downloader.handlers.s3" --hidden-import "scrapy.core.downloader.handlers.ftp" --hidden-import "scrapy.core.downloader.handlers.datauri" --hidden-import "scrapy.core.downloader.handlers.file" --hidden-import "scrapy.downloadermiddlewares.ajaxcrawl" --hidden-import "scrapy.core.downloader.contextfactory" --hidden-import "scrapy.downloadermiddlewares.redirect" --hidden-import "scrapy.downloadermiddlewares.httpcompression" --hidden-import "scrapy.downloadermiddlewares.cookies" --hidden-import "scrapy.downloadermiddlewares.httpproxy" --hidden-import "scrapy.downloadermiddlewares.stats" --hidden-import "scrapy.downloadermiddlewares.httpcache" --hidden-import "scrapy.spidermiddlewares" --hidden-import "scrapy.spidermiddlewares.httperror" --hidden-import "scrapy.spidermiddlewares.offsite" --hidden-import "scrapy.spidermiddlewares.referer" --hidden-import "scrapy.spidermiddlewares.urllength" --hidden-import "scrapy.spidermiddlewares.depth" --hidden-import "scrapy.pipelines" --hidden-import "scrapy.dupefilters" --hidden-import "queuelib" --hidden-import "scrapy.squeues" "/Users/henriquefalconer/Desktop/Poli Júnior/NTec /Projetos/Webcrawler Linkedin/python_files/macos_main.py"
```