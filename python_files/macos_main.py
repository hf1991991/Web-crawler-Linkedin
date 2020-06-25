from scrapy.crawler import CrawlerProcess
from multiprocessing import Process
from webcrawler.spiders.linkedin_spider import LinkedinSpider
import colorama
from termcolor import cprint
import os
import sys

# os.system("clear")

colorama.init()

whiteprint = lambda x: cprint(x, 'white')
checkprint = lambda x: whiteprint("✅ %s" % x)
errorprint = lambda x: whiteprint("❌ %s" % x)

excel_exists = False

def format_file_path(path):
    try:
        path = path.replace('\\', '')
        if path[0] == '/':
            path = path[1:]
        if 'Users' in path:
            path = '/'.join(path.split('/')[2:])
        while path[0] == ' ':
            path = path[1:]
        while path[-1] == ' ':
            path = path[:-1]
    except IndexError:
        pass
    return path

while not excel_exists:

    whiteprint('\nDigite o caminho do Excel (selecione o arquivo no Finder e aperte ⌘ C): ')
    excel_file = format_file_path(input())
    print()

    # Verificar se existe:
    try:
        f = open(excel_file, "r+")
        f.close()
        excel_exists = True
        checkprint("Arquivo localizado! Seguindo para o programa.\n")
    except FileNotFoundError:
        errorprint("O arquivo %s não existe. Digite novamente" % excel_file)

def execute_crawling():
    process = CrawlerProcess()
    # dispatcher.connect(set_result, signals.item_scraped)
    process.crawl(LinkedinSpider, excel_file)
    process.start()

dont_stop = True

while dont_stop:
    p = Process(target=execute_crawling)
    p.start()
    p.join()

    print()
    checkprint("Processo finalizado!\nAperte enter para abrir o Excel. (Caso já esteja aberto, feche-o antes)")
    input()

    os.system("open '/Applications/Microsoft Excel.app' '%s'" % excel_file)

    whiteprint("Quer realizar o crawl novamente? (s/n)")
    dont_stop = input() == 's'
    print()