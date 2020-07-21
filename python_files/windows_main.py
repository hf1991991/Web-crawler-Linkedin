from scrapy.crawler import CrawlerProcess
from multiprocessing import Process, freeze_support
from webcrawler.spiders.linkedin_spider import LinkedinSpider
import colorama
from termcolor import cprint
import os
import sys

colorama.init()

whiteprint = lambda x: cprint(x, 'white')
checkprint = lambda x: cprint(x, 'green')
errorprint = lambda x: cprint(x, 'red')

def format_file_path(path):
    try:
        if path[0] == '"':
            path = path[1:]
        if path[-1] == '"':
            path = path[:-1]
        while path[0] == ' ':
            path = path[1:]
        while path[-1] == ' ':
            path = path[:-1]
    except IndexError:
        pass
    return path

def execute_crawling(excel_file):
    process = CrawlerProcess()
    # dispatcher.connect(set_result, signals.item_scraped)
    process.crawl(LinkedinSpider, excel_file)
    process.start()

def is_excel_open(excel_file):
    try:
        f = open(excel_file, "r+")
        f.close()
        return False
    except Exception:
        return True

if __name__ == '__main__':
    # https://stackoverflow.com/questions/24944558/pyinstaller-built-windows-exe-fails-with-multiprocessing
    freeze_support()

    while True:

        whiteprint('\nDigite o caminho do Excel (ou selecione o arquivo no Windows Explorer e arraste-o para esta janela): ')
        excel_file = format_file_path(input())
        print()
        
        if is_excel_open(excel_file):
            errorprint("Não foi possível abrir o arquivo. Caso esteja aberto, salve e feche o Excel!")
        else:
            checkprint("Arquivo localizado! Seguindo para o programa.")
            break

    while True:

        while is_excel_open(excel_file):
            errorprint("Salve e feche o Excel!\nPara continuar, clique enter.")
            input()

        p = Process(target=execute_crawling, args=(excel_file,))
        p.start()
        p.join()

        print()
        checkprint("Processo finalizado!\nAperte enter para abrir o Excel.")
        input()

        os.system('start EXCEL.EXE "%s"' % excel_file)

        whiteprint("Quer realizar o crawl novamente? (s/n)")

        if input() != 's': break

        print()