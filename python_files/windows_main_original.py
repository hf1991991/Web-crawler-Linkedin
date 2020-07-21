from scrapy.crawler import CrawlerProcess
from webcrawler.spiders.linkedin_spider import LinkedinSpider
import colorama
from termcolor import cprint
import os

colorama.init()

excel_file = 'Links.xlsx'

# Verificar se está aberto:
excel_open = False
try:
    open(excel_file, "r+")
    excel_open = True
except Exception:
    cprint("\nNão foi possível abrir o arquivo. Salve e feche o Excel!\nAperte enter continuar.", 'red')
    input()

if excel_open:
    process = CrawlerProcess()

    process.crawl(LinkedinSpider, excel_file)
    process.start()

    cprint("\nProcesso finalizado!\nAperte enter para abrir o Excel.", "green")
    input()

    os.system("start EXCEL.EXE %s" % excel_file)