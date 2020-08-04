from scrapy.crawler import CrawlerProcess
from multiprocessing import Process
from webcrawler.spiders.linkedin_companies_spider import LinkedinSpider
import colorama
from termcolor import cprint
import os
import sys
import json

# os.system("clear")

colorama.init()

whiteprint = lambda x: cprint(x, 'white')
checkprint = lambda x: whiteprint("✅ %s" % x)
errorprint = lambda x: whiteprint("❌ %s" % x)


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


def file_exists(path):
    try:
        f = open(path, "r+")
        f.close()
        return True
    except FileNotFoundError:
        return False


def parse_json_file(path):
    try:
        f = open(path, "r+")
        data = json.loads(f.read())
        f.close()
        return data
    except json.decoder.JSONDecodeError:
        return None


def read_config_file(config_path):

    config_data = parse_json_file(config_path)

    if config_data is None:
        errorprint("O arquivo %s está mal formatado. Arrume-o e digite novamente" % config_path)
        return None

    for path in config_data['paths']:
        if not file_exists(config_data['paths'][path]):
            errorprint("O arquivo %s não existe. Entre no config.json e corrija o erro." % config_data['paths'][path])
            return None

    return config_data


def execute_crawling():
    process = CrawlerProcess(
        settings={
            'DOWNLOAD_DELAY': 2
        },
    )
    process.crawl(
        LinkedinSpider, 
        username=username,
        password=password,
        companies_excel_path=companies_excel_path, 
        cookies_path=cookies_path, 
        employees_json_path=employees_json_path
    )
    process.start()


config_exists = False
companies_excel_path = None
employees_json_path = None
cookies_path = None
username = None
password = None

while not config_exists:

    whiteprint('\nDigite o caminho do arquivo config.json (selecione o arquivo no Finder e aperte ⌘ C): ')
    config_path = format_file_path(input())
    print()

    if file_exists(config_path):
        config_exists = True
        checkprint("Configurações carregadas!\n")

        config_data = read_config_file(config_path)

        if config_data is not None:
            companies_excel_path = config_data['paths']["companies_excel"]
            employees_json_path = config_data['paths']["employees_json"]
            cookies_path = config_data['paths']["cookies"]
            username = config_data['login']["username"]
            password = config_data['login']["password"]

        else:
            config_exists = False

    else:
        errorprint("O arquivo %s não existe. Digite novamente" % config_path)

dont_stop = True

while dont_stop:
    p = Process(target=execute_crawling)
    p.start()
    p.join()

    print()
    checkprint("Processo finalizado!\nAperte enter para abrir o Excel e o arquivo JSON. (Caso já estejam abertos, feche-os antes)")
    input()

    os.system("open '/Applications/Microsoft Excel.app' '%s'" % companies_excel_path)
    os.system("open '/Applications/TextEdit.app' '%s'" % employees_json_path)

    whiteprint("Quer realizar o crawl novamente? (s/n)")
    dont_stop = input() == 's'
    print()