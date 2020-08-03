from scrapy.crawler import CrawlerProcess
from multiprocessing import Process, freeze_support
from webcrawler.spiders.linkedin_employees_spider import LinkedinSpider
import colorama
from termcolor import cprint
import os
import sys
import json

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


def file_exists(path):
    try:
        f = open(path, 'r+')
        f.close()
        return True
    except FileNotFoundError:
        return False


def parse_json_file(path):
    try:
        f = open(path, 'r+')
        data = json.loads(f.read())
        f.close()
        return data
    except json.decoder.JSONDecodeError:
        return None


def read_config_file(config_path):

    config_data = parse_json_file(config_path)

    if config_data is None:
        errorprint('O arquivo %s está mal formatado. Arrume-o e digite novamente' % config_path)
        return None

    for path in config_data['paths']:
        if not file_exists(config_data['paths'][path]):
            errorprint('O arquivo %s não existe ou está aberto. Entre no config.json e corrija o erro, ou fecheo-o.' % config_data['paths'][path])
            return None

    return config_data


def execute_crawling(username, password,  cookies_path,  employees_json_path):
    process = CrawlerProcess(
        settings={
            'DOWNLOAD_DELAY': 2
        },
    )
    process.crawl(
        LinkedinSpider, 
        username=username,
        password=password,
        cookies_path=cookies_path, 
        employees_json_path=employees_json_path
    )
    process.start()


if __name__ == '__main__':
    # https://stackoverflow.com/questions/24944558/pyinstaller-built-windows-exe-fails-with-multiprocessing
    freeze_support()

    config_exists = False
    employees_json_path = None
    cookies_path = None
    username = None
    password = None

    while not config_exists:

        # whiteprint('\nDigite o caminho do arquivo config.json (ou selecione o arquivo no Windows Explorer e arraste-o para esta janela): ')
        config_path = format_file_path('../config.json')
        print()

        if file_exists(config_path):
            config_exists = True
            checkprint('config.json encontrado!\n')

            config_data = read_config_file(config_path)

            if config_data is not None:
                employees_json_path = config_data['paths']['employees_json']
                cookies_path = config_data['paths']['cookies']
                username = config_data['login']['username']
                password = config_data['login']['password']

            else:
                config_exists = False

        else:
            errorprint('O arquivo %s não existe neste diretório ou está aberto. Feche-o e tente novamente' % config_path)

    dont_stop = True

    while dont_stop:
        p = Process(target=execute_crawling, args=(username, password, cookies_path,  employees_json_path))
        p.start()
        p.join()

        print()
        checkprint('Processo finalizado!\nAperte enter para abrir o arquivo JSON.')
        input()

        os.system('Notepad /a "%s"' % employees_json_path)

        whiteprint('Quer realizar o crawl novamente? (s/n)')
        dont_stop = input() == 's'
        print()