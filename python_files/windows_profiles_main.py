from scrapy.crawler import CrawlerProcess
from multiprocessing import Process, freeze_support
from webcrawler.spiders.profiles_linkedin_spider import ProfilesLinkedinSpider
import colorama
from termcolor import cprint
import os
import sys
import json
from datetime import datetime

colorama.init()

whiteprint = lambda x, no_new_line=False: cprint(('' if no_new_line else '\n') + x, 'magenta')
checkprint = lambda x: cprint('\n%s' % x, 'green')
errorprint = lambda x: cprint('\n%s' % x, 'red')

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


def test_file(path):
    try:
        f = open(path, 'r+')
        f.close()
        return None
    except FileNotFoundError:
        return FileNotFoundError('O arquivo %s não existe neste diretório. Tente novamente.' % path)
    except PermissionError:
        return PermissionError('O arquivo %s está aberto. Feche-o e tente novamente.' % path)


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
        errorprint('O arquivo %s está mal formatado. Arrume-o e digite novamente\nAperte enter para continuar.' % config_path)
        input()
        return None

    for path in config_data['paths']:
        config_data['paths'][path] = '../' + config_data['paths'][path]
        error = test_file(config_data['paths'][path])
        if error is not None:
            if isinstance(error, FileNotFoundError):
                error_text = 'O arquivo %s não existe. Entre no config.json e corrija o erro.' % config_data['paths'][path]
            elif isinstance(error, PermissionError):
                error_text = 'O arquivo %s está aberto. Feche-o e tente novamente.' % config_data['paths'][path]
            errorprint('%s\nAperte enter para continuar.' % error_text)
            input()
            return None

    return config_data


def execute_crawling(username, password, max_page_requests, max_connection_pages, logs_path, input_excel_path,  cookies_path,  output_json_path):
    process = CrawlerProcess(
        settings={
            'DOWNLOAD_DELAY': 2
        },
    )
    process.crawl(
        ProfilesLinkedinSpider, 
        username=username,
        password=password,
        max_page_requests=max_page_requests, 
        max_connection_pages=max_connection_pages, 
        logs_path=logs_path,
        input_excel_path=input_excel_path, 
        cookies_path=cookies_path, 
        output_json_path=output_json_path
    )
    process.start()


if __name__ == '__main__':
    # https://stackoverflow.com/questions/24944558/pyinstaller-built-windows-exe-fails-with-multiprocessing
    freeze_support()

    config_exists = False
    input_excel_path = None
    output_json_path = None
    cookies_path = None
    logs_path = None
    max_page_requests = None
    max_connection_pages = None
    username = None
    password = None

    while not config_exists:

        # whiteprint('\nDigite o caminho do arquivo config.json (ou selecione o arquivo no Windows Explorer e arraste-o para esta janela): ')
        config_path = format_file_path('../config.json')

        error = test_file(config_path)
        
        if error is None:
            config_exists = True
            checkprint('config.json carregado!')

            config_data = read_config_file(config_path)

            if config_data is not None:
                input_excel_path = config_data['paths']['input']
                output_json_path = config_data['paths']['output_perfis']
                cookies_path = config_data['paths']['cookies']
                logs_path = config_data['paths']['logs']
                max_page_requests = config_data['config']['max_paginas_por_dia']
                max_connection_pages = config_data['config']['max_paginas_de_conexoes_por_perfil']
                username = config_data['login']['username']
                password = config_data['login']['password']

            else:
                config_exists = False

        else:
            errorprint('%s\nAperte enter para continuar' % str(error).replace('../', ''))
            input()

    fresh_cookies = False

    while not fresh_cookies:
        whiteprint('Lembrete: Você já inseriu os cookies da sua sessão atual do Chrome em "%s"? (s/n)' % cookies_path)
        if input().lower() == 's': fresh_cookies = True

    dont_stop = True

    while dont_stop:
        p = Process(
            target=execute_crawling, 
            args=(
                username,
                password, 
                max_page_requests, 
                max_connection_pages, 
                logs_path, 
                input_excel_path, 
                cookies_path, 
                output_json_path
            )
        )
        p.start()
        p.join()
        
        checkprint('Processo finalizado!\nAperte enter para abrir o Excel e o arquivo JSON.')
        input()

        os.system('start EXCEL.EXE "%s"' % input_excel_path)
        os.system('start Notepad.exe "%s"' % output_json_path)

        whiteprint('Quer realizar o crawl novamente? (s/n)')
        dont_stop = input() == 's'
        print()
