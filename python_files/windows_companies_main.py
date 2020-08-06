from scrapy.crawler import CrawlerProcess
from multiprocessing import Process, freeze_support
from webcrawler.spiders.companies_linkedin_spider import CompaniesLinkedinSpider
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
        f = open(path, 'r+', encoding="utf8")
        f.close()
        return None
    except FileNotFoundError:
        return FileNotFoundError('O arquivo %s não existe neste diretório. Tente novamente.' % path)
    except PermissionError:
        return PermissionError('O arquivo %s está aberto. Feche-o e tente novamente.' % path)


def parse_json_file(path):
    try:
        f = open(path, 'r+', encoding="utf8")
        data = json.loads(f.read())
        f.close()
        return data
    except json.decoder.JSONDecodeError:
        return None


def check_open_paths(config_data):
    for path in config_data['paths']:
        config_data['paths'][path] = config_data['paths'][path]
        error = test_file(config_data['paths'][path])
        if error is not None:
            if isinstance(error, FileNotFoundError):
                error_text = 'O arquivo %s não existe. Entre no config.json e corrija o erro.' % config_data['paths'][path]
            elif isinstance(error, PermissionError):
                error_text = 'O arquivo %s está aberto. Feche-o e tente novamente.' % config_data['paths'][path]
            errorprint('%s\nAperte enter para continuar.' % error_text)
            input()
            return True
    return False


def read_config_file(config_path):

    config_data = parse_json_file(config_path)

    if config_data is None:
        errorprint('O arquivo %s está mal formatado. Arrume-o e digite novamente\nAperte enter para continuar.' % config_path)
        input()
        return None

    if check_open_paths(config_data): return None

    return config_data


def find_last_not_empty_log(logs_data):
    if (logs_data is None) or ('logs' not in logs_data): return None
    logs_data['logs'].sort(key=lambda x: x['data'], reverse=True)
    for log in logs_data['logs']:
        if len(log['dados_obtidos']) > 0:
            return log
    return None


def get_date():
    now = datetime.now()
    return now.strftime('%Y-%m-%d')


def get_companies_with_progress_to_continue(last_log):
    for company_log in last_log['dados_obtidos']:
        if company_log['ultima_pagina_de_busca_de_funcionarios_acessada'] < company_log['total_de_paginas_de_busca_de_funcionario']:
            yield company_log


def should_continue_previous_progress(logs_path, output_json_path):
    logs_data = parse_json_file(logs_path)
    last_log = find_last_not_empty_log(logs_data)
    if last_log is None: return False
    previous_progress = list(get_companies_with_progress_to_continue(last_log))
    if len(previous_progress) > 0:
        whiteprint(
            'Foi encontrado progresso anterior, de %s, com as seguintes empresas:' % last_log['data']
        )
        for company_log in previous_progress:
            whiteprint(
                ' - %s (%i páginas de funcionários restantes)' 
                % (
                    company_log['empresa'], 
                    company_log['total_de_paginas_de_busca_de_funcionario'] - company_log['ultima_pagina_de_busca_de_funcionarios_acessada']
                ), 
                no_new_line=True
            )
        whiteprint(
            '\nVocê gostaria de continuar onde parou? (s/n)\nDetalhe: O programa apenas funcionará corretamente se %s não tiver sido alterado.' % output_json_path, 
            no_new_line=True
        )
        if input() == 's':
            return True
        else:
            if get_date() == last_log['data']:
                whiteprint(
                    '\nVocê eliminará o registro da última página em que você parou, mas não seus dados obtidos em %s. Você tem certeza? (s/n)' % output_json_path, 
                    no_new_line=True
                )
                if input() != 's':
                    return True
    return False


def execute_crawling(
        username, password, continue_previous_progress, max_page_requests, 
        max_connection_pages, get_connection_data_from_profiles_with_3rd_or_higher_degree_connection, 
        logs_path, input_excel_path,  cookies_path,  output_json_path, ensure_ascii
    ):

    process = CrawlerProcess(
        settings={
            'DOWNLOAD_DELAY': 2
        },
    )
    process.crawl(
        CompaniesLinkedinSpider, 
        username=username,
        password=password,
        continue_previous_progress=continue_previous_progress,
        max_page_requests=max_page_requests, 
        max_connection_pages=max_connection_pages, 
        get_connection_data_from_profiles_with_3rd_or_higher_degree_connection=get_connection_data_from_profiles_with_3rd_or_higher_degree_connection,
        logs_path=logs_path,
        input_excel_path=input_excel_path, 
        cookies_path=cookies_path, 
        output_json_path=output_json_path,
        ensure_ascii=ensure_ascii
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
    get_connection_data_from_profiles_with_3rd_or_higher_degree_connection = None
    ensure_ascii = None
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
                output_json_path = config_data['paths']['output_empresas']
                cookies_path = config_data['paths']['cookies']
                logs_path = config_data['paths']['logs']
                max_page_requests = config_data['config']['max_paginas_por_dia']
                max_connection_pages = config_data['config']['max_paginas_de_conexoes_por_perfil']
                get_connection_data_from_profiles_with_3rd_or_higher_degree_connection = \
                    config_data['config']['obter_conexoes_de_perfis_com_grau_de_conexao_maior_que_2']
                ensure_ascii = not config_data['config']['permitir_caracteres_nao_ascii_no_output']
                username = config_data['login']['username']
                password = config_data['login']['password']

            else:
                config_exists = False

        else:
            errorprint('%s\nAperte enter para continuar' % error)
            input()

    fresh_cookies = False

    while not fresh_cookies:
        whiteprint('Lembrete: Você já inseriu os cookies da sua sessão atual do Chrome em "%s"? (s/n)' % cookies_path)
        if input().lower() == 's': fresh_cookies = True

    dont_stop = True

    while dont_stop:

        while check_open_paths(config_data):
            true = True
        
        p = Process(
            target=execute_crawling, 
            args=(
                username,
                password, 
                should_continue_previous_progress(logs_path, output_json_path),
                max_page_requests, 
                max_connection_pages, 
                get_connection_data_from_profiles_with_3rd_or_higher_degree_connection,
                logs_path, 
                input_excel_path, 
                cookies_path, 
                output_json_path,
                ensure_ascii
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
