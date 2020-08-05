from scrapy import Request
from scrapy.exceptions import CloseSpider
from scrapy.spiders.init import InitSpider
import scrapy.http as Http
from scrapy.utils.spider import iterate_spider_output

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Color, Side, Alignment, Border

import colorama
from termcolor import cprint

from ..unicode_conversion import unicode_dict
from ..ParsingException import ParsingException

import os

import json
from math import ceil, floor
from datetime import datetime
from copy import deepcopy

SYSTEM_IS_WINDOWS = os.name != 'posix'

colorama.init()

if SYSTEM_IS_WINDOWS:
    def whiteprint(x): return cprint('\n%s' % x, 'magenta')
    def warnprint(x): return cprint('\nAviso: %s' % x, 'yellow')
    def checkprint(x): return cprint('\n%s' % x, 'green')
    def errorprint(x): return cprint('\nErro: %s' % x, 'red')
else:
    def whiteprint(x): return cprint('\n%s' % x, 'white')
    def warnprint(x): return whiteprint('🟡 %s' % x)
    def checkprint(x): return whiteprint('✅ %s' % x)
    def errorprint(x): return whiteprint('❌ %s' % x)

CELL_SIDE = Side(
    border_style='thin',
    color='000000'
)

CELL_BORDER = Border(
    top=CELL_SIDE,
    bottom=CELL_SIDE,
    right=CELL_SIDE,
    left=CELL_SIDE,
)

LEFT_CELL_ALIGNMENT = Alignment(
    vertical='center',
    horizontal='left',
    wrap_text=True
)

CENTER_CELL_ALIGNMENT = Alignment(
    vertical='center',
    horizontal='center',
    wrap_text=True
)

NORMAL_FONT_CELL = Font()

BIG_FONT_CELL = Font(
    size=18
)

LINKS_TABLE_STARTING_LINE = 8

class CompaniesLinkedinSpider(InitSpider):
    name = 'linkedin_companies'
    handle_httpstatus_list = [999]

    login_page = 'https://www.linkedin.com/uas/login'

    company_urls = []

    stored_employees_search_requests = []
    stored_profile_requests = []

    first_profile_requested = False

    request_retries = {}

    current_session_connection_pages_parsed_per_profile = {}

    def __init__(self, username, password, continue_previous_progress, max_page_requests, max_connection_pages, logs_path, cookies_path, input_excel_path, output_json_path, ensure_ascii):
        self.username = username
        self.password = password
        self.continue_previous_progress = continue_previous_progress
        self.current_date = get_date()
        self.max_page_requests = max_page_requests
        self.max_connection_pages = max_connection_pages
        self.ensure_ascii = ensure_ascii
        self.parse_cookies(cookies_path)
        self.input_excel_path = input_excel_path
        self.output_json_path = output_json_path
        self.output_json_data = read_json_file(output_json_path)
        if self.output_json_data is None: self.output_json_data = { 'empresas': [] }
        self.logs_path = logs_path
        self.logs_data = read_json_file(logs_path)
        if self.logs_data is None: self.logs_data = { 'logs': [] }
        self.setup_current_log()
        self.setup_accessed_pages()

    def init_request(self):
        # Obtém os dados do excel:
        self.read_excel()
        # Arruma os links que não começam com www:
        self.fix_links_without_www()
        # A partir dos dados do excel, carrega os links:
        self.get_links_from_workbook()
        # Verifica se há e remove links duplicados:
        self.check_for_duplicate_links()
        # Aplica estilo no excel:
        self.apply_links_sheet_style()
        # Carrega os requests iniciais:
        if self.load_initial_requests() is not None: return
        # Calcula limite de acessos:
        self.calculate_max_employees_search_pages_to_access()
        # Verifica se o limite de acessos não será ultrapassado:
        if self.verify_page_access_limit() is not None: return
        # Realiza o login:
        return self.attempt_login()

    def parse_cookies(self, cookies_path):
        cookies = read_json_file(cookies_path)
        self.chrome_cookies = {}
        if cookies is None: return
        for line in cookies.split('\n'):
            if not line.startswith('#'):
                self.chrome_cookies[line.split()[-2]] = line.split()[-1]

    def load_initial_requests(self):
        if self.continue_previous_progress:
            self.resume_progress_from_last_log()
            previous_progress = list(self.get_companies_with_progress_to_continue())
            self._postinit_reqs = self.load_employee_search_requests_from_previous_progress(previous_progress)
        else:
            self.reset_previous_progress()
            if self.verify_excel_links() is not None: return not None
            self._postinit_reqs = self.start_company_requests()
        return None

    def verify_excel_links(self):
        if len(self.company_urls) == 0:
            errorprint('Não há links de empresa no Excel.\n')
            return not None
        return None

    def get_companies_with_progress_to_continue(self):
        for company_log in self.current_log['dados_obtidos']:
            if company_log['ultima_pagina_de_busca_de_funcionarios_acessada'] < company_log['total_de_paginas_de_busca_de_funcionario']:
                yield company_log

    def load_employee_search_requests_from_previous_progress(self, previous_progress):
        for company_log in previous_progress:
            self.stored_employees_search_requests.extend(
                self.load_company_employee_search_pages_requests(company_log['company_id'])
            )
        return self.start_employees_search_requests()
                
    def load_company_employee_search_pages_requests(self, company_id):
        company = self.find_company_by_id(company_id)
        company_log = self.find_company_log_by_id(company_id)

        start = company_log['ultima_pagina_de_busca_de_funcionarios_acessada'] if self.continue_previous_progress else 0 

        totalEmployeeSearchPages = ceil(company['quantidade_funcionarios'] / 10)

        for page in range(start, totalEmployeeSearchPages):
            yield self.cookie_request(
                url='https://www.linkedin.com/search/results/people/?facetCurrentCompany=%%5B"%s"%%5D&page=%i' % (company_id, page + 1),
                callback=self.store_profile_requests_of_employee_searches,
                priority=(-page)
            )

    def start_employees_search_requests(self):
        self.stored_employees_search_requests.sort(key=lambda x: -x.priority)
        self.stored_employees_search_requests = \
            self.stored_employees_search_requests[:self.max_employees_search_pages]
        self.total_employees_left_to_access = 10 * len(self.stored_employees_search_requests) + self.current_session_profiles_parsed
        for request in self.stored_employees_search_requests:
            yield request

    def reset_previous_progress(self):
        for company_log in self.current_log['dados_obtidos']:
            company_log['ultima_pagina_de_busca_de_funcionarios_acessada'] = 0
        self.save_log()

    def verify_page_access_limit(self):
        if self.max_employees_search_pages == 0: 
            warnprint(
                'A quantidade de acessos diários ao Linkedin já chegou ao limite de %i páginas.\nPara alterar esse limite, entre em config.json.\nObs.: O recomendado para contas Premium Business é de, no máximo, 150 páginas por dia.\n' 
                % self.max_page_requests
            )
            return not None
        return None

    def setup_current_log(self):
        self.current_log = self.find_current_log()
        if self.current_log is None:
            self.current_log = {
                'data': self.current_date,
                'paginas_acessadas': {
                    'empresas': 0,
                    'pesquisa_de_funcionarios': 0,
                    'perfis': 0,
                    'pesquisa_de_conexoes': 0,
                    'total': 0
                },
                'dados_obtidos': []
            }

    def find_current_log(self):
        for log in self.logs_data['logs']:
            if log['data'] == self.current_date:
                return log
        return None

    def save_log(self):
        log = self.find_current_log()
        if log is None:
            self.logs_data['logs'].append(self.current_log)
        else:
            log.update(self.current_log)
        self.logs_data['logs'].sort(key=lambda x: x['data'], reverse=True)
        save_to_file(
            self.logs_path,
            json.dumps(
                self.logs_data, 
                indent=4, 
                ensure_ascii=self.ensure_ascii
            ),
            dont_print=True
        )

    def update_current_log(self):
        self.current_log.update({
            'paginas_acessadas': {
                'empresas': self.companies_parsed,
                'pesquisa_de_funcionarios': self.employee_searches_parsed,
                'perfis': self.profiles_parsed,
                'pesquisa_de_conexoes': self.connection_pages_parsed,
                'total': self.all_pages_parsed_count()
            }
        })
        self.save_log()

    def find_company_log_by_id(self, company_id):
        for savedCompany in self.current_log['dados_obtidos']:
            if savedCompany['company_id'] == company_id:
                return savedCompany
        return None

    def create_company_log_if_necessary(self, company):
        company_obtained_data = self.find_company_log_by_id(company['company_id'])
        if company_obtained_data is None:
            self.current_log['dados_obtidos'].append({
                'empresa': company['nome'],
                'company_id': company['company_id'],
                'ultima_pagina_de_busca_de_funcionarios_acessada': 0,
                'total_de_paginas_de_busca_de_funcionario': ceil(company['quantidade_funcionarios'] / 10)
            })
            self.save_log()

    def find_last_not_empty_log(self):
        self.logs_data['logs'].sort(key=lambda x: x['data'], reverse=True)
        for log in self.logs_data['logs']:
            if len(log['dados_obtidos']) > 0:
                return log
        return None

    def resume_progress_from_last_log(self):
        last_log = self.find_last_not_empty_log()
        self.current_log['dados_obtidos'] = deepcopy(last_log['dados_obtidos'])

    def all_pages_parsed_count(self):
        return (
            self.companies_parsed 
            + self.employee_searches_parsed 
            + self.profiles_parsed 
            + self.connection_pages_parsed
        )

    def mark_new_employee_search_page_as_accessed(self, company_id):
        company_log = self.find_company_log_by_id(company_id)
        company_log['ultima_pagina_de_busca_de_funcionarios_acessada'] += 1
        self.save_log()

    def setup_accessed_pages(self):
        self.companies_parsed = self.current_log['paginas_acessadas']['empresas']
        self.employee_searches_parsed = self.current_log['paginas_acessadas']['pesquisa_de_funcionarios']
        self.profiles_parsed = self.current_log['paginas_acessadas']['perfis']
        self.connection_pages_parsed = self.current_log['paginas_acessadas']['pesquisa_de_conexoes']
        if self.continue_previous_progress:
            self.current_session_companies_parsed = self.companies_parsed
            self.current_session_employee_searches_parsed = self.employee_searches_parsed
            self.current_session_profiles_parsed = self.profiles_parsed
            self.current_session_connection_pages_parsed = self.connection_pages_parsed
        else:
            self.current_session_companies_parsed = 0
            self.current_session_employee_searches_parsed = 0
            self.current_session_profiles_parsed = 0
            self.current_session_connection_pages_parsed = 0

    # Talvez seja interessante implementar headers também
    def cookie_request(self, url, priority=0, callback=None, cookies=None, meta=None, dont_filter=False):
        return Request(
            url=url,
            priority=priority,
            callback=callback,
            dont_filter=dont_filter,
            meta=meta,
            cookies=self.chrome_cookies
        )

    def read_excel(self):
        self.workbook = load_workbook(filename=self.input_excel_path)

    def attempt_login(self):
        return self.cookie_request(url=self.login_page, callback=self.login, dont_filter=True)

    def fix_links_without_www(self):
        links_sheet = self.workbook['Empresas']
        line = LINKS_TABLE_STARTING_LINE
        link = links_sheet['C%i' % line].value
        while link is not None:
            if ('www.linkedin' not in link) and ('linkedin' in link):
                novo_link = 'https://www.linkedin' + link.split('linkedin')[1]
                warnprint(
                    'O seguinte link não contém "www.linkedin.com": %s\nModificando-o para: %s' 
                    % (link, novo_link)
                )
                links_sheet['C%i' % line] = novo_link
            line += 1
            link = links_sheet['C%i' % line].value
        self.workbook.save(self.input_excel_path)

    def count_employees_with_url(self, company_id):
        count = 0
        for company in self.output_json_data['empresas']:
            if company['company_id'] == company_id:
                for employee in company['funcionarios']:
                    if ('url' in employee) and employee['url'] is not None:
                        count += 1
        return count

    def check_for_duplicate_links(self):
        for link in self.company_urls:
            while self.company_urls.count(link) > 1:
                self.company_urls.remove(link)
                warnprint('Há uma cópia de link: %s\n' % link)

    def get_links_from_workbook(self):
        links_sheet = self.workbook['Empresas']

        line = LINKS_TABLE_STARTING_LINE
        link = links_sheet['C%i' % line].value

        while link is not None:
            self.company_urls.append(link)

            line += 1
            link = links_sheet['C%i' % line].value

    def apply_links_sheet_style(self):
        self.apply_style_to_workbook_sheet(
            sheet=self.workbook['Empresas'], 
            verification_column='C', 
            starting_line=LINKS_TABLE_STARTING_LINE, 
            columns='BCD'
        )
        self.apply_style_to_workbook_sheet(
            sheet=self.workbook['Empresas'], 
            alignment=CENTER_CELL_ALIGNMENT,
            font=BIG_FONT_CELL, 
            verification_column='C', 
            starting_line=LINKS_TABLE_STARTING_LINE, 
            columns='B'
        )

    def apply_style_to_workbook_sheet(self, sheet, verification_column, starting_line, columns, alignment=LEFT_CELL_ALIGNMENT, border=CELL_BORDER, font=NORMAL_FONT_CELL):
        line = starting_line
        link = sheet['%s%i' % (verification_column, line)].value
        while link is not None:
            for column in columns:
                cell = sheet['%s%i' % (column, line)]
                cell.alignment = alignment
                cell.border = border
                cell.font = font
            line += 1
            link = sheet['%s%i' % (verification_column, line)].value
        self.workbook.save(self.input_excel_path)

    def find_company_by_id(self, company_id):
        for company in self.output_json_data['empresas']:
            if company['company_id'] == company_id:
                return company

    def find_user_by_id(self, user_id):
        for company in self.output_json_data['empresas']:
            for employee in company['funcionarios']:
                if ('user_id' in employee) and (employee['user_id'] == user_id):
                    return employee
        return None

    def refresh_workbook_company_data(self, company_id):
        company = self.find_company_by_id(company_id)
        company_log = self.find_company_log_by_id(company_id)

        links_sheet = self.workbook['Empresas']

        line = LINKS_TABLE_STARTING_LINE
        link = links_sheet['C%i' % line].value
        
        while link is not None:

            if link == company['url']:

                links_sheet['B%i' % line] = '%i/%i' \
                    % (
                        company_log['ultima_pagina_de_busca_de_funcionarios_acessada'], 
                        company_log['total_de_paginas_de_busca_de_funcionario']
                    )
                links_sheet['D%i' % line] = company['nome']

                self.workbook.save(self.input_excel_path)

            line += 1
            link = links_sheet['C%i' % line].value

    def login(self, response):
        return Http.FormRequest.from_response(
            response,
            formdata={
                'session_key': self.username,
                'session_password': self.password,
            },
            cookies=self.chrome_cookies,
            callback=self.check_login_response
        )

    def check_login_response(self, response):
        logged_in = False

        def loginerrorprint(x): return warnprint('Login falhou. %s Acesse o linkedin com essa conta para mais detalhes.\n' % x)

        if 'Your account has been restricted' in str(response.body):
            loginerrorprint('Conta bloqueada pelo Linkedin por muitas tentativas.')
        elif 'Let&#39;s do a quick security check' in str(response.body):
            loginerrorprint('Conta pede uma verificação de se é um robô.')
        elif 'The login attempt seems suspicious.' in str(response.body):
            loginerrorprint('Conta pede que seja copiado um código do email.')
        elif 'that&#39;s not the right password' in str(response.body):
            loginerrorprint('A senha está errada.\nVerifique se o usuário e senha estão corretos.')
        elif 'We’re unable to reach you' in str(response.body):
            loginerrorprint('O Linkedin pediu uma verificação de email.')
        else:
            logged_in = True
            checkprint('Login realizado. Vamos começar o crawling!\n')

        if logged_in:
            return self.initialized()
        else:
            return

    def start_requests(self):
        return iterate_spider_output(self.init_request())

    def start_company_requests(self):
        for url in self.company_urls:
            yield self.cookie_request(
                url=url,
                callback=self.store_employees_search_requests_preserving_priority
            )

    def store_employees_search_requests_preserving_priority(self, response):
        self.stored_employees_search_requests.extend(self.parse_company(response))
        if self.current_session_companies_parsed == len(self.company_urls):
            return self.start_employees_search_requests()

    def store_profile_requests_of_employee_searches(self, response):
        self.stored_profile_requests.extend(self.parse_employees_search(response))
        if (not self.first_profile_requested) and (len(self.stored_profile_requests) > 0):
            self.first_profile_requested = True
            return self.stored_profile_requests.pop(0)

    def calculate_max_employees_search_pages_to_access(self):
        self.max_employees_search_pages = max(
            floor(
                (
                    self.max_page_requests 
                    - self.current_log['paginas_acessadas']['total'] 
                    - (
                        0 if self.continue_previous_progress \
                        else len(self.company_urls)
                    )
                ) / (
                    11 + 10 * self.max_connection_pages
                )
            ),
            0
        )

    def profile_counter(self):
        self.current_session_profiles_parsed += 1
        return {
            'counter': '(%i/%i) ' % (self.current_session_profiles_parsed, self.total_employees_left_to_access),
            'new_request': self.stored_profile_requests.pop(0) if len(self.stored_profile_requests) > 0 else None
        }

    def get_company_included_array(self, response):
        body = str(response.body.decode('utf8'))

        birthIndex = body.rindex('&quot;videosTabVisible&quot;:')
        start = body[:birthIndex].rindex('<code ')
        end = body[start:].index('</code>') + start

        while (not body[start:end].startswith('{')) and start < end:
            start += 1

        while (not body[start:end].endswith('}')) and start < end:
            end -= 1

        if start >= end:
            errorprint(
                'ERRO em get_company_included_array: não foi possivel obter dados do usuário em %s' % response.url)
            return None

        # save_to_file(
        #     response.url.split('/')[4] + '.html',
        #     convert_unicode(body[start:end], unicode_dict)
        # )

        return parse_text_to_json(body[start:end], unicode_dict, 'aa.json')['included']

    def get_search_data(self, response):
        body = str(response.body.decode('utf8'))

        birthIndex = body.rindex('&quot;com.linkedin.voyager.search.BlendedSearchCluster&quot;')
        start = body[:birthIndex].rindex('<code ')
        end = body[start:].index('</code>') + start

        while (not body[start:end].startswith('{')) and start < end:
            start += 1

        while (not body[start:end].endswith('}')) and start < end:
            end -= 1

        if start >= end:
            errorprint(
                'ERRO em get_employees_data: não foi possivel obter dados do usuário em %s' % response.url)
            return None

        # save_to_file(
        #     'employee_search.json',
        #     convert_unicode(body[start:end], unicode_dict)
        # )

        return parse_text_to_json(body[start:end], unicode_dict, 'aa.json')

    def get_search_results(self, search_data):
        for blendedSearchCluster in search_data['data']['elements']:
            if blendedSearchCluster['type'] == 'SEARCH_HITS':
                return blendedSearchCluster['elements']

        return None

    def get_big_json_included_array(self, response):
        body = str(response.body.decode('utf8'))

        birthIndex = body.rindex(',{&quot;birthDateOn')
        start = body[:birthIndex].rindex('<code ')
        end = body[start:].index('</code>') + start

        while (not body[start:end].startswith('{')) and start < end:
            start += 1

        while (not body[start:end].endswith('}')) and start < end:
            end -= 1

        if start >= end:
            whiteprint('ERRO em get_big_json_included_array: não foi possivel obter dados do usuário em %s' % response.url)
            return None

        # save_to_file(
        #     response.url.split('/')[4] + '.html',
        #     convert_unicode(body[start:end], unicode_dict)
        # )

        return parse_text_to_json(body[start:end], unicode_dict, 'aa.json')['included']

    def get_following_json_dictionary(self, response):
        body = str(response.body.decode('utf8'))

        birthIndex = body.rindex('&quot;followersCount&quot;:')
        start = body[:birthIndex].rindex('<code ')
        end = body[start:].index('</code>') + start

        while (not body[start:end].startswith('{')) and start < end:
            start += 1

        while (not body[start:end].endswith('}')) and start < end:
            end -= 1

        if start >= end:
            whiteprint('ERRO em get_following_json_dictionary: não foi possivel obter dados do usuário em %s' % response.url)
            return None

        # save_to_file(
        #     response.url.split('/')[4] + '.html',
        #     convert_unicode(body[start:end], unicode_dict)
        # )

        return parse_text_to_json(body[start:end], unicode_dict, 'aa.json')

    def get_member_badges_json_dictionary(self, response):
        body = str(response.body.decode('utf8'))

        birthIndex = body.rindex('com.linkedin.voyager.identity.profile.MemberBadges')
        start = body[:birthIndex].rindex('<code ')
        end = body[start:].index('</code>') + start

        while (not body[start:end].startswith('{')) and start < end:
            start += 1

        while (not body[start:end].endswith('}')) and start < end:
            end -= 1

        if start >= end:
            whiteprint('ERRO em get_member_badges_json_dictionary: não foi possivel obter dados do usuário em %s' % response.url)
            return None

        # save_to_file(
        #     response.url.split('/')[4] + '.html',
        #     convert_unicode(body[start:end], unicode_dict)
        # )

        return parse_text_to_json(body[start:end], unicode_dict, 'aa.json')

    def get_object_by_type(self, included_array, obj_type):
        array = []
        for obj in included_array:
            if obj['$type'] == obj_type:
                array.append(obj)
        return array

    def get_object_by_user_id(self, array, user_id):
        for obj in array:
            if obj['entityUrn'].split(':')[-1] == user_id:
                return obj

    def compare_employees(self, employee1, employee2):
        if ('url' in employee1) and ('url' in employee2):
            return employee1['url'] == employee2['url']
        else:
            return (employee1['localizacao_atual'] == employee2['localizacao_atual']) \
                and (employee1['cargo_atual'] == employee2['cargo_atual'])

    def should_employee_replace(self, old_employee, new_employee):
        if ('url' in old_employee) and ('url' in new_employee) and (old_employee['url'] == new_employee['url']):
            return False
        if 'url' in new_employee:
            return True
        return False

    def insert_employee_if_necessary(self, company_id, new_employee):
        for company in self.output_json_data['empresas']:
            if company['company_id'] == company_id:

                for existing_employee in company['funcionarios']:

                    sameEmployees = self.compare_employees(existing_employee, new_employee)
                    shouldReplace = self.should_employee_replace(existing_employee, new_employee)

                    if sameEmployees:
                        if shouldReplace:
                            existing_employee.update(new_employee)
                        return

                company['funcionarios'].append(new_employee)
                return

    def reorder_user_dict(self, user_dict):
        keys = [
            "url",
            "dados_obtidos",
            "nome",
            "sobrenome",
            "user_id",
            "cargo_atual",
            "localizacao_atual",
            "foto_de_perfil",
            "plano_de_fundo",
            "sobre",
            "premium",
            "influenciador",
            "procura_emprego",
            "seguidores",
            "conexoes",
            "habilidades",
            "linguas",
            "cursos_feitos",
            "premios",
            "estudos",
            "experiencia_profissional",
            "voluntariado",
            "projetos"
        ]

        new_user_dict = {}

        for key in keys:
            new_user_dict[key] = deepcopy(user_dict[key])

        return new_user_dict


    def convert_date(self, date):
        return {
            'mes': date['month'] if 'month' in date else None,
            'ano': date['year'] if 'year' in date else None
        } if date is not None else None
            
    def convert_date_range(self, date_range):
        return {
            'inicio': self.convert_date((date_range['start']) if ('start' in date_range) else None),
            'fim': self.convert_date((date_range['end']) if ('end' in date_range) else None)
        } if date_range is not None else None

    def stringify_date(self, date):
        return '%04i-%02i' \
            % (
                0 if (date is None) or (date['ano'] is None) else date['ano'],
                0 if (date is None) or (date['mes'] is None) else date['mes']
            )

    def stringify_date_range(self, date_range):
        return self.stringify_date(None if date_range is None else date_range['inicio'])

    def format_connections(self, connections):
        return {
            'numero_exato': connections if connections != 500 else None,
            'minimo': connections if connections == 500 else None,
            'conexoes_obtidas': []
        }

    def get_picture_url(self, image_data):
        if image_data is None: return None
        start = image_data['rootUrl']
        end = sorted(image_data['artifacts'], key=lambda x: x['width'])[-1]['fileIdentifyingUrlPathSegment']
        return start + end

    def get_company_from_profile_url(self, url):
        for company in self.output_json_data['empresas']:
            for funcionario in company['funcionarios']:
                if ('url' in funcionario) and (url == funcionario['url']):
                    return company
        return None

    # Isso pode ser ativado quando a url não começa com www:
    def check_response_status(self, response):
        if response.status == 999:
            errorprint('Status 999. O Linkedin começou a restringir pedidos.\nO crawler será encerrado automaticamente.\n')
            raise CloseSpider('Spider encerrado manualmente')

    def parse_company(self, response):
        self.check_response_status(response)

        self.companies_parsed += 1
        self.update_current_log()
        
        self.current_session_companies_parsed += 1
        
        company_included = self.get_company_included_array(response)

        company_info = self.get_object_by_type(
            company_included, 'com.linkedin.voyager.organization.Company')[0]

        company_id = company_info['companyEmployeesSearchPageUrl'].split('=')[-1]

        company = self.find_company_by_id(company_id)

        if company is None:
            company = {
                'url': str(response.url),
                'dados_obtidos': False,
                'company_id': company_id,
                'nome': company_info['name'],
                'quantidade_funcionarios': company_info['staffCount'],
                'dados_de_funcionarios_obtidos': 0,
                'funcionarios': [],
            }
            self.output_json_data['empresas'].append(company)
        
        else:
            company.update({
                'quantidade_funcionarios': company_info['staffCount'],
            })
        
        self.create_company_log_if_necessary(company)

        self.refresh_workbook_company_data(company_id)

        checkprint('(%i/%i) Dados da página da empresa %s corretamente obtidos!\n' % (self.current_session_companies_parsed, len(self.company_urls), company_info['name']))
        
        return self.load_company_employee_search_pages_requests(company_id)

    def parse_employees_search(self, response):
        self.check_response_status(response)

        self.employee_searches_parsed += 1
        self.update_current_log()

        self.current_session_employee_searches_parsed += 1

        company_id = str(response.url).split('facetCurrentCompany=')[-1].split('&page')[0][6:-6]

        company = self.find_company_by_id(company_id)
        
        self.mark_new_employee_search_page_as_accessed(company_id)

        # save_to_file(
        #     response.url.split('/')[4] + '.html',
        #     response.body
        # )

        search_data = self.get_search_data(response)

        search_results = self.get_search_results(search_data)

        all_mini_profile_data = self.get_object_by_type(
            search_data['included'], 
            'com.linkedin.voyager.identity.shared.MiniProfile'
        )

        if search_results is not None:

            for employee_data in search_results:

                employee_user_id = employee_data['targetUrn'].split(':')[-1]

                employee_mini_profile_data = self.get_object_by_user_id(all_mini_profile_data, employee_user_id)

                employee_url = employee_data['navigationUrl']

                user_data = {
                    'user_id': employee_user_id,
                    'localizacao_atual': employee_data['subline']['text'] if 'subline' in employee_data else None,
                    'cargo_atual': employee_data['headline']['text'] if 'headline' in employee_data else None,
                    'foto_de_perfil': self.get_picture_url(employee_mini_profile_data['picture']),
                    'dados_obtidos': False
                }

                if '/in/UNKNOWN' in employee_url:
                    counter = self.profile_counter()
                    if counter['new_request'] is not None:
                        yield counter['new_request']
                    whiteprint(
                        '%sNão foi possível obter url do funcionário da empresa %s. Pessoa fora de sua rede.\n'
                        % (counter['counter'], company['nome'])
                    )

                else:

                    user_data.update({
                        'url': employee_url
                    })

                    yield self.cookie_request(
                        url=employee_url,
                        callback=self.parse_profile
                    )

                self.insert_employee_if_necessary(company_id, user_data)

            company['dados_de_funcionarios_obtidos'] = self.count_employees_with_url(company_id)

            save_to_file(
                self.output_json_path,
                json.dumps(
                    self.output_json_data, 
                    indent=4, 
                    ensure_ascii=self.ensure_ascii
                ),
                dont_print=True
            )

            self.refresh_workbook_company_data(company_id)

        else:
            errorprint('Erro no parsing de lista de funcionários.\n')

    def parse_profile(self, response):
        self.check_response_status(response)

        self.profiles_parsed += 1
        self.update_current_log()

        user_dict = {
            'url': response.url,
            'dados_obtidos': False
        }

        company = self.get_company_from_profile_url(response.url)

        # Se página não tiver a seguinte string, ela provavelmente foi carregada errada:
        if 'linkedin.com/in/' not in str(response.url):
            counter = self.profile_counter()
            if counter['new_request'] is not None:
                yield counter['new_request']
            errorprint('%sEste não é um link de um perfil: %s\n' % (counter['counter'], response.url))

        # Se página não tiver a seguinte string, ela provavelmente
        # foi carregada errada, ou não é uma página válida:
        elif '{&quot;birthDateOn' not in str(response.body):
            retries = 0
            if str(response.url) in list(self.request_retries.keys()):
                retries = self.request_retries[str(response.url)]
            self.request_retries[str(response.url)] = retries + 1

            if retries < 1:
                counter = self.profile_counter()
                if counter['new_request'] is not None:
                    yield counter['new_request']
                warnprint(
                    '%sErro no parsing de %s\nAdicionando novamente à fila de links...\n' 
                    % (counter['counter'], response.url)
                )
                self.total_employees_left_to_access += 1
                yield self.cookie_request(
                    url=response.url, 
                    callback=self.parse_profile, 
                    dont_filter=True
                )
                return None
            else:
                counter = self.profile_counter()
                if counter['new_request'] is not None:
                    yield counter['new_request']
                errorprint('%sEste provavelmente não é um link de um perfil: %s' % (counter['counter'], response.url))

        else:

            try:

                # save_to_file(
                #     response.url.split('/')[4] + '.html',
                #     response.body
                # )

                included_array = self.get_big_json_included_array(response)

                if included_array is None:
                    raise ParsingException('Erro com included_array')

                user_data = self.get_object_by_type(
                    included_array, 'com.linkedin.voyager.dash.identity.profile.Profile')[0]
                education_data = self.get_object_by_type(
                    included_array, 'com.linkedin.voyager.dash.identity.profile.Education')
                positions_data = self.get_object_by_type(
                    included_array, 'com.linkedin.voyager.dash.identity.profile.Position')
                volunteer_data = self.get_object_by_type(
                    included_array, 'com.linkedin.voyager.dash.identity.profile.VolunteerExperience')
                skills_data = self.get_object_by_type(
                    included_array, 'com.linkedin.voyager.dash.identity.profile.Skill')
                honors_data = self.get_object_by_type(
                    included_array, 'com.linkedin.voyager.dash.identity.profile.Honor')
                projects_data = self.get_object_by_type(
                    included_array, 'com.linkedin.voyager.dash.identity.profile.Project')
                courses_data = self.get_object_by_type(
                    included_array, 'com.linkedin.voyager.dash.identity.profile.Course')
                languages_data = self.get_object_by_type(
                    included_array, 'com.linkedin.voyager.dash.identity.profile.Language')

                following_json = self.get_following_json_dictionary(response)

                if following_json is None:
                    raise ParsingException('Erro com following_json')

                member_badges_json = self.get_member_badges_json_dictionary(response)

                if member_badges_json is None:
                    raise ParsingException('Erro com member_badges_json')

                user_id = member_badges_json['data']['entityUrn'].split(':')[-1]

                user_dict.update({
                    'nome': user_data['firstName'] if 'firstName' in user_data else None,
                    'sobrenome': user_data['lastName'] if 'lastName' in user_data else None,
                    'user_id': user_id,
                    'cargo_atual': user_data['headline'] if 'headline' in user_data else None,
                    'localizacao_atual': user_data['locationName'] if 'locationName' in user_data else None,
                    'foto_de_perfil': self.get_picture_url(
                        None if ('profilePicture' not in user_data) or (user_data['profilePicture'] is None) \
                        else user_data['profilePicture']['displayImageReference']['vectorImage']
                    ),
                    'plano_de_fundo': self.get_picture_url(
                        None if ('backgroundPicture' not in user_data) or (user_data['backgroundPicture'] is None) \
                        else user_data['backgroundPicture']['displayImageReference']['vectorImage']
                    ),
                    'sobre': user_data['summary'] if 'summary' in user_data else None,
                    'premium': user_data['premium'] if 'premium' in user_data else None,
                    'influenciador': user_data['influencer'] if 'influencer' in user_data else None,
                    'procura_emprego': member_badges_json['data']['jobSeeker'],
                    'seguidores': following_json['data']['followersCount'],
                    'conexoes': self.format_connections(following_json['data']['connectionsCount']),
                    'habilidades': [skill['name'] for skill in skills_data],
                    'linguas': [language['name'] for language in languages_data],
                    'cursos_feitos': [course['name'] for course in courses_data],
                    'premios': sorted(
                        [
                            {
                                'nome': honor['title'] if 'title' in honor else None,
                                'instituicao': honor['issuer'] if 'issuer' in honor else None,
                                'descricao': honor['description'] if 'description' in honor else None,
                                'data': self.convert_date(
                                    honor['issuedOn'] if 'issuedOn' in honor else None,
                                )
                            } for honor in honors_data
                        ],
                        key=lambda x: self.stringify_date(x['data'])
                    ),
                    'estudos': sorted(
                        [
                            {
                                'instituicao': experience['schoolName'] if 'schoolName' in experience else None,
                                'formacao': experience['fieldOfStudy'] if 'fieldOfStudy' in experience else None,
                                'tilulo_obtido': experience['degreeName'] if 'degreeName' in experience else None,
                                'descricao': experience['description'] if 'description' in experience else None,
                                'periodo': self.convert_date_range(
                                    experience['dateRange'] if 'dateRange' in experience else None
                                ),
                            } for experience in education_data
                        ],
                        key=lambda x: self.stringify_date_range(x['periodo'])
                    ),
                    'experiencia_profissional': sorted(
                        [
                            {
                                'instituicao': experience['companyName'] if 'companyName' in experience else None,
                                'cargo': experience['title'] if 'title' in experience else None,
                                'descricao': experience['description'] if 'description' in experience else None,
                                'periodo': self.convert_date_range(
                                    experience['dateRange'] if 'dateRange' in experience else None
                                ),
                            } for experience in positions_data
                        ],
                        key=lambda x: self.stringify_date_range(x['periodo'])
                    ),
                    'voluntariado': sorted(
                        [
                            {
                                'instituicao': experience['companyName'] if 'companyName' in experience else None,
                                'papel': experience['role'] if 'role' in experience else None,
                                'causa': experience['cause'] if 'cause' in experience else None,
                                'descricao': experience['description'] if 'description' in experience else None,
                                'periodo': self.convert_date_range(
                                    experience['dateRange'] if 'dateRange' in experience else None
                                ),
                            } for experience in volunteer_data
                        ],
                        key=lambda x: self.stringify_date_range(x['periodo'])
                    ),
                    'projetos': sorted(
                        [
                            {
                                'titulo': project['title'],
                                'url': project['url'],
                                'descricao': project['description'],
                                'periodo': self.convert_date_range(
                                    project['dateRange']
                                )
                            } for project in projects_data
                        ],
                        key=lambda x: self.stringify_date_range(x['periodo'])
                    ),
                    'dados_obtidos': True
                })

                existing_employee = self.find_user_by_id(user_id)

                if existing_employee is None:
                    existing_employee = {}
                    company['funcionarios'].append(existing_employee)

                existing_employee.update(user_dict)

                existing_employee = self.reorder_user_dict(existing_employee)

                self.current_session_connection_pages_parsed_per_profile[user_id] = 0

                for page in range(self.max_connection_pages):
                    yield self.cookie_request(
                        url='https://www.linkedin.com/search/results/people/?facetConnectionOf=%%5B%%22%s%%22%%5D&page=%i' \
                            % (user_id, page + 1),
                        callback=self.parse_connections_page
                    )

                if self.max_connection_pages == 0:
                    counter = self.profile_counter()
                    if counter['new_request'] is not None:
                        yield counter['new_request']
                    checkprint('%sParsing corretamente realizado em %s da empresa %s\n' % (counter['counter'], response.url, company['nome']))

            except ParsingException:
                counter = self.profile_counter()
                if counter['new_request'] is not None:
                    yield counter['new_request']
                errorprint('%sErro no parsing de %s da empresa %s\n' % (counter['counter'], response.url, company['nome']))
            except Exception as e:
                counter = self.profile_counter()
                if counter['new_request'] is not None:
                    yield counter['new_request']
                errorprint('%sErro grave no parsing de %s da empresa %s: %s\n' % (counter['counter'], response.url, company['nome'], e))

        save_to_file(
            self.output_json_path,
            json.dumps(
                self.output_json_data, 
                indent=4, 
                ensure_ascii=self.ensure_ascii
            ),
            dont_print=True
        )

    def parse_connections_page(self, response):
        self.check_response_status(response)

        self.connection_pages_parsed += 1
        self.update_current_log()

        self.current_session_connection_pages_parsed += 1

        user_id = str(response.url).split('facetConnectionOf=')[-1].split('&page')[0][6:-6]

        self.current_session_connection_pages_parsed_per_profile[user_id] += 1

        profile = self.find_user_by_id(user_id)
        company = self.get_company_from_profile_url(profile['url'])

        # save_to_file(
        #     response.url.split('/')[4] + '.html',
        #     response.body
        # )

        search_data = self.get_search_data(response)

        search_results = self.get_search_results(search_data)
        
        all_member_badges = self.get_object_by_type(
            search_data['included'], 
            'com.linkedin.voyager.identity.profile.MemberBadges'
        )
        all_mini_profile_data = self.get_object_by_type(
            search_data['included'], 
            'com.linkedin.voyager.identity.shared.MiniProfile'
        )

        if search_results is not None:

            for connection in search_results:

                connection_user_id = connection['targetUrn'].split(':')[-1]

                connection_member_badges = self.get_object_by_user_id(all_member_badges, connection_user_id)
                connection_mini_profile_data = self.get_object_by_user_id(all_mini_profile_data, connection_user_id)

                connection_data = {
                    'url': connection['navigationUrl'],
                    'nome': connection_mini_profile_data['firstName'],
                    'sobrenome': connection_mini_profile_data['lastName'],
                    'user_id': connection_user_id,
                    'cargo_atual': connection_mini_profile_data['occupation'],
                    'localizacao_atual': connection['subline']['text'],
                    'foto_de_perfil': self.get_picture_url(connection_mini_profile_data['picture']),
                    'plano_de_fundo': self.get_picture_url(connection_mini_profile_data['backgroundImage']),
                    'premium': connection_member_badges['premium'],
                    'influenciador': connection_member_badges['influencer'],
                    'procura_emprego': connection_member_badges['jobSeeker']
                }

                found = False
                    
                for existing_connection in profile['conexoes']['conexoes_obtidas']:
                    if existing_connection['url'] == connection_data['url']:
                        found = True
                        existing_connection.update(connection_data)

                if not found:
                    profile['conexoes']['conexoes_obtidas'].append(connection_data)

            save_to_file(
                self.output_json_path,
                json.dumps(
                    self.output_json_data, 
                    indent=4, 
                    ensure_ascii=self.ensure_ascii
                ),
                dont_print=True
            )

            if self.current_session_connection_pages_parsed_per_profile[user_id] == self.max_connection_pages:
                counter = self.profile_counter()
                if counter['new_request'] is not None:
                    yield counter['new_request']
                checkprint('%sParsing corretamente realizado em %s da empresa %s\n' % (counter['counter'], profile['url'], company['nome']))

        else:
            counter = self.profile_counter()
            if counter['new_request'] is not None:
                yield counter['new_request']
            errorprint('%sErro no parsing de lista de conexões de %s da empresa %s\n' % (counter['counter'], profile['url'], company['nome']))


def get_date():
    now = datetime.now()
    return now.strftime('%Y-%m-%d')


def parse_text_to_json(text, replacements, filename):
    try:
        text = convert_unicode(text, replacements)
        return json.loads(text)
    except Exception:
        # save_to_file(
        #     filename,
        #     text
        # )
        return None


def convert_unicode(text, replacements):
    try:
        text = str(text)
        for unicode_char in list(replacements.keys()):
            for type in list(replacements[unicode_char].keys()):
                for element in replacements[unicode_char][type]:
                    text = text.replace(str(element), str(unicode_char))
    except Exception:
        errorprint('convert_unicode: não foi possível converter os caracteres unicode.\n')
    return text


def read_json_file(path):
    try:
        f = open(path, 'r+')
        data = json.loads(f.read())
        f.close()
        return data
    except json.decoder.JSONDecodeError:
        return None


def save_to_file(filename, element, dont_print=False):
    with open(filename, 'wb') as f:
        f.write(str.encode(str(element)))
    if not dont_print: whiteprint('\n💽 Texto salvo como %s\n' % filename)
