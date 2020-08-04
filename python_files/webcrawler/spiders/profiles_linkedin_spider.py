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

class ProfilesLinkedinSpider(InitSpider):
    name = 'linkedin_profiles'
    handle_httpstatus_list = [999]

    login_page = 'https://www.linkedin.com/uas/login'

    profiles_urls = []

    stored_employees_search_requests = []

    current_session_companies_parsed = 0
    current_session_employee_searches_parsed = 0
    current_session_profiles_parsed = 0
    current_session_connection_pages_parsed = 0

    request_retries = {}

    def __init__(self, username, password, max_page_requests, max_connection_pages, logs_path, cookies_path, input_excel_path, output_json_path):
        self.username = username
        self.password = password
        self.current_date = get_date()
        self.max_page_requests = max_page_requests
        self.max_connection_pages = max_connection_pages
        self.parse_cookies(cookies_path)
        self.input_excel_path = input_excel_path
        self.output_json_path = output_json_path
        self.output_json_data = read_json_file(output_json_path)
        if self.output_json_data is None: self.output_json_data = { 'perfis': [] }
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
        # Calcula limite de acessos:
        self.calculate_max_profile_pages_to_access()
        # Verifica se o limite de acessos não será ultrapassado:
        if self.verify_page_access_limit() is not None: return
        # Carrega os requests iniciais:
        if self.load_initial_requests() is not None: return
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
        if self.verify_excel_links() is not None: return not None
        self._postinit_reqs = list(self.start_profiles_requests())[:self.max_profile_pages]
        return None

    def verify_excel_links(self):
        if len(self.profiles_urls) == 0:
            errorprint('Não há links de empresa no Excel.\n')
            return not None
        return None

    def verify_page_access_limit(self):
        if self.max_profile_pages == 0: 
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
            json.dumps(self.logs_data, indent=4),
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

    def all_pages_parsed_count(self):
        return (
            self.companies_parsed 
            + self.employee_searches_parsed 
            + self.profiles_parsed 
            + self.connection_pages_parsed
        )

    def setup_accessed_pages(self):
        self.companies_parsed = self.current_log['paginas_acessadas']['empresas']
        self.employee_searches_parsed = self.current_log['paginas_acessadas']['pesquisa_de_funcionarios']
        self.profiles_parsed = self.current_log['paginas_acessadas']['perfis']
        self.connection_pages_parsed = self.current_log['paginas_acessadas']['pesquisa_de_conexoes']

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
        links_sheet = self.workbook['Perfis']
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

    def check_for_duplicate_links(self):
        for link in self.profiles_urls:
            while self.profiles_urls.count(link) > 1:
                self.profiles_urls.remove(link)
                warnprint('Há uma cópia de link: %s\n' % link)

    def get_links_from_workbook(self):
        links_sheet = self.workbook['Perfis']

        line = LINKS_TABLE_STARTING_LINE
        link = links_sheet['C%i' % line].value

        while link is not None:
            self.profiles_urls.append(link)

            line += 1
            link = links_sheet['C%i' % line].value

    def apply_links_sheet_style(self):
        self.apply_style_to_workbook_sheet(
            sheet=self.workbook['Perfis'], 
            verification_column='C', 
            starting_line=LINKS_TABLE_STARTING_LINE, 
            columns='BCD'
        )
        self.apply_style_to_workbook_sheet(
            sheet=self.workbook['Perfis'], 
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

    def refresh_workbook_profiles_data(self, user_data):
        links_sheet = self.workbook['Perfis']

        line = LINKS_TABLE_STARTING_LINE
        link = links_sheet['C%i' % line].value
        
        while link is not None:

            if link == user_data['url']:

                links_sheet['B%i' % line] = 'Sim' if user_data['dados_obtidos'] else 'Não'
                links_sheet['D%i' % line] = user_data['nome'] if 'nome' in user_data else ''

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

    def start_profiles_requests(self):
        for url in self.profiles_urls:
            yield self.cookie_request(
                url=url,
                callback=self.parse_profile
            )

    def calculate_max_profile_pages_to_access(self):
        self.max_profile_pages = max(
            floor(
                (
                    self.max_page_requests 
                    - self.current_log['paginas_acessadas']['total']
                ) / (
                    1 + self.max_connection_pages
                )
            ),
            0
        )

    def get_employees_search_elements(self, response):

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

        blendedSearchClusterCollection = parse_text_to_json(body[start:end], unicode_dict, 'aa.json')['data']['elements']

        for blendedSearchCluster in blendedSearchClusterCollection:
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

    def format_conections(self, connections):
        return {
            'numero_exato': connections if connections != 500 else None,
            'minimo': connections if connections == 500 else None
        }

    # Isso pode ser ativado quando a url não começa com www:
    def check_response_status(self, response):
        if response.status == 999:
            errorprint('Status 999. O Linkedin começou a restringir pedidos.\nO crawler será encerrado automaticamente.\n')
            raise CloseSpider('Spider encerrado manualmente')

    def parse_profile(self, response):
        self.check_response_status(response)

        self.profiles_parsed += 1
        self.update_current_log()

        self.current_session_profiles_parsed += 1

        user_dict = {
            'url': response.url,
            'dados_obtidos': False
        }

        counter = '(%i/%i) ' % (self.current_session_profiles_parsed, len(self.profiles_urls))

        # Se página não tiver a seguinte string, ela provavelmente foi carregada errada:
        if 'linkedin.com/in/' not in str(response.url):
            errorprint('%sEste não é um link de um perfil: %s\n' %
                       (counter, response.url))

        # Se página não tiver a seguinte string, ela provavelmente
        # foi carregada errada, ou não é uma página válida:
        elif '{&quot;birthDateOn' not in str(response.body):
            retries = 0
            if str(response.url) in list(self.request_retries.keys()):
                retries = self.request_retries[str(response.url)]
            self.request_retries[str(response.url)] = retries + 1

            if retries < 1:
                warnprint(
                    '%sErro no parsing de %s\nAdicionando novamente à fila de links...\n' 
                    % (counter, response.url)
                )
                self.profiles_urls.append(response.url)
                return self.cookie_request(
                    url=response.url, 
                    callback=self.parse_profile, 
                    dont_filter=True
                )
            else:
                errorprint('%sEste provavelmente não é um link de um perfil: %s' % (counter, response.url))

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

                user_dict.update({
                    'nome': user_data['firstName'] if 'firstName' in user_data else None,
                    'sobrenome': user_data['lastName'] if 'lastName' in user_data else None,
                    'user_id': member_badges_json['data']['entityUrn'].split(':')[-1],
                    'cargo_atual': user_data['headline'] if 'headline' in user_data else None,
                    'localizacao_atual': user_data['locationName'] if 'locationName' in user_data else None,
                    'sobre': user_data['summary'] if 'summary' in user_data else None,
                    'premium': user_data['premium'] if 'premium' in user_data else None,
                    'influenciador': user_data['influencer'] if 'influencer' in user_data else None,
                    'procura_emprego': member_badges_json['data']['jobSeeker'],
                    'seguidores': following_json['data']['followersCount'],
                    'conexoes': self.format_conections(following_json['data']['connectionsCount']),
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

                updated = False

                for profile in self.output_json_data['perfis']:
                    if profile['url'] == response.url:
                        updated = True
                        profile.update(user_dict)

                if not updated:
                    self.output_json_data['perfis'].append(user_dict)

                save_to_file(
                    self.output_json_path,
                    json.dumps(self.output_json_data, indent=4),
                    dont_print=True
                )

                checkprint('%sParsing corretamente realizado em %s\n' % (counter, response.url))

            except ParsingException:
                errorprint('%sErro no parsing de %s\n' % (counter, response.url))
            except Exception as e:
                errorprint('%sErro grave no parsing de %s: %s\n' % (counter, response.url, e))

        self.refresh_workbook_profiles_data(user_dict)


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
