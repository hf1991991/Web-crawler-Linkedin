from scrapy import Request
from scrapy.exceptions import CloseSpider
from scrapy.spiders import CrawlSpider, Rule
from scrapy.spiders.init import InitSpider
import scrapy.http as Http
from scrapy.utils.spider import iterate_spider_output

import colorama
from termcolor import cprint

from ..unicode_conversion import unicode_dict

import os

import json
from math import ceil

SYSTEM_IS_WINDOWS = os.name != 'posix'

colorama.init()

if SYSTEM_IS_WINDOWS:
    def whiteprint(x): return cprint(x, 'magenta')
    def warnprint(x): return cprint('Aviso: %s' % x, 'yellow')
    def checkprint(x): return cprint(x, 'green')
    def errorprint(x): return cprint('Erro: %s' % x, 'red')
else:
    def whiteprint(x): return cprint(x, 'white')
    def warnprint(x): return whiteprint("🟡 %s" % x)
    def checkprint(x): return whiteprint("✅ %s" % x)
    def errorprint(x): return whiteprint("❌ %s" % x)

class LinkedinSpider(InitSpider):
    name = "linkedin_employees"
    handle_httpstatus_list = [999]

    parsed_profile_urls = []

    request_retries = {}

    login_page = 'https://www.linkedin.com/uas/login'

    def __init__(self, username, password, employees_json_path=None, cookies_path=None):
        self.username = username
        self.password = password
        self.parse_cookies(cookies_path)
        self.employees_json_path = employees_json_path
        self.employees_json_data = read_json_file(employees_json_path)
        self.profile_requests = list(self.load_profile_requests())
        self.profile_requests.sort(key=lambda elem: -elem.priority)

    def init_request(self):
        if self.check_profile_requests_size() is not None: return
        return self.attempt_login()

    def load_profile_requests(self):
        for company in self.employees_json_data['empresas']:
            priority = 0
            for employee in company['funcionarios']:
                if ('url' in employee) and (not employee['dados_obtidos']):
                    yield self.cookie_request(
                        url=employee['url'],
                        callback=self.parse_profile,
                        priority=priority
                    )
                    priority -= 1

    def parse_cookies(self, cookies_path):
        cookies = read_json_file(cookies_path)
        self.chrome_cookies = {}
        if cookies is None: return
        for line in cookies.split('\n'):
            if not line.startswith('#'):
                self.chrome_cookies[line.split()[-2]] = line.split()[-1]

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

    def check_profile_requests_size(self):
        if len(self.profile_requests) == 0:
            print()
            warnprint('Não há urls de perfis de funcionários para serem pesquisados. Tente rodar o webcrawler de empresas novamente.\n')
            return not None
        return None

    def attempt_login(self):
        return self.cookie_request(url=self.login_page, callback=self.login, dont_filter=True)

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

        print()

        if "Your account has been restricted" in str(response.body):
            loginerrorprint('Conta bloqueada pelo Linkedin por muitas tentativas.')
        elif "Let&#39;s do a quick security check" in str(response.body):
            loginerrorprint("Conta pede uma verificação de se é um robô.")
        elif "The login attempt seems suspicious." in str(response.body):
            loginerrorprint("Conta pede que seja copiado um código do email.")
        elif "that&#39;s not the right password" in str(response.body):
            loginerrorprint("A senha está errada.\nVerifique se o usuário e senha estão corretos.")
        elif "We’re unable to reach you" in str(response.body):
            loginerrorprint('O Linkedin pediu uma verificação de email.')
        else:
            logged_in = True
            checkprint("Login realizado. Vamos começar o crawling!\n")

        if logged_in:
            return self.initialized()
        else:
            return

    def start_requests(self):
        self._postinit_reqs = self.start_url_requests()
        return iterate_spider_output(self.init_request())

    def start_url_requests(self):
        for request in self.profile_requests:
            yield request

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

        return parse_text_to_json(body[start:end], unicode_dict, 'aa.json')["included"]

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
            whiteprint(
                'ERRO em get_following_json_dictionary: não foi possivel obter dados do usuário em %s' % response.url)
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

    def format_conections(self, connections):
        return {
            'numero_exato': connections if connections != 500 else None,
            'minimo': connections if connections == 500 else None
        }

    def get_company_id_from_profile_response(self, response):
        for company in self.employees_json_data['empresas']:
            for funcionario in company['funcionarios']:
                if ('url' in funcionario) and (response.url == funcionario['url']):
                    return company['company_id']
        return None

    def find_company_by_id(self, company_id):
        for company in self.employees_json_data['empresas']:
            if company['company_id'] == company_id:
                return company

    # Isso pode ser ativado quando a url não começa com www:
    def check_response_status(self, response):
        if response.status == 999:
            print()
            errorprint('Status 999. O Linkedin começou a restringir pedidos.\nO crawler será encerrado automaticamente.\n')
            raise CloseSpider('Spider encerrado manualmente')

    def parse_profile(self, response):
        self.check_response_status(response)

        user_dict = None

        # Pode haver um erro quando é igual a None:
        company_id = self.get_company_id_from_profile_response(response)

        company = self.find_company_by_id(company_id)

        self.parsed_profile_urls.append(response.url)

        counter = '(%i/%i) ' % (len(self.parsed_profile_urls), len(self.profile_requests))

        print()

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
                warnprint('%sErro no parsing de %s\nAdicionando novamente à fila de links...\n' % (
                    counter, response.url))
                new_request = self.cookie_request(url=response.url, callback=self.parse_profile, dont_filter=True)
                self.profile_requests.append(new_request)
                return new_request
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
                    raise Exception('Erro com included_array')

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
                    raise Exception('Erro com following_json')

                user_dict = {
                    'nome': user_data['firstName'] if 'firstName' in user_data else None,
                    'sobrenome': user_data['lastName'] if 'lastName' in user_data else None,
                    'cargo_atual': user_data['headline'] if 'headline' in user_data else None,
                    'localizacao_atual': user_data['locationName'] if 'locationName' in user_data else None,
                    'sobre': user_data['summary'] if 'summary' in user_data else None,
                    'seguidores': following_json['data']['followersCount'],
                    'conexoes': self.format_conections(following_json['data']['connectionsCount']),
                    'habilidades': [skill['name'] for skill in skills_data],
                    'linguas': [language['name'] for language in languages_data],
                    'cursos_feitos': [course['name'] for course in courses_data],
                    'premios': [
                        {
                            'nome': honor['title'] if 'title' in honor else None,
                            'instituicao': honor['issuer'] if 'issuer' in honor else None,
                            'descricao': honor['description'] if 'description' in honor else None,
                            'data': self.convert_date(
                                honor['issuedOn'] if 'issuedOn' in honor else None,
                            )
                        } for honor in honors_data
                    ],
                    'estudos': [
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
                    'experiencia_profissional': [
                        {
                            'instituicao': experience['companyName'] if 'companyName' in experience else None,
                            'cargo': experience['title'] if 'title' in experience else None,
                            'descricao': experience['description'] if 'description' in experience else None,
                            'periodo': self.convert_date_range(
                                experience['dateRange'] if 'dateRange' in experience else None
                            ),
                        } for experience in positions_data
                    ],
                    'voluntariado': [
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
                    'projetos': [
                        {
                            'titulo': project['title'],
                            'url': project['url'],
                            'descricao': project['description'],
                            'periodo': self.convert_date_range(
                                project['dateRange']
                            )
                        } for project in projects_data
                    ],
                    'dados_obtidos': True
                }

                for company_index in range(len(self.employees_json_data['empresas'])):
                    if self.employees_json_data['empresas'][company_index]['company_id'] == company_id:

                        matches = False

                        for employee_index in range(len(self.employees_json_data['empresas'][company_index]['funcionarios'])):
                            if ('url' in self.employees_json_data['empresas'][company_index]['funcionarios'][employee_index]) \
                                and (self.employees_json_data['empresas'][company_index]['funcionarios'][employee_index]['url'] == str(response.url)):

                                matches = True

                                self.employees_json_data['empresas'][company_index]['funcionarios'][employee_index].update(user_dict)

                        if not matches:
                            self.employees_json_data['empresas'][company_index]['funcionarios'].append(user_dict)

                checkprint('%sParsing corretamente realizado em %s da empresa %s\n' % (counter, response.url, company['nome']))

            except Exception:
                errorprint('%sErro no parsing de %s da empresa %s\n' % (counter, response.url, company['nome']))

        save_to_file(
            self.employees_json_path,
            json.dumps(self.employees_json_data, indent=4),
            dont_print=True
        )


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
        print()
        errorprint('convert_unicode: não foi possível converter os caracteres unicode.\n')
    return text


def read_json_file(path):
    try:
        f = open(path, "r+")
        data = json.loads(f.read())
        f.close()
        return data
    except json.decoder.JSONDecodeError:
        return None


def save_to_file(filename, element, dont_print=False):
    with open(filename, 'wb') as f:
        f.write(str.encode(str(element)))
    if not dont_print: whiteprint('\n💽 Texto salvo como %s\n' % filename)
