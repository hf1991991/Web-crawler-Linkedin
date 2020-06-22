from scrapy import Request
from scrapy.spiders import CrawlSpider, Rule
from scrapy.spiders.init import InitSpider
import scrapy.http as Http
from scrapy.utils.spider import iterate_spider_output

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import colors
from openpyxl.styles import Font, Color, Side, Alignment, Border
from openpyxl.utils import get_column_letter

import json

CELL_SIDE = Side(
    border_style="thin",
    color="000000"
)

CELL_BORDER = Border(
    top=CELL_SIDE,
    bottom=CELL_SIDE,
    right=CELL_SIDE,
    left=CELL_SIDE,
)

CELL_ALIGNMENT = Alignment(
    vertical="center",
    horizontal="left",
    wrap_text=True
)

class LinkedinSpider(InitSpider):
    name = "linkedin"
    allowed_domains = ["linkedin.com"]

    workbook_filename = 'Links.xlsx'
    workbook = None

    user_name = None
    passwd = None
    user_line_on_excel = None

    start_urls = []

    unicode_dict = {}

    login_page = 'https://www.linkedin.com/uas/login'
        
    def init_request(self):
        # Obtém os dados do excel:
        self.workbook = load_workbook(filename=self.workbook_filename)
        # Arruma dados da tabela de usuários do excel:
        self.fix_users_sheet_data()
        # A partir dos dados do excel, associa valores às variaveis de login, assim como à dos links:
        if self.get_login_data_from_workbook() is not None: return
        self.get_links_from_workbook()
        # Aplica estilo no excel:
        self.apply_style_to_workbook_sheet(sheet=self.workbook['Links'], starting_line=3, columns="BCDEFGH")
        self.apply_style_to_workbook_sheet(sheet=self.workbook['Usuários'], starting_line=3, columns="BCDEF")
        # Lê os valores de conversão unicode:
        self.read_unicode_conversion()
        # Realiza o login:
        return Request(url=self.login_page, callback=self.login, dont_filter=True)

    def fix_users_sheet_data(self):
        users_sheet = self.workbook['Usuários']
        line = 3
        while users_sheet['B%i' % line].value is not None or users_sheet['C%i' % line].value:
            if users_sheet['D%i' % line].value is None:
                users_sheet['D%i' % line] = 0
            if (users_sheet['E%i' % line].value != 'Sim') and (users_sheet['E%i' % line].value != 'Não') and (users_sheet['E%i' % line].value != 'Não testado'):
                users_sheet['E%i' % line] = 'Não testado'
            if users_sheet['F%i' % line].value is None:
                users_sheet['F%i' % line] = '---'
            line += 1
        self.workbook.save(self.workbook_filename)


    def get_login_data_from_workbook(self):
        # print('GET_LOGIN_DATA_FROM_WORKBOOK')

        def has_been_tested(item):
            return item['does_it_work'] == 'Sim'

        def times_used(item):
            return item['times_used']
        
        users_sheet = self.workbook['Usuários']
        possible_users = []
        line = 3

        while True:
            login = {
                'email': users_sheet['B%i' % line].value,
                'password': users_sheet['C%i' % line].value,
                'times_used': users_sheet['D%i' % line].value,
                'does_it_work': users_sheet['E%i' % line].value,
                'line': line
            }
            if login['email'] is None or login['password'] is None:
                break
            if login['does_it_work'] != 'Não':
                possible_users.append(login)
            line += 1

        if len(possible_users) == 0:
            print('\nNão há mais usuários válidos para serem utilizados.\nEntre na tabela do Excel para adicionar um usuário, ou arrumar algum que tenha gerado um erro.\n')
            self.workbook.save(self.workbook_filename)
            return 'Não há mais usuários válidos para serem utilizados'
        
        possible_users.sort(key=times_used)
        possible_users.sort(key=has_been_tested)

        self.user_name = possible_users[0]['email']
        self.passwd = possible_users[0]['password']
        self.user_line_on_excel = possible_users[0]['line']

        users_sheet['D%i' % self.user_line_on_excel] = possible_users[0]['times_used'] + 1
        self.workbook.save(self.workbook_filename)
        return None

    def get_links_from_workbook(self):
        # print('GET_LINKS_FROM_WORKBOOK')
        links_sheet = self.workbook['Links']
        line = 3
        link = links_sheet['C%i' % line].value
        while link is not None:
            self.start_urls.append(link)
            line += 1
            link = links_sheet['C%i' % line].value

    def apply_style_to_workbook_sheet(self, sheet, starting_line, columns):
        # print('APPLY_STYLE_TO_WORKBOOK')
        line = starting_line
        link = sheet['C%i' % line].value
        while link is not None:
            for column in columns:
                cell = sheet['%s%i' % (column, line)]
                cell.alignment = CELL_ALIGNMENT
                cell.border = CELL_BORDER
            line += 1
            link = sheet['C%i' % line].value
        self.workbook.save(self.workbook_filename)

    def write_on_workbook(self, url, user_dict):
        # print('WRITE_ON_WORKBOOK')
        links_sheet = self.workbook['Links']
        column_association = {
            'D': 'first_name',
            'E': 'last_name',
            'F': 'occupation',
            'G': 'location',
            'H': 'about',
        }
        line = 3
        link = links_sheet['C%i' % line].value
        while link is not None:
            if link == url:
                links_sheet['B%i' % line] = line - 5
                if user_dict is not None:
                    for column in column_association:
                        text = user_dict[column_association[column]]
                        if text == None:
                            text = '---'
                        links_sheet['%s%i' % (column, line)] = text
                else:
                    true = True
                    # Implementar mensagem de conta não existe
                self.workbook.save(self.workbook_filename)
                return
            line += 1
            link = links_sheet['C%i' % line].value
        print('ERRO em write_on_workbook: %s não foi encontrado' % url)

    def read_unicode_conversion(self):
        with open('unicode_conversion.json', 'r') as f:
            self.unicode_dict = json.loads(f.read())

    def login(self, response):
        return Http.FormRequest.from_response(
            response,
            formdata={
                'session_key': self.user_name,
                'session_password': self.passwd,
            },
            callback = self.check_login_response,
            meta={
                'proxy': None
            }
        )

    def set_error_message_on_users_sheet(self, error_text):
        users_sheet = self.workbook['Usuários']
        if error_text == None:
            users_sheet['E%i' % self.user_line_on_excel] = 'Sim'
            error_text = '---'
        else:
            users_sheet['E%i' % self.user_line_on_excel] = 'Não'
        users_sheet['F%i' % self.user_line_on_excel] = error_text
        self.workbook.save(self.workbook_filename)

    def response_is_ban(self, request, response):
        return b'Let&#39;s do a quick security check' in response.body

    def exception_is_ban(self, request, exception):
        return None

    def check_login_response(self, response):
        error_text = None

        print("\nLogin utilizado:\n - Email: %s\n - Senha: %s\n" % (self.user_name, self.passwd))
        if "Your account has been restricted" in str(response.body):
            error_text = 'Conta bloqueada pelo Linkedin por muitas tentativas. Troque esta conta por outra, ou remova esta linha do Excel.'
            print("Login falhou. Conta bloqueada pelo Linkedin por muitas tentativas.\nTente criar uma nova conta.\n")
        elif "Let&#39;s do a quick security check" in str(response.body):
            # error_text = 'É necessário realizar uma verificação de segurança. Entre em %s e faça login com essa conta para resolver um captcha.' % self.login_page
            print("Login falhou. É necessário realizar uma verificação de segurança\nO programa deve mudar de proxy agora.\n")
        elif "Email or phone" in str(response.body):
            error_text = 'A conta ou a senha parecem estar erradas. Verifique se o usuário e senha estão corretos.'
            print("Login falhou. A conta ou a senha estão erradas.\nVerifique se o usuário e senha estão corretos.\n")
        else:
            print("\nLogin realizado. Vamos começar o crawling!\n")

        self.set_error_message_on_users_sheet(error_text)

        if error_text == None:
            return self.initialized() 
        else:
            return self.init_request()

    def start_requests(self):
        self._postinit_reqs = self.start_requests_without_proxy_change()
        return iterate_spider_output(self.init_request())

    def start_requests_without_proxy_change(self):
        # print('START_SPLASH_REQUESTS')
        for url in self.start_urls:
            # O seguinte código faz com que todos os Requests depois do login não mudem de proxy:
            yield Request(
                url=url, 
                callback=self.parse,
                meta={
                    'proxy': None
                }
            )

    def get_user_data_string(self, response):
        user_data_string = '{&quot;birthDateOn' + str(response.body).split(',{&quot;birthDateOn')[-1]
        end = 1
        partial = user_data_string[:end]
        while (partial.count('{') != partial.count('}')) and (partial.count('{') < 200) and (len(user_data_string) > end):
            end += 1
            partial = user_data_string[:end]
            # if partial.endswith('{') or partial.endswith('}'):
            #     print(partial.count('{'), partial.count('}'))
        if partial.count('{') != partial.count('}'):
            print('ERRO em get_user_data_string: não foi possivel obter dados do usuário em %s' % response.url)
            return None
        return partial

    def parse(self, response):
        # Descomente as seguintes linhas para obter em texto a resposta real de html:
        # page = response.url.split("/")[4]
        # filename = 'profile-%s.html' % page
        # with open(filename, 'wb') as f:
        #     f.write(response.body)
        # self.log('Saved file %s' % filename)

        user_dict = None

        # Se página não tiver a seguinte string, a conta não existe:
        if '{&quot;birthDateOn' not in str(response.body):
            print('\nConta não existe em %s\n' % response.url)
        else:
            filename = 'profile-%s.json' % response.url.split("/")[4]

            user_data_string = self.get_user_data_string(response)

            if user_data_string != None:
                user_data = parse_text_to_json(user_data_string, self.unicode_dict, filename)


                if user_data is not None:

                    user_dict = {
                        'first_name': user_data['firstName'],
                        'last_name': user_data['lastName'],
                        'occupation': user_data['headline'],
                        'location': user_data['locationName'],
                        'about': user_data['summary'],
                    }

                    print('\nParsing corretamente realizado em %s\n' % response.url)
                else:
                    print('\nErro no parsing de %s\n' % response.url)

        self.write_on_workbook(response.url, user_dict)


def parse_text_to_json(text, replacements, filename):
    try:
        text = convert_unicode(text, replacements)
        # save_to_file(
        #     filename, 
        #     text
        # )
        return json.loads(text)
    except Exception:
        print('\nERRO em parse_text_to_json.\n')
        return None

def convert_unicode(text, replacements):
    try:
        text = str(text)
        for unicode_char in list(replacements.keys()):
            for type in list(replacements[unicode_char].keys()):
                for element in replacements[unicode_char][type]:
                    text = text.replace(str(element), str(unicode_char))
    except Exception:
        print('\nERRO em convert_unicode.\n')
    return text

def save_to_file(filename, element):
    element = str(element).replace("'", '"').replace('"s ', "'s ").replace('True', 'true').replace('False', 'false').replace('None', 'null')
    with open(filename, 'wb') as f:
        f.write(str.encode(str(element)))
    print('\nSaved file %s\n' % filename)
