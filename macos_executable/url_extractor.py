import json


def read_json_file(path):
    try:
        f = open(path, "r+")
        data = json.loads(f.read())
        f.close()
        return data
    except json.decoder.JSONDecodeError:
        return None


def read_txt_file(path):
    try:
        f = open(path, "r+")
        data = f.read()
        f.close()
        return data
    except Exception:
        return None


def create_list(txt):
    return txt.split('\n') if txt is not None else []


def save_to_file(filename, element, dont_print=False):
    with open(filename, 'wb') as f:
        f.write(str.encode(str(element)))
    if not dont_print: print('\n💽 Texto salvo como %s\n' % filename)


urls = create_list(read_txt_file('urls.txt'))

print(urls)

for company in read_json_file('../new_windows_executable/output.json')['empresas']:
    for employee in company['funcionarios']:
        if ('url' in employee) and (employee['url'] not in urls):
            urls.append(employee['url'])

save_to_file(
    'urls.txt',
    '\n'.join(urls)
)
