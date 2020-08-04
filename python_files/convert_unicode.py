from webcrawler.unicode_conversion import unicode_dict

def convert_unicode(text, replacements):
    try:
        text = str(text)
        for unicode_char in list(replacements.keys()):
            for type in list(replacements[unicode_char].keys()):
                for element in replacements[unicode_char][type]:
                    text = text.replace(str(element), str(unicode_char))
    except Exception:
        print('convert_unicode: não foi possível converter os caracteres unicode.\n')
    return text


def save_to_file(filename, element, dont_print=False):
    with open(filename, 'wb') as f:
        f.write(str.encode(str(element)))
    if not dont_print: print('\n💽 Texto salvo como %s\n' % filename)


text = '{&quot;data&quot;:{&quot;premium&quot;:false,&quot;influencer&quot;:false,&quot;entityUrn&quot;:&quot;urn:li:fs_memberBadges:ACoAACHAYngBHR7LUufQU0EPRSbbpQzJLEHTlos&quot;,&quot;openLink&quot;:false,&quot;jobSeeker&quot;:false,&quot;$type&quot;:&quot;com.linkedin.voyager.identity.profile.MemberBadges&quot;},&quot;included&quot;:[]}'

save_to_file(
    'merber_badges_json.json',
    convert_unicode(text, unicode_dict)
)