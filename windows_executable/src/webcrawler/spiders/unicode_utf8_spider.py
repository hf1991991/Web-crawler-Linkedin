from scrapy import Request, Spider

import json

class UnicodeUtf8Spyder(Spider):
    name = "unicode"

    start_urls = ['https://dev.w3.org/html5/html-author/charref']

    def parse(self, response):
        dictn = {}

        for char_data in response.css('tr'):
            character = char_data.css('td.character::text').get()[1:]
            dec = char_data.css('td.dec code::text').get().split()
            hexa = char_data.css('td.hex code::text').get().split()
            named = char_data.css('td.named code::text').get().split()

            if (character == '"'):
                character = 'double_quote'
            elif (character == "'"):
                character = 'single_quote'
            
            dictn[character] = {
                "named": named,
                "hex": hexa,
                "dec": dec,
            }
        
        filename = 'unicode_conversion.json'
        with open(filename, 'wb') as f:
            f.write(str.encode(str(dictn).replace("'", '"').replace('\\x', '\\\\x').replace('single_quote', "'").replace('double_quote', '\\"')))
        self.log('Saved file %s' % filename)
