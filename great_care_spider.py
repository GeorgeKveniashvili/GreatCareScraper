import openpyxl
from scrapy.spiders import CrawlSpider, Rule
from scrapy.linkextractors import LinkExtractor

grat_care_url = 'https://www.iwantgreatcare.org/search'
allowed_extractor = ['/doctors/', '/optometrists/', '/nurses/', '/dentists/', '/physiotherapists/', '/dietitians/', '/occupationaltherapists/']
title_xpath = '//*[@id="entity-name-score-container"]/h1'
specialises_xpath = '//*[@id="specialies-container"]/div/ul'
#//*[@id="information"]/div[2]/div[2]/div/ul

works_xpaths = ['//*[@id="works-at-container"]/div/ul', '/html/body/div[1]/main/div/span/div[2]/div/div[2]/div[1]/ul',
                '//*[@id="information"]/div[2]/div[1]/div/ul', '//*[@id="information"]/div[1]/div[1]/div/ul']
postcode_xpath = 'sss'
file_path = "test101.xlsx"
file_index = 2

class GreatCareSpider(CrawlSpider):
    name = 'greatcarespider'
    allowed_domains = ['iwantgreatcare.org']
    start_urls = [grat_care_url]

    rules = (
        # Extract and follow all links!
        Rule(LinkExtractor(allow=('/doctors/')), callback='parse_item', follow=True),
    )

    def parse_item(self, response):
        global file_index

        for title in response.xpath(title_xpath):
            full_name_formated = split_name(title.css('::text').get())
            name = full_name_formated[0]
            surname = full_name_formated[1]

            specialises_selector = response.xpath(specialises_xpath)
            specialises_list = specialises_selector.css('li::text').getall()
            specialises = ""
            for item in specialises_list:
                specialises += str(item)
            
            profile = "Empty"

            works = extract_works(response)

            write_file(file_index, name, surname, specialises, profile, works, postcode_xpath, response.url)
            file_index += 1
        self.log('crawling'.format(response.url))

def extract_works(response):
    for works_xpath in works_xpaths:
        works_selector = response.xpath(works_xpath)
        if works_selector is None:
            continue
        works_list = works_selector.css('a::text').getall()
    works = ""
    for item in works_list:
        works += str(item)
    return works


def write_file(index, name, surname, specialises, profile, works, postcode, domain):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.get_sheet_by_name('Sheet')

    ws.cell(row=index, column=1).value = name
    ws.cell(row=index, column=2).value = surname
    ws.cell(row=index, column=3).value = specialises
    ws.cell(row=index, column=4).value = profile
    ws.cell(row=index, column=5).value = works
    ws.cell(row=index, column=6).value = postcode
    ws.cell(row=index, column=7).value = domain

    wb.save(file_path)
    wb.close()

def split_name(full_name):

    if(len(full_name.split(' ')) > 2):
        name = ' '.join(full_name.split(' ', 2)[:2])
        surname = ' '.join(full_name.split(' ', 2)[2:])
        return [name, surname]
    name = full_name.split(' ', 1)[0]
    surname = full_name.split(' ', 1)[1]
    return [name, surname]