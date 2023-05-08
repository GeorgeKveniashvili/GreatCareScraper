import openpyxl
from scrapy.spiders import CrawlSpider, Rule
from scrapy.linkextractors import LinkExtractor

grat_care_url = 'https://www.iwantgreatcare.org/search'
title_xpath = '//*[@id="entity-name-score-container"]/h1'
file_path = "test101.xlsx"
file_index = 2

class GreatCareSpider(CrawlSpider):
    name = 'greatcarespider'
    allowed_domains = ['iwantgreatcare.org']
    start_urls = [grat_care_url]

    rules = (
        # Extract and follow all links!
        Rule(LinkExtractor(allow=['/doctors/', '/optometrists/', '/nurses/', '/dentists/', 
                                  '/physiotherapists/', '/dietitians/', '/occupationaltherapists/']), 
                                  callback='parse_item', follow=True),
    )

    def parse_item(self, response):
        global file_index

        for title in response.xpath(title_xpath):
            full_name_formated = split_name(title.css('::text').get())
            name = full_name_formated[0]
            surname = full_name_formated[1]

            write_file(file_index, name, surname, response.url)
            file_index += 1
        self.log('crawling'.format(response.url))


def write_file(index, name, surname, domain):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.get_sheet_by_name('Sheet')

    ws.cell(row=index, column=1).value = name
    ws.cell(row=index, column=2).value = surname
    ws.cell(row=index, column=3).value = domain

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