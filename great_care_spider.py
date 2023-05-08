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
        Rule(LinkExtractor(allow=('/doctors/')), callback='parse_item', follow=True),
    )

    def parse_item(self, response):
        global file_index

        for title in response.xpath(title_xpath):
            ExcelWorker.write_file(file_index, title.css('::text').get(), response.url)
            file_index += 1
        self.log('crawling'.format(response.url))

class ExcelWorker():
    def write_file(index, name, surname):
        wb = openpyxl.load_workbook(file_path)
        ws = wb.get_sheet_by_name('Sheet')

        ws.cell(row=index, column=1).value = name
        ws.cell(row=index, column=2).value = surname

        wb.save(file_path)
        wb.close()