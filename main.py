import requests
from scrapy.selector import Selector
from selenium import webdriver
from selenium.webdriver.firefox.options import Options

headers = {
    'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.106 Safari/537.36'
    }

class amazonApi:
    def open_browser(self) :
        print('opening browser..')
        options = Options()
        options.add_argument("--incognito")
        self.driver = webdriver.Firefox()

    def get_page_source(self,url):
        self.driver.get(url)
        page_source = self.driver.page_source
        res = Selector(text=page_source)
        return res

    def main(self):
        self.open_browser()
        asin_list = ['B08J65DST5','B08J6Z2']
        for asin in asin_list:
            url = 'https://www.amazon.com/dp/{}'.format(asin)
            response = self.get_page_source(url)
            title = response.css('title::text').get()
            print(title)



if __name__=='__main__':
    api = amazonApi()
    api.main()