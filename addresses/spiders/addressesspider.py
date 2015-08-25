# -*- coding: utf-8 -*-
import scrapy
from selenium import webdriver
from lxml.html import fromstring
import xlrd
import urllib
from xlrd import open_workbook
from xlutils.copy import copy

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import time

class AddressesspiderSpider(scrapy.Spider):
    name = "addressesspider"
    allowed_domains = ["addresses.com"]
    start_urls = (
        'http://www.addresses.com/',
    )

    def read_xls(self,file,col_no=5):
        # Open the workbook
        xl_workbook = xlrd.open_workbook(file)
        # List sheet names, and pull a sheet by name
        sheet_names = xl_workbook.sheet_names()
        print('Sheet Names', sheet_names)

        xl_sheet = xl_workbook.sheet_by_name(sheet_names[0])
        num_cols = xl_sheet.ncols   # Number of columns

        row = xl_sheet.row(0)  # 1st row

        num_cols = xl_sheet.ncols   # Number of columns

        row = []
        for row_idx in range(1, xl_sheet.nrows):    # Iterate through rows
            row_obj = xl_sheet.row(row_idx)
            row.append(row_obj)

        return row


    def parse(self, response):
        file='propertyownerlist.xlsx'
        addresses = self.read_xls(file)
        """
            Modifying the existing xls file
        """
        rb = open_workbook(file)
        wb = copy(rb)
        sheet = wb.get_sheet(0)
        sheet.write(0,23,'Phone')

        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-javascript")
            options.add_argument("--disable-java")
            options.add_argument("--disable-plugins")
            options.add_argument("--disable-popup-blocking")
            options.add_argument("--disable-images")
            driver = webdriver.Chrome('c://chromedriver.exe',chrome_options = options)

            """
                Working with version 1.9.8 which supports a resource timeout. 2.0.1 seems to ignore the value provided.
            """

            #driver = webdriver.PhantomJS(executable_path='C:/phantomjs-2.0.0-windows/bin/phantomjs.exe', service_args=['--load-images=no','--proxy-server=%s' % PROXY[random.randrange(len(PROXY))]])

            for ctr,address in enumerate(addresses,start=507):
                #Getting col values from xls file
                address_col = address[5].value
                city_col = address[7].value
                st = address[8].value
                name_col = address[1].value

                print address_col, city_col, st
                print type(city_col)

                city =  city_col+" "+st
                city = urllib.quote_plus(city)
                add = urllib.quote_plus(address_col)
                driver.get("http://www.addresses.com/addr/{}/{}".format(add,city))
                time.sleep(1.5)

                html = driver.page_source
                hxs = fromstring(html)
                results = hxs.xpath('//div[@class="wp_result_details detail_column"]')
                for result in results:
                    name = result.xpath('./div/a/text()')
                    """
                        Getting full name from addresse.com search with query .address so as to match with Col B in Excel.
                    """
                    full_name = ''
                    if name:
                        print "sitename" + name[0]
                        if all(x in name_col.split() for x in name[0].split()):
                            phone = result.xpath('./div/div[@class="listing_header"]/text()')
                            if phone:
                                phone = phone[0]
                                print phone
                                sheet.write(ctr,23,phone)
                                wb.save('file.xls')

                            print "matched"
                            print ctr


                    address = result.xpath('normalize-space(./div[@class="listing_detail"]/text())')
                    phone = result.xpath('./div/div[@class="listing_header"]/text()')

        except Exception as e:
            print e

