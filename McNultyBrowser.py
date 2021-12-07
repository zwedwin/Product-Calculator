import xlwings
import json
import os
import time
from pprint import pprint
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import NoSuchElementException

class McNultyBrowser:

    def __init__(self):
        options = Options()
        options.headless = True
        prefs = {"profile.managed_default_content_settings.images": 2}
        options.add_experimental_option("prefs", prefs)
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        caps = DesiredCapabilities.CHROME
        caps['goog:loggingPrefs'] = {'performance': 'ALL'}
        self.driver = webdriver.Chrome(options = options, desired_capabilities = caps)
        self.driver.get('https://primemover.mcnulty.us/index.lasso')
        self.final_link_list = []
        self.link_list = []
        self.link_info = []

    def login(self, username, password):
        self.username = username
        self.password = password
        self.driver.find_element_by_name("mcnulty_username").send_keys(self.username)
        self.driver.find_element_by_name("mcnulty_password").send_keys(self.password)
        self.driver.find_element_by_css_selector('body > div > div > form > button').click()
        time.sleep(1)
        if self.driver.current_url != 'https://primemover.mcnulty.us/main.lasso':
            raise ValueError("Invalid Username or Password")

    def search(self,CAGE):
        self.driver.get('https://primemover.mcnulty.us/wraps/list.lasso')
        search_bar = self.driver.find_element_by_xpath('//*[@id="dataTable_wrapRateListing_filter"]/label/input')
        search_bar.send_keys(CAGE)
        i = 1
        while True:
            try:
                product_element = self.driver.find_element_by_xpath('//*[@id="dataTable_wrapRateListing"]/tbody/tr['+str(i)+']/td[5]/a')
                self.link_list.append(product_element.get_attribute('href'))
                i+=1
            except NoSuchElementException:
                break
        self.set_link_info()

    def set_link_info(self):
        for url in self.link_list:
            self.driver.get(url)
            self.wrapID = url.replace('https://primemover.mcnulty.us/wraps/detail.lasso?id=','')
            self.driver.find_element_by_xpath('/html/body/div[2]/div/div[2]/ul/li[5]/a').click()
            data_url = self.find_data_url()
            self.driver.get(data_url)
            text = self.driver.find_element_by_xpath('/html/body/pre').text
            self.json_wrap_data = json.loads(text)
            other_info = self.find_info_other('other4ColumData')
            code_info = self.find_info_codes('other4ColumData')
            info = other_info + code_info
            self.link_info.append(info)

    def get_WRAP(self,link):
        self.url = link
        self.wrapID = self.url.replace('https://primemover.mcnulty.us/wraps/detail.lasso?id=','')
        #go to product page
        self.driver.get(self.url)
        #go to 4Column Formatter tab
        self.driver.find_element_by_xpath('/html/body/div[2]/div/div[2]/ul/li[5]/a').click()
        self.data_url = self.find_data_url()
        self.driver.get(self.data_url)
        #find and format raw json data
        text = self.driver.find_element_by_xpath('/html/body/pre').text
        self.json_wrap_data = json.loads(text)
        #turn json data into specific data for fringe,overhead, and ga
        plant_fringe = self.find_info_fringe('displayPlant')
        plant_overhead = self.find_info_overhead('displayPlant')
        plant_ga = self.find_info_ga('displayPlant')
        field_fringe = self.find_info_fringe('displayField')
        field_overhead = self.find_info_overhead('displayField')
        field_ga = self.find_info_ga('displayField')
        other_info = self.find_info_other('other4ColumData')
        code_info = self.find_info_codes('other4ColumData')
        opmarg_info = self.find_info_opmarg('other4ColumData')
        self.empties_to_zeros(plant_fringe)
        self.empties_to_zeros(field_fringe)

        WRAP_type = other_info[4]

        self.workbook = xlwings.Book('4-columncalculatorscratch.xlsx')
        self.worksheet = self.workbook.sheets('4 Column Calculator')

        self.paste_data("G17","H27",plant_fringe,field_fringe)
        self.paste_data("G31","H34",plant_overhead,field_overhead)
        self.paste_data("G38","H42",plant_ga,field_ga)
        self.workbook.app.calculate()

        if WRAP_type == 'Development':
            self.workbook_final = xlwings.Book('1-Company CAGE Development WRAP Analysis_April 2021 v00dg.xlsx')
            self.worksheet_final = self.workbook_final.sheets('Company Name')
        elif WRAP_type == 'Services':
            self.workbook_final = xlwings.Book('CompanyCAGEServicesWRAPAnalysis_December2020v00dg.xlsx')
            self.worksheet_final = self.workbook_final.sheets('BAE')

        if len(self.json_wrap_data['wrapRate']['other4ColumData']['companyData_name']) < 31:
            self.worksheet_final.title = self.json_wrap_data['wrapRate']['other4ColumData']['companyData_name']
        else:
            self.worksheet_final.title = self.json_wrap_data['wrapRate']['other4ColumData']['companyData_name'][0:29]
        self.paste_info("B2","B8",other_info)
        self.paste_info("H2","H4",code_info)
        self.paste_info("M31","M32",opmarg_info)
        self.paste_final_data("B16","I42")
        self.set_formulas("B27","I27","B16")
        self.set_formulas("B34","I34","B30")
        self.set_formulas("B42","I42","B37")
        try:
            xlwings.Range('C47').value = str(float(self.json_wrap_data['companyData_materialHandelingRatePercent'])*100)+'%'
        except TypeError:
            xlwings.Range('C47').value = '0.00%'
        self.workbook_final.app.calculate()
        self.json_other_data = self.json_wrap_data['wrapRate']['other4ColumData']
        self.workbook_final_name = self.json_other_data['companyData_name'] + ' ' + self.json_other_data['companyData_codeCage'] + ' ' + self.json_other_data['companyData_wrapType'] +' WRAP Analysis_May2021 v00zl.xlsx'
        self.workbook_final.save(os.path.join(os.path.join(os.environ['USERPROFILE'],'Desktop'),self.workbook_final_name))
        self.workbook_final.close()
        self.workbook.close()

    def find_data_url(self):
        log = self.driver.get_log('performance')
        url_pattern = 'https://primemover.mcnulty.us/wraps/ajax/wrap_end_result.lasso?wrapID=' + self.wrapID
        for entry in log:
            message = json.loads(entry['message'])
            if message['message']['method'] == 'Network.requestWillBeSent':
                if url_pattern in message['message']['params']['request']['url']:
                    data_url = message['message']['params']['request']['url']
                    return data_url

    def find_info_fringe(self,index):
        data_order_fringe = ['fringe_finalHolidayVacationSick','fringe_pto','fringe_employerFICA','fringe_workersComp','fringe_healthInsurance','fringe_lifeInsurance',
        'fringe_pension','fringe_adAndD','fringe_salaryContinuation','fringe_disability','fringe_miscellaneous']
        fringe_data = ['']*11
        i = 0
        for key in data_order_fringe:
            fringe_data[i] = self.json_wrap_data['wrapRate'][index][key]
            i+=1
        return fringe_data

    def find_info_overhead(self,index):
        data_order_overhead = ['overhead_salariesWages','overhead_depreciation','overhead_occupancyAllocation','overhead_miscellaneous']
        overhead_data = ['']*4
        i = 0
        for key in data_order_overhead:
            overhead_data[i] = self.json_wrap_data['wrapRate'][index][key]
            i+=1
        return overhead_data

    def find_info_ga(self,index):
        data_order_ga = ['ga_salariesWages','ga_ird','ga_bp','ga_employeeStock','ga_miscellaneous']
        ga_data = ['']*5
        i = 0
        for key in data_order_ga:
            ga_data[i] = self.json_wrap_data['wrapRate'][index][key]
            i+=1
        return ga_data

    def find_info_other(self,index):
        data_order_other = ['companyData_name','companyData_costCenter','companyData_location1','companyData_location2',
        'companyData_wrapType','companyData_industrySector','fiscalYear']
        other_data = ['']*7
        i = 0
        for key in data_order_other:
            other_data[i] = self.json_wrap_data['wrapRate'][index][key]
            i+=1
        return other_data

    def find_info_codes(self,index):
        data_order_codes = ['companyData_codeDACIS','companyData_codeCage',
        'companyData_codeDUNS']
        code_data = ['']*3
        i = 0
        for key in data_order_codes:
            code_data[i] = self.json_wrap_data['wrapRate'][index][key]
            i+=1
        return code_data

    def find_info_opmarg(self,index):
        data_order_opmarg = ['companyData_operatingMarginPercent1','companyData_operatingMarginPercent2']
        opmarg_data = ['']*2
        i = 0
        for key in data_order_opmarg:
            opmarg_data[i] = self.json_wrap_data['wrapRate'][index][key]
            i+=1
        return opmarg_data

    def empties_to_zeros(self,list):
        for i in range(len(list)):
            if list[i] == '':
                list[i] = '0.00%'

    def paste_data(self,start_cell,end_cell,plant_data,field_data):
        start_cell_letter = start_cell[0].lower()
        end_cell_letter = end_cell[0].lower()
        start_cell_number = int(start_cell.strip(start_cell_letter.upper()))
        end_cell_number = int(end_cell.strip(end_cell_letter.upper()))
        letters = []
        for i in range(ord(start_cell_letter),(ord(end_cell_letter)+1)):
            letter = chr(i)
            letters.append(letter)
        for letter in letters:
            if letter.upper() in start_cell:
                for i in range(start_cell_number,(end_cell_number+1)):
                    xlwings.Range(letter.upper()+str(i)).value = plant_data[i-start_cell_number]
            if letter.upper() in end_cell:
                for i in range(start_cell_number,(end_cell_number+1)):
                    xlwings.Range(letter.upper()+str(i)).value = field_data[i-start_cell_number]

    def paste_info(self,start_cell,end_cell,info):
        self.workbook_final.activate()
        letter = start_cell[0]
        start_num = int(start_cell.strip(letter.upper()))
        end_num = int(end_cell.strip(letter.upper()))
        for i in range(start_num,end_num+1):
            xlwings.Range(letter+str(i)).value = info[i-start_num]

    def paste_final_data(self,start_cell,end_cell):
        start_cell_letter = start_cell[0].lower()
        end_cell_letter = end_cell[0].lower()
        start_cell_number = int(start_cell.strip(start_cell_letter.upper()))
        end_cell_number = int(end_cell.strip(end_cell_letter.upper()))
        letters = []
        for i in range(ord(start_cell_letter),(ord(end_cell_letter)+1)):
            letter = chr(i)
            letters.append(letter)
        cell_values = []
        self.workbook.activate()
        for letter in letters:
                for i in range(start_cell_number,(end_cell_number+1)):
                    this_letter = chr(ord(letter)+1)
                    value = xlwings.Range(this_letter.upper()+str(i+1)).value
                    cell_values.append(value)
        self.workbook_final.activate()
        i = 0
        for letter in letters:
            for j in range(start_cell_number,(end_cell_number+1)):
                xlwings.Range(letter.upper()+str(j)).value = cell_values[i]
                i+=1

    def set_formulas(self,start_cell,end_cell,top_cell):
        start_cell_letter = start_cell[0].lower()
        end_cell_letter = end_cell[0].lower()
        top_cell_letter = top_cell[0].lower()
        cell_number = int(start_cell.strip(start_cell_letter.upper()))
        top_cell_number = int(top_cell.strip(top_cell_letter.upper()))
        letters = []
        for i in range(ord(start_cell_letter),(ord(end_cell_letter)+1)):
            letter = chr(i)
            letters.append(letter)
        for letter in letters:
            xlwings.Range(letter.upper()+str(cell_number)).value = '=SUM('+letter.upper()+str(top_cell_number)+':'+letter.upper()+str(cell_number-1)+')'
