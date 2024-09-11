import urllib.request
import pandas as pd
import os
import re
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import *
import urllib


df = pd.read_excel('odczynniki.xls', engine='xlrd')

manu = {'thermo':'https://www.thermofisher.com/order/catalog/product/',
            'sigma':'https://www.sigmaaldrich.com/PL/pl',
            'vwr':'https://pl.vwr.com/store/'}

driver = webdriver.Chrome()

#disables search engine selection pop-up

chrome_options = Options()
chrome_options.add_argument("--disable-search-engine-choice-screen")
driver = webdriver.Chrome(options=chrome_options)

#'table_search' function iterates the df over keys ('manu' dict.),
#returns lists to 'data_pass' function 'data_pass' returns
# a list of lists -> [link ("manu" dict.), cat nr, lot nr (from df)]

def data_pass():

    web_cat_lot=[]

    def table_search(key, link):
        for index, row in df.iterrows():
                try:
                    if len(re.findall(key,row['web']))==1:
                        web_cat_lot.append([link, str(row['cat nr']), str(row['lot nr'])])
                except TypeError:
                    print('row of "web" column is empty')

    for item in manu.items():
        table_search(item[0], item[1])

    return web_cat_lot

def wait_until(time:int, sele, CSS=True):

    if CSS==True:
        element = WebDriverWait(driver, time).\
            until(EC.visibility_of_element_located((By.CSS_SELECTOR, sele)))
    else:
        element = WebDriverWait(driver, time).\
            until(EC.visibility_of_element_located((By.XPATH, sele)))
    
    return element

#'shadow_element' function gets an element from shadow-root

def shadow_element(entry_selector: str, element_selector: str):

    shadow_entry = driver.find_element(By.CSS_SELECTOR, entry_selector)
    shadow_root = driver.execute_script("return arguments[0].shadowRoot", shadow_entry)
    element = shadow_root.find_element(By.CSS_SELECTOR, element_selector)

    return element


def shadow_elements(entry_selector: str, elements_selector: str):

    shadow_entry = driver.find_element(By.CSS_SELECTOR, entry_selector)
    shadow_root = driver.execute_script("return arguments[0].shadowRoot", shadow_entry)
    elements = shadow_root.find_elements(By.CSS_SELECTOR, elements_selector)

    return elements

#'parser' navigates a website depending of manufacturer 
# (code prepared for Thermofisher so far)

def parser():

    web_cat_lot = data_pass()

    for i in web_cat_lot:
        
        if len(re.findall('thermo', i[0]))==1:
            adress = i[0]+i[1]
        
            driver.get(adress)
            driver.maximize_window()
            
            #acceptation of 'cookies'
            try:
                cookie = driver.find_element(
                    By.ID,'''truste-consent-button''').click()

            except NoSuchElementException:
                pass          
            
        #assigning product name to "header_text" variable
            try:
                header = driver.find_element(
                    By.XPATH, '''//*[@id="root"]/div/div[1]/div[2]/div[2]/div[1]/span''')
                header_text = header.text
                      
            except NoSuchElementException:
                header = driver.find_element(
                    By.XPATH, '''//*[@id="root"]/div/div/div[4]\
                        /div/div/div/div[3]/div[1]/div/div[2]/h1''')
                header_text = header.accessible_name

        
            selector = '''#certificates > div.pdp-certificates-search > \
                        div.pdp-certificates-search__inputs > div > core-search'''
            try:
                wait_until(10, selector)

                iframe = driver.find_element(By.CSS_SELECTOR,selector)
                ActionChains(driver).scroll_to_element(iframe).perform()
            except TimeoutException:
                try:
                    iframe = driver.find_element(By.CSS_SELECTOR,selector)
                    ActionChains(driver).scroll_to_element(iframe).perform()

                except NoSuchElementException:
                    selector = '''#root > div > div > div.p-tabs > div.p-tabs__content > div:nth-child(7) > div > div > div:nth-child(1) > div:nth-child(2) > div.pdp-documents__document-section > div > div.pdp-documents__search > div.pdp-documents__search-inputs > div > div.c-search-bar.pdp-documents__search-bar.pdp-documents__search-bar--desktop > input'''
                    try:
                        wait_until(10, selector)
                        
                        iframe = driver.find_element(By.CSS_SELECTOR, selector)
                        ActionChains(driver).scroll_to_element(iframe).perform()

                    except TimeoutException:
                        try:
                            iframe = driver.find_element(By.CSS_SELECTOR, selector)
                            ActionChains(driver).scroll_to_element(iframe).perform()
                        except NoSuchElementException:
                            print('html structure has changed or website crashed A')
                             
            #finds search bar and enters lot nr
            
            try:
                wait_until(10, selector)
                element = shadow_element(
                            entry_selector = selector,
                            element_selector = 'div > div > input'
                                        )
            except (NoSuchElementException, AttributeError):
                try:
                    element = shadow_element(
                            entry_selector = selector,
                            element_selector = 'div > div > input'
                                        )
                except (NoSuchElementException, AttributeError):
                    selector = '''//*[@id="root"]/div/div/div[5]/div[2]/div[7]/div/div/div[1]/div[1]/div[2]/div/div[1]/div[2]/div/div[2]/input'''
                    try:
                        wait_until(10, selector, CSS=False)
                        element = driver.find_element(By.XPATH, selector)
                    except TimeoutException:
                        try:
                            element = driver.find_element(By.XPATH, selector)
                        except NoSuchElementException:
                            print('html structure has changed or website crashed B')
           
            driver.execute_script("arguments[0].click();", element)
            element.send_keys(i[2])
            element.send_keys(Keys.RETURN)
            
        
            #getting link to CoA
            try:
                selector = '''//*[@id="certificates"]
                    /div[2]/div[1]/div[2]/span[1]'''
                
                element = wait_until(10, selector, CSS=False)

                link = driver.find_element(By.XPATH, selector)
                
            except TimeoutException:
                try:
                    link = driver.find_element(By.XPATH, selector)

                except NoSuchElementException:
                    selector = '''//*[@id="root"]/div/div/div[5]
                                /div[2]/div[7]/div/div/div[1]/div[1]
                                /div[2]/div/div[2]/div[2]/div/span[1]
                                /a/span[2]'''
                    try:
                        element = wait_until(10, selector, CSS=False)
                        link = driver.find_element(By.XPATH, selector)
                    except TimeoutException:
                        try:
                            link = driver.find_element(By.XPATH, selector)
                        except NoSuchElementException:
                            print('html structure has changed or website crashed C')
         
            window_before = driver.window_handles[0]
            driver.execute_script("arguments[0].click();", link)
            window_after = driver.window_handles[1]
            driver.switch_to.window(window_after)
            string_url = driver.current_url
            driver.close()
            driver.switch_to.window(window_before)

            new_dir = 'output/' + header_text + '_' + 'nr_cat_' + i[1] + '/' + 'CoA/'
            
            try:
                os.makedirs(new_dir)

            except FileExistsError:
                pass
           
            url_path = new_dir + 'CoA_nr_lot_' + i[2]
            urllib.request.urlretrieve(string_url, url_path)

parser()