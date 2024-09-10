import urllib.request
import pandas as pd
import os
import re
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import *
import time
import urllib


df = pd.read_excel('odczynniki.xls', engine='xlrd')

dostawcy = {'thermo':'https://www.thermofisher.com/order/catalog/product/',
            'sigma':'https://www.sigmaaldrich.com/PL/pl',
            'vwr':'https://pl.vwr.com/store/'}

driver = webdriver.Chrome()

#wyłączenie okna z wyborem wyszukiwarki

chrome_options = Options()
chrome_options.add_argument("--disable-search-engine-choice-screen")
driver = webdriver.Chrome(options=chrome_options)

#funkcja 'table_search' wyszukuje klucze 'keys' ze słownika 'dostawcy' 
#iterując po data frame, zwraca listy do funkcji data_pass
#funkcja data_pass zwraca listę list -> [link (słownik "dostawcy"), cat nr, lot nr (z data frame)]
def data_pass():

    web_cat_lot=[]

    def table_search(key, link):
        for index, row in df.iterrows():
                try:
                    if len(re.findall(key,row['web']))==1:
                        web_cat_lot.append([link, str(row['cat nr']), str(row['lot nr'])])
                except:
                    print('pusty wiersz kolumny "web"')
            
                
    for item in dostawcy.items():
        table_search(item[0], item[1])

    return web_cat_lot

#funkcja 'shadow_element' wyciąga element html z shadow-root
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

#funkcja parser nawiguje po stronie internetowej w zależności
# od producenta (zrobiono wersję dla thermoscientific)
def parser():

    web_cat_lot = data_pass()

    for i in web_cat_lot:
        
        if len(re.findall('thermo', i[0]))==1:
            adress = i[0]+i[1]
        
            driver.get(adress)
            driver.maximize_window()
            
            #zatwierdzenie "cookie"
            try:
                cookie = driver.find_element(
                    By.ID,'''truste-consent-button''').click()

            except NoSuchElementException:
                pass           
            
            #zapisanie do zmiennej "header_text" nazwy odczynnika
            try:
                header = driver.find_element(
                    By.XPATH, '''//*[@id="root"]/div/div[1]/div[2]/
                    div[2]/div[1]/span''')
                header_text = header.text
                      
            except NoSuchElementException:
                header = driver.find_element(
                    By.XPATH, '''//*[@id="root"]/div/div/div[4]/div/
                    div/div/div[3]/div[1]/div/div[2]/h1''')
                header_text = header.accessible_name  

            try:
                selector = '''#certificates > div.pdp-certificates-search > 
                        div.pdp-certificates-search__inputs > div > core-search'''
                iframe = driver.find_element(By.CSS_SELECTOR,selector)
                ActionChains(driver).scroll_to_element(iframe).perform()

            except NoSuchElementException:
                selector = '''#root > div > div > div.p-tabs > div.p-tabs
                __content > div:nth-child(7) > div > div > div:nth-child(1)
                  > div:nth-child(2) > div.pdp-documents__document-section 
                  > div > div.pdp-documents__search > div.pdp-documents__search
                  -inputs > div > div.c-search-bar.pdp-documents__search-bar.pdp
                  -documents__search-bar--desktop > input'''
            
                iframe = driver.find_element(By.CSS_SELECTOR, selector)
                ActionChains(driver).scroll_to_element(iframe).perform()
                             
            #wyszukanie search bar'u i wprowadzenie nr lot
            
            try:
                element = shadow_element(
                            entry_selector = selector,
                            element_selector = 'div > div > input'
                                        )
            except AttributeError:
                element = driver.find_element(
                    By.XPATH, '''//*[@id="root"]/div/div/div[5]/div[2]
                    /div[7]/div/div/div[1]/div[1]/div[2]/div/div[1]/div[2]
                    /div/div[2]/input''')
            element.click()
            element.send_keys(i[2])
            element.send_keys(Keys.RETURN)
            time.sleep(2)
        
            #pobranie linka pliku CoA
            try:
                link = driver.find_element(
                    By.XPATH, '''//*[@id="certificates"]
                    /div[2]/div[1]/div[2]/span[1]''')
                
            except NoSuchElementException:
                link = driver.find_element(By.XPATH, 
                        '''//*[@id="root"]/div/div/div[5]
                        /div[2]/div[7]/div/div/div[1]/div[1]
                        /div[2]/div/div[2]/div[2]/div/span[1]
                        /a/span[2]''')
            
            window_before = driver.window_handles[0]
            link.click()
            window_after = driver.window_handles[1]
            driver.switch_to.window(window_after)
            String_url = driver.current_url
            driver.close()
            driver.switch_to.window(window_before)

            new_dir = 'output/' + header_text + '_' + 'nr_cat_' + i[1] + '/' + 'CoA/'
            
            try:
                os.makedirs(new_dir)

            except FileExistsError:
                pass
           
            url_path = new_dir + 'CoA_nr_lot_' + i[2]
            urllib.request.urlretrieve(String_url, url_path)

parser()           