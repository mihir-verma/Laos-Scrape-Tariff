import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import MVermaResLaos as M

M.Hello()
M.yesnoloop('Have you created a folder \'Output\\Laos\\\'? (y/n) ')

try:
    init = int(input('\nGive the index from where to start extraction\n(Leave blank for default): '))
except:
    init = 1
    
service = webdriver.chrome.service.Service(r'C:/Program Files (x86)/SeleniumWrapper/chromedriver.exe')
service.start()
chrome_options = Options()
chrome_options.headless = True
chrome_options.add_argument('--lang=en-US')
chrome_options.add_argument('--start-fullscreen')
#chrome_options.add_argument('--log-level=3')
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Remote(service.service_url, desired_capabilities = chrome_options.to_capabilities())
print('\nGetting Laos Main Page')
driver.get('https://www.laotradeportal.gov.la/index.php?r=searchMeasures/index')
print('Selecting Language as English')
M.to_eng(driver)

print('\nFetching All Measure Types')
select_drop =  Select(driver.find_element_by_id('MeasuresSearchForm_measuretypes'))
all_options = M.get_text(select_drop.options)

for i in range(init,len(all_options)): #,len(all_options)):
    
    print('\nSelecting Measure Type Index', i)
    select_drop =  Select(driver.find_element_by_id('MeasuresSearchForm_measuretypes'))    
    select_drop.select_by_index(i)

    submit = driver.find_element_by_xpath('//*[@id="_measures_search_form"]/div[2]')
    submit.submit()
    
    driver2 = webdriver.Remote(service.service_url, desired_capabilities = chrome_options.to_capabilities())
    
    desc = driver.find_element_by_xpath('//*[@id="measures-grid"]/div[1]').text
    
    all_ids = ['Ids:']
    sub_min_table = pd.DataFrame(columns=[0,2])
    sub_min_table.loc[0] = 'Sub-Agency Name', 'Enforced By'
    total_hs = ['No. of HS Codes']

    if desc == "":
        print('\nNo Data on Page')
        total_num=[0]
        pass
    else:
        total_num, pages = M.find_pages(desc)
    
        for j in range(1,pages+1): #(1,pages+1):
            
            print('\nWorking on page', j, '(Measure Type', i, ')')
            if pages!=1:
            
                try:
                    id_list_element = driver.find_element_by_xpath('//*[@id="measures-grid"]/div[3]')
                except:
                    id_list_element = driver.find_element_by_xpath('//*[@id="measures-grid"]/div[2]')
                    
                page_nav = driver.find_element_by_xpath('//*[@id="yw0"]')
                bi = page_nav.find_element_by_partial_link_text(str(j))
                driver.execute_script('arguments[0].click();', bi)
                
                wait = WebDriverWait(driver,25)
                element = wait.until(EC.staleness_of(id_list_element))
            
            print('Fetching the IDs information')
            try:
                id_list_element = driver.find_element_by_xpath('//*[@id="measures-grid"]/div[3]')
            except:
                id_list_element = driver.find_element_by_xpath('//*[@id="measures-grid"]/div[2]')
            
            sub_min_table_pagewise = M.get_table(driver, '//*[@id="measures-grid"]/table')
            sub_min_table_pagewise.drop([1,3,4,5,6], axis=1, inplace=True)
            sub_min_table_pagewise.drop(0, inplace=True)
            sub_min_table = pd.concat([sub_min_table, sub_min_table_pagewise], ignore_index=True)
             
            id_elements = id_list_element.find_elements_by_xpath('./*')
            ids = []
            hs_codes = []
            for id_element in id_elements:
                id = id_element.get_attribute('innerHTML')
                ids.append(id)
                
                print('\nOpening page for ID', id, 'for fetching Total HS Codes information')
                driver2.get(M.link_maker(id,1))
                print('Selecting English Option')
                M.to_eng(driver2)
                
                desc2 = driver2.find_element_by_xpath('//*[@id="commodity-description-list"]/div[1]').text
                if desc2 == "":
                    total_num2=[0]
                else:
                    total_num2, _p = M.find_pages(desc2)
                
                print('Information Fetched')
                hs_codes.append(total_num2[0])
                
            all_ids = all_ids + ids
            total_hs = total_hs + hs_codes     

    ministry_df = pd.DataFrame(columns=range(4))
    ministry_df[[0,1]] = sub_min_table
    ministry_df[2] = all_ids  
    ministry_df[3] = total_hs  
    
    print('\nSaving the Output File')
    path=r'../Output/Laos/Lao_MeasureType_'+str(i)+'.xlsx' 
    writer = pd.ExcelWriter(path, engine='openpyxl') 
    try:
        writer.book = load_workbook(path)
        try:
            writer.book.remove(writer.book['Main'])
            writer.save()
        except:
            pass
    except:   
        pass    
    
    final_df = pd.concat([M.timestamp(), M.blanks(1), pd.Series([all_options[i]]),pd.Series([str(total_num[0])+' total Sub-Ministries']), M.blanks(1),ministry_df], ignore_index=True)
    final_df.to_excel(writer, header=False,index=False, sheet_name='Main')
    writer.save()
    
    writer.close()
    
    print('\n'+M.timestamp().loc[0,1]+'\n')
    
driver.quit()
M.complete_msg()