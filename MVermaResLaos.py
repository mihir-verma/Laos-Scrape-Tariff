from datetime import datetime
import pandas as pd
import os

def linebr(n):
    str = '-'
    print(n*str)

def Hello():
    print()
    linebr(58)
    print("Welcome to Laos Controls Auto Extraction Program.\n")
    print('\nThis program is developed by MIHIR VERMA (VAD Office, IND)')
    print('\nPlease contact in case of any doubt/error\n(Mihir.Verma@thomsonreuters.com)\n\nThank You.')
    linebr(58)
    print('\n\npress any key to continue . . .')
    input('\n\n') 

def yesnoloop(str):
    while(True):
        inp = input(str)
        if inp=='y':
            break
        elif inp=='n':
            print('Complete the task first and re-run the program')
            return sys.exit()
        else:
            print('Give an appropriate response')
            continue

def link_maker(id,page):
    return 'https://www.laotradeportal.gov.la/index.php?r=searchMeasures/view&id='+str(id)+'&page='+str(page)

def to_eng(driver):
    driver.find_element_by_xpath('//*[@id="mainMbMenu"]/nav/div/div[1]/div/div/a').click()
    driver.find_element_by_xpath('//*[@id="mainMbMenu"]/nav/div/div[1]/div/div/div/div/ul/li[1]/a').click()

def timestamp():
    stamp = datetime.now().strftime('%Y-%m-%d %H:%M')
    timestamp_df = pd.DataFrame(columns=range(2))
    timestamp_df.loc[0] = ['Extraction Date:', stamp]
    return timestamp_df

def blanks(num):
    blanks_df = pd.DataFrame(index=range(num))
    return blanks_df

def get_text(element_list):
    text_list = []
    for element in element_list:
        text_list.append(element.text)
    return text_list
    
def find_pages(desc):
    per_page = int(desc[13:15])
    total_num = [int(s) for s in desc.split() if s.isdigit()]
    pages = -(-total_num[0] // per_page)
    return total_num, pages
    
def get_table(driver, table_xpath):

    table_element = driver.find_element_by_xpath(table_xpath)

    all_rows = table_element.find_elements_by_tag_name('tr')
    
    n_cols=0

    for row in all_rows:
        n = len(row.find_elements_by_xpath('./*'))
        if n>n_cols:
            n_cols=n

    table_out = pd.DataFrame(columns=range(n_cols))
    
    for i, row in enumerate(all_rows):
        elements = row.find_elements_by_xpath("./*")
        t_elements = get_text(elements)
        table_out.loc[i, list(range(len(t_elements)))] = t_elements
    
    return table_out
    
def output_files():
    path, dirs, files = next(os.walk('Output'))
    return files
    
def complete_msg():
    print()
    linebr(31)
    print('Program Completed Successfully!')
    linebr(31)
    print()