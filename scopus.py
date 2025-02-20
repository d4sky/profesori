import time
from pathlib import Path

import pandas as pd

from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

import time

from bs4 import BeautifulSoup, Tag

import requests

import sys

sys.path.append('C:/Python/')
#sys.path.append('/mnt/e/python/')
from nutils import ExcelTable

#headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"}

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    # Add more headers if needed
}

def Write_page(content, path_for_outfile):
    with open(path_for_outfile, "w", encoding="utf-8") as file:
        file.write(content)
    print(f"Page saved as '{path_for_outfile}'")

def Get_page(url_link, out_file_path, mode = "request", sTime = 10):
    # request-simple, request-headers, request-session
    # selenium-simple, selenium-options
    print(f"Getting content from URL:")
    print(url_link)
    
    if   mode.startswith("request"):
        if   "headers" in mode:
            response = requests.get(url_link, headers=headers)
        elif "session" in mode:
            session = requests.Session()
            session.headers.update(headers)
            response = session.get(url_link)
        else: # "simple" in mode
            response = requests.get(url_link)

        if response.status_code == 200:
            Write_page(response.text, out_file_path)
        else:
            print(f"Failed to retrieve the page. Status code: {response.status_code}")
            
    elif mode.startswith("selenium"):
        if "options" in mode:
            options = Options()
            options.headless = True
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")

            #driver = webdriver.Chrome(options=options)
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        else:
            driver = webdriver.Chrome()
            
        driver.get(url_link)
        time.sleep(sTime) 

        Write_page(driver.page_source, out_file_path)

        driver.quit()

def Search_author(meno, priezvisko, out_file):
    driver = webdriver.Chrome()

    #driver.get("https://www.scopus.com/freelookup/form/author.uri?zone=TopNavBar&origin=NO%20ORIGIN%20DEFINED")
    driver.get("https://www.scopus.com/freelookup/form/author.uri")

    # Fill in a form field
    input_field = driver.find_element(By.NAME, "searchterm1")
    input_field.send_keys(priezvisko)

    input_field = driver.find_element(By.NAME, "searchterm2")
    input_field.send_keys(meno)

    # Submit the form (if there's a button)
    #submit_button = driver.find_element(By.NAME, "submit_button_name")
    #submit_button.click()

    # Press RETURN to submit
    input_field.send_keys(Keys.RETURN)

    time.sleep(2)  # Adjust based on response time
    Write_page(driver.page_source, out_file)

    driver.quit()

def find_exact_classes(tag, class_list):
    return tag.has_attr('class') and set(tag['class']) == set(class_list)

def parse_html_content(html_or_path, tag, search_terms, atrID=''):
    if isinstance(html_or_path, Tag):
        #soup = BeautifulSoup(html_or_path, "html.parser")  # Parse HTML string
        soup = html_or_path
    else:
        with open(html_or_path, "r", encoding="utf-8") as file:
            soup = BeautifulSoup(file, "html.parser")  # Parse from file
        
    if atrID == '':    
        # Find all <a> elements with class "docTitle"
        if len(search_terms) == 1:
            elements = soup.find_all(tag, class_=search_terms[0])
        else:
            elements = soup.find_all(lambda tag: find_exact_classes(tag, search_terms))
    else:
        elements = soup.find_all(tag, {atrID: search_terms[0]})
    
    return elements
    
def Get_author_page(inp_file_path, out_file_pattern):
    links = parse_html_content(inp_file_path, 'a', ["docTitle"])
    # Extract href 
    outIDs = []
    for link in links:
        href = link.get("href")
        items = href.split('&')
        if len(items)>1:
            pairs = items[0].split('ID=')
            if len(pairs)>1:
                authorID = pairs[1]
                print(f"Trying... ID {authorID}")
                url = f"https://www.scopus.com/authid/detail.uri?authorId={authorID}"
                print(url)
                #out_file_path = out_file_pattern.format(len(outIDs) + 1)
                out_file_path = out_file_pattern.format(authorID)
                Get_page(url, out_file_path, "selenium-options", 2)
                outIDs.append(authorID)
        else:
            print("Link is wrong")
            
    return outIDs
         
def Parse_for_hindex(inp_file_path, cid = 1):
    clss = ["Typography-module__lVnit", 
            "Typography-module__ix7bs",
            "Typography-module__Nfgvc",
            "Button-module__Imdmt"
            ]
            
    cls = clss[cid]
    out_list = []
    spans = parse_html_content(inp_file_path, "span", [cls])
    for span in spans:
        out_list.append(span.text.strip())
    return out_list    
    
def Check_and_get_first_element(to_check_object, tag_name, item_list, atribute_id=''):    
    elems = parse_html_content(to_check_object, tag_name, item_list, atribute_id)
    ret_value = ""
    if len(elems) > 0: ret_value = elems[0]
    return ret_value    
    
def Parse_for_affiliation(inp_file_path):
    affil_element = Check_and_get_first_element(inp_file_path, "span", ["authorInstitution"], "data-testid")
    if affil_element:
        instit_element = Check_and_get_first_element(affil_element, "span", ["Typography-module__lVnit", "Typography-module__Nfgvc", "Button-module__Imdmt"])
        if instit_element: instit_element = instit_element.text
        
        place_element = Check_and_get_first_element(affil_element, "span", ["Typography-module__lVnit", "Typography-module__Nfgvc"])
        if place_element: place_element = place_element.text
        
        return  instit_element + place_element
    else:
        print("No records for Affilation found")
        return "No Affiliation"

def Get_first_n_values(values, n=1):
    ret_values = []
    for i in range(n):
        if len(values) > i: 
            ret_values.append(values[i])
        else:
            ret_values.append("na")
    return tuple(ret_values) 

rok = "2024_1"
zoznam_path = f"zoznam_{rok}.xlsx"
udaje_path  = f"udaje_{rok}.xlsx"

zoznam = ExcelTable(zoznam_path)
udaje  = ExcelTable(udaje_path)

zoznam.Add_empty_column("Found", False)
udaje.Add_empty_columns(["ScopusID", "Meno", "Priezvisko", "Citations", "Documents", "H-index", "Datum", "Affiliation"], False)

udaje.Just_save()
zoznam.Just_save()

data_directory = Path(f"data_{rok}")
data_directory.mkdir(parents=True, exist_ok=True)

for index, row in zoznam.row_iterator():
    print("x"*100)
    print(row)
    print("-"*100)
    meno       = row["Meno"]
    priezvisko = row["Priezvisko"]
    search_page_path = f"data_{rok}/{meno}_{priezvisko}_s.html"
    author_page_pattern = f"data_{rok}/{meno}_{priezvisko}" + "_{}.html"

    if not Path(search_page_path).exists():
        print(f"Going to search for: {meno} {priezvisko}")
        Search_author(meno, priezvisko, search_page_path)
        
    if Path(search_page_path).exists():
        print(f"File exists! Going to parse {meno} {priezvisko}")
        if pd.isna(row["Found"]):
            print(f"Processing {meno} {priezvisko}")
            authorsIDs = Get_author_page(search_page_path, author_page_pattern)
            print('#'*111)
            print(authorsIDs)
            print('#'*111)
            
            zoznam.DF.at[index, "Found"] = ','.join(authorsIDs)
            zoznam.Just_save()
        else:
            print(f"{meno} {priezvisko} has already IDs: {row["Found"]}")
            authorsIDs = str(row["Found"]).split(',')

        for ai, aID in enumerate(authorsIDs):
            authorID = int(float(aID))
            #if not udaje.DF['ScopusID'].isin([authorID]).any():
            #if not udaje.DF['ScopusID'].astype(int).isin([authorID]).any():
            #if udaje.DF['ScopusID'].apply(int).isin([authorID]).any():
            #idxs = udaje.index[udaje['ScopusID'].astype(int) == authorID].tolist()
            idxs = [i for i, sID in enumerate(udaje.DF['ScopusID']) if int(sID) == authorID]
            H_indexes = udaje.DF.loc[idxs, "H-index"].tolist()
            if len(idxs) == 0 or (len(H_indexes) == 1 and H_indexes[0]=='na'):
                if len(H_indexes) == 1 and H_indexes[0]=='na':
                    pass
                author_page_path = author_page_pattern.format(authorID)
                if Path(author_page_path).exists():
                    print(f"Will get the information about {meno} {priezvisko} {authorID}")
                    items = Parse_for_hindex(author_page_path, 1)
                    citations, documents, h_index = Get_first_n_values(items, 3)
                    datum = datetime.today().strftime("%Y-%m-%d")  # Format: YYYY-MM-DD
                    affiliation = Parse_for_affiliation(author_page_path)

                    print(citations, documents, h_index, affiliation)
                    new_row = [authorID, meno, priezvisko, citations, documents, h_index, datum, affiliation]
                    udaje.Append_row(new_row)
                    udaje.Just_save()
    else:
        pass
    print()

sys.exit()



