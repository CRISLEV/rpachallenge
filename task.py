import os
from RPA.Browser.Selenium import Selenium
from datetime import datetime
from calendar import monthrange
from dateutil.relativedelta import relativedelta
from openpyxl import Workbook
import re
import urllib.request

browser_lib = Selenium()

search_phrase = os.getenv("SEARCH_PHRASE", None)
news_section = os.getenv("NEWS_SECTION", None)
months_for_search = int(os.getenv("MONTHS_FOR_SEARCH", None))

def open_the_website(url):
    browser_lib.open_available_browser(url)


def enter_search_phrase():
    search_bar = browser_lib.find_element('//button[@data-test-id="search-button"]')
    browser_lib.click_button(search_bar)
    search_input = browser_lib.find_element('//input[@data-testid="search-input"]')
    browser_lib.input_text(search_input, search_phrase)
    browser_lib.press_keys(search_input, "ENTER")


def apply_filters():
    # Section filter
    multiselect_btn = browser_lib.find_element('//button[@data-testid="search-multiselect-button"]')
    browser_lib.click_button(multiselect_btn)
    multiselect_items = browser_lib.find_elements('//li[@class="css-1qtb2wd"]')
    for item in multiselect_items:
        if (news_section.lower() in item.text.lower()):
            browser_lib.click_button(item)
            break

    # Time filter
    input_dt = datetime.today()
    search_date_btn = browser_lib.find_element('//button[@data-testid="search-date-dropdown-a"]')
    browser_lib.click_button(search_date_btn)
    specific_dates_btn = browser_lib.find_element('//button[@value="Specific Dates"]')
    browser_lib.click_button(specific_dates_btn)
    start_date_input = browser_lib.find_element('//input[@data-testid="DateRange-startDate"]')
    browser_lib.input_text(start_date_input, get_start_date(input_dt, months_for_search))
    end_date_input = browser_lib.find_element('//input[@data-testid="DateRange-endDate"]')
    browser_lib.input_text(end_date_input, get_end_date(input_dt))
    browser_lib.press_keys(end_date_input, "ENTER")


def get_start_date(current_date, months):
    format_string = "%m/%d/%Y"
    if (months<=1):
        res = current_date.replace(day=1)
        return res.strftime(format_string)
    else:
        res = current_date - relativedelta(months=(months - 1))
        res = res.replace(day=1)
        return res.strftime(format_string)
    

def get_end_date(current_date):
    lastday = monthrange(current_date.year, current_date.month)[1]
    enddate = datetime(year=current_date.year, month= current_date.month, day=lastday)
    format_string = "%m/%d/%Y"
    return enddate.strftime(format_string)

def display_info():
    try:
        search_date_btn = browser_lib.find_element('//button[@data-testid="search-show-more-button"]')
        while(browser_lib.does_page_contain_button(search_date_btn)):
            browser_lib.click_button(search_date_btn)
            search_date_btn = browser_lib.find_element('//button[@data-testid="search-show-more-button"]')
    finally:
        return
    

def get_news_info():
    news_info = []
    search_results = browser_lib.find_elements('//li[@class="css-1l4w6pd"]')
    i=0
    while i < len(search_results):
        try:
            title_res = browser_lib.find_elements('//h4[@class="css-2fgx4k"]',search_results[i])
            description_res = browser_lib.find_elements('//p[@class="css-16nhkrn"]',search_results[i])
            date_res = browser_lib.find_elements('//span[@class="css-17ubb9w"]',search_results[i])
            # img src
            img_loc = browser_lib.find_elements('//img[@class="css-rq4mmj"]',search_results[i])
            img_src = browser_lib.get_element_attribute(img_loc[i],"src")
            img_src_arr = re.findall("\/.+\.jpg", img_src)
            img_name = img_src_arr[0].split("/")
            download_image(img_src, os.path.dirname(__file__)+'/images/', img_name[len(img_name)-1])
            # search phrase coincidences
            title_desc = title_res[i].text + description_res[i].text
            search_phrase_coincidences = re.findall(search_phrase.lower(), title_desc.lower())
            # info contains amount of money
            contains_amount_money = re.search("\$\d{1,3}(\.\d+)*|\$\d{1,3}(,\d{3})*(\.\d+)*|\d dollars|\d usd", title_desc.lower())

            news_info.append({
                "title": title_res[i].text,
                "description": description_res[i].text if description_res[i].text!=None else "",
                "date": date_res[i].text,
                "img": img_name[len(img_name)-1],
                "coincidences" : len(search_phrase_coincidences),
                "amountmoney": "True" if contains_amount_money else "False"
            })
        finally:
            i+=1

    return news_info
        

def download_image(url, file_path, file_name):
    try:
        full_path = file_path + file_name
        urllib.request.urlretrieve(url, full_path)
    finally:
        return


def create_excel_file(news_info):
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "Title"
    sheet["B1"] = "Description"
    sheet["C1"] = "Date"
    sheet["D1"] = "Img name"
    sheet["E1"] = "Search Coincidences!"
    sheet["F1"] = "Info has amount of money"
    cell = 2
    for info in news_info:
        sheet[f"A{cell}"] = info["title"]
        sheet[f"B{cell}"] = info["description"]
        sheet[f"C{cell}"] = info["date"]
        sheet[f"D{cell}"] = info["img"]
        sheet[f"E{cell}"] = info["coincidences"]
        sheet[f"F{cell}"] = info["amountmoney"]
        cell += 1
        
    workbook.save(filename=os.path.dirname(__file__)+"/news_info.xlsx")


def store_screenshot(filename):
    browser_lib.screenshot(filename=filename)
    

def main():
    try:
        open_the_website("www.nytimes.com")
        enter_search_phrase()
        apply_filters()
        display_info()
        news_info = get_news_info()
        create_excel_file(news_info)
        store_screenshot("output/screenshot.png")
    finally:
        browser_lib.close_all_browsers()


# Call the main() function, checking that we are running as a stand-alone script:
if __name__ == "__main__":
    main()