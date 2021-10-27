import os, time, win32clipboard, pywintypes
from csv import DictWriter
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

class DownloadClass:

    def __init__(self):

        self.flag = 1

        if not os.path.exists("./result.csv"):
            field_names = ['Playlist Link', 'Track URL', 'Source', 'Collection']
            dict={'Playlist Link': 'Playlist Link', 'Track URL': 'Track URL', 'Source': 'Source', 'Collection': 'Collection'}
            with open('./result.csv', 'a') as f_object:
                dictwriter_object = DictWriter(f_object, fieldnames = field_names)
                dictwriter_object.writerow(dict)
                f_object.close()

        self.browser = webdriver.Chrome(ChromeDriverManager().install())
        self.browser.maximize_window()
        self.other_browser = webdriver.Chrome(ChromeDriverManager().install())
        self.other_browser.maximize_window()
        txt = open("./Playlists.txt", "r", encoding="utf-8")
        page_list = txt.read().split("\n")
        txt.close()

        for page_url in page_list:
            print("\n--->", page_url)
            self.goto_page(page_url)
            field_names = ['Playlist Link', 'Track URL', 'Source', 'Collection']
            dict={'Playlist Link': '', 'Track URL': '', 'Source': '', 'Collection': ''}
            with open('./result.csv', 'a+', newline='') as f_object:
                dictwriter_object = DictWriter(f_object, fieldnames = field_names)
                dictwriter_object.writerow(dict)
                f_object.close()

        print("\nFinish Scraping\n")
    
    def goto_page(self, page_url):

        print("\nLoading...\n")
        self.browser.get(page_url)
        WebDriverWait(self.browser, 90).until(ec.visibility_of_element_located((By.CSS_SELECTOR, ".e8ea6a219247d88aa936a012f6227b0d-scss")))
        
        if self.flag:
            try:
                self.browser.find_element_by_css_selector(".banner-close-button").click()
            except:
                self.browser.find_element_by_css_selector("#onetrust-accept-btn-handler").click()
            self.flag = 0

        loop = 1

        while loop:

            try:
                item = self.browser.find_element_by_xpath(f"//span[text()='{loop}']")
                item = item.find_element_by_xpath("..")
            except:
                break

            ActionChains(self.browser).context_click(item).perform()
            share = self.browser.find_element_by_xpath('//div[@id="context-menu"]/ul/li[8]')    
            ActionChains(self.browser).move_to_element(share).perform()
            time.sleep(0.1)
            self.browser.find_element_by_xpath('//div[@id="context-menu"]/ul/li[8]/div/ul/li[1]').click()
            self.OpenClipboardWithEvilRetries()
            link = win32clipboard.GetClipboardData()
            win32clipboard.CloseClipboard()
            print("Track Link:", link)

            ActionChains(self.browser).context_click(item).perform()
            self.browser.find_element_by_xpath('//div[@id="context-menu"]/ul/li[5]').click()
            WebDriverWait(self.browser, 90).until(ec.visibility_of_element_located((By.CSS_SELECTOR, '.GenericModal')))
            source = self.browser.find_element_by_css_selector(".c083bbfb4d7a747031cad5f3e881d043-scss").text.split(": ")[-1]
            self.browser.find_element_by_css_selector("._088c26b2bc9b5b413ea463c89333da3a-scss").click()
            print("Source:", source)

            row = item.find_element_by_xpath("..").find_element_by_xpath("..")
            name = row.find_element_by_css_selector(".da0bc4060bb1bdb4abb8e402916af32e-scss").text
            artist_url = row.find_element_by_css_selector("._966e29b71d2654743538480947a479b3-scss a").get_attribute("href")
            self.other_browser.get(artist_url)
            try:
                WebDriverWait(self.other_browser, 30).until(ec.visibility_of_element_located((By.CSS_SELECTOR, ".e8ea6a219247d88aa936a012f6227b0d-scss")))
                found = self.other_browser.find_element_by_xpath(f'//div[text()="{name}"]').find_element_by_xpath("..").find_element_by_xpath("..").find_element_by_xpath("..")
                try:
                    collection = found.find_element_by_css_selector(".d47b790d001ed769adcd9ddfc0e83acc-scss").text
                except:
                    collection = ""
            except:
                collection = ""
            print("Collection:", collection)

            field_names = ['Playlist Link', 'Track URL', 'Source', 'Collection']
            dict={'Playlist Link': page_url, 'Track URL': link, 'Source': source, 'Collection': collection}
            with open('./result.csv', 'a+', newline='', encoding="utf-8") as f_object:
                dictwriter_object = DictWriter(f_object, fieldnames = field_names)
                dictwriter_object.writerow(dict)
                f_object.close()
                
            loop = loop + 1

    def OpenClipboardWithEvilRetries(self, retries=10, delay=0.1):
        while True:
            try:
                return win32clipboard.OpenClipboard()
            except pywintypes.error as e:
                if e.winerror!=5 or retries==0:
                    raise
                retries = retries - 1
                time.sleep(delay)

if __name__ == "__main__":
    DownloadClass()