import os
import time
import traceback

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import createTicketFunctions
import numpy as np

import result_excel


class Automation():
    """Go to IDQ and analyze the CMID, download the needed asins with given types. Then move on the selection if needed
    and download the requested attributes. Finaly go to sim ticketing website and create the ticket."""


    def selenium(self):
        profile_name = 'Mail_Automation'
        profile_path = rf'{os.path.expanduser("~")}\Chrome Profiles\{profile_name}'
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument(f'user-data-dir={profile_path}')
        chrome_options.add_argument(f'--profile-directory={profile_name}')
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        chrome_options.add_experimental_option("detach", True)
        browser = webdriver.Chrome( options=chrome_options)
        browser.maximize_window()
        return browser

    def Idq(self,webdriver,MCID,ticketType):
        """Go to IDQ and download the files according to the ticketType and CMID. The reason that there are
        that many "try" statement is that these catalog improvemnets might not exists in all of the given MCID's"""

        webdriver.get(f'https://idq-hawk.a2z.com/search?aggregationLevel=asin&asin_type=Retail&asin_type=FBA&asin_type=3P&c=ASIN&c=MKT&c=TITL&c=IDQGV2&c=IDQSV2&c=LEAF&c=KYWR&c=QUAR&c=INDX&c=WEBA&c=IMG&c=BRAN&c=BITS&c=THCL&c=HBP&c=HDES&c=A%2B&c=ZMAI&c=BPNO&c=QIMG&c=IAI&c=RAIC4&c=RAIC10&c=AVRA&c=QREV&c=GVS&c=IMPS&c=INVE&c=ORD&c=REPL&c=MNFC&marketplace_id=338851&merchant_customer_id={MCID}&showPercentages=true&size=1000')
        WebDriverWait(webdriver, 100).until(EC.presence_of_element_located((By.ID, "gradeAnalyserButton")))
        webdriver.find_element(By.ID, "gradeAnalyserButton").click()
        time.sleep(5)

        if ticketType =='leafNode':
            try:
                WebDriverWait(webdriver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="asinsCount-grading_has_leaf_node"]')))
                asinCount=webdriver.find_element(By.ID,'asinsCount-grading_has_leaf_node').text.replace(',','')
                webdriver.find_element(By.XPATH,'//*[@id="check-grading_has_leaf_node"]').click()
                webdriver.find_element(By.XPATH,'//*[@id="gradeAnalyserDownload"]').click()
                time.sleep(5)
            except:
                print('No leafNode found')
                asinCount=""

        elif ticketType == "keyword":
            try:
                WebDriverWait(webdriver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="asinsCount-grading_has_keywords"]')))
                keywordCount=webdriver.find_element(By.XPATH,'//*[@id="asinsCount-grading_has_keywords"]').text.replace(',','')
                webdriver.find_element(By.XPATH,'//*[@id="check-grading_has_keywords"]').click()
                webdriver.find_element(By.XPATH,'//*[@id="gradeAnalyserDownload"]').click()
                time.sleep(5)
                return keywordCount
            except:
                print('No keyword found')
                asinCount=""

        elif ticketType == "bulletpoints":
            bpCount=0
            try:
                WebDriverWait(webdriver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="check-Catalog-Improvements"]')))
                time.sleep(2)
                try:
                    print("Looking for Add 3 or more bullet points ")
                    bpCount+=int(webdriver.find_element(By.XPATH,'//*[@id="asinsCount-grading_has_bullet_points"]').text.replace(',',''))
                    webdriver.find_element(By.XPATH,'//*[@id="check-grading_has_bullet_points"]').click()
                    print("YES, 3 or more bullet points have been found")
                except Exception as e:
                    print("NO, 3 or more bullet points have not been found")
                try:
                    print("Looking for Add 2 or more bullet points ")
                    bpCount+=int(webdriver.find_element(By.XPATH,'//*[@id="asinsCount-grading_has_1_bullet_point"]').text.replace(',',''))
                    webdriver.find_element(By.XPATH,'//*[@id="check-grading_has_1_bullet_point"]').click()
                    print("YES, 2 or more bullet points have been found")
                except Exception as e:
                    print("NO, 2 or more bullet points have not been found")
                try:
                    print("Looking for Add 1 or more bullet points ")
                    bpCount+=int(webdriver.find_element(By.XPATH,'//*[@id="asinsCount-grading_has_2_bullet_points"]').text.replace(',',''))
                    webdriver.find_element(By.XPATH,'//*[@id="check-grading_has_2_bullet_points"]').click()
                    print("YES, 1 or more bullet points have been found")
                except:
                    print("NO, 1 or more bullet points have not been found")

                print(bpCount)
                webdriver.find_element(By.XPATH,'//*[@id="gradeAnalyserDownload"]').click()
                time.sleep(5)
                return bpCount
            except:
                print('No bullet point found')
                asinCount=""
        elif ticketType == "title":
            try:
                WebDriverWait(webdriver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="asinsCount-grading_has_brand_in_title_start"]')))
                titleCount=webdriver.find_element(By.XPATH,'//*[@id="asinsCount-grading_has_brand_in_title_start"]').text.replace(',','')
                webdriver.find_element(By.XPATH,'//*[@id="check-grading_has_brand_in_title_start"]').click()
                webdriver.find_element(By.XPATH,'//*[@id="gradeAnalyserDownload"]').click()
                time.sleep(5)
                return titleCount
            except:
                print('No title found')
                asinCount=""

        elif ticketType == "image":
            try:
                WebDriverWait(webdriver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="asinsCount-grading_has_main_img"]')))
                keywordCount=webdriver.find_element(By.XPATH,'//*[@id="asinsCount-grading_has_main_img"]').text.replace(',','')
                webdriver.find_element(By.XPATH,'//*[@id="check-grading_has_main_img"]').click()
                webdriver.find_element(By.XPATH,'//*[@id="gradeAnalyserDownload"]').click()
                time.sleep(5)
                return keywordCount
            except:
                print('No add main image found')
                asinCount=""
        elif ticketType == "product_description":
            try:
                WebDriverWait(webdriver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="asinsCount-grading_has_description_or_aplus"]')))
                keywordCount=webdriver.find_element(By.XPATH,'//*[@id="asinsCount-grading_has_description_or_aplus"]').text.replace(',','')
                webdriver.find_element(By.XPATH,'//*[@id="check-grading_has_description_or_aplus"]').click()
                webdriver.find_element(By.XPATH,'//*[@id="gradeAnalyserDownload"]').click()
                time.sleep(5)
                return keywordCount
            except:
                print('No add main image found')
                asinCount=""


        return asinCount

    def selection(self,webdriver,ticketType):
        """Fill asin column, select MP, adjust display columns and go to search
        In the next page select actions and start the export task and go to task
        In final page wait till the status is completed and download the results"""

        #Checking if the files are downloaded
        file_names = os.listdir("downloads")
        while len(file_names) < 1:
            file_names = os.listdir("downloads")
            time.sleep(2)
        #find asin
        df=pd.read_csv('downloads/latest_snapshot.csv', encoding='UTF-16', sep='\t')
        asin_data=df['ASIN']
        asins_list=''
        for asin in asin_data:
            asins_list=asins_list+" "+asin
        if len(asins_list) > 500:
            middle_index=len(asins_list)//2
            asin_first_half=asins_list[:middle_index]
            asins_second_half=asins_list[middle_index:]

        #To have atleast 2 asins to add Selection ASIN section
        if len(asins_list) < 11:
            asins_list=asins_list+asins_list


        #set reset
        webdriver.get('https://selection.amazon.com/')
        WebDriverWait(webdriver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[text()="Select marketplaces"]')))
        webdriver.find_element(By.XPATH, '//span[text()="Reset"]').click()

        #set asin DONT FORGET TO CHANGE THIS
        if len(asins_list)>500:
            time.sleep(2)
            webdriver.find_element(By.XPATH,'//*[contains(@placeholder, "asin1, asin2")]').send_keys(asin_first_half)
            time.sleep(2)
            webdriver.find_element(By.XPATH,'//*[contains(@placeholder, "asin1, asin2")]').send_keys(asins_second_half)
        else:
            webdriver.find_element(By.XPATH,'//*[contains(@placeholder, "asin1, asin2")]').send_keys(asins_list)


        #set MP
        webdriver.find_element(By.XPATH,'//span[text()="Select marketplaces"]').click()
        WebDriverWait(webdriver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@title="www.amazon.com.tr"]')))
        webdriver.find_element(By.XPATH,'//*[@title="www.amazon.com.tr"]').click()

        #Close dropdown
        webdriver.find_element(By.XPATH,'//div[@data-test-id="HomePage__Content"]').click()

        #set columns
        webdriver.find_element(By.CLASS_NAME,'ColumnsPicker__EditButton').click()
        time.sleep(2)
        clear_all_elements=webdriver.find_elements(By.XPATH,'//span[text()="Clear all"]')
        clear_all_elements[1].click()

        #pick column according to the ticketType
        if ticketType=='bulletpoints':
            webdriver.find_element(By.XPATH, '//span[text()="Start typing the attribute name"]').click()
            webdriver.find_element(By.CLASS_NAME, 'AttributesPicker').find_element(By.TAG_NAME, 'input').send_keys('bullet_point')
            WebDriverWait(webdriver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//span[@data-value="bullet_point"]')))
            webdriver.find_element(By.XPATH, '//span[@data-value="bullet_point"]').click()
            #Set product description column
            webdriver.find_element(By.XPATH, '//span[text()="Start typing the attribute name"]').click()
            webdriver.find_element(By.XPATH, '//span[text()="Start typing the attribute name"]').click()
            webdriver.find_element(By.CLASS_NAME, 'AttributesPicker').find_element(By.TAG_NAME, 'input').send_keys('product_description')
            WebDriverWait(webdriver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//span[@data-value="product_description"]')))
            webdriver.find_element(By.XPATH, '//span[@data-value="product_description"]').click()

        #set item_name column
        if ticketType=='title':
            webdriver.find_element(By.XPATH, '//span[text()="Start typing the attribute name"]').click()
            webdriver.find_element(By.CLASS_NAME, 'AttributesPicker').find_element(By.TAG_NAME, 'input').send_keys(
                'item_name')
            WebDriverWait(webdriver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//span[@data-value="item_name"]')))
            webdriver.find_element(By.XPATH, '//span[@data-value="item_name"]').click()

            #set Brand name column
            webdriver.find_element(By.XPATH, '//span[text()="Start typing the attribute name"]').click()
            webdriver.find_element(By.XPATH, '//span[text()="Start typing the attribute name"]').click()
            webdriver.find_element(By.CLASS_NAME, 'AttributesPicker').find_element(By.TAG_NAME, 'input').send_keys(
                'brand')
            WebDriverWait(webdriver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//span[@data-value="brand"]')))
            webdriver.find_element(By.XPATH, '//span[@data-value="brand"]').click()

        #Close drop down
        webdriver.find_element(By.XPATH, '//span[text()="Choose columns"]').click()

        #click Apply
        webdriver.find_element(By.XPATH, '//span[text()="Apply"]').click()

        #click search
        webdriver.find_element(By.XPATH,'//span[text()="Search"]').click()

        #Click actions and export search results
        time.sleep(10)
        WebDriverWait(webdriver, 10).until(EC.element_to_be_clickable((By.XPATH, '//span[text()="Actions"]')))
        webdriver.find_element(By.XPATH,'//span[text()="Actions"]').click()
        WebDriverWait(webdriver, 10).until(EC.presence_of_element_located((By.XPATH, '//span[text()="Export search results"]')))
        webdriver.find_element(By.XPATH,'//span[text()="Export search results"]').click()

        #Click the export task
        WebDriverWait(webdriver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[contains(@href, "/export/")]')))
        webdriver.find_element(By.XPATH,'//*[contains(@href, "/export/")]').click()

        #Wait till export is completed
        WebDriverWait(webdriver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@data-test-id="ExportStatusIndicator"]')))
        status_element=webdriver.find_element(By.XPATH,'//*[@data-test-id="ExportStatusIndicator"]')
        #Waiting 6 minutes to  export to be completed
        WebDriverWait(webdriver, 300).until(EC.presence_of_element_located((By.XPATH,'//span[text()="Completed"]')))
        webdriver.find_element(By.XPATH,'//span[text()="Result"]').click()

    def find_sku_sources(self,ticketType):
        """Takes asins from latest snapshot file. Then finds the same asins in the amazon inventory file and takes the SKU's."""

        #Checking if the file is downloaded
        file_names = os.listdir("downloads")
        while len(file_names) < 1:
            file_names = os.listdir("downloads")
            time.sleep(2)

        #Taking asins from latest Snapshot file
        df=pd.read_csv('downloads/latest_snapshot.csv', encoding='UTF-16', sep='\t')
        asin_data=df['ASIN']
        asin_list=[]

        for asin in asin_data:
            asin_list.append(asin)

        #Finding SKU's of the asins inside the amazon inventory file
        df=pd.read_excel('Needed resources/Amazon_inventory.xlsx')
        matching_asins=df[df['asin'].isin(asin_list)]
        matched_values=matching_asins.values
        asin_sku_dict={}
        for row in matched_values:
            asin_sku_dict[row[1]]=str(row[0])
        print(asin_sku_dict)

        if ticketType=="image":
            #Finding images that are needed for the asins and attaching them to the dictionary
            asin_image_dict={}
            df = pd.read_excel('Needed resources/Trendyol.xlsx')
            for asin, sku in asin_sku_dict.items():
                print(f'checking sku and asins are {sku} {asin}')
                sku=[sku]
                matching_skus = df[df['Barkod'].isin(sku)]
                matched_values=matching_skus.values
                if len(matched_values) !=0:
                    images=matched_values[0][19:25]
                    asin_image_dict[asin]=images

            #Finding the latests snapshot and attaching the images
            df = pd.read_csv('downloads/latest_snapshot.csv',encoding='UTF-16', sep='\t')

            # Adding columns
            df['Images1'] = np.nan
            df['Images2'] = np.nan
            df['Images3'] = np.nan
            df['Images4'] = np.nan
            df['Images5'] = np.nan
            df['Images6'] = np.nan

            for asin, image_list in asin_image_dict.items():
                asin=[asin]
                matching_asins = df[df['ASIN'].isin(asin)]

                #finding index and adding images to the column
                index=matching_asins.index[0]
                count=6
                for image in image_list:
                    df.iloc[index,count]=image
                    count += 1
            #Create new excel file and return the asin count
            df.to_excel('downloads/Image Links.xlsx',index=False)
            return len(asin_image_dict)
        elif ticketType=='product_description':
            #Finding product descriptions that are needed for the asins and attaching them to the dictionary
            asin_productDescription_dict={}
            df = pd.read_excel('Needed resources/Trendyol.xlsx')
            for asin, sku in asin_sku_dict.items():
                sku=[sku]
                matching_skus = df[df['Barkod'].isin(sku)]
                matched_values=matching_skus.values
                if len(matched_values) !=0:
                    productDescription=matched_values[0][12]
                    asin_productDescription_dict[asin]=productDescription

            #Finding the latests snapshot and attaching the product description
            df = pd.read_csv('downloads/latest_snapshot.csv',encoding='UTF-16', sep='\t')

            # Adding columns
            df['Product Description'] = np.nan

            for asin, productDescription in asin_productDescription_dict.items():
                asin=[asin]
                matching_asins = df[df['ASIN'].isin(asin)]

                #finding index and adding product descrtiption to the column
                index=matching_asins.index[0]
                df.iloc[index,6]=productDescription
            #Create new excel file and return the asin count
            df.to_excel('downloads/Product Descriptions.xlsx',index=False)
            return len(asin_productDescription_dict)

    def removeFile(self):
        """Remove the files in the downloads directory"""
        file_names=os.listdir("downloads")
        file_paths=[os.path.abspath(os.path.join("downloads",file_name)) for file_name in file_names]
        for file in file_paths:
            os.remove(file)

    def takeMCID(self):
        df=pd.read_excel('Needed resources/CID.xlsx')
        CID_SellerName_dict={}
        for index, row in df.iterrows():
            CID_SellerName_dict[row[0]]=row[1]
        return CID_SellerName_dict

    def process_automation(self):
        MCID_list=self.takeMCID()
        webdriver=self.selenium()
        #Making sure the directory is empty
        self.removeFile()
        df=result_excel.df
        ticketType_list=["leafNode","keyword","bulletpoints","title","image","product_description"]

        MCID_count=0
        for MCID, sellerName in MCID_list.items():

            df=result_excel.add_row(MCID,df)
            print(f"Working on the CID {MCID_count+1}")
            for position in range(1,len(ticketType_list)+1):
                try:

                    ticketType=ticketType_list[position-1]
                    print(f"Working on {MCID} and ticketType is {ticketType}")
                    asinCount=self.Idq(webdriver,MCID,ticketType)
                    if asinCount != "":
                        if ticketType=="bulletpoints" or ticketType=="title":
                            self.selection(webdriver,ticketType)
                            #print('filler')
                        elif ticketType=="image" or ticketType=='product_description':
                            asinCount=self.find_sku_sources(ticketType)
                        createTicketFunctions.createTicket(webdriver,MCID,asinCount,sellerName,ticketType,MCID_count,position,df)
                        print('CYCLE ENDED MOVING ON')
                    else:
                        reason=f"Could not find { ticketType} for the MCID {MCID}, in IDQ, moving on!"
                        print(reason)
                        result_excel.add_no(MCID_count,position,df)
                        result_excel.add_reason(MCID_count, position+6,df,reason)
                except Exception as E:
                    error=f'Error encountered, moving to the nex ticket type.{traceback.print_exc()}'
                    print(E)
                    traceback.print_exc()
                    result_excel.add_error(MCID_count, position, df)
                    result_excel.add_reason(MCID_count, position+6, df,error)
                    time.sleep(10)

                self.removeFile()
            MCID_count += 1
        result_excel.to_excel(df)
        result_excel.paint_excel()



if __name__ == "__main__":
    obj = Automation()
    obj.process_automation()
