
# imports
import ast  # parsing list from a string
import random  # for random time
import undetected_chromedriver as uc
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep


# webdriver configs
options = webdriver.ChromeOptions()
options.add_argument('--disable-infobars')
options.add_argument('--disable-web-security')
options.add_argument('--allow-running-insecure-content')
options.add_argument('--disable-popup-blocking')
options.add_argument('--disable-notifications')
options.add_argument('--headless')

import openpyxl
from openpyxl.styles import Alignment ,Font
import pandas as pd


class scrape():
    # init method
    def __init__(self):
        self.driver = uc.Chrome(options=options)
        self.driver.set_window_size(1920, 1080)

    def exit(self):
        self.driver.quit()

    def goto_url(self, url):
        self.driver.get(url)

    def scrape_table(self):
        data = self.driver.execute_script(''' 
        window.dataArray = []
        try{
        let data = document.getElementsByTagName('tr');
        for(let keys in data){const td = data[keys]; console.log(td);
         tabledata = td.querySelectorAll('td')
            if(tabledata[0] == undefined){
             console.log("it is undefined")
            }
            else{
                let name  = tabledata[0].innerText;
                let latest_previous_price = tabledata[1].innerText;
                let low_high = tabledata[2].innerText;

                // positive negative color
                let positive_negative_color  = tabledata[3].children[0].getAttribute('class');
                if(positive_negative_color == 'colorGreen'){
                    positive_negative = `${tabledata[3].innerText}G`
                }
                else if(positive_negative_color == 'colorRed'){
                    positive_negative = `${tabledata[3].innerText}R`
                }
                else{
                    positive_negative = tabledata[3].innerText
                }

                // time and date
                let time_date = tabledata[4].innerText

                // three month and color 

                let three_month_color = tabledata[5].children[0].getAttribute('class')
                if(three_month_color == 'colorGreen'){
                    three_month = `${tabledata[3].innerText}G`
                }
                else if(three_month_color == 'colorRed'){
                    three_month = `${tabledata[3].innerText}R`
                }
                else{
                    three_month = tabledata[3].innerText
                }
                
                // six month and color 

                let six_month_color = tabledata[6].children[0].getAttribute('class')
                if(six_month_color == 'colorGreen'){
                    six_month  = `${tabledata[3].innerText}G`
                }
                else if(six_month_color == 'colorRed'){
                    six_month = `${tabledata[3].innerText}R`
                }
                else{
                    six_month = tabledata[3].innerText
                }

                // one year and color 

                let one_year_color = tabledata[7].children[0].getAttribute('class')
                if(one_year_color == 'colorGreen'){
                    one_year  = `${tabledata[3].innerText}G`
                }
                else if(one_year_color == 'colorRed'){
                    one_year = `${tabledata[3].innerText}R`
                }
                else{
                    one_year = tabledata[3].innerText
                }

                let Array = [name,latest_previous_price,low_high,positive_negative,time_date,three_month,six_month,one_year];
                dataArray.push(Array)
            }
        }
        }catch(err){
            console.log(dataArray)
            // console.log(err)
        }
        finally{
        console.log(dataArray)
        return dataArray  

        }
        ''')

        for value in data:
            raw_data = {'NAME': [value[0]], 'LATEST PRICE/ PREVIOUS CLOSE': [value[1]], 'LOW / HIGH': [value[2]],
                        '+/-/%': [value[3]], 'TIME/DATE': [value[4]], '3 MO.+/-%': [value[5]], '6 MO.+/-%': [value[6]], '1 YEAR +/-%': [value[7]]}
            self.save_data(raw_data)

    def save_data(self, raw_data):
        data_frame = pd.DataFrame(raw_data, columns=['NAME', 'LATEST PRICE/ PREVIOUS CLOSE', 'LOW / HIGH', '+/-/%', 'TIME/DATE',
                                                     '3 MO.+/-%', '6 MO.+/-%', '1 YEAR +/-%'])
        try:
            read_csv = pd.read_csv('output.csv')
            empty = False
        except Exception as e:
            empty = True

    # checking if iles not exist create one and
    # add header and i already present not to add header
        if empty == True:
            data_frame.to_csv('output.csv', mode='a', index=False, header=True)
        else:
            data_frame.to_csv('output.csv', mode='a',
                              index=False, header=False)

    # appending data in csv
        data_frame.append(raw_data, ignore_index=True)

    def scrape_news(self):
        self.driver.get('https://www.marketwatch.com/')
        final_data = self.driver.execute_script('''
        data = document.getElementsByClassName('list list--bullets')[0]
        li_data = data.getElementsByTagName('li')
        final_data = Array.from(li_data).map(elem => elem.innerText)
        console.log(final_data)
        return final_data
        
        ''')
        for data in final_data:
            raw_data = {'NAME': [data], 'LATEST PRICE/ PREVIOUS CLOSE': '-', 'LOW / HIGH': '-',
                        '+/-/%': '-', 'TIME/DATE': '-', '3 MO.+/-%': '-', '6 MO.+/-%': '-', '1 YEAR +/-%': '-'}
            self.save_data(raw_data)
            
    

        print(final_data)

    def convert_xls(self,name):
        try:
            xsv = pd.read_csv(name)
            xls = pd.ExcelWriter(f'{name[:-4]}.xlsx')
            xsv.to_excel(xls, index=False)
            xls.save()
        except Exception as e:
            print("ouptut.csv not found")
        finally:
            sleep(2)
            self.style_excel('output.xlsx')


    def style_excel(self,filename):
        wb = openpyxl.load_workbook(filename)
        sheet = wb['Sheet1']
        wb.active = sheet

        align = Alignment(
        horizontal='center',
        vertical='center',
        text_rotation=0,
        wrap_text=True,
        shrink_to_fit=True,
        # indent=0
        )

        columns = ["A", "B", "C", "D", "E", "F", "G", "H"]

        for column in columns:
            sheet.column_dimensions[column].width = 15
            if column == "A":
                sheet.column_dimensions[column].width = 30
            else:
                pass
            for cell in sheet[column]:
                cell.alignment = align

    # color pallette
        my_color = openpyxl.styles.colors.Color(rgb="00F7E00F")
        my_color_fill = openpyxl.styles.fills.PatternFill(
        patternType="solid", fgColor=my_color
        )
        my_color_green = openpyxl.styles.colors.Color(rgb="0032CD32")
        my_color_red = openpyxl.styles.colors.Color(rgb="00FF3333")
    

    # adding colors to the text according to the columns
        colored_columns = ["D", "F", "G", "H"]
        for column in colored_columns:
            for cell in sheet[column]:
                if(cell.value == '-'):
                    pass
                else:
                    color_name = cell.value[-1:]
                    if(color_name == 'R'):
                        cell.font = Font(color=my_color_red)
                        cell.value = cell.value[:-1]
                    elif(color_name == 'G'):
                        cell.font = Font(color=my_color_green)
                        cell.value = cell.value[:-1]

        # coloring heading row
        first_rows = ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1"]

        for row in first_rows:
            sheet[row].fill = my_color_fill

        wb.save(filename)


if __name__ == "__main__":
    try:
        sc = scrape()
        for i in range(1, 12):
            url = f"https://markets.businessinsider.com/index/components/s&p_500?p={i}"
            sc.goto_url(url)
            sleep(4)
            # sc.scrape_table()
            # sc.scrape_news()
            
    except Exception as e:
        # sc.exit()
        print(e)
    finally:
        sc.convert_xls('output.csv')
        print("Done")
        sc.exit()
