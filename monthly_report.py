from selenium import webdriver
import chromedriver_autoinstaller
import time
import xlsxwriter
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By

chromedriver_autoinstaller.install()  # Check if the current version of chromedriver exists
# and if it doesn't exist, download it automatically,
# then add chromedriver to path

driver = webdriver.Chrome()
driver.maximize_window()
driver.get("https://www.kireilife.net/")
time.sleep(2)
user_id = "kireilife1009"
password = "BrawnKona2018"

id_name = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div/div[2]/div/ul/li[5]/div[2]/form/div[2]/input')
id_name.send_keys(user_id)
id_pass = driver.find_element(By.XPATH,
                              '/html/body/div[1]/div/div[1]/div/div[2]/div/ul/li[5]/div[2]/form/div[2]/div[1]/input')
id_pass.send_keys(password)

log_in = driver.find_element(By.XPATH,
                             value='/html/body/div[1]/div/div[1]/div/div[2]/div/ul/li[5]/div[2]/form/div[2]/button')
log_in.click()
time.sleep(15)

driver.find_element(By.XPATH, value='/html/body/div[1]/div/div[1]/div/div[8]/div/ul/li[2]/a/img').click()
driver.find_element(By.XPATH, value='/html/body/div[1]/div/div[1]/div/div[8]/div/table[2]/tbody/tr[2]/td[2]/a').click()
time.sleep(3)

driver.find_element(By.XPATH, value='/html/body/div/div[2]/div[1]/button').click()
time.sleep(2)

c_number = []
c_edate = []
c_bill = []
c_tax = []
c_power = []

e_number = []
e_edate = []
e_power = []
e_bill = []
e_tax = []
e_sdate = []

#change the first number in () according to the User ID. You can find the num in read_me
for x in range(87,0,-1):
    path = '//*[@id="contractInfo'
    path += str(x)
    path += ' "]/table/tbody/tr[4]/td[1]'
    print(x)
    td = driver.find_element(By.XPATH, value= path)
    number = td.text
    if "/" in number:
        num = number.split()[0]
        c_number.append(num)
        button = '//*[@id="contractInfo'
        button += str(x)
        button += ' "]/table/tbody'
        time.sleep(1)
        element = driver.find_element(By.XPATH, value=button)
        driver.execute_script("arguments[0].click();", element)
        time.sleep(3)

        edate = driver.find_element(By.XPATH, value='//*[@id="tab1tbl"]/div/table/tbody/tr[3]/td[1]').text
        if edate == '2023/09':
            c_edate.append(edate)

            power = driver.find_element(By.XPATH, value='//*[@id="tab1tbl"]/div/table/tbody/tr[3]/td[3]').text
            c_power.append(power)

            bill = driver.find_element(By.XPATH, value='//*[@id="tab1tbl"]/div/table/tbody/tr[3]/td[4]').text
            c_bill.append(bill)

            tax = driver.find_element(By.XPATH,value='//*[@id="tab1tbl"]/div/table/tbody/tr[3]/td[5]').text
            c_tax.append(tax)
        else:
            edate = driver.find_element(By.XPATH, value='//*[@id="tab1tbl"]/div/table/tbody/tr[2]/td[1]').text
            c_edate.append(edate)

            power = driver.find_element(By.XPATH, value='//*[@id="tab1tbl"]/div/table/tbody/tr[2]/td[3]').text
            c_power.append(power)

            bill = driver.find_element(By.XPATH, value='//*[@id="tab1tbl"]/div/table/tbody/tr[2]/td[4]').text
            c_bill.append(bill)

            tax = driver.find_element(By.XPATH, value='//*[@id="tab1tbl"]/div/table/tbody/tr[2]/td[5]').text
            c_tax.append(tax)

    else:
        e_number.append(number)
        button = '//*[@id="contractInfo'
        button += str(x)
        button += ' "]/table/tbody//button'
        element = driver.find_element(By.XPATH, value=button)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", element)
        time.sleep(5)
        driver.find_element(By.XPATH, value='/html/body/div/div[2]/div[1]/div/div[2]/div[1]/div[2]/div/div').click()
        time.sleep(2)
        dropdown_xpath = "//*[@id='kenshinKounyuSelect']/select"
        dropdown_element = driver.find_element(By.XPATH, value= dropdown_xpath)
        dropdown = Select(dropdown_element)

        #change the date according to your need
        desired_option = '2023年9月分'

        # Check if the desired option is already selected
        if dropdown.first_selected_option.text == desired_option:
            time.sleep(1)
            sedate = driver.find_element(By.XPATH,
                                         value='/html/body/div/div[2]/div[2]/div[5]/div[3]/div/div/div[1]/div[2]/table[2]/tbody/tr[1]/td[1]').text
            b, c = sedate.split('～')
            e_sdate.append(b)
            e_edate.append(c)
            power = driver.find_element(By.XPATH,
                                        value='/html/body/div/div[2]/div[2]/div[5]/div[3]/div/div/div[1]/div[4]/table[4]/tbody/tr/td').text
            e_power.append(power)
            bill = driver.find_element(By.XPATH,
                                       value='/html/body/div[1]/div[2]/div[2]/div[5]/div[3]/div/div/div[1]/div[4]/table[5]/tbody/tr/td').text
            e_bill.append(bill)
            tax = driver.find_element(By.XPATH,
                                      value="//*[@id='F4']/table[6]//th[text()='うち消費税等相当額']/following-sibling::td").text
            e_tax.append(tax)
            time.sleep(2)
        else:
            dropdown.select_by_visible_text(desired_option)
            time.sleep(2)
            sedate = driver.find_element(By.XPATH,
                                         value='/html/body/div/div[2]/div[2]/div[5]/div[3]/div/div/div[1]/div[2]/table[2]/tbody/tr[1]/td[1]').text
            b, c = sedate.split('～')
            e_sdate.append(b)
            e_edate.append(c)
            power = driver.find_element(By.XPATH,
                                        value='/html/body/div/div[2]/div[2]/div[5]/div[3]/div/div/div[1]/div[4]/table[4]/tbody/tr/td').text
            e_power.append(power)
            bill = driver.find_element(By.XPATH,
                                       value='/html/body/div[1]/div[2]/div[2]/div[5]/div[3]/div/div/div[1]/div[4]/table[5]/tbody/tr/td').text
            e_bill.append(bill)
            tax = driver.find_element(By.XPATH,
                                      value="//*[@id='F4']/table[6]//th[text()='うち消費税等相当額']/following-sibling::td").text
            e_tax.append(tax)
            time.sleep(2)

    driver.find_element(By.XPATH, value='/html/body/div/div[2]/div[1]/button').click()
    time.sleep(2)
c_earn = []
c_earn.append(c_number)
c_earn.append(c_edate)
c_earn.append(c_power)
c_earn.append(c_bill)
c_earn.append(c_tax)


e_earn = []
e_earn.append(e_number)
e_earn.append(e_sdate)
e_earn.append(e_edate)
e_earn.append(e_power)
e_earn.append(e_bill)
e_earn.append(e_tax)

print(c_earn)
print(e_earn)
# print(result)
workbook = xlsxwriter.Workbook('result.xlsx')
worksheet1 = workbook.add_worksheet('Sheet1')
worksheet2 = workbook.add_worksheet('Sheet2')

row = 1
for col, data in enumerate(c_earn):
    worksheet1.write_column(row, col, data)

for col, data in enumerate(e_earn):
    worksheet2.write_column(row, col, data)

workbook.close()
