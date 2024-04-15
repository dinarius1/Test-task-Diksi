'''
Второй способ решение задачи с помощью библиотеки RPA (Robotic process automation). Более простой и понятный способ решения, но производительность кода по сравнению с первым способом ниже. В файле results.png указано за сколько выполняется код - 5189 милисекунд.
'''

import rpa as r
import pandas as pd

# Инициализируем РПА процесс
r.init(visual_automation=True, chrome_browser=True, turbo_mode=True)

# Читаем данные из файла
df= pd.read_excel('challenge.xlsx')


r.url('https://www.rpachallenge.com/')
r.wait(10)
r.click('//button[text()="Start"]')

# Проходимся по каждой строке из файла, и заполняем сразу же формы в соотвествии с названием элемента на веб-странице
for index,row in df.iterrows():
    r.type('//input[@ng-reflect-name="labelFirstName"]',row['First Name'])
    r.type('//input[@ng-reflect-name="labelLastName"]',row['Last Name '])
    r.type('//input[@ng-reflect-name="labelCompanyName"]',row['Company Name'])
    r.type('//input[@ng-reflect-name="labelRole"]',row['Role in Company'])
    r.type('//input[@ng-reflect-name="labelAddress"]',row['Address'])
    r.type('//input[@ng-reflect-name="labelEmail"]',row['Email'])
    r.type('//input[@ng-reflect-name="labelPhone"]',str(row['Phone Number']))
    r.click('//input[@value="Submit"]')


r.snap('/html/body/app-root/div[2]','results.png')
r.close()
