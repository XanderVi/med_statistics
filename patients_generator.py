'''
Standart cells:
A2, A3... - number of patient (1, 2, 3...)
B2, B3... - visit date (dd.mm.yyyy)
C2, C3... - type of visit (1, 2 or 3)
D2, D3... - consultation type (1, 2, or 3)
E2, E3... - full name of the patient
F2, F3... - gender (male or female)
G2, G3... - date of birth (dd.mm.yyyy)
H2, H3... - phone number
I2, I3... - ICPC-2-E code
J2, J3... - MKB-10 code
K2, K3... - ICPC-2-E process code
L2, L3... - category (1-11)
'''

import random as r
import openpyxl

name = ['Артем', 'Богдан', 'Валерий', 'Дмитрий', 'Евгений', 
        'Кирилл', 'Михаил', 'Николай', 'Петр', 'Ярослав',
        'Анна', 'Виктория', 'Евгения', 'Людмила', 'Мария',
        'Екатерина', 'Елена', 'Валентина', 'Татьяна', 'Инна']
surname = ['Хлебопек', 'Бортко', 'Заец', 'Тетеренко', 'Карандаш',
           'Розуменко', 'Голуб', 'Сербул', 'Нестеренко', 'Кресал', 
           'Петросян', 'Алешко', 'Микитюк', 'Парсон', 'Железняк']
gender = ['чол.', 'жiн.']
phone_codes = ['050', '063', '068', '073', '093', '099']
letters = [chr(i) for i in range(65, 65+26)]

main_wb = openpyxl.load_workbook('main.xlsx')
sheet = main_wb['Sheet1']

for i in range(10000):
    sheet['A'+str(i+3)] = str(i+1)
    date_1 = r.randint(1,28)
    month_1 = r.randint(1,12)
    date_1 = str(date_1) if date_1 > 9 else '0' + str(date_1)
    month_1 = str(month_1) if month_1 > 9 else '0' + str(month_1)
    sheet['B'+str(i+3)] = date_1+'.'+month_1+'.'+'2018'
    sheet['C'+str(i+3)] = str(r.randint(1,3))
    sheet['D'+str(i+3)] = str(r.randint(1,3))
    sheet['E'+str(i+3)] = str(r.choice(surname))+' '+str(r.choice(name))
    sheet['F'+str(i+3)] = str(r.choice(gender))
    date_2 = r.randint(1,28)
    month_2 = r.randint(1,12)
    date_2 = str(date_2) if date_2 > 9 else '0' + str(date_2)
    month_2 = str(month_2) if month_2 > 9 else '0' + str(month_2)
    sheet['G'+str(i+3)] = date_2+'.'+\
                          month_2+'.'+\
                          str(r.randint(1918,2017))
    sheet['H'+str(i+3)] = str(r.choice(phone_codes))+'-'+\
                          str(r.randint(0,9))+\
                          str(r.randint(0,9))+\
                          str(r.randint(0,9))+'-'+\
                          str(r.randint(0,9))+\
                          str(r.randint(0,9))+'-'+\
                          str(r.randint(0,9))+\
                          str(r.randint(0,9))
    sheet['I'+str(i+3)] = str(r.choice(letters))+str(r.randint(0,99))+' '+\
                          str(r.choice(letters))+str(r.randint(0,99))+' '+\
                          str(r.choice(letters))+str(r.randint(0,99))
    sheet['J'+str(i+3)] = str(r.choice(letters))+str(r.randint(0,99))+' '+\
                          str(r.choice(letters))+str(r.randint(0,99))+'.'+str(r.randint(0,9))
    sheet['K'+str(i+3)] = str(r.randint(30,60))+' '+\
                          str(r.randint(30,60))+' '+\
                          str(r.randint(30,60))
    sheet['L'+str(i+3)] = str(r.randint(0,10))

main_wb.save('main.xlsx')