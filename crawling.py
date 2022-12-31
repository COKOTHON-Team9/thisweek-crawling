import requests
from bs4 import BeautifulSoup
import pymysql
from selenium import webdriver
import openpyxl

wb = openpyxl.Workbook()
sheet = wb.active

request = requests.get("http://ticket.interpark.com/contents/Ranking?pType=W")

html = request.text

soup = BeautifulSoup(html, 'html.parser')
links = soup.select('ul > li > a')
title = soup.select(".prdName")
date = soup.select(".prdDuration")
image = soup.select('ul > li > a > img')


# for link in links:
#     if link.has_attr('href'):
#         if link.get('href').find('Ranking') != -1:
#             print(link)
genre_list = []
for link in links:
    if link.has_attr('href'):
        if link.get('href').find('Ranking') != -1:
            genre_list.append(link.text)
            #print(link.text)
print(genre_list)


title_list = []
date_list = []
for link in title:
    title_list.append(link.text)
    # print(link.text)

for link in date:
    date_list.append(link.text)
    # print(link.text)

print(title_list)
print(date_list)

musical_title_list = []
for i in range(0, 15):
    musical_title_list.append(title_list[i])
print(musical_title_list)

concert_title_list = []
for i in range(15, 30):
    concert_title_list.append(title_list[i])
print(concert_title_list)

role_title_list = []
for i in range(30, 45):
    role_title_list.append(title_list[i])
print(role_title_list)

classic_title_list = []
for i in range(45, 60):
    classic_title_list.append(title_list[i])
print(classic_title_list)

sports_title_list = []
for i in range(60, 67):
    sports_title_list.append(title_list[i])
print(sports_title_list)

leisure_title_list = []
for i in range(67, 82):
    leisure_title_list.append(title_list[i])
print(leisure_title_list)

exhibit_title_list = []
for i in range(82, 97):
    exhibit_title_list.append(title_list[i])
print(exhibit_title_list)



musical_date_list = []
for i in range(0, 15):
    musical_date_list.append(date_list[i])
print(musical_date_list)

concert_date_list = []
for i in range(15, 30):
    concert_date_list.append(date_list[i])
print(concert_date_list)

role_date_list = []
for i in range(30, 45):
    role_date_list.append(date_list[i])
print(role_date_list)

classic_date_list = []
for i in range(45, 60):
    classic_date_list.append(date_list[i])
print(classic_date_list)

sports_date_list = []
for i in range(60, 67):
    sports_date_list.append(date_list[i])
print(sports_date_list)

leisure_date_list = []
for i in range(67, 82):
    leisure_date_list.append(date_list[i])
print(leisure_date_list)

exhibit_date_list = []
for i in range(82, 97):
    exhibit_date_list.append(date_list[i])
print(exhibit_date_list)

#sheet.merge_cells('A4:A6')
sheet["A1"].value = "period"
sheet["B1"].value = "title"
sheet["C1"].value = "genre"
sheet["D1"].value = "picture"

for i in range(0, 15):
    section_period = "A" + str(i+2)
    sheet[section_period].value = musical_date_list[i]
for i in range(0, 15):
    section_title = "B" + str(i+2)
    sheet[section_title].value = musical_title_list[i]
for i in range(0, 15):
    section_genre = "C" + str(i+2)
    sheet[section_genre].value = "뮤지컬"

for i in range(0, 15):
    section_period = "A" + str(i+17)
    sheet[section_period].value = concert_date_list[i]
for i in range(0, 15):
    section_title = "B" + str(i+17)
    sheet[section_title].value = concert_title_list[i]
for i in range(0, 15):
    section_genre = "C" + str(i+17)
    sheet[section_genre].value = "콘서트"

for i in range(0, 15):
    section_period = "A" + str(i+32)
    sheet[section_period].value = role_date_list[i]
for i in range(0, 15):
    section_title = "B" + str(i+32)
    sheet[section_title].value = role_title_list[i]
for i in range(0, 15):
    section_genre = "C" + str(i+32)
    sheet[section_genre].value = "연극"

for i in range(0, 15):
    section_period = "A" + str(i+47)
    sheet[section_period].value = classic_date_list[i]
for i in range(0, 15):
    section_title = "B" + str(i+47)
    sheet[section_title].value = classic_title_list[i]
for i in range(0, 15):
    section_genre = "C" + str(i+47)
    sheet[section_genre].value = "클래식/무용"

for i in range(0, 7):
    section_period = "A" + str(i+62)
    sheet[section_period].value = sports_date_list[i]
for i in range(0, 7):
    section_title = "B" + str(i+62)
    sheet[section_title].value = sports_title_list[i]
for i in range(0, 7):
    section_genre = "C" + str(i+62)
    sheet[section_genre].value = "스포츠"

for i in range(0, 15):
    section_period = "A" + str(i+69)
    sheet[section_period].value = leisure_date_list[i]
for i in range(0, 15):
    section_title = "B" + str(i+69)
    sheet[section_title].value = leisure_title_list[i]
for i in range(0, 15):
    section_genre = "C" + str(i+69)
    sheet[section_genre].value = "레저"

for i in range(0, 15):
    section_period = "A" + str(i+84)
    sheet[section_period].value = exhibit_date_list[i]
for i in range(0, 15):
    section_title = "B" + str(i+84)
    sheet[section_title].value = exhibit_title_list[i]
for i in range(0, 15):
    section_genre = "C" + str(i+84)
    sheet[section_genre].value = "전시/행사"


image_list = []
for link in image:
    image_list.append(link['src'])

for i in range(97):
    section_picture = "D" + str(i+2)
    sheet[section_picture].value = image_list[13+i]

wb.save("crawling.xlsx")
wb.close()

