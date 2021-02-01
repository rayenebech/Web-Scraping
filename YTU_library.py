
#Exporting to Excel
import pandas as pd
from openpyxl import load_workbook
import requests
from bs4 import BeautifulSoup
index= requests.get('http://www.bologna.yildiz.edu.tr/index.php?r=program/bachelor')
soup= BeautifulSoup(index.text,'html.parser')
links= soup.select('a')
departments={}
size= len(links)
for i in range(35,size-2):
    departments[links[i].text]='http://www.bologna.yildiz.edu.tr'+links[i].get('href')
for x in departments:
    departement= requests.get(departments[x])
    depsoup= BeautifulSoup(departement.text,'html.parser')    
    courses= depsoup.select('#semesters')
    codes= courses[0].find_all('a')
    coursedict={}
    courselist=[]
    for code in codes:
        if code.get('href') != None:
            coursedict[code.text]='http://www.bologna.yildiz.edu.tr/'+code.get('href')
    for y in coursedict:    
        course= requests.get(coursedict[y])
        soup= BeautifulSoup(course.text,'html.parser')
        resources= soup.select('.textcontent li')
        coursedict[y] = [coursedict[y]]
        for book in resources:
            coursedict[y].append(book.text)
    
    df= pd.DataFrame.from_dict(coursedict,orient='index')
    book = load_workbook('E:\studentunion\Library\YtuLib.xlsx')
    writer = pd.ExcelWriter('E:\studentunion\Library\YtuLib.xlsx', engine = 'openpyxl')
    writer.book = book
    #with pd.ExcelWriter('E:\studentunion\Library\YtuLib.xlsx') as writer:  
    df.to_excel(writer, sheet_name='sheet', startrow=1)
    writer.save()
    writer.close()
    
    
    
    