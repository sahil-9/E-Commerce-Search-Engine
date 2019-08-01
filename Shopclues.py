from bs4 import BeautifulSoup
import xlrd
import urllib.request as ur
import re
from xlutils.copy import copy
from _sqlite3 import Row

print('Starting program...')
fname = "E:\\Eclipse neon\\Eclipse python\\Final SDL\\src\\Test.xls"  # opening file containing review 
                                                                                # opening xlsx file
book = xlrd.open_workbook(fname)                           
sheet = book.sheet_by_index(0)
wb = copy(book)
wsheet = wb.get_sheet(0)

def searchprod(prod_name):
        
    title1 = []
    title = []
    links = []
    temp = []
    price = []
    temp1 = 0
    out2 = "Rs."
    j = 0
    k = 0
    
    prod_name = prod_name.replace(' ', '%20')
    url = "http://www.shopclues.com/search?q=" + prod_name 
    
    html = ur.urlopen(url).read()
    soup = BeautifulSoup(html, "html.parser")
    
    for link in soup.find_all('img'):
        t = link.get('title')
        title1.append(t)
        
    for i in range(len(title1)):
        name = title1[i]
        if name != None :
            title.append(name)
        else:
            print("in else")
    
    temp = soup.find_all('div', {'class' : 'img_section'})
    if(temp):
        for t1 in range(len(temp)):
            href1 = temp[t1].findParent('a')['href']
            links.append(href1)
    
    for p in soup.find_all('span', {'class' : 'p_price'}):
        temp1 = re.sub(out2, ' ', p.text)
        temp1 = temp1.strip(" ")
        price.append(temp1)
        
    r, c, l = 16, 0, 0
    for i in range(len(title)):
        if(l < 10):
            wsheet.write(r, c, title[i])
            wsheet.write(r, (c + 1), price[j])
            wsheet.write(r, (c + 2), links[k])
            item_reviews(r, (c + 3), links[k])
            r += 1
            j += 1
            k += 1
            l += 1
    
    wb.save('E:\\Eclipse neon\\Eclipse python\\Final SDL\\src\\Test.xls')
    
def item_reviews(row, col, item_url):
    html = ur.urlopen(item_url).read()
    soup = BeautifulSoup(html, "html.parser")
    reviews = []
    review = []
    count = 0
    o1 = "Was this review helpful"
    a = soup.find('div', {'class' : 'rnr_lists'})
    r = row 
    c = col       
        
    if(a):
        
        for link in a.findAll('p'):
            reviews.append(link.text)
        
        k = 0
        if(reviews):
            for i in range(len(reviews)):
                temp2 = reviews[i]
                temp2 = re.sub(o1, ' ', temp2)
                if temp2 != " ?":
                    review.append(temp2)
                else:
                    print("in append")
        
            for j in range(len(review)):
                if(j < 10):
                    wsheet.write(r, c, review[j])
                    count += 1
                    c += 1            
         
                wsheet.write(r, 13, int(count))
                
        else:
            print("no reviews")
            for j in range(3, 13):
                wsheet.write(r, j, ' ')
            wsheet.write(r, 13, 0) 
        
    else:
        print("no reviews")
        for j in range(3, 13):
            wsheet.write(r, j, ' ')
        wsheet.write(r, 13, 0) 
        
        
# searchprod('shirt') 
print("Ending program.....")       
    
