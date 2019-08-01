import xlwt
from bs4 import BeautifulSoup
import requests as re
import urllib.request as ur
import json

print('Starting program...')

workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('A Test Sheet')
worksheet.write(0, 0, 'Title')
worksheet.write(0, 1, 'Price')
worksheet.write(0, 2, 'Link')
worksheet.write(0, 3, 'R1')
worksheet.write(0, 4, 'R2')
worksheet.write(0, 5, 'R3')
worksheet.write(0, 6, 'R4')
worksheet.write(0, 7, 'R5')
worksheet.write(0, 8, 'R6')
worksheet.write(0, 9, 'R7')
worksheet.write(0, 10, 'R8')
worksheet.write(0, 11, 'R9')
worksheet.write(0, 12, 'R10')
worksheet.write(0, 13, 'total_count')

def search_prod(prod_name):
    prod_name = prod_name.replace(' ', '%20')
    
    html = ur.urlopen("https://www.flipkart.com/search?q=" + prod_name + "&otracker=start&as-show=on&as=off").read()
    soup = BeautifulSoup(html, "html.parser")

    price = []  # empty list for storing price
        
    priceClass = soup.findAll('div', {'class':'_1vC4OE'})
    if(priceClass):
        for cost in priceClass:
            if(len(price) < 15):
                price.append(cost.text)
            else:
                break
    else:
        print("error while getting price")
        
    r, c, i = 1, 0, 0
    data = soup.findAll('a', {'class':'_2cLu-l'})
    if(data):  # grid layout
        print("grid layout")
        for link in data:
            if(i < 15):
                href = 'https://www.flipkart.com' + link.get('href')  # link of the product
                worksheet.write(r, c, link.string)  # title of the product
                worksheet.write(r, (c + 1), price[i])
                worksheet.write(r, (c + 2), href)
                item_reviews(href, r, (c + 3))
                i += 1
                r += 1
                c = 0
            else:
                break
    else:  # list layout
        print("list layout")
        a = 0  
        title = []  # for list layout (e.g. microwave)
        titles = soup.findAll('div', {'class':'_3wU53n'})
        if(titles):
            for name in titles: 
                title.append(name.text)
        else:
            print("no title in list layout")
            
        links = soup.findAll('a', {'class':'_1UoZlX'})
        if(links):
            for link in links:
                if(i < 15):
                    href = 'https://www.flipkart.com' + link.get('href')  # link of the product
                    worksheet.write(r, c, title[a])
                    worksheet.write(r, (c + 1), price[i])
                    worksheet.write(r, (c + 2), href)
                    item_reviews(href, r, (c + 3))
                    a += 1
                    i += 1
                    r += 1
                    c = 0
                else:
                    break
        else:
            print("no links in list layout")
            
    workbook.save('E:\\Eclipse neon\\Eclipse python\\Final SDL\\src\\Test.xls');        
            
def item_reviews(item_url, row, col):
    html = ur.urlopen(item_url).read()
    soup = BeautifulSoup(html, "html.parser")
    a = soup.find('div', {'class': 'swINJg _3nrCtb'})
    r = row
    c = col
    if(a):
        link = a.findParent('a')['href']
        
        if(link):
            review_url = 'https://www.flipkart.com' + link
            retrive_reviews(review_url, row, col)
        
        else:
            print("no reviews")  # set total reviews = 0 in xl
            for i in range(3, 13):
                worksheet.write(r, i, ' ')
            worksheet.write(r, 13, 0)
    else:
        print("no reviews")  # set total reviews = 0 in xl
        for i in range(3, 13):
            worksheet.write(r, i, ' ')
        worksheet.write(r, 13, 0)


def retrive_reviews(url, row, col):
    productId = url[url.find("=") + 1:len(url)]
    headers = {
        'content-type': "application/json",
        'user-agent': "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
        'x-user-agent': "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36 FKUA/website/41/website/Desktop",
    }

    data = '{"requestContext":{"productId":"' + productId + '"}}'
    req = re.request("POST", "https://www.flipkart.com/api/3/page/dynamic/product-reviews", data=data, headers=headers)
    if(req.text):
        abc = json.loads(req.text)
        r = row
        c = col
        count = 0
        for i in range(0, 10):
            try :
                x1 = abc['RESPONSE']['data']['product_review_page_default_1']['data'][i]['value']['text']  # check this value and then set in column
                worksheet.write(r, c, x1)
                count += 1
                c += 1
            except (UnicodeEncodeError, IndexError):
                # worksheet.write(0, 0,"ERROR IN WRITING REVIEW TO THE FILE!!!!!!!!!")
                continue
        worksheet.write(r, 13, int(count))
    else:
        print("error while getting reviews")
        

# search_prod(prod_name=input('Enter your search query: '))
# search_prod("sneakers")

print('Program exited with no error...')
