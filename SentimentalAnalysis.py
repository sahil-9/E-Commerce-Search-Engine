from textblob import TextBlob
import xlrd
from xlutils.copy import copy

def sentimental_analysis():
    global positive_counter                                     #declarations
    global negative_counter
    global neutral_counter
    global semipos_counter
    global positive
    global negative
    global semipos
    global neutral
    global compound
    global count
    global senti
    global i
    global j
   
    fname = "E:\\Eclipse neon\\Eclipse python\\Final SDL\\src\\Test.xls"
                                                                #opening file containing review 
    book  = xlrd.open_workbook(fname)                           #opening xlsx file
    sheet = book.sheet_by_index(0)
    wb = copy(book)
    wsheet = wb.get_sheet(0)
    
    wsheet.write(0,14,'Positive Count')
    wsheet.write(0,15,'Negative Count')
    wsheet.write(0,16,'Neutral Count')
    wsheet.write(0,17,'Overall Sentiment')
    
    i=1
    for i in range(1,26):
        senti = 0
        positive = 0
        negative = 0
        neutral = 0
        positive_counter = 0
        negative_counter = 0
        neutral_counter = 0
        j=3
        for j in range(3,13):
            sentence = sheet.cell_value(i,j)
            if sentence.strip(" ") == "":
                break
            else:
                blob = TextBlob(sentence)
                senti = blob.sentiment.polarity
                
                if senti > 0:                                        #determining the polarity
                    positive = positive + senti
                    positive_counter+=1
                elif senti == 0:
                    neutral = neutral + senti
                    neutral_counter+=1
                elif senti < 0:
                    negative = negative + senti        
                    negative_counter+=1
                
            j+=1       
            
                
        wsheet.write(i,14,positive_counter)
        wsheet.write(i,15,negative_counter)
        wsheet.write(i,16,neutral_counter)
        if positive_counter > negative_counter and positive_counter >= neutral_counter:
            wsheet.write(i,17,"Positive")
        elif negative_counter > positive_counter and negative_counter >= neutral_counter:
            wsheet.write(i,17,"Negative")
        elif neutral_counter > positive_counter and neutral_counter > negative_counter:
            wsheet.write(i,17,"Neutral")
        elif positive_counter == negative_counter:
            wsheet.write(i,17,"Neutral")    
        else:
            wsheet.write(i,17,"No Reviews")
        
        
        i+=1
    
    wb.save('E:\\Eclipse neon\\Eclipse python\\Final SDL\\src\\Test.xls')
        
sentimental_analysis()
