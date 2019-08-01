import xlrd
import xlutils
import matplotlib.pyplot as plt
from xlutils.copy import copy

print('program started')

fname = 'C:/Users/Admin/Desktop/Test.xls'
                                                         # opening file containing review 
book = xlrd.open_workbook(fname)  # opening xlsx file
sheet = book.sheet_by_index(0)
wb = copy(book)
wsheet = wb.get_sheet(0)
 
wsheet.write(0, 14, 'Positive Count')
wsheet.write(0, 15, 'Negative Count')
wsheet.write(0, 16, 'Neutral Count')
wsheet.write(0, 17, 'Overall Sentiment')
wsheet.write(0, 18, 'Positive_Percentage')
wsheet.write(0, 19, 'Negative_Percentage')
wsheet.write(0, 20, 'Neutral_Percentage')

plt.title('Analysis')

for i in range (1, 16): 
    try:  # for ZeroDivisionError error
        a = (sheet.cell_value(i, 14) / sheet.cell_value(i, 13)) * 100  # taking values from sheet and calculating percentage
        b = (sheet.cell_value(i, 15) / sheet.cell_value(i, 13)) * 100
        c = (sheet.cell_value(i, 16) / sheet.cell_value(i, 13)) * 100
    except ZeroDivisionError:
        wsheet.write(i, 18, 0)
        wsheet.write(i, 19, 0)
        wsheet.write(i, 20, 0)
        continue
     
    wsheet.write(i, 18, a)  # writing values
    wsheet.write(i, 19, b)
    wsheet.write(i, 20, c)
    
    if a != 0 and b != 0 and c != 0:
        sizes = [a, b, c]
        if a > b and a > c:  # comparing values
            explode = (0.1, 0, 0)   
        elif b > c and b > a:
            explode = (0, 0.1, 0)
        elif c > b and c > a:
            explode = (0, 0, 0.1)
        else:  # for zero values
            explode = (0, 0, 0)

        labels = 'Positive', 'Negative', 'Neutral'
        cols = ['r', 'b', 'g']

        plt.pie(sizes, colors=cols, shadow=True, explode=explode, startangle=-90 , autopct='%1.1f%%')  # specify labels hera as labels = labels
        plt.legend(labels, loc="upper right")
        plt.axis('equal')  # ensure proper axial circle(can be oval)
        plt.tight_layout()  # fit chart in centre          
        loc = "C:/Users/Admin/Desktop/product" + str(i) + ".png"
        plt.savefig(loc)
        plt.close()
    
    elif a == 0 or b == 0 or c == 0:
        if a == 0:
            if b > c:
                if c == 0:
                    sizes = [b]
                    explode = (0, 0)
                else:
                    sizes = [b, c]
                    explode = (0.1, 0)
            else:
                if b == 0:
                    sizes = [c]
                    explode = (0, 0)
                else:
                    sizes = [b, c]
                    explode = (0, 0.1)
            
            labels = 'Negative', 'Neutral'
            cols = ['b', 'g']
            sizes = [b, c]
            
            plt.pie(sizes, colors=cols, shadow=True, explode=explode, startangle=-90 , autopct='%1.1f%%')  # specify labels hera as labels = labels
            plt.legend(labels, loc="upper right")
            plt.axis('equal')  # ensure proper axial circle(can be oval)
            plt.tight_layout()  # fit chart in centre          
            loc = "C:/Users/Admin/Desktop/product" + str(i) + ".png"
            plt.savefig(loc)
            plt.close()
            
        if b == 0:
            if a > c:
                if c == 0:
                    explode = (0, 0)
                    sizes = [a]
                else:
                    sizes = [a, c]
                    explode = (0.1, 0)
            else:
                if a == 0:
                    sizes = [c]
                    explode = (0, 0)
                else:
                    sizes = [a, c]
                    explode = (0, 0.1)
            
            labels = 'Positive', 'Neutral'
            cols = ['r', 'g']
            
            
            plt.pie(sizes, colors=cols, shadow=True, explode=explode, startangle=-90 , autopct='%1.1f%%')  # specify labels hera as labels = labels
            plt.legend(labels, loc="upper right")
            plt.axis('equal')  # ensure proper axial circle(can be oval)
            plt.tight_layout()  # fit chart in centre          
            loc = "C:/Users/Admin/Desktop/product" + str(i) + ".png"
            plt.savefig(loc)
            plt.close()
            
        if c == 0:
            if a > b:
                if b == 0:
                    sizes = [a]
                    explode = (0, 0)
                else:
                    sizes = [a, b]
                    explode = (0.1, 0)
            else:
                if a == 0:
                    sizes = [b]
                    explode = (0, 0)
                else:
                    sizes = [a, b]
                    explode = (0, 0.1)
            
            labels = 'Positive', 'Negative'
            cols = ['r', 'b']
            sizes = [a, b]
            
            plt.pie(sizes, colors=cols, shadow=True, explode=explode, startangle=-90 , autopct='%1.1f%%')  # specify labels hera as labels = labels
            plt.legend(labels, loc="upper right")
            plt.axis('equal')  # ensure proper axial circle(can be oval)
            plt.tight_layout()  # fit chart in centre          
            loc = "C:/Users/Admin/Desktop/product" + str(i) + ".png"
            plt.savefig(loc)
            plt.close()
    
    elif (a + b + c) == 0:
        continue    

       
wb.save('C:/Users/Admin/Desktop/Test.xls')

print('program ended')
