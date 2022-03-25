import string
import re
import pdfquery
import os
import pypyodbc

directory=r'C:\Users\lwetzel\Documents\ClientWork\Commercial\GP\Antron PDF Carpet Scrape\Mannington\PDF'
        
def connect_to_database():
    #START CONNECTION TO DATABASE
    print("Before Connecting To Database")
    cnxn = pypyodbc.connect(r"Driver={SQL Server};Server=COMMERCIAL\COMMERCIAL;Database=i360_Invista_Fibre_Scrape;Trusted_Connection=yes;")
    cur = cnxn.cursor()
    print("Connected to database")
    return cur
   

def get_value_for_text(string_to_search_for,pdf):
    dimension=''
    try:
        label = pdf.pq('LTTextLineHorizontal:contains("'+string_to_search_for+'")')
        left_corner = float(label.attr('x0'))
        bottom_corner = float(label.attr('y0'))

        dimension = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner-15, left_corner+500, bottom_corner+20)).text().replace(string_to_search_for,'')
        print(string_to_search_for+" : "+dimension)
        return dimension
    except:
        print('Could Not find '+ string_to_search_for)
        return ''

        

def get_value_for_text_shortened(string_to_search_for,pdf):
    dimension=''

    try:
        label = pdf.pq('LTTextLineHorizontal:contains("'+string_to_search_for+'")')
        left_corner = float(label.attr('x0'))
        bottom_corner = float(label.attr('y0'))

        dimension = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner-15, left_corner+200, bottom_corner+20)).text().replace(string_to_search_for,'')
        print(string_to_search_for+" : "+dimension)
        return dimension
    except:
        print('Could Not find '+ string_to_search_for)
        return ''


def main():
    cur =connect_to_database()
    
    for filename in os.listdir(directory):
        full_file_name=directory+'\\'+filename
        pdf = pdfquery.PDFQuery(full_file_name)
        pdf.load()

        # Product_code=''
        Backing_size=''
        Fiber=''
        Dye_method=''
        Stitches=''
        Pile_Density=''
        Total_Weight=''

        
        strings_to_search_for=['PRIMARY BACKING','FACE FIBER','DYE METHOD']
        for string_to_search_for in strings_to_search_for:
            if string_to_search_for=='PRIMARY BACKING' :
                    Backing_size=get_value_for_text(string_to_search_for,pdf)   

            if string_to_search_for=='FACE FIBER':
                    Fiber=get_value_for_text(string_to_search_for,pdf)                   

            if string_to_search_for=='DYE METHOD':
                    Dye_method=get_value_for_text(string_to_search_for,pdf)

        #print(Product_code)

        strings_to_search_for=['STITCHES PER INCH','DENSITY','TUFTED YARN WEIGHT']
        for string_to_search_for in strings_to_search_for:
            if string_to_search_for=='STITCHES PER INCH':
                Stitches=get_value_for_text_shortened(string_to_search_for,pdf)    

            if string_to_search_for=='DENSITY':
                Pile_Density=get_value_for_text_shortened(string_to_search_for,pdf)                  

            if string_to_search_for=='TUFTED YARN WEIGHT':
                Total_Weight=get_value_for_text_shortened(string_to_search_for,pdf)

            

        print('\n')

main()
