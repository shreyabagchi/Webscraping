import bs4
from urllib.request import Request, urlopen
from bs4 import BeautifulSoup as soup
import xlsxwriter


reqList = ['https://www.selectscience.net/products/ortho-vision-analyzer/?prodID=197538#tab-1',
           'https://www.selectscience.net/products/vitros-250+350--chemistry-system/?prodID=209470',
           'https://www.selectscience.net/products/ortho-provue/?prodID=195222',
          ]

#my_url='https://www.selectscience.net/products/ortho-vision-analyzer/?prodID=197538#tab-1'
#uClient = uReq(my_url)

i=1
# Workbook() takes one, non-optional, argument  
# which is the filename that we want to create. 
workbook = xlsxwriter.Workbook('orthoPdtReviews.xlsx')
for reqline in reqList:
    req = Request(reqline, headers={'User-Agent': 'Mozilla/5.0'})
    webpage = urlopen(req).read()
    page_soup = soup(webpage, "html.parser")
    containers = page_soup.findAll("div",{"class":"col-md-12 landmark"})
    review=[]
    for container in containers:
        review.append(container.text)
   
    
  
    # The workbook object is then used to add new  
    # worksheet via the add_worksheet() method. 
    worksheet = workbook.add_worksheet("sheet"+ str(i)) 
    row=5
    col=0
    i=i+1
    # Use the worksheet object to write 
    # data via the write() method.
    headers="Product Name"
    worksheet.write('A1',headers) 
    worksheet.write('A2',page_soup.h1.text) 
    worksheet.write('A4','Reviews') 

    for inrev in review:
        worksheet.write(row,col,inrev)
        row=row+1
        
# Finally, close the Excel file 
# via the close() method. 
workbook.close()
