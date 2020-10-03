import openpyxl as xl
from selenium import webdriver

wb = xl.load_workbook(
    'C:\\Users\\Mayur\\Documents\\Code\\Projects\\PC_Update\\PCbuild2.xlsx')
ws = wb.get_sheet_by_name('Details')


def Diff(li1, li2):
    return (list(list(set(li1) - set(li2)) + list(set(li2) - set(li1))))


class Part:
    def __init__(self, urls, cells):
        self.url = urls
        self.cell = cells
        self.driver = webdriver.Chrome(
            'C:\\Users\\Mayur\\Documents\\Code\\Tools\\chromedriver.exe')

    def price_update(self, url, cell):
        price = []

        for i in url:
            self.driver.get(i)
            block = self.driver.find_element_by_id('priceblock_ourprice')
            price.append(block.text)

        #TODO:Remove '₹' and ' ' from price and copy to revised
        revisedprice = price

        revisedprice.sort()

        ws[cell] = revisedprice[0]

        self.driver.close()


#TODO:Add More Parts
processer = Part([
    'https://www.amazon.in/Intel-i3-9100F-Processor-Discrete-Graphics/dp/B07R7Q3JZH'
], 'G4')
processer.price_update(processer.url, processer.cell)

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
gpu = Part([
    'https://www.amazon.in/MSI-GTX-1650-XS-4G/dp/B07QPVNPB3/ref=sr_1_1?crid=2L0P0JNIKXTXA&dchild=1&keywords=gtx+1650&qid=1601554684&sprefix=hard+disk+%2Caps%2C-1&sr=8-1',
    'https://www.amazon.in/GeForce-Phoenix-Overclocked-Graphics-PH-GTX1650-O4G/dp/B07QJDT7GR/ref=sr_1_2?crid=2L0P0JNIKXTXA&dchild=1&keywords=gtx+1650&qid=1601554684&sprefix=hard+disk+%2Caps%2C-1&sr=8-2',
    'https://www.amazon.in/GeForce-1-Click-128-bit-DIRECTX-Graphic/dp/B07QVLCWPF/ref=sr_1_3?crid=2L0P0JNIKXTXA&dchild=1&keywords=gtx+1650&qid=1601554684&sprefix=hard+disk+%2Caps%2C-1&sr=8-3',
    'https://www.amazon.in/Galax-GeForce-Click-GDDR6-Graphic/dp/B08FDYYMWX/ref=sr_1_4?crid=2L0P0JNIKXTXA&dchild=1&keywords=gtx+1650&qid=1601554684&sprefix=hard+disk+%2Caps%2C-1&sr=8-4',
    'https://www.amazon.in/Zotac-Gaming-Geforce-GDDR6-Graphics/dp/B086T66Z63/ref=sr_1_5?crid=2L0P0JNIKXTXA&dchild=1&keywords=gtx+1650&qid=1601554684&sprefix=hard+disk+%2Caps%2C-1&sr=8-5',
    'https://www.amazon.in/GeForce-128-bit-Gaming-Graphics-ZT-T16500F-10L/dp/B07QF1H9YR/ref=sr_1_8?crid=2L0P0JNIKXTXA&dchild=1&keywords=gtx+1650&qid=1601554684&sprefix=hard+disk+%2Caps%2C-1&sr=8-8',
    'https://www.amazon.in/MSI-GeForce-GTX-1650-OCV1/dp/B08GQ29HRP/ref=sr_1_10?crid=2L0P0JNIKXTXA&dchild=1&keywords=gtx+1650&qid=1601554684&sprefix=hard+disk+%2Caps%2C-1&sr=8-10'
], 'G7')
gpu.price_update(gpu.url, gpu.cell)

wb.save('Updated.xlsx')
