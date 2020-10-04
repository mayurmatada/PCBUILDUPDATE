import openpyxl as xl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import threading

wb = xl.load_workbook(
    'C:\\Users\\Mayur\\Documents\\Code\\Projects\\PC_Update\\PCbuild2.xlsx')
ws = wb.get_sheet_by_name('Details')
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1920x1080")


def Diff(li1, li2):
    return (list(list(set(li1) - set(li2)) + list(set(li2) - set(li1))))


class Part:
    def __init__(self, urls, cells):
        self.url = urls
        self.cell = cells
        self.driver = webdriver.Chrome(
            'C:\\Users\\Mayur\\Documents\\Code\\Tools\\chromedriver.exe',
            options=chrome_options)

    def price_update(self, url, cell):
        global ws
        price = []

        for i in url:
            self.driver.get(i)
            block = self.driver.find_element_by_id('priceblock_ourprice')
            price.append(block.text)
        revisedprice = []

        for a in price:
            k = a.replace('₹ ', '')
            revisedprice.append(k)

        revisedprice.sort()
        for j in cell:
            ws[j] = revisedprice[0]

        self.driver.close()


processer = Part([
    'https://www.amazon.in/Intel-i3-9100F-Processor-Discrete-Graphics/dp/B07R7Q3JZH'
], ['G4', 'C4', 'E4'])

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
gpu = Part([
    'https://www.amazon.in/MSI-GTX-1650-XS-4G/dp/B07QPVNPB3/ref=sr_1_1?crid=2L0P0JNIKXTXA&dchild=1&keywords=gtx+1650&qid=1601554684&sprefix=hard+disk+%2Caps%2C-1&sr=8-1',
    'https://www.amazon.in/GeForce-Phoenix-Overclocked-Graphics-PH-GTX1650-O4G/dp/B07QJDT7GR/ref=sr_1_2?crid=2L0P0JNIKXTXA&dchild=1&keywords=gtx+1650&qid=1601554684&sprefix=hard+disk+%2Caps%2C-1&sr=8-2',
    'https://www.amazon.in/GeForce-1-Click-128-bit-DIRECTX-Graphic/dp/B07QVLCWPF/ref=sr_1_3?crid=2L0P0JNIKXTXA&dchild=1&keywords=gtx+1650&qid=1601554684&sprefix=hard+disk+%2Caps%2C-1&sr=8-3',
    'https://www.amazon.in/Galax-GeForce-Click-GDDR6-Graphic/dp/B08FDYYMWX/ref=sr_1_4?crid=2L0P0JNIKXTXA&dchild=1&keywords=gtx+1650&qid=1601554684&sprefix=hard+disk+%2Caps%2C-1&sr=8-4',
    'https://www.amazon.in/Zotac-Gaming-Geforce-GDDR6-Graphics/dp/B086T66Z63/ref=sr_1_5?crid=2L0P0JNIKXTXA&dchild=1&keywords=gtx+1650&qid=1601554684&sprefix=hard+disk+%2Caps%2C-1&sr=8-5',
    'https://www.amazon.in/GeForce-128-bit-Gaming-Graphics-ZT-T16500F-10L/dp/B07QF1H9YR/ref=sr_1_8?crid=2L0P0JNIKXTXA&dchild=1&keywords=gtx+1650&qid=1601554684&sprefix=hard+disk+%2Caps%2C-1&sr=8-8',
    'https://www.amazon.in/MSI-GeForce-GTX-1650-OCV1/dp/B08GQ29HRP/ref=sr_1_10?crid=2L0P0JNIKXTXA&dchild=1&keywords=gtx+1650&qid=1601554684&sprefix=hard+disk+%2Caps%2C-1&sr=8-10'
], ['G7'])
#------------------------------------------------------------------------------------------------
ram = Part([
    'https://www.amazon.in/DDR4-2400-MHZ-Year-Warranty/dp/B07RF9TJXF/ref=sr_1_6?crid=2L0P0JNIKXTXA&dchild=1&keywords=ram+2400mhz+crucial%3B&qid=1601813557&sprefix=hard+disk+%2Caps%2C-1&sr=8-6',
    'https://www.amazon.in/Crucial-CT4G4DFS824A-2400-MTps-288-Pin/dp/B019FRDFU0/ref=sr_1_5?crid=2L0P0JNIKXTXA&dchild=1&keywords=ram+2400mhz+crucial%3B&qid=1601813557&sprefix=hard+disk+%2Caps%2C-1&sr=8-5',
    'https://www.amazon.in/Crucial-8gb-Ddr4-Memory-Module/dp/B01BIWK9OK/ref=sr_1_4?crid=2L0P0JNIKXTXA&dchild=1&keywords=ram+2400mhz+crucial%3B&qid=1601813557&sprefix=hard+disk+%2Caps%2C-1&sr=8-4'
], ['G5', 'E5', 'C5'])
#--------------------------------------------------------------------------------------------
motherboard = Part([
    'https://www.amazon.in/Intel-Micro-Motherboards-H310M-R2-0/dp/B07GNQK5NP'
], ['C6', 'E6', 'G6'])
#-------------------------------------------------------------------------------------------
case = Part([
    'https://www.amazon.in/Ant-Esports-ICE-130AG-Motherboard-Preinstalled/dp/B08D6G6LMK/ref=sr_1_15?crid=X1ZQL00JUEGL&dchild=1&keywords=computer+case&qid=1601814412&refinements=p_36%3A200000-&rnid=1318502031&s=computers&sprefix=Compu%2Ccomputers%2C279&sr=1-15',
    'https://www.amazon.in/Antec-NX130-Cabinet-Computer-Preinstalled/dp/B08FHTX1MD/ref=sr_1_19?crid=X1ZQL00JUEGL&dchild=1&keywords=computer+case&qid=1601814412&refinements=p_36%3A200000-&rnid=1318502031&s=computers&sprefix=Compu%2Ccomputers%2C279&sr=1-19',
    'https://www.amazon.in/CHIPTRONEX-X410B-GAMING-CABINET-WITHOUT/dp/B07D2BB1L4/ref=sr_1_11?dchild=1&keywords=computer+cabinets&pd_rd_r=d4136365-9cf7-4cd6-9dd7-31b8a5f144e3&pd_rd_w=bLh6G&pd_rd_wg=DEisK&pf_rd_p=e98d1fad-4664-4b4d-a930-8932601ebadf&pf_rd_r=4Y8H7DYDA6H5PT1B8BXC&qid=1601814482&refinements=p_36%3A200000-&rnid=1318502031&sr=8-11',
    'https://www.amazon.in/CHIPTRONEX-MX3-RGB-Cabinet-Tempered/dp/B08BJDWKWF/ref=sr_1_40?dchild=1&keywords=computer+cabinets&pd_rd_r=d4136365-9cf7-4cd6-9dd7-31b8a5f144e3&pd_rd_w=bLh6G&pd_rd_wg=DEisK&pf_rd_p=e98d1fad-4664-4b4d-a930-8932601ebadf&pf_rd_r=4Y8H7DYDA6H5PT1B8BXC&qid=1601814514&refinements=p_36%3A200000-&rnid=1318502031&sr=8-40'
], ['C8', 'E8', 'G8'])
#----------------------------------------------------------------------------------------------------------
ssd = Part([
    'https://www.amazon.in/WDS120G2G0A-120GB-2-5-inch-Internal-Green/dp/B076XWDN6V/ref=sr_1_3?crid=2L0P0JNIKXTXA&dchild=1&keywords=ssd+120gb&qid=1601814807&sprefix=hard+disk+%2Caps%2C-1&sr=8-3',
    'https://www.amazon.in/Crucial-BX500-120GB-2-5-inch-CT120BX500SSD1/dp/B07G3KRZBY/ref=sr_1_4?crid=2L0P0JNIKXTXA&dchild=1&keywords=ssd+120gb&qid=1601814807&sprefix=hard+disk+%2Caps%2C-1&sr=8-4',
    'https://www.amazon.in/Kingston-SSDNow-Internal-SA400S37-120GIN/dp/B079T88WY5/ref=sr_1_5?crid=2L0P0JNIKXTXA&dchild=1&keywords=ssd+120gb&qid=1601814807&sprefix=hard+disk+%2Caps%2C-1&sr=8-5',
    'https://www.amazon.in/Western-Digital-120GB-Internal-WDS120G2G0B/dp/B078WYRR9S/ref=sr_1_7?crid=2L0P0JNIKXTXA&dchild=1&keywords=ssd+120gb&qid=1601814807&sprefix=hard+disk+%2Caps%2C-1&sr=8-7'
], ['C10', 'E10', 'G10'])
#--------------------------------------------------------------------------------------------
hdd = Part([
    'https://www.amazon.in/Sata-Western-Digital-Product-Warranty/dp/B08HN47GDW/ref=sr_1_28?dchild=1&keywords=internal+hard+disk&qid=1601815548&refinements=p_36%3A150000-&rnid=1318502031&sr=8-28',
    'https://www.amazon.in/500-Gb-Green-Hard-Disk/dp/B08CBZY9R1/ref=sr_1_51?dchild=1&keywords=internal+hard+disk&qid=1601815581&refinements=p_36%3A150000-&rnid=1318502031&sr=8-51'
], ['C11', 'E11', 'G11'])
#-----------------------------------------------------------------------------------------------
psu = Part([
    'https://www.amazon.in/500-Gb-Green-Hard-Disk/dp/B08CBZY9R1/ref=sr_1_51?dchild=1&keywords=internal+hard+disk&qid=1601815581&refinements=p_36%3A150000-&rnid=1318502031&sr=8-51',
    'https://www.amazon.in/Antec-VP450P-450W-Power-Supply/dp/B006TM8XPW/ref=sr_1_4?dchild=1&keywords=PSU+450W&qid=1601815804&sr=8-4',
    'https://www.amazon.in/GIGABYTE-GP-P450B-Bronze-Power-Supply/dp/B08DK7YPX4/ref=sr_1_7?dchild=1&keywords=PSU+450W&qid=1601815804&sr=8-7',
    'https://www.amazon.in/Antec-VP450P-Plus-efficient-Certified/dp/B00006HS81/ref=sr_1_8?dchild=1&keywords=PSU+450W&qid=1601815804&sr=8-8',
    'https://www.amazon.in/Corsair-Bronze-Certified-Non-Modular-Supply/dp/B07YVVXYFN/ref=sr_1_1?dchild=1&keywords=PSU+450W&qid=1601815804&sr=8-1'
], ['C9', 'E9', 'G9'])

t_a = threading.Thread(target=processer.price_update,
                       args=(
                           processer.url,
                           processer.cell,
                       ))
t_b = threading.Thread(target=gpu.price_update, args=(
    gpu.url,
    gpu.cell,
))
t_c = threading.Thread(target=ram.price_update, args=(
    ram.url,
    ram.cell,
))
t_d = threading.Thread(target=motherboard.price_update,
                       args=(
                           motherboard.url,
                           motherboard.cell,
                       ))

t_a.start()
t_b.start()
t_c.start()
t_d.start()
t_a.join()
t_e = threading.Thread(target=case.price_update, args=(
    case.url,
    case.cell,
))
t_e.start()
t_b.join()
t_f = threading.Thread(target=ssd.price_update, args=(
    ssd.url,
    ssd.cell,
))
t_f.start()
t_c.join()
t_g = threading.Thread(target=hdd.price_update, args=(
    hdd.url,
    hdd.cell,
))
t_g.start()
t_h = threading.Thread(target=psu.price_update, args=(
    psu.url,
    psu.cell,
))
t_h.start()
t_d.join()
t_e.join()
t_f.join()
t_g.join()
t_h.join()

print("Done!")

wb.save('Updated.xlsx')
