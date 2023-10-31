from RPA.Browser.Selenium import Selenium
import pandas as pd
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from PyPDF2 import PdfFileMerger
from PIL import Image
import os
import time as t
import shutil


lib = Selenium()
    
curr_dir = os.getcwd()
if "ss_pdf" not in list(os.listdir(curr_dir)):
    os.mkdir("ss_pdf")

if "merged_files" not in list(os.listdir(curr_dir)):
    os.mkdir("merged_files")

def open_browser():
    u_rl = "https://robotsparebinindustries.com"
    lib.open_available_browser(u_rl)
    lib.click_element("xpath:/html/body/div/header/div/ul/li[2]/a")

def screenshot(k):
    lib.screenshot("xpath:/html/body/div/div", f"output/screenshot_{k}.png")

    
def csv():
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)

def pdff(k):
    t.sleep(2)
    pran=lib.get_element_attribute("id:receipt","outerHTML")
    pdf=PDF()
    pdf.html_to_pdf(pran,f"output/receipts_{k}.pdf")


def fill_all_excel():
    csv()
    df = pd.read_csv("orders.csv")
    it = int(len(list(df["Head"])))
    i = 0
    while i<it:
        try:
            lib.click_element("xpath:/html/body/div/div/div[2]/div/div/div/div/div/button[1]") #ok button
            lib.click_element("id:head")
            lib.select_from_list_by_value("id:head",str(list(df["Head"])[i]))
            lib.click_element("id:head")
            lib.click_element("id:id-body-{0}".format(str(list(df["Body"])[i])))
            lib.input_text("xpath:/html/body/div/div/div[1]/div/div[1]/form/div[3]/input",str(list(df["Legs"])[i]))
            lib.input_text("id:address",str(list(df["Address"])[i]))
            lib.click_element("xpath:/html/body/div/div/div[1]/div/div[1]/form/button[1]")
            lib.click_element("xpath:/html/body/div/div/div[1]/div/div[1]/form/button[2]")
            screenshot(i)
            pdff(i)
            lib.click_element("xpath:/html/body/div/div/div[1]/div/div[1]/div/button")
            i=i+1
        except:
            if i == 0:
                lib.reload_page()
                i = 0
            else:
                lib.reload_page()
                i-=1


def convert_to_pdf():
    df = pd.read_csv("orders.csv")
    it = int(len(list(df["Head"])))
    k = 0
    dir_name = "ss_pdf"
    ss_path = os.path.join(curr_dir, dir_name)
    while k<it:
        image1 = Image.open(r'output/screenshot_{0}.png'.format(k))
        im1 = image1.convert('RGB')
        im1.save(r'{0}/screenshot_{1}.pdf'.format(ss_path,k))
        k = k+1
        
def merge_pdfs():
    df = pd.read_csv("orders.csv")
    it = len(list(df["Head"]))
    dir_name = "ss_pdf"
    ss_path = os.path.join(curr_dir, dir_name)

    dir_name1 = "merged_files"
    merge_path = os.path.join(curr_dir, dir_name1)

    for t in range(it):
        pdfs = ['{0}/screenshot_{1}.pdf'.format(ss_path,t), 'output/receipts_{0}.pdf'.format(t)]
        merger = PdfFileMerger()
        for pdf in pdfs:
            merger.append(pdf)
        merger.write("{0}/merged_file_{1}.pdf".format(merge_path,t))
        merger.close()       

def make_zip():
    shutil.make_archive("output/compressed_pdfs", 'zip', "merged_files")
    

def main():
    open_browser()
    fill_all_excel()
    convert_to_pdf()
    merge_pdfs()
    make_zip()

if __name__ == "__main__":
    main()
