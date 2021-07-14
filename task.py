"""Template robot with Python."""

from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Robocloud.Secrets import Secrets

secrets = Secrets()
USER_NAME = secrets.get_secret("credentials")["username"]
PASSWORD = secrets.get_secret("credentials")["password"]

url = "https://robotsparebinindustries.com/"
browser = Selenium()
downloads = HTTP()
excel = Files()
pdf = PDF()


def openWebsite():
    browser.open_available_browser(url)
    browser.input_text("xpath://html/body/div/div/div/div/div[1]/form/div[1]/input", USER_NAME)
    browser.input_text('xpath://*[@id="password"]', PASSWORD)
    browser.submit_form()
    browser.wait_until_page_contains_element('xpath://*[@id="sales-form"]')


def downloadFile():
    downloads.download("https://robotsparebinindustries.com/SalesData.xlsx",overwrite=True)


def getDateExcel():
    excel.open_workbook("SalesData.xlsx")
    salesReps = excel.read_worksheet_as_table(header=True)
    excel.close_workbook()
    for salesRep in salesReps:
        submitTheForm(salesRep)


def submitTheForm(salesRep):
    browser.input_text('xpath://*[@id="firstname"]', salesRep["First Name"])
    browser.input_text('xpath://*[@id="lastname"]', salesRep["Last Name"])
    browser.input_text('xpath://*[@id="salesresult"]', salesRep["Sales"])
    targetAs = str(salesRep["Sales Target"])
    browser.select_from_list_by_value('xpath://*[@id="salestarget"]', targetAs)
    browser.submit_form()


def collectTheResults():
    browser.screenshot('xpath://*[@id="root"]/div/div/div/div[2]/div[1]', "output/sales_summary.png")


def exportPDF():
    browser.wait_until_element_is_visible('xpath://*[@id="sales-results"]')
    salesResultsHTML = browser.get_element_attribute('xpath://*[@id="sales-results"]', "outerHTML")
    pdf.html_to_pdf(salesResultsHTML, "output/sales_results.pdf")


def logOut():
    browser.click_button('xpath://*[@id="logout"]')
    browser.close_browser()


def main():
    try:
        openWebsite()
        downloadFile()
        getDateExcel()
        collectTheResults()
        exportPDF()
    finally:
        logOut()    


if __name__ == "__main__":
    main()
