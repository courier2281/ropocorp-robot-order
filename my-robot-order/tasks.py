from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Robocorp.Vault import Vault
import shutil
from PIL import Image

_secret = Vault().get_secret("mcredentials")

USER_NAME = _secret["musername"]
PASSWORD = _secret["mpassword"]

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """

    browser.configure(
        slowmo=0,
    )

    open_robot_order_website()
    navigate_to_order_page()
    orders = get_orders()
    generate_robot_orders(orders)
    archive_receipts()   

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/")

def navigate_to_order_page():
    page = browser.page()
    page.locator("xpath=//a[contains(@href, '#/robot-order')]").click()

def get_orders():
    download_csv_file()

    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number","Head","Body","Legs","Address"]
    )

    return orders

def download_csv_file():
    """Downloads csv file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def generate_robot_orders(orders):
    """Read data from csv and generate robot orders"""
    page = browser.page()
    page.set_default_timeout(1000)

    for row in orders:
        retrycount = 0
        success = False
        exceptiontxt = ""

        while not success and retrycount < 10:
            try:
                fill_the_form(row)
                success = True
            except:
                retrycount + 1
                if page.locator(".alert"):
                    exceptiontxt = page.locator(".alert", ).inner_text
                    print(exceptiontxt)
                navigate_to_order_page()
                continue
        
        if not success:
            raise Exception("Failed to make order " + row["Order number"] +": " + exceptiontxt)

def close_annoying_modal():
    page = browser.page()
    page.set_default_timeout(1000)
    
    try:
        page.locator(".btn-danger").click()
    except:
        ()

def fill_the_form(order_details):
    """Completes one robot order"""
    page = browser.page()
    page.set_default_timeout(1000)

    close_annoying_modal()

    page.locator("#head").select_option(order_details["Head"])
    page.locator("#id-body-"+order_details["Body"]).click()
    page.locator("xpath=//label[contains(.,'3. Legs:')]/../input").fill(order_details["Legs"])
    page.locator("#address").fill(order_details["Address"])

    page.locator("#order").click()

    screenshot_robot(order_details["Order number"])
    store_receipt_as_pdf(order_details["Order number"])

    page.locator("#order-another").click()

def screenshot_robot(order_number):
    """Take a screenshot of the page"""
    page = browser.page()
    page.locator("#robot-preview-image").screenshot(path="output/robot_screenshots/robot_"+order_number+".png")

def store_receipt_as_pdf(order_number):
    """Export the data to a pdf file and embed robot image"""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf_file = "output/order_receipts/order_receipt_"+order_number+".pdf"
    pdf.html_to_pdf(receipt_html, pdf_file)
    
    image_path = "output/robot_screenshots/robot_"+order_number+".png"
   
    embed_screenshot_to_receipt(image_path, pdf_file)

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_files_to_pdf(files=[screenshot], target_document=pdf_file, append=True)

def resize_screenshot(image_path):
    with Image.open(image_path) as img:
        img = img.resize((int(float(img.size[0])*0.2), int(float(img.size[1])*0.2)))
        img.save(image_path)

def archive_receipts():
    # Create a zip archive from the specified directory
    shutil.make_archive("output/robot_orders", 'zip', root_dir="output/order_receipts")