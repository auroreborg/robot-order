from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF


import shutil
from pathlib import Path


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure()
    open_robot_order_website()
    ordertable = get_orders()
    #print_row(ordertable)
    close_annoying_modal()
    fill_the_form(ordertable)
    convert_pdfs_to_zip(pdf_folder_path = "output/Receipts/PDF", output_zip_path= "output\Receipts\ZippedPDFs")

    
 
def open_robot_order_website():
    """Open browser and navigates to given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """Download CSV file from the given URL"""
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)
    """Read the CSV file as a datatable"""
    table = Tables()
    ordertable = table.read_table_from_csv("orders.csv", header=True)
    return ordertable

# def print_row(ordertable):
#     for row in ordertable:
#         print(row)

def close_annoying_modal():
    page = browser.page()
    page.click("text=OK")

def fill_the_form(ordertable):    
    page = browser.page()
    for row in ordertable:
            page.select_option("#head", row['Head'])
            page.click("#id-body-" + row['Body'])
            page.fill("input.form-control[placeholder='Enter the part number for the legs']", row['Legs'])
            page.fill("#address", row['Address'])
            page.click("#preview")
            
            order_and_handle_error()
            file_path_pdf = store_receipt_as_pdf(row['Order number'])
            file_path_screenshot = screenshot_robot(row['Order number'])
            
            embed_screenshot_to_receipt(file_path_screenshot, file_path_pdf)

            page.click("#order-another")  
            close_annoying_modal()


def order_and_handle_error(retry_limit=3):
    """Click order button, check for error message, if appears click order button again"""
    page = browser.page()
    retries = 0
    page.click("#order")
    while page.is_visible("div.alert.alert-danger[role='alert']") and retries < retry_limit:
            page.click("#order")
            retries = retries + 1
            
def store_receipt_as_pdf(order_number):
    page = browser.page()
    pdf = PDF()
    file_path_pdf = (f"output/Receipts/Pdf/Receipt_Ordernumber_{order_number}.pdf")
    receipt_html = page.locator("#receipt").inner_html()
    pdf.html_to_pdf(receipt_html, file_path_pdf)
    return file_path_pdf

def screenshot_robot(order_number):
    page = browser.page()
         
    file_path_screenshot = (f"output/Receipts/Pictures/Screenshot_Ordernumber_{order_number}.png")
    robot_picture = page.locator("#robot-preview-image")
    robot_picture.screenshot(path=file_path_screenshot)

    # img = Image.open(file_path_screenshot)
    # resized_image = img.resize((100,200))
    # resized_image.save(file_path_screenshot)

    return  file_path_screenshot

def embed_screenshot_to_receipt(file_path_screenshot, file_path_pdf):
    pdf = PDF() 
    pdf.add_watermark_image_to_pdf(
        image_path=file_path_screenshot,
        source_path=file_path_pdf,
        output_path=file_path_pdf
    )


def convert_pdfs_to_zip(pdf_folder_path, output_zip_path):
    pdf_folder_path = Path(pdf_folder_path)     # Ensure the PDF folder path is a Path object
    output_zip_folder = Path(output_zip_path).parent     # Ensure the output folder for the zip exists
    
    output_zip_folder.mkdir(parents=True, exist_ok=True)
    
    shutil.make_archive(base_name=output_zip_path, format='zip', root_dir=pdf_folder_path)     # Create a zip file containing all PDFs in the specified folder

