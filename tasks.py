from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import time

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    
    browser.configure(slowmo=250,)
    open_robot_order_website()
    close_annoying_modal()
    download_orders()
    fill_the_form()
    archive_receipts()

def open_robot_order_website():
    """Navigates to the given URL"""
    
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_orders():
    """Downloads orders csv file"""
    
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def fill_the_form():
    """Fill in the robot order form, create pdf file from receipt, 
    create screenshot from the robot preview and embed screenshot to pdf"""
    
    orders = get_orders()
    order_number = 0
   
    for row in orders:
        page = browser.page()
        page.select_option("#head", str(row["Head"]))
        page.click(f"input.form-check-input[name='body'][value='{str(row['Body'])}']")
        page.fill("input.form-control[placeholder='Enter the part number for the legs']", str(row['Legs']))
        page.fill("#address", row['Address'])
        page.click("text=Preview")
        page.click("id=order")
        
        check_error()
        
        order_number = int(row['Order number'])
        store_receipt_as_pdf(order_number)
        pdf_file = store_receipt_as_pdf(order_number)

        screenshot_robot(order_number)
        screenshot = screenshot_robot(order_number)
        
        embed_screenshot_to_receipt(screenshot, pdf_file)
        
        page.click("id=order-another")
        
        close_annoying_modal()

def check_error():
     """Error check for clicking order button"""
     page = browser.page()
     max_attempts = 5
     attempt = 0
     error_is_visible = page.is_visible("div.alert.alert-danger[role='alert']")

     while attempt < max_attempts and error_is_visible:
            if error_is_visible:
                time.sleep(0.2)
                page.click("id=order")
                time.sleep(0.2)
                
            error_is_visible = page.is_visible("div.alert.alert-danger[role='alert']")
            attempt += 1

def get_orders():
    """Reads CSV file content as a table"""
    library = Tables()
    orders = library.read_table_from_csv("orders.csv", header=True)
    return orders

def close_annoying_modal():
    """Close the annoying modal popup"""
    page = browser.page()
    page.click("button:text('OK')")

def store_receipt_as_pdf(order_number):
    """Saves receipt as PDF"""
    page = browser.page()
    pdf = PDF()
    output_filename = f"output/receipts/receipt_{order_number}.pdf"
    
    receipt_html = page.locator("id=receipt").inner_html()
    pdf.html_to_pdf(receipt_html, output_filename)
    return output_filename

def screenshot_robot(order_number):
    """Saves screenshot of robot preview"""
    page = browser.page()
    output_filename = f"output/screenshots/robot_{order_number}.png"

    robot_picture = page.locator("#robot-preview-image")
    robot_picture.screenshot(path=output_filename)
    return output_filename

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Appends the screenshot of the robot into the receipt PDF file"""
    pdf = PDF()
    pdf.add_files_to_pdf(files=[screenshot], target_document=pdf_file, append=True)

def archive_receipts():
    """Create zip file from all of the receipts created"""
    archive = Archive()
    files_dir = "./output/receipts"
    zip_file_name = "./output/receipts/receipts.zip"

    archive.archive_folder_with_zip(folder=files_dir, archive_name=zip_file_name)