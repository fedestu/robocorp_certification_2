import os
from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
import zipfile

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    create_folders()
    open_robot_order_website()
    close_annoying_modal()
    download_orders()
    orders = get_orders()
    fill_and_order(orders)
    archive_receipts()


def create_folders():
    os.makedirs('output/receipts', exist_ok=True)
    os.makedirs('output/screenshots', exist_ok=True)

def open_robot_order_website():
    """Navigates to the given URL"""    
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_orders():
    """Downloads csv file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    """Read data from csv and return the values"""
    csv = Tables()
    data = csv.read_table_from_csv("orders.csv")

    return data

def close_annoying_modal():
    """
    Closes the pop-up modal that appears on the robot order website.
    """
    page = browser.page()
    page.click("text=OK")

def store_receipt_as_pdf(order_number):
    """
    Stores the order receipt as a PDF file.

    Args:
        order_number (str): The order number to be used in the file name.

    Returns:
        str: The file system path to the PDF file.
    """
    page = browser.page()
    sales_results_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    receipt_path = f"output/receipts/receipt_{order_number}.pdf"
    pdf.html_to_pdf(sales_results_html, receipt_path)
    return receipt_path

def screenshot_robot(order_number):
    """
    Takes a screenshot of the robot preview and saves it.

    Args:
        order_number (str): The order number to be used in the file name.

    Returns:
        str: The file system path to the screenshot.
    """
    page = browser.page()
    screenshot_path = f"output/screenshots/robot_{order_number}.png"
    page.screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """
    Embeds the robot screenshot to the receipt PDF file.

    Args:
        screenshot (str): The file system path to the screenshot.
        pdf_file (str): The file system path to the PDF file.
    """  
    pdf = PDF()    
    pdf.add_files_to_pdf(files=[pdf_file, screenshot], target_document=pdf_file, append=True)

def fill_and_order(orders):
    """
    Fills the order form and submits it for each order in the list.

    Args:
        orders (list): A list of orders to be processed.
    """
    page = browser.page()
    max_attempts = 3    
    print("Voy a llenar el formulario")
    for row in orders:
        body = str(row["Body"])
        page.select_option("#head", str(row["Head"]))
        page.click(f"#id-body-{body}")
        page.fill('input[placeholder="Enter the part number for the legs"]', str(row["Legs"]))
        page.fill("#address", str(row["Legs"]))
        attempts = 0
        while attempts < max_attempts:
            try:
                page.click("#order")
                print("Voy a sacar foto")
                screenshot_path = screenshot_robot(row['Order number'])

                if page.is_visible(".alert.alert-danger"):
                    raise Exception("Submit failed, retrying...")
                
                receipt_path = store_receipt_as_pdf(row['Order number'])
                embed_screenshot_to_receipt(screenshot_path, receipt_path)
                
                page.click("#order-another")
                close_annoying_modal()
                break
            except Exception as e:
                print(f"Error al enviar el pedido: {e}")
                attempts += 1
                if attempts >= max_attempts:
                    print(f"Max attempts reached for order: {row}")
                    break               

def archive_receipts():
    """
    Archives all receipt PDF files into a single ZIP file.

    This function searches for all PDF files in the 'output/receipts' directory
    and compresses them into a single ZIP file stored in the 'output' directory.
    """
    source_folder = 'output/receipts'  # Directorio donde están los archivos PDF
    zip_path = 'output/receipts_archive.zip'  # Ruta del archivo ZIP que se creará

    # Crear un archivo ZIP y añadir todos los archivos PDF del directorio especificado
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for foldername, subfolders, filenames in os.walk(source_folder):
            for filename in filenames:
                # Asegúrate de que solo se añaden archivos PDF
                if filename.endswith('.pdf'):
                    file_path = os.path.join(foldername, filename)
                    zipf.write(file_path, os.path.relpath(file_path, source_folder))
