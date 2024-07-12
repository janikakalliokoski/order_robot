import zipfile
import os

from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF


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
        slowmo=500,
    )
    open_robot_order_website()
    get_orders()
    orders = read_csv_data()
    for row in orders:
        close_annoying_modal()
        fill_in_order_form(row)
        submit_order_form(row["Order number"])
    archive_receipts()

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def close_annoying_modal():
    """Closes the annoying modal that pops up"""
    page = browser.page()
    page.click("button:text('OK')")

def fill_in_order_form(order_rep):
    """Fills in the order form"""
    page = browser.page()
    page.select_option("#head", str(order_rep["Head"]))
    page.click(f"#id-body-{order_rep['Body']}")
    legs_locator = page.get_by_placeholder("Enter the part number for the legs")
    legs_locator.fill(order_rep["Legs"])
    page.fill("#address", order_rep["Address"])

def get_orders():
    """Downloads csv file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def read_csv_data():
    """Read data from csv and fill in the sales form"""
    csv = Tables()
    worksheet = csv.read_table_from_csv("orders.csv")
    return worksheet

def submit_order_form(row):
    """Submit the order form"""
    page = browser.page()
    page.click("button:text('Order')")
    succesful_order = page.query_selector("#receipt")
    while not succesful_order:
        page.click("button:text('Order')")
        succesful_order = page.query_selector("#receipt")
    store_receipt_as_pdf(row)
    screenshot_robot(row)
    embed_screenshot_to_receipt(f"robot-{row}.png", f"receipt-{row}.pdf")
    page.click("button:text('Order another robot')")

def store_receipt_as_pdf(order_number):
    """Stores the receipt as a PDF"""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(receipt_html, f"output/receipt-{order_number}.pdf")

def screenshot_robot(order_number):
    """Takes a screenshot of the robot"""
    page = browser.page()
    robot_preview_image = page.locator("#robot-preview-image")
    robot_preview_image.screenshot(path=f"output/robot-{order_number}.png")

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.open_pdf(f"output/{pdf_file}")
    pdf.add_files_to_pdf(files=[f"output/{screenshot}"], target_document=f"output/{pdf_file}", append=True)
    pdf.close_all_pdfs()

def archive_receipts():
    """Creates a ZIP archive of the receipts and the images"""
    with zipfile.ZipFile("output/robot-orders.zip", "w") as zipf:
        for file in os.listdir("output"):
            if file.endswith(".pdf"):
                zipf.write(f"output/{file}")
