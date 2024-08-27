import shutil
from pathlib import Path

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
        slowmo=10,
    )
    open_robot_order_website()
    orders = get_orders()
    submit_orders(orders)
    archive_receipts()


def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def download_orders_file():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def get_orders():
    download_orders_file()
    library = Tables()
    orders = library.read_table_from_csv("orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"])
    return orders


def close_annoying_modal():
    page = browser.page()
    page.click("button:text('OK')")


def fill_the_form(order):
    page = browser.page()
    page.select_option("#head", str(order["Head"]))
    page.click(f"//label/input[@value='{order['Body']}']")
    page.get_by_placeholder("Enter the part number for the legs").fill(order["Legs"])
    page.fill("#address", order["Address"])
    page.click("#preview")


def submit_form():
    page = browser.page()
    page.click("#order")
    while page.get_attribute(".alert", "class", timeout=2000) == "alert alert-danger":
        page.click("#order")


def get_order_number():
    page = browser.page()
    order_number = page.locator(".badge-success").text_content()
    return order_number


def screenshot_robot(order_number):
    page = browser.page()

    images_path = Path("output/images")
    images_path.mkdir(parents=True, exist_ok=True)

    robot_div = page.locator("#robot-preview-image")
    robot_div.screenshot(path=images_path / f"{order_number}.png")


def store_receipt_as_pdf(order_number):
    page = browser.page()
    pdf = PDF()

    receipt_html = page.locator("#receipt").inner_html()
    receipts_path = Path("output/receipts")
    receipts_path.mkdir(parents=True, exist_ok=True)
    pdf.html_to_pdf(receipt_html, receipts_path / f"{order_number}.pdf")


def merge_receipt_with_image(order_number):
    pdf = PDF()
    images_path = Path("output/images")
    receipts_path = Path("output/receipts")
    pdf.add_files_to_pdf(
        files=[receipts_path / f"{order_number}.pdf", images_path / f"{order_number}.png"],
        target_document=receipts_path / f"{order_number}.pdf",
    )


def go_to_next_robot():
    page = browser.page()
    page.click("#order-another")


def submit_orders(orders):
    for order in orders:
        close_annoying_modal()
        fill_the_form(order)
        submit_form()
        order_number = get_order_number()
        screenshot_robot(order_number)
        store_receipt_as_pdf(order_number)
        merge_receipt_with_image(order_number)
        go_to_next_robot()


def archive_receipts():
    output_path = Path("output")
    receipts_path = output_path / "receipts"
    receipts_archive = output_path / "receipts.zip"
    shutil.make_archive(receipts_archive.stem, "zip", output_path, receipts_path.name)
    shutil.move(f"{receipts_archive.stem}.zip", output_path / "receipts.zip")
