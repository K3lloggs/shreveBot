import os
import re
import time
import pytz
import win32com.client
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright, expect
from dotenv import load_dotenv

#############################################
# 1. ENV & CONSTANTS
#############################################

load_dotenv()

# Adjust if needed or store them in your .env file
USERNAME = os.getenv("SCRL_USERNAME")
PASSWORD = os.getenv("SCRL_PASSWORD")

# For Outlook parsing
EASTERN = pytz.timezone('America/New_York')
CUTOFF_DATE = EASTERN.localize(datetime(2024, 12, 18))

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
SAVE_DIR = os.path.join(CURRENT_DIR, "IMAGE_FOLDER")
os.makedirs(SAVE_DIR, exist_ok=True)


#############################################
# 2. LABEL MAP DICTIONARIES
#    (Map mail-parser labels -> WP values)
#############################################

CATEGORIES = {
    "fine jewelry": "Fine Jewelry",
    "earrings": "Earrings",
    "bracelets": "Bracelets",
    "tennis bracelets": "Tennis Bracelets",
    "necklaces": "Necklaces",
    "rings": "Rings",
    "engagement rings": "Engagement Rings",
    "antique and estate": "Antique & Estate",
    "antique & estate": "Antique & Estate",  # handle slight variation
    "pin & brooch": "Brooches",             # or "Pins", adapt as needed
    # Add more as needed...
}

METALS = {
    "yellow gold": "18KT Yellow Gold",
    "white gold": "18KT White Gold",
    "platinum": "Platinum",
    "gold": "18KT Gold",
    "18kt yellow gold": "18KT Yellow Gold",
    # ... expand as needed
}

STONES = {
    "pearl": "Pearl",
    "diamond": "Diamond",
    "ruby": "Ruby",
    "emerald": "Emerald",
    "lapis": "Lapis",
    "aquamarine": "Aquamarine",
    "moonstone": "Moonstone",
    "enamel": "Enamel",
    "sapphire": "Sapphire",
    "torquoise": "Turquoise",  # or fix spelling -> "turquoise"
    # ... expand as needed
}

BRANDS = {
    "cartier": "Cartier",
    "vintage cartier": "Cartier",  # unify under "Cartier"
    "harry winston": "Harry Winston",
    "breguet": "Breguet",
    # Add more as needed
}

CUTS = {
    "round brilliant cut": "Round Brilliant",
    "emerald cut": "Emerald",
    # ... expand as needed
}


#############################################
# 3. MAIL PARSER: OUTLOOK
#############################################

def extract_sku(subject):
    """
    Extracts the SKU from the subject line.
    Example:
      Subject: "FW: SKU6934724" -> "6934724"
      or "FW: M1234567" -> "M1234567"
    """
    cleaned = re.sub(r'^(FW:|Fw:|FWD:|Fwd:)\s*', '', subject).strip()
    if cleaned.startswith('SKU'):
        cleaned = cleaned.replace('SKU', '').strip()
    
    # e.g. "M1234567"
    if cleaned.lower().startswith('m') and len(cleaned) == 8 and cleaned[1:].isdigit():
        return cleaned.upper()
    elif cleaned.isdigit() and len(cleaned) == 7:
        return cleaned
    return None

def extract_labels(body):
    """
    Finds the first line in the email body with comma-separated labels,
    ignoring typical lines like 'Subject:', 'From:', 'To:', etc.
    Returns a list of normalized labels (lower-cased, stripped).
    """
    for line in body.split('\n'):
        if (',' in line and 
            '@' not in line and 
            'Subject:' not in line and 
            'From:' not in line and 
            'To:' not in line):
            labels = [label.strip().lower() for label in line.split(',')]
            if len(labels) > 1:
                return labels
    return []

def save_attachments(message, sku):
    """
    Saves image attachments for the given email message to a subfolder 
    in SAVE_DIR named after the SKU.
    Returns a list of saved file paths.
    """
    if message.Attachments.Count == 0:
        return []
    
    sku_dir = os.path.join(SAVE_DIR, sku)
    os.makedirs(sku_dir, exist_ok=True)
    
    saved_files = []
    for attachment in message.Attachments:
        fname = attachment.FileName.lower()
        # skip "Outlook signature" attachments or non-images
        if fname.startswith('outlook'):
            continue
        if fname.endswith(('.jpg', '.jpeg', '.png', '.gif')):
            file_path = os.path.join(sku_dir, attachment.FileName)
            try:
                if os.path.exists(file_path):
                    continue
                attachment.SaveAsFile(file_path)
                saved_files.append(file_path)
                print(f"Saved: {file_path}")
            except Exception as e:
                print(f"Error saving {attachment.FileName}: {str(e)}")
    
    return saved_files

def fetch_mail_data():
    """
    Connect to Outlook, iterate over messages in 'Inbox',
    filter them, extract SKU, labels, and save attachments.
    Return a list of dictionaries:
      [
         {
           'sku': '6934724',
           'labels_str': 'fine jewelry, bracelets, diamond, ...',
           'image_folder': 'C:\\...\\IMAGE_FOLDER\\6934724'
         },
         ...
      ]
    """
    outlook = win32com.client.Dispatch('outlook.application').GetNamespace("MAPI")
    inbox = outlook.Folders('cclose@shrevecrumpandlow.com').Folders('Inbox')
    messages = inbox.Items

    mail_data_list = []
    message_count = 0

    for msg in messages:
        message_count += 1
        try:
            # Filter: must be from Brian Walker, must be forwarded mail, not replies, etc.
            if (not msg.SenderName 
                or "brian walker" not in msg.SenderName.lower()
                or not msg.Subject.lower().startswith(('fw:', 'fwd:', 'fw'))
                or msg.Subject.lower().startswith('re:')):
                continue
            
            # Check date
            received_date = EASTERN.localize(msg.ReceivedTime.replace(tzinfo=None))
            if received_date < CUTOFF_DATE:
                continue
            
            # Extract SKU
            sku = extract_sku(msg.Subject)
            if sku:
                labels = extract_labels(msg.Body)
                saved_files = save_attachments(msg, sku)
                
                # Build a single comma-separated string for the label parser 
                # (like "fine jewelry, earrings, pearl, diamond, white gold")
                labels_str = ", ".join(labels)

                mail_data_list.append({
                    "sku": sku,
                    "labels_str": labels_str,
                    "image_folder": os.path.join(SAVE_DIR, sku)
                })

                print(f"\nSKU: {sku}")
                print(f"Labels: {labels_str}")
                print(f"Images saved: {len(saved_files)}")
        
        except Exception as e:
            print(f"Error processing message: {str(e)}")

    print(f"\nTotal messages scanned: {message_count}")
    print("\nSKU Summary:")
    for item in mail_data_list:
        print(f"\nSKU: {item['sku']}")
        print(f"Labels: {item['labels_str']}")
        # We can check the image folder size or count if desired

    return mail_data_list


#############################################
# 4. MAP LABELS -> WP FIELDS
#############################################

def parse_labels_to_attributes(labels_str: str) -> dict:
    """
    Takes a comma-separated string of labels (e.g. "fine jewelry, earrings, pearl, diamond, white gold")
    and maps them into categories, metals, stones, brands, cuts.
    """
    labels = [lbl.strip().lower() for lbl in labels_str.split(",")]
    
    categories = []
    metals = []
    stones = []
    brands = []
    cuts = []

    for label in labels:
        # Category
        if label in CATEGORIES:
            categories.append(CATEGORIES[label])
        # Metal
        if label in METALS:
            metals.append(METALS[label])
        # Stone
        if label in STONES:
            stones.append(STONES[label])
        # Brand
        if label in BRANDS:
            brands.append(BRANDS[label])
        # Cut
        if label in CUTS:
            cuts.append(CUTS[label])

    return {
        "categories": categories,
        "metals": metals,
        "stones": stones,
        "brands": brands,
        "cuts": cuts,
    }


def build_product_data_from_parser(sku: str, labels_str: str, image_folder: str) -> dict:
    """
    Construct the product_data dict used by fill_product_form() in Playwright.
    """
    attr = parse_labels_to_attributes(labels_str)

    product_data = {
        "name": f"SKU {sku}",
        "price": "0.00",
        "sku": sku,
        "categories": attr["categories"],
        "metals": attr["metals"],
        "stones": attr["stones"],
        "brands": attr["brands"],
        "cuts": attr["cuts"],
        "images_folder": image_folder,
    }
    return product_data


#############################################
# 5. PLAYWRIGHT AUTOMATION
#############################################

def fill_product_form(page, product_data):
    """
    Fills out the product form fields (name, price, SKU, images, categories, etc.)
    within the #content iframe. Then publishes the product.
    """
    content_frame = page.locator("#content iframe").content_frame()

    # 1. Product name
    content_frame.get_by_label("Product name").click()
    content_frame.get_by_label("Product name").fill(product_data["name"])

    # 2. Price
    content_frame.get_by_label("Regular price ($)").click()
    content_frame.get_by_label("Regular price ($)").fill(product_data["price"])

    # 3. Inventory / SKU
    content_frame.get_by_role("link", name="Inventory").click()
    content_frame.get_by_label("SKU", exact=True).click()
    content_frame.get_by_label("SKU", exact=True).fill(product_data["sku"])

    # 4. Upload images
    content_frame.get_by_role("link", name="Add product gallery images").click()
    content_frame.get_by_role("tab", name="Upload files").click()
    
    # If the WP media uploader has <input type="file" multiple>, do it directly
    if os.path.isdir(product_data["images_folder"]):
        image_files = sorted(
            [os.path.join(product_data["images_folder"], f) 
             for f in os.listdir(product_data["images_folder"]) 
             if f.lower().endswith((".jpg", ".jpeg", ".png", ".gif"))]
        )
        if image_files:
            for img_path in image_files:
                content_frame.locator("input[type='file']").set_input_files(img_path)

    # Attempt to close the dialog
    try:
        content_frame.get_by_role("button", name=" Close dialog").click()
    except:
        pass

    # 5. Categories
    for cat in product_data["categories"]:
        # Example: "Primary Make Primary Fine Jewelry"
        try:
            content_frame.get_by_text(f"Primary Make Primary {cat}", exact=True).click()
        except:
            print(f"[WARNING] Category '{cat}' not found.")

    # 6. Metals
    for metal in product_data["metals"]:
        try:
            content_frame.get_by_text(f"Primary Make Primary {metal}", exact=True).click()
        except:
            print(f"[WARNING] Metal '{metal}' not found.")

    # 7. Stones
    for stone in product_data["stones"]:
        try:
            content_frame.get_by_text(f"Primary Make Primary {stone}", exact=True).click()
        except:
            print(f"[WARNING] Stone '{stone}' not found.")

    # 8. Brands
    for brand in product_data["brands"]:
        try:
            content_frame.get_by_text(f"Primary Make Primary {brand}", exact=True).click()
        except:
            print(f"[WARNING] Brand '{brand}' not found.")

    # 9. Cuts
    for cut in product_data["cuts"]:
        try:
            content_frame.get_by_text(f"Primary Make Primary {cut}", exact=True).click()
        except:
            print(f"[WARNING] Cut '{cut}' not found.")

    # 10. Publish
    content_frame.get_by_role("button", name="Publish", exact=True).click()


def run_playwright_automation(playwright: Playwright, mail_data_list):
    """
    Logs into the WP admin, iterates through mail_data_list (parsed from Outlook),
    and creates a new product for each.
    """
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()

    # 1. Log in
    page.goto("https://shrevecrumpandlow.com/account?vgfa_redirect_to_login=1")
    page.get_by_label("Username or email address *").fill(USERNAME)
    page.get_by_label("Password *Required").fill(PASSWORD)
    page.get_by_role("button", name="Log in").click()

    # 2. Navigate to Admin Portal -> Products
    page.get_by_role("link", name="Admin Portal").click()
    page.get_by_role("link", name="Products", exact=True).click()

    for item in mail_data_list:
        # 3. Click "Add new product"
        page.locator("#content iframe").content_frame().get_by_role("link", name="Add new product").click()

        # 4. Convert email data -> product_data
        product_data = build_product_data_from_parser(
            sku = item["sku"],
            labels_str = item["labels_str"],
            image_folder = item["image_folder"]
        )

        # 5. Fill the form & publish
        fill_product_form(page, product_data)

        # 6. Go back to product list for the next
        page.get_by_role("link", name="Products", exact=True).click()

    # End
    context.close()
    browser.close()


#############################################
# 6. MAIN ENTRY POINT
#############################################

def main():
    # 1. Pull new emails & parse them
    mail_data_list = fetch_mail_data()
    if not mail_data_list:
        print("No new mail items found to upload.")
        return

    # 2. Use Playwright to upload each new SKU
    with sync_playwright() as playwright:
        run_playwright_automation(playwright, mail_data_list)


if __name__ == "__main__":
    main()
