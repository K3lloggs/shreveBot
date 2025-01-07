import win32com.client
import re
from datetime import datetime
import pytz
import os

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
SAVE_DIR = os.path.join(CURRENT_DIR, "IMAGE_FOLDER")
os.makedirs(SAVE_DIR, exist_ok=True)

outlook = win32com.client.Dispatch('outlook.application').GetNamespace("MAPI")
inbox = outlook.Folders('cclose@shrevecrumpandlow.com').Folders('Inbox')
messages = inbox.Items

eastern = pytz.timezone('America/New_York')
cutoff_date = eastern.localize(datetime(2024, 12, 18))

def extract_sku(subject):
    cleaned = re.sub(r'^(FW:|Fw:|FWD:|Fwd:)\s*', '', subject).strip()
    if cleaned.startswith('SKU'):
        cleaned = cleaned.replace('SKU', '').strip()
    
    if cleaned.lower().startswith('m') and len(cleaned) == 8 and cleaned[1:].isdigit():
        return cleaned.upper()
    elif cleaned.isdigit() and len(cleaned) == 7:
        return cleaned
    return None

def extract_labels(body):
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
    if message.Attachments.Count == 0:
        return []
    
    sku_dir = os.path.join(SAVE_DIR, sku)
    os.makedirs(sku_dir, exist_ok=True)
    
    saved_files = []
    for attachment in message.Attachments:
        if attachment.FileName.lower().startswith('outlook'):
            continue
            
        if attachment.FileName.lower().endswith(('.jpg', '.jpeg', '.png', '.gif')):
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

sku_data = {}
message_count = 0

for msg in messages:
    message_count += 1
    try:
        if (not msg.SenderName or 
            "brian walker" not in msg.SenderName.lower() or
            not msg.Subject.lower().startswith(('fw:', 'fwd:', 'fw')) or
            msg.Subject.lower().startswith('re:')):
            continue
        
        received_date = eastern.localize(msg.ReceivedTime.replace(tzinfo=None))
        if received_date < cutoff_date:
            continue
        
        sku = extract_sku(msg.Subject)
        if sku:
            labels = extract_labels(msg.Body)
            saved_files = save_attachments(msg, sku)
            
            sku_data[sku] = {
                'labels': labels,
                'received_date': received_date,
                'sender': msg.SenderName,
                'image_files': saved_files
            }
            
            print(f"\nSKU: {sku}")
            print(f"Labels: {', '.join(labels)}")
            print(f"Images saved: {len(saved_files)}")
    
    except Exception as e:
        print(f"Error processing message: {str(e)}")

print(f"\nTotal messages: {message_count}")
print("\nSKU Summary:")
for sku, data in sku_data.items():
    print(f"\nSKU: {sku}")
    print(f"Labels: {', '.join(data['labels'])}")
    print(f"Images: {len(data['image_files'])}")