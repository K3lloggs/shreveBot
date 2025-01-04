import win32com.client
import re
from datetime import datetime
import pytz
import os

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
SAVE_DIR = os.path.join(CURRENT_DIR, "IMAGE_FOLDER")

if not os.path.exists(SAVE_DIR):
   os.makedirs(SAVE_DIR)

outlook = win32com.client.Dispatch('outlook.application').GetNamespace("MAPI")
inbox = outlook.Folders('cclose@shrevecrumpandlow.com').Folders('Inbox')
messages = inbox.Items

message_count = 0
eastern = pytz.timezone('America/New_York')
cutoff_date = eastern.localize(datetime(2024, 12, 18))

def extract_sku(subject):
   cleaned = re.sub(r'^(FW:|Fw:|FWD:|Fwd:)\s*', '', subject).strip()
   
   seven_digits = r'^(\d{7})$'
   m_pattern = r'^M(\d{7})$'
   sku_pattern = r'^SKU\s(\d{7})$'
   
   match = (re.match(seven_digits, cleaned) or 
           re.match(m_pattern, cleaned) or 
           re.match(sku_pattern, cleaned))
   
   if match:
       return match.group(1)
   return None

def extract_labels(body):
   """Extract category labels from email body"""
   for line in body.split('\n'):

       if (',' in line and 
           '@' not in line and 
           'Subject:' not in line and 
           'From:' not in line and 
           'To:' not in line):
          
           line_labels = [label.strip().lower() for label in line.split(',')]
           if len(line_labels) > 1:  
               return line_labels
   return []

def save_attachments(message, sku):
   """Save attachments using absolute paths"""
   if message.Attachments.Count == 0:
       return []
   
   sku_dir = os.path.join(SAVE_DIR, sku)
   if not os.path.exists(sku_dir):
       os.makedirs(sku_dir)
       print(f"Created directory: {sku_dir}")
   
   saved_files = []
   for attachment in message.Attachments:
       if attachment.FileName.lower().endswith(('.jpg', '.jpeg', '.png', '.gif')):
           file_path = os.path.join(sku_dir, attachment.FileName)
           try:
               if os.path.exists(file_path):
                   print(f"File already exists: {file_path}")
                   continue
                   
               print(f"Attempting to save to: {file_path}")
               attachment.SaveAsFile(file_path)
               saved_files.append(file_path)
               print(f"Successfully saved: {file_path}")
           except Exception as e:
               print(f"Error saving attachment {attachment.FileName}")
               print(f"Attempted path: {file_path}")
               print(f"Error details: {str(e)}")
   
   return saved_files


sku_data = {}

for msg in messages:
   message_count += 1
   try:
       if not msg.SenderName:  
           continue
           
       if ("brian walker" not in msg.SenderName.lower() or 
           not msg.Subject.lower().startswith(('fw:', 'fwd:', 'fw'))):
           continue
       
       if msg.Subject.lower().startswith('re:'):
           continue
           
       received_date = eastern.localize(msg.ReceivedTime.replace(tzinfo=None))
       if received_date < cutoff_date:
           continue
           
       subject = msg.Subject.strip()
       sku = extract_sku(subject)
       
       if sku:
          
           labels = extract_labels(msg.Body)
           
           
           sku_data[sku] = {
               'labels': labels,
               'received_date': received_date,
               'sender': msg.SenderName,
           }
           
           print(f"\nProcessing email with SKU: {sku}")
           print(f"Labels found: {', '.join(labels)}")
           print(f"Attachment count: {msg.Attachments.Count}")
           
           saved_files = save_attachments(msg, sku)
           
           print(f"\nSKU: {sku}")
           print(f"From: {msg.SenderName}")
           print(f"Subject: {subject}")
           if saved_files:
               print(f"Saved Images: {len(saved_files)}")
               for file in saved_files:
                   print(f"  - {file}")
               sku_data[sku]['image_files'] = saved_files
           else:
               print("No images saved")
               sku_data[sku]['image_files'] = []
           print("-" * 50)
           
   except Exception as e:
       print(f"Error processing message: {str(e)}")

print(f"\nTotal messages checked: {message_count}")
print("\nSKU Summary:")
for sku, data in sku_data.items():
   print(f"\nSKU: {sku}")
   print(f"Labels: {', '.join(data['labels'])}")
   print(f"Images: {len(data['image_files'])}")