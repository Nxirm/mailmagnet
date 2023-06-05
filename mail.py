import os
import re
import sys
import urllib.parse

import email
import argparse
import datetime
from termcolor import colored
from colorama import Fore,Back, Style
from colorama import init
init(autoreset =True)
from tabulate import tabulate

logo = """



███╗   ███╗ █████╗ ██╗██╗         ███╗   ███╗ █████╗  ██████╗ ███╗   ██╗███████╗████████╗
████╗ ████║██╔══██╗██║██║         ████╗ ████║██╔══██╗██╔════╝ ████╗  ██║██╔════╝╚══██╔══╝
██╔████╔██║███████║██║██║         ██╔████╔██║███████║██║  ███╗██╔██╗ ██║█████╗     ██║   
██║╚██╔╝██║██╔══██║██║██║         ██║╚██╔╝██║██╔══██║██║   ██║██║╚██╗██║██╔══╝     ██║   
██║ ╚═╝ ██║██║  ██║██║███████╗    ██║ ╚═╝ ██║██║  ██║╚██████╔╝██║ ╚████║███████╗   ██║   
╚═╝     ╚═╝╚═╝  ╚═╝╚═╝╚══════╝    ╚═╝     ╚═╝╚═╝  ╚═╝ ╚═════╝ ╚═╝  ╚═══╝╚══════╝   ╚═╝   
                                                                                         





MMWWMMWWWMMWWMM........................MMWWWWdKWMMWW
WWWWWWWWWWWWWWW kWWWWWWWWWWWWWWWWWWWWK WWOONXxOWWWWW
WWWWWWWWWWWWWWW OMM'lllll,0NlllllloMMX WWxcWWWWWWWWW
WWWWWWWWWWWWWWW OMW WMMMM.kNoooooooMMX WWWWOOOOWWWWW
WWWWWWWWWWWWWWW OMM loooo kNoooooooMMX WXkkWONWWWWWW
MMWWWMWWXONkkNW OMMNNNNNNNMMNNNNNNNMMX WMWWWMWWWMMWW
WWWMWWKkOWWNx:, OMKllllllllllllllll0MX ,:xNWWWMWWWMM
MMWWWMWWWM0   ' OM0llllllllllllllll0MX '   0MWWWMMWW
WWMMWWWMWWk.No:.;xXMMMMMMMMMMMMMMMMXx:.;lx.kWWMMWWMM
MMWWWNlkWMO.NOOOOd:;cxNMMMMMMMMNxc,;lxxxxx.kMWWWMMWW
WWMMWWONWWk.NOOOOOOOOo;;cl;;lc;,cdxxxxxxxx.kWWMMWWMM
MMWWMMWWWMO.NOOOOOOOOd;,;oxxo;,,lxxxxxxxxx.kMWWWMMWW
WWMMWWWMMWk.NOOOOOd;';oxxxxxxxxo;',lxxxxxx.kWWMMWWMM
MMWWMMWWWMO.NOOo;';dxxxxxxxxxxxxxxd;',cxxx.OMWWWMMWW
WWMMWWWMMWk.x;,:dxxxxxxxxxxxxxxxxxxxxd:,,c.kWWMMWWMM
MMWWMMWWWMK,,ccccccccccccccccccc,.ccccccc,,KMWWWMMWW
WWMMWWWMWWWMMWWMMWWWMWWWMMWWWMWWW,xWWMMWWWMWWWMMWWMM

"""
print(Fore.GREEN+(logo))
print("\n")
print(Fore.GREEN+"Command-line tool designed to extract and analyze emails headers as well as attachments.")
print(Fore.GREEN+"           By:Nupur Bhosle(Nxirm)")

def extract_attachments(email_file, output_dir):
   
    with open(email_file, 'rb') as f:
        msg = email.message_from_binary_file(f)

    attachments_details = []
    
    if msg.get_content_maintype() == 'multipart':
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None:
                continue

            
            if part.get_filename():
                # Extract the attachment filename
                filename = part.get_filename()
                filename = os.path.basename(filename)
                filename = filename.replace('/', '_')

                # Save the attachment 
                attachment_path = os.path.join(output_dir, filename)
                with open(attachment_path, 'wb') as attachment_file:
                    attachment_file.write(part.get_payload(decode=True))
                print(Fore.GREEN+ "Attachment Saved Successfully")
                print(f"Path of attached file : {attachment_path}")
                
                # attachment details
                file_size = os.path.getsize(attachment_path)
                file_type = part.get_content_type()
                file_modified = datetime.datetime.fromtimestamp(os.path.getmtime(attachment_path))

                attachments_details.append([filename, file_size, file_type, file_modified])

    else:
        print(Fore.RED+ "No attachments found in the email.")
    urls = re.findall(r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', str(msg))
    #extract urld
    if urls:
        print(Fore.GREEN+"URLs found in the email:")
        print(tabulate([[url] for url in urls], headers=["URL"], tablefmt="fancy_grid"))
    else:
        print("No URLs found in the email.")

    return attachments_details
#header analysis
def extract_header_details(msg):
    header_details = [
        ['Sender', msg['From']],
        ['Recipients', msg['To']],
        ['Subject', msg['Subject']],
        ['Date', msg['Date']],
        ['Reply-To', msg['Reply-To']],
        ['Content-Type', msg['Content-Type']],
        ['MIME Version', msg['MIME-Version']],
        ['X-Mailer', msg['X-Mailer']],
        ['DKIM Signature', msg['DKIM-Signature']],
        ['SPF', msg['SPF']],
        ['DMARC', msg['DMARC']],
    ]
    return header_details

if __name__ == '__main__':
    #cmd line instructions
    parser = argparse.ArgumentParser(description='Extract attachments from an email file.')
    parser.add_argument('email_file', help='Path to the email file')
    parser.add_argument('output_dir', help='Path to the output directory')
    args = parser.parse_args()

    with open(args.email_file, 'rb') as f:
        msg = email.message_from_binary_file(f)

    # header
    header_details = extract_header_details(msg)
    print(tabulate(header_details, headers=['Field', 'Value'], tablefmt='fancy_grid'))

    #attachment
    attachments_details = extract_attachments(args.email_file, args.output_dir)
    attachments_table = tabulate(attachments_details, headers=['Filename', 'Size', 'File Type', 'Last Modified'],
                                 tablefmt='fancy_grid')
    print(Fore.GREEN+"\nAttachment Details:\n")
    print(attachments_table)
    
  

